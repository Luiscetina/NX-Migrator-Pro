using Microsoft.Win32;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Launcher
{
    // P/Invoke for reliable key state detection
    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        internal static extern short GetAsyncKeyState(int vKey);
    }

    public partial class MainForm : Form
    {
        private const string PythonVersion = "3.13.7";
        private const string MainScript = "main.py";
        private readonly string[] Dependencies = { "psutil", "pywin32", "ttkbootstrap", "WMI" };
        private const bool HideConsole = true;
        private string ConfigPath => Path.Combine(Application.StartupPath, Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".config");

        private TextBox logTextBox;
        private ProgressBar progressBar;
        private bool shouldShowWindow = true;

        public MainForm()
        {
            // Check if config exists and HideConsole is true - if so, start completely hidden
            if (HideConsole && File.Exists(ConfigPath))
            {
                // Config exists, we can run silently without showing window
                shouldShowWindow = false;
                this.FormBorderStyle = FormBorderStyle.None;
                this.WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
                this.Opacity = 0;
                this.Size = new System.Drawing.Size(0, 0);
                this.Location = new System.Drawing.Point(-10000, -10000);
            }
            else
            {
                // First run or debug mode - show window for setup/debugging
                shouldShowWindow = true;
                this.Text = "main - Initializing...";
                this.Size = new System.Drawing.Size(550, 220);
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;

                logTextBox = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Dock = DockStyle.Fill,
                    Font = new System.Drawing.Font("Consolas", 9F),
                    BackColor = System.Drawing.Color.FromArgb(30, 30, 30),
                    ForeColor = System.Drawing.Color.FromArgb(220, 220, 220),
                    BorderStyle = BorderStyle.None
                };
                this.Controls.Add(logTextBox);

                progressBar = new ProgressBar { Dock = DockStyle.Bottom, Height = 10, Visible = false };
                this.Controls.Add(progressBar);
            }
        }

        protected override async void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            if (!IsAdministrator())
            {
                RelaunchAsAdmin();
                return;
            }
            await Task.Delay(100);
            await StartProcess();
        }

        private class LauncherConfig
        {
            public string PythonPath { get; set; }
            public string PythonVersion { get; set; }
            public string[] InstalledDependencies { get; set; }
            public DateTime LastVerified { get; set; }
        }

        private LauncherConfig LoadConfig()
        {
            if (!File.Exists(ConfigPath)) return null;
            try
            {
                string json = File.ReadAllText(ConfigPath);
                return System.Text.Json.JsonSerializer.Deserialize<LauncherConfig>(json);
            }
            catch { return null; }
        }

        private void SaveConfig(string pythonPath, string[] deps)
        {
            var config = new LauncherConfig
            {
                PythonPath = pythonPath,
                PythonVersion = PythonVersion,
                InstalledDependencies = deps,
                LastVerified = DateTime.Now
            };
            try
            {
                string json = System.Text.Json.JsonSerializer.Serialize(config, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(ConfigPath, json);
            }
            catch { }
        }

        private async Task StartProcess()
        {
            Log("Launcher initialized");
            Log("Administrator privileges: OK");

            // Find the script file (case-insensitive search)
            string scriptPath = FindScriptFile();
            if (string.IsNullOrEmpty(scriptPath))
            {
                string errorMsg = $"Script file not found!\\n\\nExpected: {MainScript}\\nLocation: {Application.StartupPath}\\n\\nPlease ensure the Python script is in the same folder as this launcher.";
                Log($"ERROR: Script file not found!");
                Log($"Expected: '{MainScript}'");
                Log($"Location: '{Application.StartupPath}'");
                ShowErrorMessage(errorMsg);
                Application.Exit();
                return;
            }

            Log($"Found script: {Path.GetFileName(scriptPath)}");

            // Check if Shift key is held for force reinstall (check native key state for reliability)
            bool forceRecheck = false;
            try
            {
                // Check using GetAsyncKeyState for more reliable detection
                const int VK_SHIFT = 0x10;
                forceRecheck = (NativeMethods.GetAsyncKeyState(VK_SHIFT) & 0x8000) != 0;
            }
            catch
            {
                // Fallback to Control.ModifierKeys
                forceRecheck = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            }
            
            var config = LoadConfig();
            string pythonPath = null;

            // Use cached config if valid and not forcing recheck
            if (!forceRecheck && config != null && File.Exists(config.PythonPath) && config.PythonVersion == PythonVersion)
            {
                // Check if dependencies match
                bool depsMatch = config.InstalledDependencies != null &&
                                 Dependencies.Length == config.InstalledDependencies.Length &&
                                 Dependencies.OrderBy(x => x).SequenceEqual(config.InstalledDependencies.OrderBy(x => x));

                if (depsMatch)
                {
                    pythonPath = config.PythonPath;
                    Log($"Using cached Python: {pythonPath}");
                    Log("Dependencies already installed (cached)");
                    Log("TIP: Hold SHIFT during launch to force reinstall");

                    // Quick launch - window is already hidden if config existed
                    await LaunchScript(pythonPath, scriptPath);
                    Application.Exit();
                    return;
                }
                else
                {
                    Log("Dependencies changed - need to reinstall...");
                    pythonPath = config.PythonPath;
                    // Continue to full installation below
                }
            }

            if (forceRecheck) Log("Force recheck enabled (Shift held)");

            // Full installation process
            // Window is already visible to show installation progress
            if (config == null)
            {
                Log("First launch detected - running setup...");
            }

            Log($"Checking for Python {PythonVersion}...");
            pythonPath = FindPythonPath();

            if (string.IsNullOrEmpty(pythonPath))
            {
                Log("Python not found. Installing...");
                UpdateTitle("Installing Python...");
                bool installed = await InstallPython();
                if (installed)
                {
                    Log("Installation complete. Waiting for registration...");

                    // Try multiple times with increasing delays to find Python after installation
                    int maxRetries = 10;
                    for (int i = 0; i < maxRetries; i++)
                    {
                        await Task.Delay(2000); // Wait 2 seconds between attempts
                        pythonPath = FindPythonPath();

                        if (!string.IsNullOrEmpty(pythonPath))
                        {
                            Log($"Python detected after {(i + 1) * 2} seconds");
                            break;
                        }

                        if (i < maxRetries - 1)
                            Log($"Python not yet registered, waiting... (attempt {i + 2}/{maxRetries})");
                    }

                    if (string.IsNullOrEmpty(pythonPath))
                    {
                        Log("ERROR: Python installed but not detected in registry or PATH");
                        Log("This may be a timing issue. Try running the launcher again.");
                    }
                    else
                    {
                        // Verify installed version matches requested version
                        if (!await VerifyInstalledPythonVersion(pythonPath))
                        {
                            Log($"WARNING: Installed Python version may not match {{PythonVersion}}");
                        }
                    }
                }
                else
                {
                    Log("ERROR: Python installation failed");
                    MessageBox.Show("Failed to install Python.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                    return;
                }
            }
            else
            {
                Log($"Python found: {pythonPath}");
                // Verify version for existing installation
                await VerifyInstalledPythonVersion(pythonPath);
            }

            if (!string.IsNullOrEmpty(pythonPath))
            {
                

                // Verify tkinter/IDLE installation
                Log("Verifying tkinter/IDLE module...");
                if (!await IsTkinterAvailable(pythonPath))
                {{
                    Log("ERROR: tkinter module not found");
                    ShowErrorMessage(
                        "The tkinter module is missing from your Python installation.\\n\\n" +
                        "Tkinter is required for this application's graphical interface.\\n\\n" +
                        "To fix this:\\n" +
                        "1. Open 'Add or Remove Programs' in Windows Settings\\n" +
                        "2. Find your Python installation\\n" +
                        "3. Click 'Modify'\\n" +
                        "4. Ensure 'tcl/tk and IDLE' is checked\\n" +
                        "5. Complete the modification\\n\\n" +
                        "Then restart this launcher.");
                    Application.Exit();
                    return;
                }}
                Log("✓ Tkinter module is available");

                

                // Update pip
                Log("Updating pip...");
                await UpdatePip(pythonPath);

                if (Dependencies.Length > 0)
                {
                    Log("Installing dependencies...");
                    foreach (var dep in Dependencies)
                        await InstallPackage(pythonPath, dep);
                }

                // Save config after successful setup
                SaveConfig(pythonPath, Dependencies);
                Log("Configuration saved for future launches");

                await LaunchScript(pythonPath, scriptPath);
                Application.Exit();
            }
            else
            {
                Log("ERROR: Could not locate Python");
                ShowErrorMessage("Python not found after installation attempt.\\n\\nPlease install Python manually from python.org");
                Application.Exit();
            }
        }

        private string FindScriptFile()
        {
            // First try exact match with full path
            string exactPath = Path.Combine(Application.StartupPath, MainScript);
            if (File.Exists(exactPath))
                return exactPath;

            // Try case-insensitive search in startup directory
            string directory = Application.StartupPath;
            try
            {
                string[] files = Directory.GetFiles(directory, "*.py");
                foreach (string file in files)
                {
                    if (Path.GetFileName(file).Equals(MainScript, StringComparison.OrdinalIgnoreCase))
                    {
                        Log($"Found script with different casing: {Path.GetFileName(file)}");
                        return file;
                    }
                }
            }
            catch { }

            return null;
        }

        private void ShowErrorMessage(string message)
        {
            // Always show errors - the main form is always available now
            MessageBox.Show(this, message, "Launcher Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private string FindPythonPath()
        {
            try
            {
                foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
                {
                    using (var baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64))
                    using (var key = baseKey.OpenSubKey(@"Software\Python\PythonCore"))
                    {
                        if (key != null)
                        {
                            var versions = key.GetSubKeyNames().OrderByDescending(v => v);
                            foreach (var ver in versions)
                            {
                                using (var pathKey = key.OpenSubKey($@"{ver}\InstallPath"))
                                {
                                    if (pathKey?.GetValue("") is string path)
                                    {
                                        string pythonExe = Path.Combine(path, "pythonw.exe");
                                        if (File.Exists(pythonExe))
                                            return pythonExe;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Registry search error: {ex.Message}");
            }

            var pathVar = Environment.GetEnvironmentVariable("PATH");
            if (pathVar != null)
            {
                var pathsToCheck = pathVar.Split(';')
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Select(x => x.Trim())
                    .OrderBy(p => p.Contains("WindowsApps") ? 1 : 0); // Prefer non-WindowsApps paths

                foreach (var p in pathsToCheck)
                {
                    try
                    {
                        var pythonExe = Path.Combine(p, "pythonw.exe");
                        if (File.Exists(pythonExe))
                        {
                            // Verify it's not a stub/alias by checking file size
                            var fileInfo = new FileInfo(pythonExe);
                            if (fileInfo.Length > 50000) // Real Python exe is much larger than stubs
                                return pythonExe;
                        }
                    }
                    catch { } // Skip invalid paths
                }
            }

            return null;
        }

        private async Task InstallPackage(string pythonPath, string package)
        {
            string pipPath = Path.Combine(Path.GetDirectoryName(pythonPath), "Scripts", "pip.exe");
            if (!File.Exists(pipPath))
            {
                Log($"ERROR: pip not found at {pipPath}");
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = pipPath,
                Arguments = $"install {package} --no-warn-script-location --timeout 300",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try
            {
                using (var process = Process.Start(startInfo))
                {
                    Log($"Installing {package}...");

                    // Read output asynchronously to show progress
                    var outputTask = Task.Run(async () =>
                    {
                        while (!process.StandardOutput.EndOfStream)
                        {
                            string line = await process.StandardOutput.ReadLineAsync();
                            if (!string.IsNullOrWhiteSpace(line))
                            {
                                // Filter out excessive pip output, show important lines
                                if (line.Contains("Collecting") || line.Contains("Downloading") ||
                                    line.Contains("Installing") || line.Contains("Successfully") ||
                                    line.Contains("Requirement already satisfied"))
                                {
                                    Log($"  {line.Trim()}");
                                }
                            }
                        }
                    });

                    var errorTask = Task.Run(async () =>
                    {
                        string stderr = await process.StandardError.ReadToEndAsync();
                        return stderr;
                    });

                    // Add timeout to prevent hanging (180 seconds per package for large packages)
                    var completedTask = await Task.WhenAny(
                        Task.Run(() => process.WaitForExit()),
                        Task.Delay(180000)
                    );

                    if (!process.HasExited)
                    {
                        Log($"WARNING: Package '{package}' installation timed out, killing process");
                        process.Kill();
                        return;
                    }

                    // Wait for output tasks to complete
                    await Task.WhenAll(outputTask, errorTask);
                    string errors = await errorTask;

                    if (process.ExitCode == 0)
                    {
                        Log($"✓ Package '{package}' ready");
                    }
                    else
                    {
                        Log($"WARNING: Package '{package}' install exit code {process.ExitCode}");
                        if (!string.IsNullOrWhiteSpace(errors))
                            Log($"pip error: {errors.Trim()}");
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"ERROR installing '{package}': {ex.Message}");
            }
        }

        private async Task LaunchScript(string pythonPath, string scriptPath)
        {
            Log($"Launching script...");
            UpdateTitle("Running...");

            var startInfo = new ProcessStartInfo
            {
                FileName = pythonPath,
                Arguments = $"\"{scriptPath.Replace("\"", "\\\\\"")}\" ",
                WorkingDirectory = Application.StartupPath,
                UseShellExecute = false,
                CreateNoWindow = HideConsole,
                WindowStyle = HideConsole ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal,
                RedirectStandardOutput = false,
                RedirectStandardError = false
            };

            try
            {
                using (var process = Process.Start(startInfo))
                {
                    if (process == null)
                    {
                        Log("Failed to start script");
                        return;
                    }

                    if (HideConsole)
                    {
                        Log("Script started successfully. Launcher will now hide.");
                        // Hide the launcher window now that the script is running
                        this.Invoke((Action)delegate { this.Hide(); });
                    }
                    else
                    {
                        Log("Script started. Launcher window visible for debugging.");
                    }

                    await Task.Run(() => process.WaitForExit());
                    Log("Script exited");
                }
            }
            catch (Exception ex)
            {
                Log($"Launch error: {ex.Message}");
                MessageBox.Show($"Error running script:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task<bool> InstallPython()
        {
            if (shouldShowWindow && progressBar != null)
                this.Invoke((Action)delegate { progressBar.Visible = true; });
            string url = $"https://www.python.org/ftp/python/{PythonVersion}/python-{PythonVersion}-amd64.exe";
            string installerPath = Path.Combine(Path.GetTempPath(), $"python-{PythonVersion}.exe");

            try
            {
                Log("Downloading Python installer...");
                using (var client = new WebClient())
                {
                    // Set timeout (default is 100 seconds, increase for large downloads)
                    client.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)");

                    client.DownloadProgressChanged += (s, e) =>
                    {
                        if (shouldShowWindow && progressBar != null)
                            this.Invoke((Action)delegate { progressBar.Value = e.ProgressPercentage; });
                    };

                    // Add timeout wrapper (10 minutes for large Python installer)
                    var downloadTask = client.DownloadFileTaskAsync(new Uri(url), installerPath);
                    var timeoutTask = Task.Delay(600000); // 10 minutes

                    if (await Task.WhenAny(downloadTask, timeoutTask) == timeoutTask)
                    {
                        client.CancelAsync();
                        throw new Exception("Download timed out after 10 minutes");
                    }

                    await downloadTask; // Ensure any exceptions are thrown
                }

                // Verify file was downloaded
                if (!File.Exists(installerPath))
                {
                    Log("ERROR: Installer download failed");
                    return false;
                }

                Log($"Downloaded installer ({new FileInfo(installerPath).Length / 1024 / 1024} MB)");

                // Determine install scope based on admin rights
                bool isAdmin = IsAdministrator();
                string installScope = isAdmin ? "InstallAllUsers=1" : "InstallAllUsers=0";

                Log(isAdmin ? "Installing Python for all users..." : "Installing Python for current user...");

                var startInfo = new ProcessStartInfo
                {
                    FileName = installerPath,
                    Arguments = $"/passive {installScope} PrependPath=1 Include_pip=1 Include_test=0 Include_tcltk=1",
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                };

                using (var process = Process.Start(startInfo))
                {
                    await Task.Run(() => process.WaitForExit());

                    try { File.Delete(installerPath); }
                    catch { }

                    return process.ExitCode == 0;
                }
            }
            catch (Exception ex)
            {
                Log($"Installation error: {ex.Message}");
                try { if (File.Exists(installerPath)) File.Delete(installerPath); }
                catch { }
                return false;
            }
            finally
            {
                if (shouldShowWindow && progressBar != null)
                    this.Invoke((Action)delegate { progressBar.Visible = false; });
            }
        }

        private void Log(string msg)
        {
            string entry = $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}";

            // Write to UI if window is shown
            if (shouldShowWindow && logTextBox != null)
            {
                if (this.InvokeRequired)
                    this.Invoke(new Action(() => logTextBox.AppendText(entry)));
                else
                    logTextBox.AppendText(entry);
            }

            // Always write to log file for debugging
            try
            {
                string logPath = Path.Combine(Application.StartupPath, Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".log");
                File.AppendAllText(logPath, entry);
            }
            catch (Exception logEx)
            {
                // If we can't write to app directory, try temp directory
                try
                {
                    string tempLogPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(Application.ExecutablePath) + ".log");
                    File.AppendAllText(tempLogPath, entry);
                    File.AppendAllText(tempLogPath, $"[WARNING] Could not write to main log: {logEx.Message}{Environment.NewLine}");
                }
                catch { } // Give up silently if both fail
            }
        }

        private void UpdateTitle(string text)
        {
            if (!shouldShowWindow) return;

            if (this.InvokeRequired)
                this.Invoke(new Action(() => this.Text = text));
            else
                this.Text = text;
        }

        private async Task<bool> VerifyInstalledPythonVersion(string pythonPath)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = pythonPath,
                    Arguments = "-c \"import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}')\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(startInfo))
                {
                    string output = (await process.StandardOutput.ReadToEndAsync()).Trim();
                    await Task.Run(() => process.WaitForExit());

                    if (process.ExitCode == 0 && !string.IsNullOrEmpty(output))
                    {
                        Log($"Python version detected: {output}");
                        // Check if major.minor matches (micro version can differ)
                        if (output.StartsWith(PythonVersion.Substring(0, PythonVersion.LastIndexOf('.'))))
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log($"Version verification failed: {ex.Message}");
            }
            return false;
        }

        private async Task<bool> IsTkinterAvailable(string pythonPath)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = pythonPath,
                    Arguments = "-c \"import tkinter\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using (var process = Process.Start(startInfo))
                {
                    await Task.Run(() => process.WaitForExit());
                    return process.ExitCode == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        

        private async Task UpdatePip(string pythonPath)
        {
            string pipPath = Path.Combine(Path.GetDirectoryName(pythonPath), "Scripts", "pip.exe");
            if (!File.Exists(pipPath))
            {
                Log("pip.exe not found, skipping update");
                return;
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = pipPath,
                Arguments = "install --upgrade pip",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try
            {
                using (var process = Process.Start(startInfo))
                {
                    await Task.Run(() => process.WaitForExit());
                    if (process.ExitCode == 0)
                        Log("✓ pip updated successfully");
                    else
                        Log("pip update completed with warnings");
                }
            }
            catch (Exception ex)
            {
                Log($"pip update failed: {{ex.Message}}");
            }
        }

        

        private bool IsAdministrator()
        {
            try
            {
                using (var identity = WindowsIdentity.GetCurrent())
                {
                    var principal = new WindowsPrincipal(identity);
                    return principal.IsInRole(WindowsBuiltInRole.Administrator);
                }
            }
            catch
            {
                return false;
            }
        }

        private void RelaunchAsAdmin()
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = Application.ExecutablePath,
                Verb = "runas",
                UseShellExecute = true
            };

            try
            {
                Process.Start(startInfo);
            }
            catch (Win32Exception)
            {
                MessageBox.Show("Administrator privileges required.", "Admin Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            Application.Exit();
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}