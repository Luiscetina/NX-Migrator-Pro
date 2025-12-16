@echo off
title NX Migrator Pro - Setup & Run

:: Always run in the exact folder where the .bat file is located
cd /d "%~dp0"

:: Request Administrator privileges
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ================================================
    echo   REQUESTING ADMINISTRATOR PRIVILEGES...
    echo ================================================
    echo This tool needs admin rights for disk operations.
    echo Click "Yes" in the UAC prompt.
    echo.
    powershell "Start-Process '%~f0' -Verb RunAs -WorkingDirectory '%~dp0'"
    exit /b
)

cls
echo ================================================
echo     NX Migrator Pro - Automatic Setup & Run
echo ================================================
echo Working folder: %cd%
echo Running as Administrator.
echo.

echo [1/3] Checking/Creating virtual environment...
if not exist "venv" (
    echo Creating virtual environment...
    py -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create venv.
        pause
        exit /b 1
    )
    echo Virtual environment created.
) else (
    echo Virtual environment already exists.
)

echo.
echo [2/3] Activating virtual environment...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate venv.
    pause
    exit /b 1
)

echo Activated venv Python: 
where python
python --version

echo.
echo [3/3] Installing/upgrading packages...
pip install --upgrade pip
echo Installing core requirements...
pip install --upgrade ttkbootstrap pywin32 WMI psutil
echo Installing all dependencies from requirements.txt...
pip install -r requirements.txt --upgrade
echo.
echo Post-install: Configuring pywin32...
python venv\Scripts\pywin32_postinstall.py -install

echo.
echo ================================================
echo   SETUP COMPLETE! Launching main.py...
echo ================================================
echo.

venv\Scripts\python.exe main.py

echo.
echo ================================================
echo   main.py has finished.
echo ================================================
echo.
echo Virtual environment is still active.
echo Run again with: venv\Scripts\python.exe main.py
echo.
pause