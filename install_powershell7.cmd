@echo off
echo ====================================================
echo    PowerShell 7 Installation and Configuration
echo ====================================================
echo.

:: Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.6 or higher from https://python.org
    pause
    exit /b 1
)

echo Python detected. Starting PowerShell 7 manager...
echo.

:: Run the PowerShell manager
python powershell_manager.py

echo.
echo ====================================================
echo Installation process completed.
echo Check the generated report for details.
echo ====================================================
pause 