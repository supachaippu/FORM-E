@echo off
title FORM E - Document Comparison Tool
cd /d "%~dp0"

echo ===================================================
echo   FORM E - Document Comparison Tool (Windows Launcher)
echo ===================================================
echo.
echo Checking Python installation...
echo.

:: Verify if 'python' command is a real functioning Python interpreter
set "PY_OUTPUT="
for /f "tokens=*" %%i in ('python -c "print('REAL')" 2^>nul') do set "PY_OUTPUT=%%i"
if "%PY_OUTPUT%"=="REAL" (
    set "PYTHON_CMD=python"
    goto :run
)

:: Verify if 'py' command is a real functioning Python interpreter
set "PY_OUTPUT="
for /f "tokens=*" %%i in ('py -c "print('REAL')" 2^>nul') do set "PY_OUTPUT=%%i"
if "%PY_OUTPUT%"=="REAL" (
    set "PYTHON_CMD=py"
    goto :run
)

:: If we reach here, Python is either not installed or is just the Windows Store dummy stub
echo [ERROR] Python is not installed or not configured in your PATH!
echo.
echo TO FIX THIS:
echo 1. Download Python for Windows: https://www.python.org/downloads/
echo 2. Run the installer and check the box at the bottom:
echo    "[X] Add Python.exe to PATH"
echo 3. Click "Install Now"
echo 4. Close this terminal and double-click run_compare.bat again.
echo.
pause
exit /b

:run
echo Starting the program, please wait...
%PYTHON_CMD% compare_gui.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] An error occurred while running the program.
    pause
)
