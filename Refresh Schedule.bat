@echo off
title CMF Schedule Refresh
echo.
echo  ============================================
echo   CMF Production Schedule - Refreshing...
echo  ============================================
echo.

:: Move to the folder where this .bat file lives
cd /d "%~dp0"

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python is not installed.
    echo  Download it from https://www.python.org and check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

:: Install required packages silently if missing
echo  Checking dependencies...
pip install openpyxl Pillow --quiet --disable-pip-version-check

echo.
echo  Running builder...
echo.
python cmf_schedule_builder.py

if errorlevel 1 (
    echo.
    echo  !! Something went wrong. See error above.
) else (
    echo.
    echo  ============================================
    echo   Done! Open "CMF WIP - Schedule.xlsx"
    echo  ============================================
)

echo.
pause
