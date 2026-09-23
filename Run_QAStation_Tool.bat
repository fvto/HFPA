@echo off
chcp 65001 >nul
title QAStation ^& HFPA Data Processing Tool
color 0A

echo =====================================================================
echo                QAStation ^& HFPA Data Processing Tool
echo =====================================================================
echo.
echo [1/3] Checking Python environment...

python --version >nul 2>&1
if %errorlevel% neq 0 (
    color 0C
    echo [ERROR] Python is not found in PATH!
    echo Please install Python 3.8+ or add it to system PATH.
    echo.
    pause
    exit /b 1
)

echo [2/3] Running QAStation ^& HFPA pipeline...
echo [CHU Y] Vui long DONG cac file Excel va PowerPoint trong thu muc Output va Database truoc khi chay!
echo.
python "%~dp0process_qastation.py"

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [ERROR] The data processing pipeline encountered an error!
    echo Please check the error message above.
    echo.
    pause
    exit /b 1
)

echo.
echo =====================================================================
echo [3/3] Execution completed successfully!
echo Results are saved in: "%~dp0Output"
echo =====================================================================
echo.

:: Open Output folder in File Explorer
explorer "%~dp0Output"

pause
