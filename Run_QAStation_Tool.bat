@echo off
chcp 65001 >nul
cd /d "%~dp0"
title HFPA and QAStation Proactive Multi-Agent System
color 0B

echo ===============================================================================
echo                HFPA and QAStation Proactive Multi-Agent System
echo ===============================================================================
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

echo [OK] Python environment detected.
echo.
echo ===============================================================================
echo  CHOOSE EXECUTION MODE:
echo ===============================================================================
echo  [1] Run Proactive Multi-Agent Pipeline (Fast In-Memory + Executive Brief) [Default]
echo  [2] Check Ingestion Health and 4-Factory Completeness (Non-blocking check)
echo  [3] Start Continuous Directory Watcher Agent (Daemon mode)
echo  [4] Run Legacy Pipeline (process_qastation.py)
echo ===============================================================================
echo.
set /p MODE="Enter choice (1, 2, 3, 4) [Press ENTER for 1]: "
if defined MODE set "MODE=%MODE: =%"

if "%MODE%"=="" set MODE=1
if "%MODE%"=="1" goto run_agent
if "%MODE%"=="2" goto run_check
if "%MODE%"=="3" goto run_watch
if "%MODE%"=="4" goto run_legacy

echo [!] Lua chon khong hop le, tu dong chay che do [1]...
goto run_agent

:run_agent
echo.
echo [*] Executing Proactive Multi-Agent Pipeline...
python -m agents.hfpa_agent_orchestrator --run
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [ERROR] Proactive Pipeline encountered an issue!
    echo.
    pause
    exit /b 1
)
goto finish

:run_check
echo.
echo [*] Running Ingestion Readiness and Health Scan...
python -m agents.hfpa_agent_orchestrator --check
echo.
pause
exit /b 0

:run_watch
echo.
echo [*] Starting Ingestion Watcher Daemon (Polls every 10s)...
echo [*] Nhan Ctrl+C de dung qua trinh giam sat.
python -m agents.hfpa_agent_orchestrator --watch --interval 10
pause
exit /b 0

:run_legacy
echo.
echo [*] Running Legacy Pipeline (process_qastation.py)...
python "%~dp0process_qastation.py"
if %errorlevel% neq 0 (
    color 0C
    echo.
    echo [ERROR] Legacy Pipeline encountered an issue!
    echo.
    pause
    exit /b 1
)
goto finish

:finish
color 0A
echo.
echo ===============================================================================
echo [3/3] Execution completed successfully!
echo Results are saved in: "%~dp0Output"
echo ===============================================================================
echo.

:: Open Output folder in File Explorer
explorer "%~dp0Output"

pause
