@echo off
REM ═══════════════════════════════════════════════════════════════
REM  AI Velocity - Setup Automatic SharePoint Sync (Every Hour)
REM  This creates a Windows Task Scheduler task that:
REM    1. Runs Selenium to download CSV from SharePoint
REM    2. Pushes the CSV to GitHub
REM    3. Render auto-deploys with fresh data
REM ═══════════════════════════════════════════════════════════════

echo.
echo  ======================================
echo   AI Velocity Auto-Sync Scheduler Setup
echo  ======================================
echo.

REM Get the Python path
for /f "delims=" %%i in ('where python 2^>nul') do set PYTHON_PATH=%%i
if not defined PYTHON_PATH (
    echo ERROR: Python not found in PATH!
    pause
    exit /b 1
)
echo Found Python: %PYTHON_PATH%

REM Set script path
set SCRIPT_PATH=%~dp0auto_sync_and_push.py
set PROJECT_DIR=%~dp0..
echo Script: %SCRIPT_PATH%
echo Project: %PROJECT_DIR%

echo.
echo Creating scheduled task: AI_Velocity_Sync (runs every 1 hour)...
echo.

REM Delete existing task if any
schtasks /delete /tn "AI_Velocity_Sync" /f >nul 2>&1

REM Create new scheduled task — runs every 1 hour, starting now
schtasks /create /tn "AI_Velocity_Sync" /tr "\"%PYTHON_PATH%\" \"%SCRIPT_PATH%\"" /sc HOURLY /mo 1 /st %time:~0,5% /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo  ========================================
    echo   SUCCESS! Task scheduled.
    echo  ========================================
    echo.
    echo   Task Name:  AI_Velocity_Sync
    echo   Frequency:  Every 1 hour
    echo   Script:     %SCRIPT_PATH%
    echo.
    echo   The sync will:
    echo     1. Open Edge browser (using your SSO session)
    echo     2. Download CSV from SharePoint
    echo     3. Git push to GitHub
    echo     4. Render auto-deploys in ~2 min
    echo.
    echo   To check status:  schtasks /query /tn "AI_Velocity_Sync"
    echo   To run now:        schtasks /run /tn "AI_Velocity_Sync"
    echo   To stop/delete:    schtasks /delete /tn "AI_Velocity_Sync" /f
    echo.
    echo   Logs saved to: %PROJECT_DIR%\data\sync_log.txt
    echo.
) else (
    echo.
    echo  ERROR: Failed to create scheduled task.
    echo  Try running this script as Administrator.
    echo.
)

pause
