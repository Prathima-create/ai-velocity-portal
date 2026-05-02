@echo off
REM ═══════════════════════════════════════════════════════════════
REM  AI Velocity — Setup Automatic Hourly Sync
REM  Creates a Windows Task Scheduler task that:
REM    1. Opens Edge (headless) with your SSO session
REM    2. Fetches latest data from SharePoint
REM    3. Pushes updated CSV to GitHub
REM    4. Render auto-deploys with fresh data
REM ═══════════════════════════════════════════════════════════════

echo.
echo  ======================================
echo   AI Velocity Auto-Sync Setup
echo  ======================================
echo.

REM Find Python
for /f "delims=" %%i in ('where python 2^>nul') do set PYTHON_PATH=%%i
if not defined PYTHON_PATH (
    echo ERROR: Python not found in PATH
    pause
    exit /b 1
)
echo Python: %PYTHON_PATH%

set SCRIPT_PATH=%~dp0auto_sync_and_push.py
echo Script: %SCRIPT_PATH%

echo.
echo Creating scheduled task: AI_Velocity_Sync (every 1 hour)...

REM Remove old task if exists
schtasks /delete /tn "AI_Velocity_Sync" /f >nul 2>&1

REM Create task — runs every hour
schtasks /create /tn "AI_Velocity_Sync" /tr "\"%PYTHON_PATH%\" \"%SCRIPT_PATH%\"" /sc HOURLY /mo 1 /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo  ========================================
    echo   DONE! Sync scheduled every hour.
    echo  ========================================
    echo.
    echo   To run now:     schtasks /run /tn "AI_Velocity_Sync"
    echo   To check:       schtasks /query /tn "AI_Velocity_Sync"
    echo   To remove:      schtasks /delete /tn "AI_Velocity_Sync" /f
    echo   Logs:           data\sync_log.txt
    echo.
) else (
    echo.
    echo  ERROR: Could not create task. Try running as Administrator.
    echo.
)

pause
