@echo off
echo ============================================================
echo  FinCom QC Dashboard - Auto Deploy Watcher
echo ============================================================
echo.
echo  This will watch your SharePoint CSV files for changes
echo  and auto-deploy the dashboard every 2 minutes.
echo.
echo  Press Ctrl+C to stop.
echo.
echo ============================================================
echo.

cd /d "%~dp0\.."
python scripts/auto_deploy_dashboard.py %*

pause
