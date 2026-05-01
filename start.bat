@echo off
echo.
echo  ========================================
echo   AI Velocity Portal - AI Wins Dashboard
echo  ========================================
echo.

:: Install dependencies
echo [1/2] Installing backend dependencies...
pip install fastapi uvicorn pydantic python-dotenv --quiet 2>nul
echo       Done!
echo.

:: Start the server
echo [2/2] Starting FastAPI server...
echo.
echo  ┌──────────────────────────────────────────┐
echo  │  Dashboard:  http://localhost:3000        │
echo  │  API:        http://localhost:3000/api    │
echo  │  API Docs:   http://localhost:3000/docs   │
echo  │                                           │
echo  │  Data: SharePoint CSV (data/submissions)  │
echo  └──────────────────────────────────────────┘
echo.
echo  Press Ctrl+C to stop the server.
echo.

:: Open browser
start http://localhost:3000

:: Run the server
python -m uvicorn backend.main:app --host 0.0.0.0 --port 3000
