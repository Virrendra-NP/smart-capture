@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" weekly_dashboard_app.py
if errorlevel 1 (
    echo.
    echo ERROR: Script failed.
    pause
)
