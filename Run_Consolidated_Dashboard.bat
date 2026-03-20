@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" consolidated_dashboard_app.py
if errorlevel 1 (
    echo.
    echo ERROR: Script failed.
    pause
)
