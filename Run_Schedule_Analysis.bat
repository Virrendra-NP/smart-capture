@echo off
cd /d "%~dp0"
echo Launching JKD Schedule Analytics Engine...
".venv\Scripts\python.exe" create_schedule_dashboard.py
if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch dashboard tool.
    pause
)
