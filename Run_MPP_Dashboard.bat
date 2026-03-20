@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" mpp_to_excel_app.py
if errorlevel 1 (
    echo.
    echo ERROR: Script failed. Check mpp_app_debug.txt for details.
    pause
)
