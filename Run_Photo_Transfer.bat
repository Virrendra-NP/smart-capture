@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" photo_to_excel_app.py
if errorlevel 1 (
    echo.
    echo ERROR: Script failed.
    pause
)
