@echo off
cd /d "%~dp0"
echo Launching JKD Smart Site Portal...
echo -------------------------------------------------------------------
echo NOTE: Make sure your iPhone/iPad is on the SAME WI-FI as this PC.
echo Look for the "Network URL" (e.g., http://192.168.1.5:8501)
echo Type that address into Safari on your iPhone/iPad!
echo -------------------------------------------------------------------
echo.
".venv\Scripts\python.exe" -m streamlit run site_photo_portal.py
if errorlevel 1 (
    echo.
    echo ERROR: Failed to launch portal. Check if Streamlit is installed.
    pause
)
