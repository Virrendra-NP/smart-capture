@echo off
cd /d "%~dp0"
echo Launching JKD Smart Site Portal (PUBLIC)...
echo -------------------------------------------------------------------
echo STEP 1: Starting the AI Portal Engine...
start /b ".venv\Scripts\python.exe" -m streamlit run site_photo_portal.py --server.port 8501 --server.address 0.0.0.0

echo STEP 2: Creating a SECURE PUBLIC LINK for your iPhone/iPad...
echo -------------------------------------------------------------------
echo.
echo PLEASE WAIT 10-20 SECONDS...
echo.
echo LOOK FOR A LINK THAT ENDS WITH: .trycloudflare.com
echo COPY THAT LINK TO YOUR IPHONE AND IPAD!
echo.
cloudflared.exe tunnel --url http://localhost:8501
if errorlevel 1 (
    echo.
    echo ERROR: Could not create public tunnel. 
    pause
)
