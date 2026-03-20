@echo off
cd /d "%~dp0"
echo 🚀 JKD SMART CAPTURE - QUICK SYNC
echo -------------------------------------------------------------------
echo Sending updates to your iPad app...
echo.

:: 1. Add and Commit
git add .
git commit -m "Live Update via Sync_Now.bat"

:: 2. Silent Push
git push origin main

if errorlevel 1 (
    echo.
    echo ERROR: Sync failed. Make sure you are connected to the internet.
    pause
) else (
    echo.
    echo ✅ SUCCESS! Your iPad app will update in 60 seconds.
    timeout /t 5
)
