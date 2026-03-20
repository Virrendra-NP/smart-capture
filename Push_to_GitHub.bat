@echo off
setlocal
cd /d "%~dp0"

echo -------------------------------------------------------------------
echo 🚀 JKD SMART CAPTURE - GLOBAL DEPLOYMENT (By Virrendra)
echo -------------------------------------------------------------------
echo.

:: 1. Initialize Git if not already done
if not exist ".git" (
    echo Initializing Local Project...
    git init
)

:: 2. Set Branch to main (Modern Standard)
git branch -M main

:: 3. Stage and Commit
echo Adding Local Files...
git add .
git commit -m "Official Deployment by Virrendra"

echo.
echo -------------------------------------------------------------------
echo STEP: CONNECT TO YOUR GITHUB ACCOUNT
echo.
echo 1. Go to https://github.com/new
echo 2. Name it "smart-capture" and Click "Create Repository"
echo 3. Copy the URL (e.g., https://github.com/USERNAME/smart-capture.git)
echo -------------------------------------------------------------------
echo.

set /p github_url="PASTE YOUR GITHUB REPO LINK HERE: "

:: 4. Link Remote
git remote add origin %github_url% 2>nul
git remote set-url origin %github_url%

:: 5. Push!
echo.
echo -------------------------------------------------------------------
echo SENDING CODE TO THE CLOUD...
echo -------------------------------------------------------------------
git push -u origin main

if errorlevel 1 (
    echo.
    echo ERROR: Could not push to GitHub. 
    echo Check your internet connection and make sure your repo is empty on GitHub.
    pause
) else (
    echo.
    echo SUCCESS! Your code is now in the Cloud.
    echo Next: Go to share.streamlit.io and login with GitHub to get your link!
    pause
)
