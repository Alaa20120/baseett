@echo off
set GIT_PATH="C:\Program Files\Git\cmd\git.exe"
echo [1/4] Initializing Git and checking remote...
%GIT_PATH% init
%GIT_PATH% remote remove origin >nul 2>&1
%GIT_PATH% remote add origin https://github.com/Alaa20120/baseett.git

echo [2/4] Adding files (excluding node_modules via .gitignore)...
%GIT_PATH% add .

echo [3/4] Committing changes...
%GIT_PATH% commit -m "Fix: Resolve 404 errors, fix config paths, and update Service Worker"

echo [4/4] Pushing to GitHub...
%GIT_PATH% branch -M main
%GIT_PATH% push -u origin main -f

echo.
echo Done! If asked for credentials, use your GitHub Personal Access Token.
pause
