@echo off
title Indian Info Tech - Server
cd /d "%~dp0"
echo ========================================================
echo   Starting Indian Info Tech Database & Website Server...
echo ========================================================
echo.
echo   Permanent Excel Database : database.xlsx
echo   Google Sheets Live Sync  : ACTIVE
echo.
echo   Opening website at http://localhost:5000 ...
start http://localhost:5000
echo.
echo   Server is running! Keep this window open while using the website.
echo   Press Ctrl+C to stop.
echo.
node server-express.js
pause
