@echo off
title Pak QR Print - Print Agent
cd /d "%~dp0"
echo Pak QR Print - Print Agent shuru ho raha hai...
echo (Band karne ke liye is window mein Ctrl+C dabayein.)
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js nahi mila! Pehle https://nodejs.org se Node.js (LTS) install karein, phir ye file dobara chalayein.
  pause
  exit /b 1
)
node agent.js
echo.
echo Agent band ho gaya hai. Dobara chalane ke liye ye file dobara chalayein.
pause
