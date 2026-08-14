@echo off
title Asset Management – ZF
color 1F
echo.
echo  =====================================================
echo   ZF Asset Management Dashboard
echo  =====================================================
echo.
echo  Starting server...
echo.

:: Kill anything already on port 5000
for /f "tokens=5" %%p in ('netstat -aon 2^>nul ^| findstr ":5000 "') do (
    taskkill /F /PID %%p >nul 2>&1
)

:: Give it a moment then open Edge
timeout /t 2 /nobreak >nul
start microsoft-edge:http://localhost:5000

:: Start Flask (this window stays open as the server)
python server.py

echo.
echo  Server stopped. Press any key to close.
pause >nul
