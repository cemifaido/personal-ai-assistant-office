@echo off
title Office AI Copilot Server
echo ========================================================
echo Avvio del server locale HTTPS (Word, Excel, Outlook)...
echo ========================================================
cd /d "%~dp0"
node server.js
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ATTENZIONE: Il server si e' interrotto o la porta 3000 era occupata.
)
pause
