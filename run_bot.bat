@echo off
title Medusa AI - Telegram Bot
cd /d "%~dp0"

echo ========================================================
echo               Medusa AI Telegram Bot
echo ========================================================
echo.

:: Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in your PATH.
    pause
    exit /b 1
)

:: Run the bot
python bot.py

pause

