@echo off
title Xsafe Product Crawler
echo ===================================================
echo   LAUNCHING XSAFE HIGH-SPEED PRODUCT CRAWLER...
echo ===================================================
echo.
python "%~dp0xsafe_crawler_gui.py"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to launch crawler. Please check if Python is installed and in your PATH.
    pause
)
