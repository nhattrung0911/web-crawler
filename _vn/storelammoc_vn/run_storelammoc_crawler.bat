@echo off
title Store Lam Moc Product Crawler
echo ===================================================
echo   LAUNCHING STORE LAM MOC HIGH-SPEED PRODUCT CRAWLER...
echo ===================================================
echo.
python "%~dp0storelammoc_crawler_gui.py"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to launch crawler. Please check if Python is installed and in your PATH.
    pause
)
