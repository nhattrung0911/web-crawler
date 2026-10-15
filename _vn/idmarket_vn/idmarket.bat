@echo off
cd /d "%~dp0"
python idmarket.py %*
if errorlevel 1 pause
