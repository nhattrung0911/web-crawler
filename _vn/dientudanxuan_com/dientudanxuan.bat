@echo off
cd /d "%~dp0"
python dientudanxuan.py %*
if errorlevel 1 pause
