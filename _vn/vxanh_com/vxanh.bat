@echo off
cd /d "%~dp0"
python vxanh.py %*
if errorlevel 1 pause
