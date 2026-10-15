@echo off
cd /d "%~dp0"
python super_mro.py %*
if errorlevel 1 pause
