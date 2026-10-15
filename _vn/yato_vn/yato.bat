@echo off
cd /d "%~dp0"
python yato.py %*
if errorlevel 1 pause
