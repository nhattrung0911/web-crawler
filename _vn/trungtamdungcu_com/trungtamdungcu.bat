@echo off
cd /d "%~dp0"
python trungtamdungcu.py %*
if errorlevel 1 pause
