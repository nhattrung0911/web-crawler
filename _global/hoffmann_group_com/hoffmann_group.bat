@echo off
cd /d "%~dp0"
python hoffmann_group.py %*
if errorlevel 1 pause
