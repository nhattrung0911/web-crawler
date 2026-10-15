@echo off
cd /d "%~dp0"
python trungtamthietbi.py %*
if errorlevel 1 pause
