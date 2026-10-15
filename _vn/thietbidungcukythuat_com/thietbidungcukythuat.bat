@echo off
cd /d "%~dp0"
python thietbidungcukythuat.py %*
if errorlevel 1 pause
