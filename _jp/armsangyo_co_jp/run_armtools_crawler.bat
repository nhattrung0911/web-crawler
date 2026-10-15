@echo off
cd /d "%~dp0"
REM GUI has its own console panel; run headless so closing the window ends everything.
start "" pythonw armtools_crawler_gui.py
exit
