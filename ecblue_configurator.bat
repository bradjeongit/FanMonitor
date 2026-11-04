@echo off
REM ECblue configurator: pick IP and ID, read registers with meanings, speed graph, speed and auto-run settings.
cd /d "%~dp0"
python tools\ecblue_configurator.py
if errorlevel 1 pause
