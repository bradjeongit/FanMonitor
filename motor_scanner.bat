@echo off
REM Motor scanner: find motor IDs on a gateway IP, view/change ECblue settings and comm-loss auto run.
cd /d "%~dp0"
python tools\motor_scanner.py
if errorlevel 1 pause
