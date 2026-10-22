@echo off
REM SUNINUS MOTOR TEST: copy old-motor RPM (or a manual RPM) to up to 5 new motors.
cd /d "%~dp0"
python tools\motor_tester.py
if errorlevel 1 pause
