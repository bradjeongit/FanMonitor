@echo off
REM ECblue old-motor tester: 172.27.0.95:502, motor ID 180 and 176 (fixed).
cd /d "%~dp0"
python tools\ecblue_tester.py
if errorlevel 1 pause
