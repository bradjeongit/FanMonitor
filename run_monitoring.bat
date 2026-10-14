@echo off
REM Monitoring mode: read-only, never sends control signals to motors.
cd /d "%~dp0"
if exist BLDC_MOTOR_MONITOR.exe (
  start "" BLDC_MOTOR_MONITOR.exe --monitoring
) else (
  python run.py --monitoring
)
