@echo off
REM Build Windows executable (requires Python 3.9+)
python -m pip install --upgrade pyinstaller
pyinstaller --noconfirm motor_monitor.spec
copy /Y run_monitoring.bat dist\BLDC_MOTOR_MONITOR\ >nul
echo.
echo Done: dist\BLDC_MOTOR_MONITOR\BLDC_MOTOR_MONITOR.exe
pause
