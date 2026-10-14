@echo off
REM Build Windows executable (requires Python 3.9+)
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
if errorlevel 1 goto fail
python -m PyInstaller --noconfirm motor_monitor.spec
if errorlevel 1 goto fail
if not exist dist\BLDC_MOTOR_MONITOR\BLDC_MOTOR_MONITOR.exe goto fail
copy /Y run_monitoring.bat dist\BLDC_MOTOR_MONITOR\ >nul
echo.
echo Done: dist\BLDC_MOTOR_MONITOR\BLDC_MOTOR_MONITOR.exe
pause
exit /b 0
:fail
echo.
echo BUILD FAILED - see messages above.
pause
exit /b 1
