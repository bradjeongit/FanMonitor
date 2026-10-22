@echo off
REM Build motor scanner single-file executable (requires Python 3.9+)
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
if errorlevel 1 goto fail
python -m PyInstaller --noconfirm motor_scanner.spec
if errorlevel 1 goto fail
if not exist dist\MOTOR_SCANNER.exe goto fail
echo.
echo Done: dist\MOTOR_SCANNER.exe
pause
exit /b 0
:fail
echo.
echo BUILD FAILED - see messages above.
pause
exit /b 1
