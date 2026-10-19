@echo off
REM Build mini controller single-file executable (requires Python 3.9+)
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
if errorlevel 1 goto fail
python -m PyInstaller --noconfirm mini_control.spec
if errorlevel 1 goto fail
if not exist dist\MINI_CONTROL.exe goto fail
echo.
echo Done: dist\MINI_CONTROL.exe
pause
exit /b 0
:fail
echo.
echo BUILD FAILED - see messages above.
pause
exit /b 1
