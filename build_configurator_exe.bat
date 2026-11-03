@echo off
REM Build ECblue configurator single-file executable (requires Python 3.9+)
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
if errorlevel 1 goto fail
python -m PyInstaller --noconfirm ecblue_configurator.spec
if errorlevel 1 goto fail
if not exist dist\ECBLUE_CONFIGURATOR.exe goto fail
echo.
echo Done: dist\ECBLUE_CONFIGURATOR.exe
pause
exit /b 0
:fail
echo.
echo BUILD FAILED - see messages above.
pause
exit /b 1
