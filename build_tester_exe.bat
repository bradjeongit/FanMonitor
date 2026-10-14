@echo off
REM Build ECblue tester single-file executable (requires Python 3.9+)
cd /d "%~dp0"
python -m pip install --upgrade pyinstaller
pyinstaller --noconfirm ecblue_tester.spec
echo.
echo Done: dist\ECBLUE_TESTER.exe
pause
