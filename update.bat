@echo off
REM Update this folder to the latest version on GitHub (main branch), then download the latest
REM program executables (built by GitHub Actions, release "exe-latest") next to this file.
REM Config.dll, settings files and logs are not touched (they are excluded from git).
setlocal
cd /d "%~dp0"
set "FAIL=0"
set "BASE=https://github.com/bradjeongit/FanMonitor/releases/download/exe-latest"

echo [1/2] Updating program files...
git pull --ff-only origin main
if errorlevel 1 (
  echo   Source update failed. Local files may have been edited. Ask for help before forcing.
  set "FAIL=1"
)

echo.
echo [2/2] Downloading latest executables...
for %%E in (MOTOR_SCANNER MOTOR_TESTER MINI_CONTROL ECBLUE_TESTER) do call :getexe %%E

echo.
if "%FAIL%"=="0" goto ok
echo Update finished with errors - see messages above.
goto end
:ok
echo Update complete.
:end
pause
exit /b

REM ---- download one exe: keep the current file if anything fails
:getexe
set "NAME=%~1.exe"
set "TMPF=%NAME%.download"
if exist "%TMPF%" del /q "%TMPF%"
curl -fsSL --retry 2 -o "%TMPF%" "%BASE%/%NAME%" 2>nul
if not errorlevel 1 goto check
if exist "%TMPF%" del /q "%TMPF%"
powershell -NoProfile -Command "try { [Net.ServicePointManager]::SecurityProtocol = 'Tls12'; Invoke-WebRequest -UseBasicParsing -Uri '%BASE%/%NAME%' -OutFile '%TMPF%' } catch { exit 1 }" >nul 2>&1
:check
if not exist "%TMPF%" goto dlfail
for %%A in ("%TMPF%") do if %%~zA LSS 100000 goto dlfail
move /y "%TMPF%" "%NAME%" >nul 2>&1
if errorlevel 1 goto inuse
echo   %NAME%: updated
exit /b
:dlfail
if exist "%TMPF%" del /q "%TMPF%"
echo   %NAME%: download failed - kept current file (no internet, or not published yet)
set "FAIL=1"
exit /b
:inuse
del /q "%TMPF%" >nul 2>&1
echo   %NAME%: could not replace - close the program and run update.bat again
set "FAIL=1"
exit /b
