@echo off
REM Update this folder to the latest version on GitHub (main branch).
REM Config.dll and logs are not touched (they are excluded from git).
cd /d "%~dp0"
git pull --ff-only origin main
if errorlevel 1 (
  echo.
  echo Update failed. Local files may have been edited. Ask for help before forcing.
) else (
  echo.
  echo Update complete.
)
pause
