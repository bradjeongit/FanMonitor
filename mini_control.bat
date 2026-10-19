@echo off
REM ECblue mini controller: add fans by IP/ID, set speed % or follow a reference fan.
cd /d "%~dp0"
python tools\mini_control.py
if errorlevel 1 pause
