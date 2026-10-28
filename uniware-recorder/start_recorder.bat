@echo off
rem Next to get_recorder.bat on the other PC: fetches the newest actions
rem (skipped when offline) and opens the actions dashboard.
cd /d "%~dp0"
set "APP=%~dp0SnapySjaak\uniware-recorder"
if not exist "%APP%\.venv\Scripts\python.exe" (echo Run get_recorder.bat first. & pause & exit /b 1)
git -C "%~dp0SnapySjaak" pull --ff-only -q || echo Could not fetch the newest version -- starting the one on this PC.
cd /d "%APP%"
title Uniware actions dashboard
".venv\Scripts\python.exe" dashboard.py
pause
