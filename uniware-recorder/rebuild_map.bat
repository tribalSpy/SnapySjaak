@echo off
cd /d "%~dp0"
rem Re-reads all recorded screenshots with the current recognition.
".venv\Scripts\python.exe" rebuild_map.py
pause
