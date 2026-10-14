@echo off
cd /d "%~dp0"
".venv\Scripts\python.exe" make_report.py
start "" "data\map.html"
