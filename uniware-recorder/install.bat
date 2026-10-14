@echo off
rem One time: makes .venv and installs what the recorder needs.
cd /d "%~dp0"
where py >nul 2>nul && (py -3.12 -m venv .venv || py -3 -m venv .venv) || python -m venv .venv
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Install failed. winsdk needs Python 3.9 - 3.12: install Python 3.12 from python.org and run this again.
) else (
  echo.
  echo Installed. Start recording with run_recorder.bat
)
pause
