@echo off
REM Media poller for the BEAST PC (Campagnes photos / videos).
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Making the Python environment...
  py -3.12 -m venv .venv || python -m venv .venv
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
rem New requirements after a pull_code (quick when nothing changed).
".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
:loop
".venv\Scripts\python.exe" poller.py
echo Poller stopped -- restarting in 15 seconds (close this window to stop).
timeout /t 15 /nobreak >nul
goto loop
