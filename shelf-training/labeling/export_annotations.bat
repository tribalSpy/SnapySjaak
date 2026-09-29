@echo off
setlocal
cd /d "%~dp0"
call "%~dp0..\.venv\Scripts\activate.bat"
python export_annotations.py %*
endlocal
