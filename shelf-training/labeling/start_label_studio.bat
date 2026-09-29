@echo off
setlocal
cd /d "%~dp0"
call "%~dp0..\.venv\Scripts\activate.bat"
if not exist "%~dp0..\data\dataset" mkdir "%~dp0..\data\dataset"
for %%I in ("%~dp0..\data\dataset") do set LOCAL_FILES_DOCUMENT_ROOT=%%~fI
set LABEL_STUDIO_LOCAL_FILES_SERVING_ENABLED=true
echo Serving local files from %LOCAL_FILES_DOCUMENT_ROOT%
label-studio start
endlocal
