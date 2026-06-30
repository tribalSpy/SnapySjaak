@echo off
REM Dubbelklik op deze .bat om de stickers te genereren.
cd /d "%~dp0"
where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    py -3 "Maak_Stickers.py"
) else (
    where python >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        python "Maak_Stickers.py"
    ) else (
        echo Python is niet geinstalleerd. Download van https://www.python.org/downloads/
        pause
        exit /b 1
    )
)
