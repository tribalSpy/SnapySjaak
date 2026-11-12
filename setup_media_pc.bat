@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem Sets up the BEAST PC for the Campagnes photos / videos (media-poller-app):
rem Python 3.12, ffmpeg (with the RTX's NVENC encoder), the poller's own
rem Python environment and its config.json. Safe to run again: every step
rem skips what is already there.
rem
rem After this: fill in media-poller-app\config.json and start
rem media-poller-app\run_poller.bat (see GPU_PC_SETUP.md, section 5b).

echo ============================================
echo  BEAST PC setup -- Campagnes media poller
echo ============================================
echo.

rem ---- 1. Python 3.12 ---------------------------------------------------
echo [1/5] Python 3.12 ...
set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 (
    py -3.12 --version >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=py -3.12"
)
if not defined PYTHON_CMD (
    echo   Not found -- installing Python 3.12 with winget...
    winget install --id Python.Python.3.12 -e --accept-source-agreements --accept-package-agreements
    where py >nul 2>nul
    if not errorlevel 1 (
        py -3.12 --version >nul 2>nul
        if not errorlevel 1 set "PYTHON_CMD=py -3.12"
    )
)
if not defined PYTHON_CMD (
    echo [ERROR] Python 3.12 is still not available. Install it from
    echo https://www.python.org/downloads/windows/ ^(tick "Add python.exe to PATH"^),
    echo open a NEW terminal and run this script again.
    pause
    exit /b 1
)
echo   using !PYTHON_CMD!

rem ---- 2. ffmpeg --------------------------------------------------------
echo.
echo [2/5] ffmpeg ...
set "FFMPEG_OK="
where ffmpeg >nul 2>nul
if not errorlevel 1 set "FFMPEG_OK=1"
if not defined FFMPEG_OK (
    echo   Not found -- installing ffmpeg ^(Gyan build, has NVENC^) with winget...
    winget install --id Gyan.FFmpeg -e --accept-source-agreements --accept-package-agreements
    rem winget adds it to PATH for NEW terminals only; look for it directly.
    for /f "delims=" %%F in ('dir /b /s "%LOCALAPPDATA%\Microsoft\WinGet\Packages\Gyan.FFmpeg*\ffmpeg.exe" 2^>nul') do set "FFMPEG_EXE=%%F"
    if defined FFMPEG_EXE (
        echo   installed: !FFMPEG_EXE!
        set "FFMPEG_OK=1"
    ) else (
        echo   [WARNING] ffmpeg was not found after the install. Open a NEW terminal
        echo   and run this script again, or install it by hand and set
        echo   "ffmpeg_path" in media-poller-app\config.json.
    )
) else (
    for /f "delims=" %%F in ('where ffmpeg') do if not defined FFMPEG_EXE set "FFMPEG_EXE=%%F"
    echo   found: !FFMPEG_EXE!
)
if defined FFMPEG_EXE (
    "!FFMPEG_EXE!" -hide_banner -loglevel error -f lavfi -i color=c=black:s=256x256:d=1 -c:v h264_nvenc -f null - >nul 2>nul
    if errorlevel 1 (
        echo   NVENC test: not available -- videos will use the CPU encoder ^(slower^).
        echo   Update the NVIDIA driver if this is the RTX PC.
    ) else (
        echo   NVENC test: OK -- videos are encoded on the RTX.
    )
)

rem ---- 3. Python environment -------------------------------------------
echo.
echo [3/5] Python environment media-poller-app\.venv ...
if not exist "media-poller-app\.venv\Scripts\python.exe" (
    !PYTHON_CMD! -m venv media-poller-app\.venv
    if errorlevel 1 (
        echo [ERROR] Could not make the virtual environment.
        pause
        exit /b 1
    )
) else (
    echo   already exists.
)
"media-poller-app\.venv\Scripts\python.exe" -m pip install --upgrade pip >nul
"media-poller-app\.venv\Scripts\python.exe" -m pip install -r media-poller-app\requirements.txt
if errorlevel 1 (
    echo [ERROR] pip install failed -- see above.
    pause
    exit /b 1
)

rem ---- 4. config.json ---------------------------------------------------
echo.
echo [4/5] config.json ...
if not exist "media-poller-app\config.json" (
    copy /y "media-poller-app\config.example.json" "media-poller-app\config.json" >nul
    if defined FFMPEG_EXE (
        "media-poller-app\.venv\Scripts\python.exe" -c "import json,sys;p='media-poller-app/config.json';c=json.load(open(p,encoding='utf-8'));c['ffmpeg_path']=sys.argv[1].replace('\\','/');json.dump(c,open(p,'w',encoding='utf-8'),indent=2)" "!FFMPEG_EXE!"
    )
    echo   made media-poller-app\config.json -- fill it in now ^(opens in Notepad^).
    start "" notepad "media-poller-app\config.json"
) else (
    echo   already there ^(not changed^).
)
if not exist "media-poller-app\music" mkdir "media-poller-app\music"

rem ---- 5. quick test ----------------------------------------------------
echo.
echo [5/5] Test render ^(no server, no Drive^) ...
"media-poller-app\.venv\Scripts\python.exe" -c "from PIL import Image,ImageDraw;im=Image.new('RGB',(2000,2000),(245,245,240));d=ImageDraw.Draw(im);[d.ellipse((400+i*300,700,700+i*300,1000),fill=(120+i*30,60,160)) for i in range(4)];im.save('media-poller-app/work_test.jpg')"
"media-poller-app\.venv\Scripts\python.exe" media-poller-app\poller.py --test-recipe zoom_in --photo media-poller-app\work_test.jpg --title "Test" --duration 4 --out media-poller-app\test_output
if errorlevel 1 (
    echo   [WARNING] The test render failed -- see above.
) else (
    echo   Test video in media-poller-app\test_output\
)
del /q "media-poller-app\work_test.jpg" >nul 2>nul

echo.
echo ============================================
echo  Done. Next:
echo   1. Fill in media-poller-app\config.json:
echo        server_url, api_key ^(= SHADOW_LLM_POLLER_API_KEY on Render^),
echo        agent_name, the Drive login, caption_model ^(an Ollama model^).
echo   2. Drive: Library + Output folders in a Shared Drive with the
echo      service account as Content manager ^(or use the OAuth login^);
echo      paste both links in SnappySjaak - Campagnes - Settings.
echo   3. Start media-poller-app\run_poller.bat
echo ============================================
pause
