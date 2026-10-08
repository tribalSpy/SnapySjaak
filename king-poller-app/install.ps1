# King poller installer -- started by install.bat. Asks only what it can't
# know (the King folders), fills in the rest, tests it, and makes the poller
# start by itself at every logon. Safe to run again: it keeps the previous
# answers as defaults. ASCII only on purpose (Windows PowerShell 5.1 reads a
# BOM-less script as ANSI).

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms

$source = Split-Path -Parent $MyInvocation.MyCommand.Path
$target = Join-Path $env:LOCALAPPDATA "KingPoller"
$taskName = "King poller"

function Say($text, $color = "Gray") { Write-Host $text -ForegroundColor $color }
function Step($text) { Write-Host ""; Write-Host "== $text ==" -ForegroundColor Cyan }

function New-TopForm { New-Object System.Windows.Forms.Form -Property @{ TopMost = $true } }

function Ask-YesNo($question) {
  $answer = [System.Windows.Forms.MessageBox]::Show((New-TopForm), $question, "King poller setup", "YesNo", "Question")
  return $answer -eq "Yes"
}

function Pick-Folder($description, $default) {
  Say $description "White"
  $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
  $dialog.Description = $description
  $dialog.ShowNewFolderButton = $true
  if ($default -and (Test-Path $default)) { $dialog.SelectedPath = $default }
  if ($dialog.ShowDialog((New-TopForm)) -eq "OK" -and $dialog.SelectedPath) {
    return $dialog.SelectedPath
  }
  # Network paths (\\SERVER\share) can also just be typed.
  $typed = Read-Host "Type the folder path instead (Enter = $default)"
  if ([string]::IsNullOrWhiteSpace($typed)) { return $default }
  return $typed.Trim()
}

# The real python.exe (its full path), not the Microsoft Store placeholder.
function Find-Python {
  foreach ($call in @(@("py", "-3"), @("python"))) {
    try {
      if (-not (Get-Command $call[0] -ErrorAction SilentlyContinue)) { continue }
      $arguments = @($call | Select-Object -Skip 1) + @("-c", "import sys; print(sys.executable)")
      $exe = (& $call[0] @arguments 2>$null | Select-Object -First 1)
      if ($exe -and (Test-Path $exe.Trim()) -and ($exe -notmatch "WindowsApps")) { return $exe.Trim() }
    } catch { }
  }
  return $null
}

function Refresh-Path {
  $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
}

Clear-Host
Say "King poller setup" "Green"
Say "This PC will deliver the 'Import naar King' exports from the app into King's import folder."
Say "You will be asked for 2 folders. Everything else is filled in for you."

# --- 1. Python ---------------------------------------------------------------
Step "1/5  Python"
$python = Find-Python
if (-not $python) {
  if (-not (Ask-YesNo "Python is needed and is not installed on this PC.`n`nInstall it now? (takes 1-2 minutes)")) {
    Say "Setup stopped: Python is required." "Red"; exit 1
  }
  if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Say "This PC has no winget. Install Python from https://www.python.org/downloads/ (tick 'Add python.exe to PATH'), then run install.bat again." "Yellow"
    Start-Process "https://www.python.org/downloads/"
    exit 1
  }
  Say "Installing Python..."
  winget install -e --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements --silent
  Refresh-Path
  $python = Find-Python
  if (-not $python) {
    $guess = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Filter python.exe -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($guess) { $python = $guess.FullName }
  }
  if (-not $python) { Say "Python was installed but can't be found yet. Restart the PC and run install.bat again." "Red"; exit 1 }
}
Say "Python: $python" "Green"

# --- 2. Copy the program -----------------------------------------------------
Step "2/5  Copying the poller to $target"
try { schtasks /End /TN $taskName 2>$null | Out-Null } catch { }
New-Item -ItemType Directory -Force -Path $target | Out-Null
foreach ($file in @("poller.py", "config.defaults.json", "README.md")) {
  $from = Join-Path $source $file
  if (Test-Path $from) { Copy-Item $from (Join-Path $target $file) -Force }
}
Get-ChildItem $target | Unblock-File -ErrorAction SilentlyContinue

$configPath = Join-Path $target "config.json"
$previous = @{}
if (Test-Path $configPath) {
  try { (Get-Content $configPath -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $previous[$_.Name] = $_.Value } } catch { }
}

# --- 3. The folders ----------------------------------------------------------
Step "3/5  King folders"
$importDefault = $previous["import_dir"]
$importDir = Pick-Folder "Choose the folder King's scheduler reads the import files from (ask finance / King if unsure)." $importDefault
while (-not $importDir -or -not (Test-Path $importDir)) {
  Say "That folder doesn't exist or can't be reached from this PC: $importDir" "Red"
  $importDir = Pick-Folder "Choose King's import folder again." $importDefault
}
Say "Import folder: $importDir" "Green"

$pdfDefault = if ($previous["pdf_dir"]) { $previous["pdf_dir"] } else { Join-Path $importDir "pdf" }
if (Ask-YesNo "The invoice PDFs will go into:`n`n$pdfDefault`n`nIs that OK? (No = choose another folder)") {
  $pdfDir = $pdfDefault
} else {
  $pdfDir = Pick-Folder "Choose the folder for the invoice PDFs." $pdfDefault
}
if (-not (Test-Path $pdfDir)) { New-Item -ItemType Directory -Force -Path $pdfDir | Out-Null }
Say "PDF folder: $pdfDir" "Green"

$kingDefault = if ($previous["king_pdf_dir"]) { $previous["king_pdf_dir"] } else { $pdfDir }
Say ""
Say "King must find the PDFs itself. If King runs on another server, that server may see this" "White"
Say "folder under a different name (for example D:\KingImport\pdf instead of \\SERVER\KingImport\pdf)." "White"
$kingPdfDir = Read-Host "How does the KING SERVER see the PDF folder? (Enter = $kingDefault)"
if ([string]::IsNullOrWhiteSpace($kingPdfDir)) { $kingPdfDir = $kingDefault }

# --- 4. Save + test ----------------------------------------------------------
Step "4/5  Saving and testing"
$config = [ordered]@{
  agent_name           = "king-export-$($env:COMPUTERNAME.ToLower())"
  pc_name              = $env:COMPUTERNAME
  import_dir           = $importDir
  pdf_dir              = $pdfDir
  king_pdf_dir         = $kingPdfDir.Trim()
  archive_wait_minutes = 30
  poll_interval_seconds = 30
}
# UTF-8 without BOM (Windows PowerShell's Set-Content would add one).
[System.IO.File]::WriteAllText($configPath, ($config | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))

# Starter: restarts the poller by itself if it ever stops.
$runner = @"
@echo off
title King poller (keep this window open)
cd /d "%~dp0"
:loop
"$python" poller.py
echo King poller stopped -- restarting in 30 seconds...
timeout /t 30 /nobreak >nul
goto loop
"@
Set-Content -Path (Join-Path $target "run_poller.bat") -Value ($runner -replace "`r?`n", "`r`n") -Encoding ASCII -NoNewline

Push-Location $target
& $python poller.py --check
$checkOk = $LASTEXITCODE -eq 0
Pop-Location
if (-not $checkOk) {
  Say ""
  Say "The test above failed (see the FAIL line). Fix it and run install.bat again." "Red"
  Say "Nothing else was changed." "Red"
  exit 1
}

# --- 5. Start automatically --------------------------------------------------
Step "5/5  Start automatically"
$runPath = Join-Path $target "run_poller.bat"
schtasks /Create /TN $taskName /TR "`"$runPath`"" /SC ONLOGON /RL LIMITED /F | Out-Null
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath("Desktop")) "King poller.lnk"))
$shortcut.TargetPath = $runPath
$shortcut.WorkingDirectory = $target
$shortcut.WindowStyle = 7
$shortcut.Save()
Start-Process -FilePath $runPath -WorkingDirectory $target -WindowStyle Minimized

Say ""
Say "Done! The King poller is running (minimized window 'King poller')." "Green"
Say "It starts by itself every time this PC logs on; there is also a 'King poller' shortcut on the desktop."
Say "In the app (Inkoop Controle > Import naar King) it now shows as online."
Say "King PDF folder to use in the app's King settings: $($kingPdfDir.Trim())"
