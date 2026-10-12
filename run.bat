@echo off
chcp 65001 >nul
title Restart notes_gemini


echo Stopping old process...
rem Run from the folder of this .bat (/d also switches drive)
cd /d "%~dp0"
powershell -Command "Get-Process py -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*notes_gemini*' -or $_.Path -like '*python*' } | Stop-Process -Force -ErrorAction SilentlyContinue"


echo Starting notes_gemini...
start "" powershell -NoExit -Command "py notes_gemini.py"


echo Opening browser...
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"

echo Done!
