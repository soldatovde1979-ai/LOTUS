@echo off
chcp 65001 >nul
title Restart notes_gemini


echo Stopping old process...
cd c:\Projects\LOTUS
powershell -Command "Get-Process py -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*notes_gemini*' -or $_.Path -like '*python*' } | Stop-Process -Force -ErrorAction SilentlyContinue"


echo Starting notes_gemini...
start "" powershell -NoExit -Command "py notes_gemini.py 8766"


echo Opening browser...
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8766"

echo Done!
