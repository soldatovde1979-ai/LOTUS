@echo off
chcp 65001 >nul
title Restart notes_gemini


echo Starting notes_gemini...
rem Run from the folder of this .bat (/d also switches drive)
cd /d "%~dp0"
rem Прошлый экземпляр сервер останавливает сам (kill_stale_server в notes_gemini.py)
start "" powershell -NoExit -Command "py notes_gemini.py"


echo Opening browser...
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"

echo Done!
