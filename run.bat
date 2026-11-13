@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title Restart notes_gemini


echo Starting notes_gemini...
rem Run from the folder of this .bat (/d also switches drive)
cd /d "%~dp0"

rem AI-шлюз gw-proxy (OpenCoder, порт 8787): без него AI-триаж и сводка падают
rem с WinError 10061 «конечный компьютер отверг запрос». Поднимаем, если не слушает.
rem Путь к лаунчеру: C:\OpenCoder\ZOO_ (в старой инструкции указан ZOO — проверяем оба).
netstat -ano | findstr ":8787 " | findstr "LISTENING" >nul
if errorlevel 1 (
  set "GW=C:\OpenCoder\ZOO_\gw-proxy.cmd"
  if not exist "!GW!" set "GW=C:\OpenCoder\ZOO\gw-proxy.cmd"
  if exist "!GW!" (
    echo Starting gw-proxy...
    start "gw-proxy" /min "!GW!"
    timeout /t 3 /nobreak >nul
  ) else (
    echo [i] gw-proxy not found - AI triage will fail with WinError 10061
  )
) else (
  echo gw-proxy already listening on 8787
)

rem Прошлый экземпляр сервер останавливает сам (kill_stale_server в notes_gemini.py)
start "" powershell -NoExit -Command "py notes_gemini.py"


echo Opening browser...
timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"

echo Done!
endlocal
