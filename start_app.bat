@echo off

set url="http://127.0.0.1:5000/"

REM python app.py
start cmd /k "myenv\Scripts\python.exe app.py"

REM Wait a few seconds for the server to start
timeout /t 3 >nul

REM open in Chrome
start chrome %url%
