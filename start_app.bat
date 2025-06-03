@echo off

set url=http://127.0.0.1:5000/

rem python app.py
myenv\Scripts\python.exe app.py

rem open in Chrome
start chrome %url%

pause
