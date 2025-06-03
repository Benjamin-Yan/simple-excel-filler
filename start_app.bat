@echo off

set url="http://127.0.0.1:5000/"

rem open in Chrome
start chrome %url%

python app.py

pause
