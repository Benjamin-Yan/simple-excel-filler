@echo off

python -m venv myenv

REM starts a new cmd session to activate and run everything
call cmd /c "myenv\Scripts\activate.bat && pip install -r requirements.txt && deactivate"
