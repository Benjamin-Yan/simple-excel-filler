@echo off

python -m venv myenv

myenv\Scripts\activate

pip install -r requirements.txt

deactivate
