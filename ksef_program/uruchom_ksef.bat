@echo off
cd /d "%~dp0"
py -m pip install -r requirements.txt
py ksef_zestawienie_gui.py
pause
