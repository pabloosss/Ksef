@echo off
cd /d %~dp0
python -m pip install -r requirements_ksef_zestawienie.txt
python -m playwright install chromium
python ksef_program_zestawienie_v2.py
pause
