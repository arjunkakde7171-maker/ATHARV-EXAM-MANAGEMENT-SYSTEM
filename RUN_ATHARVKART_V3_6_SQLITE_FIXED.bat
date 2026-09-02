@echo off
cd /d "%~dp0"
echo Starting ATHARVKART V3.6 SQLite Lock Fixed...
python -m pip install -r requirements.txt
python app.py
pause
