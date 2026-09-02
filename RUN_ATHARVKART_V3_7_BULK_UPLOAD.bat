@echo off
cd /d "%~dp0"
echo Starting ATHARVKART V3.7 Bulk Upload...
python -m pip install -r requirements.txt
python app.py
pause
