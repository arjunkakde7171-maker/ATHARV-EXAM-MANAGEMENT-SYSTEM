@echo off
cd /d "%~dp0"
echo Installing required packages...
py -m pip install -r requirements.txt
echo Starting ATHARV EXAM MANAGEMENT SYSTEM V3...
py app.py
pause
