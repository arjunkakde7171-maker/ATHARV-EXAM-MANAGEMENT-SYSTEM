@echo off
title ATHARVKART V3.7 FINAL
cd /d "%~dp0"
if not exist app.py (echo ERROR: app.py missing&pause&exit /b 1)
if not exist requirements.txt (echo ERROR: requirements.txt missing&pause&exit /b 1)
where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
%PY% --version
if errorlevel 1 (echo ERROR: Python not found&pause&exit /b 1)
%PY% -m pip install -r requirements.txt
if errorlevel 1 (echo ERROR: Dependency installation failed&pause&exit /b 1)
%PY% -m py_compile app.py
if errorlevel 1 (echo ERROR: app.py failed syntax check&pause&exit /b 1)
echo Starting ATHARVKART at http://127.0.0.1:5000
start "" http://127.0.0.1:5000
%PY% app.py
pause
