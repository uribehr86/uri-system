@echo off
chcp 65001 > nul
cd /d d:\uri-system
call venv\Scripts\activate.bat
python flask_app.py
pause
