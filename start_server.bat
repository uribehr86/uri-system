@echo off
chcp 65001 > nul
cd /d "c:\uri system scan\uri-system"
"C:\Program Files\WindowsApps\PythonSoftwareFoundation.Python.3.12_3.12.2800.0_x64__qbz5n2kfra8p0\python3.12.exe" flask_app.py
pause
