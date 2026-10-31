@echo off
chcp 65001 >nul
cd /d "%~dp0"
python update.py
python -m pip install -q requests
python move_bot.py
if errorlevel 1 pause
