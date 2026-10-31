@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -q pywin32
python setup_odata_move.py
