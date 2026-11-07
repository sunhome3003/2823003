@echo off
chcp 65001 >nul
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8
python order_auto.py > log.txt 2>&1
type log.txt
pause
