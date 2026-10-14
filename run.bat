@echo off
cd /d "%~dp0"
.venv\Scripts\python.exe -W ignore scraper.py %*
pause
