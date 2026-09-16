@echo off
cd /d "%~dp0"
python ltspice_csv_converter.py %*
pause
