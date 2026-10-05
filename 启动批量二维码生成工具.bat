@echo off
cd /d "%~dp0"
python qr_generator.py
if errorlevel 1 pause