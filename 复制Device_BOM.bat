@echo off
chcp 65001 >nul
cd /d "%~dp0"
python copy_device_bom.py
