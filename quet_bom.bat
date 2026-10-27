@echo off
chcp 65001 >nul
cd /d "%~dp0"
python -m pip install -q -r requirements.txt
set /p DIR=Nhap duong dan folder BOM (vi du D:\boho\BOM): 
python tools\check_import.py "%DIR%"
python tools\scan_boms.py "%DIR%"
echo.
echo Xong. Gui 2 file bom_check_report.md va bom_scan_report.md cho Claude.
pause
