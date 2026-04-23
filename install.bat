@echo off
chcp 65001 >nul

echo === breeding install ===

echo [1/5] pip upgrade...
C:\Python311\python.exe -m pip install --upgrade pip

echo [2/5] installing packages...
C:\Python311\python.exe -m pip install requests pywin32

echo [3/5] pywin32 post-install...
C:\Python311\python.exe C:\Python311\Scripts\pywin32_postinstall.py -install

echo [4/5] installing openpyxl...
C:\Python311\python.exe -m pip install openpyxl requests

echo [5/5] git pull...
cd /d %~dp0
git pull

echo.
echo === credentials ===
setx ARCGIS_BREEDING_USER "ЛОГИН_ЗДЕСЬ"
setx ARCGIS_BREEDING_PASS "ПАРОЛЬ_ЗДЕСЬ"

echo.
echo Done. Restart console before running breedingSync.
pause
