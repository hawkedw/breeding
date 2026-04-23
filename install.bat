@echo off
chcp 65001 >nul

echo === breeding install ===

echo [1/4] pip upgrade...
C:\Python311\python.exe -m pip install --upgrade pip

echo [2/4] installing packages...
C:\Python311\python.exe -m pip install requests pywin32

echo [3/4] pywin32 post-install...
C:\Python311\python.exe C:\Python311\Scripts\pywin32_postinstall.py -install

echo [4/4] pywin32 post-install...
C:\Python311\python.exe -m pip install openpyxl requests

echo.
echo === credentials ===
setx ARCGIS_BREEDING_USER "ЛОГИН_ЗДЕСЬ"
setx ARCGIS_BREEDING_PASS "ПАРОЛЬ_ЗДЕСЬ"

echo.
echo Done. Restart console before running breedingSync.
pause
