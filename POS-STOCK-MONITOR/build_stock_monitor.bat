@echo off
setlocal

cd /d "%~dp0"
title Build Stock Monitor

echo [1/2] Installing build dependencies...
py -3 -m pip install --upgrade pyinstaller openpyxl pywin32
if errorlevel 1 goto :fail

echo [2/2] Building StockMonitor.exe...
py -3 -m PyInstaller --clean --noconfirm "StockMonitor.spec"
if errorlevel 1 goto :fail

echo.
echo Build completed successfully.
echo Output: "%~dp0dist"
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1
