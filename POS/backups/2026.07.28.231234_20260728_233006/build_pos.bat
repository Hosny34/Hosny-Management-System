@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
title Build Hosny POS

set "PY_CMD=py -3.10"
%PY_CMD% -V >nul 2>&1
if errorlevel 1 (
    set "PY_CMD=py -3"
    !PY_CMD! -V >nul 2>&1
    if errorlevel 1 goto :missing_python
    echo Python 3.10 was not found. Falling back to the installed Python launcher target.
)

echo [1/2] Installing build dependencies...
%PY_CMD% -m pip install --upgrade pyinstaller certifi openpyxl pywin32
if errorlevel 1 goto :fail

echo [2/2] Building HosnyPOS.exe...
%PY_CMD% -m PyInstaller --clean --noconfirm "HosnyPOS.spec"
if errorlevel 1 goto :fail

echo Copying CA bundle beside HosnyPOS.exe...
%PY_CMD% -c "import certifi,shutil,pathlib; d=pathlib.Path('dist'); d.mkdir(exist_ok=True); shutil.copyfile(certifi.where(), d / 'cacert.pem')"
if errorlevel 1 goto :fail

echo.
echo Build completed successfully.
echo Output: "%~dp0dist"
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 0

:missing_python
echo.
echo No usable Python launcher target was found.
echo Install Python and make sure the Windows py launcher is available, then run this file again.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1

:fail
echo.
echo Build failed.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1
