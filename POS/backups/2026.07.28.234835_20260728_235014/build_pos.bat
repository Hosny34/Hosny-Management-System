@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
title Build Hosny POS

set "PY_CMD="
call :try_python "py -3.10"
if not defined PY_CMD call :try_python "py -3.9"
if not defined PY_CMD call :try_python "py -3.8"
if not defined PY_CMD call :try_python "py -3.7"
if not defined PY_CMD call :try_python "py -3"
if not defined PY_CMD goto :missing_python
echo Using Python command: %PY_CMD%

echo [1/2] Installing build dependencies...
%PY_CMD% -m pip install --upgrade pyinstaller certifi openpyxl pywin32
if errorlevel 1 goto :fail

echo [2/2] Building HosnyPOS.exe...
%PY_CMD% -m PyInstaller --clean --noconfirm "HosnyPOS.spec"
if errorlevel 1 goto :fail

echo Copying CA bundle beside HosnyPOS.exe...
%PY_CMD% -c "import certifi,shutil,pathlib; d=pathlib.Path('dist'); d.mkdir(exist_ok=True); shutil.copyfile(certifi.where(), d / 'cacert.pem')"
if errorlevel 1 goto :fail

echo Verifying generated executable can start...
"%~dp0dist\HosnyPOS.exe" --smoke-import
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

:try_python
set "CANDIDATE=%~1"
%CANDIDATE% -c "import sys, tkinter; print(sys.version)" >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=%CANDIDATE%"
)
exit /b 0
