@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
title Build Hosny POS - Windows 7 64-bit

set "TCL_LIBRARY="
set "TK_LIBRARY="

set "PY_CMD="
call :try_python "py -3.7-64"
if not defined PY_CMD call :try_python "py -3.7"
if not defined PY_CMD call :try_python "python"
if not defined PY_CMD goto :missing_python

echo [0/4] Checking Python 3.7.3 64-bit...
%PY_CMD% -c "import platform,sys; print(sys.version); print(platform.architecture()[0])"
if errorlevel 1 goto :missing_python
%PY_CMD% -c "import tkinter; t=tkinter.Tcl(); print('Tcl/Tk OK:', t.eval('info patchlevel'))"
if errorlevel 1 goto :bad_tcl

echo [1/4] Preparing pip for Python 3.7...
%PY_CMD% -m pip --version >nul 2>&1
if errorlevel 1 (
    echo pip was not found. Bootstrapping pip with ensurepip...
    %PY_CMD% -m ensurepip --default-pip
    if errorlevel 1 goto :pip_missing
)
%PY_CMD% -m pip install --upgrade "pip<24.1" "setuptools<68" "wheel<0.42"
if errorlevel 1 goto :fail

echo [2/4] Installing pinned Win7 x64 build dependencies...
%PY_CMD% -m pip install --upgrade --force-reinstall -r "requirements-win7-x64.txt"
if errorlevel 1 goto :fail

echo [3/4] Building HosnyPOS.exe with PyInstaller 4.10...
%PY_CMD% -m PyInstaller --clean --noconfirm "HosnyPOS-win7-x64.spec"
if errorlevel 1 goto :fail

echo Copying CA bundle beside HosnyPOS.exe...
%PY_CMD% -c "import certifi,shutil,pathlib; d=pathlib.Path('dist'); d.mkdir(exist_ok=True); shutil.copyfile(certifi.where(), d / 'cacert.pem')"
if errorlevel 1 goto :fail

echo [4/4] Verifying generated executable is 64-bit x64...
%PY_CMD% "verify_pe_arch.py" "dist\HosnyPOS.exe" x64
if errorlevel 1 goto :bad_arch

echo.
echo Build completed successfully.
echo Output: "%~dp0dist\HosnyPOS.exe"
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 0

:missing_python
echo.
echo Python 3.7.3 64-bit was not found.
echo Install the Windows x64 Python 3.7.3 release, then run this file again.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1

:bad_arch
echo.
echo Build finished, but the executable is not 64-bit x64.
echo Make sure the Windows x64 Python 3.7.3 installer is installed and selected by py -3.7-64.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1

:bad_tcl
echo.
echo Python 3.7.3 was found, but Tcl/Tk is not installed correctly.
echo Re-run the Python 3.7.3 installer, choose Modify, and enable tcl/tk and IDLE.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1

:fail
echo.
echo Build failed.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1

:pip_missing
echo.
echo pip is not installed and ensurepip could not install it.
echo Re-run the Python 3.7.3 installer, choose Modify, and enable pip.
if /I not "%HOSNY_AUTO_UPDATE%"=="1" pause
exit /b 1

:try_python
set "CANDIDATE=%~1"
%CANDIDATE% -c "import platform,sys; ok=sys.version_info[:3]>=(3,7,3) and sys.version_info[:2]==(3,7) and platform.architecture()[0]=='64bit'; raise SystemExit(0 if ok else 1)" >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=%CANDIDATE%"
)
exit /b 0
