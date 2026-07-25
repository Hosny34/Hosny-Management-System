@echo off
setlocal EnableDelayedExpansion

cd /d "%~dp0"
title Build Hosny POS - Windows 7 32-bit

set "PY_CMD=py -3.7-32"

echo [0/4] Checking Python 3.7.3 32-bit...
%PY_CMD% -c "import platform,sys; ok=sys.version_info[:3]>=(3,7,3) and sys.version_info[:2]==(3,7) and platform.architecture()[0]=='32bit'; print(sys.version); print(platform.architecture()[0]); raise SystemExit(0 if ok else 1)"
if errorlevel 1 goto :missing_python

echo [1/4] Preparing pip for Python 3.7...
%PY_CMD% -m pip install --upgrade "pip<24.1" "setuptools<68" "wheel<0.42"
if errorlevel 1 goto :fail

echo [2/4] Installing pinned Win7 x86 build dependencies...
%PY_CMD% -m pip install --upgrade --force-reinstall -r "requirements-win7-x86.txt"
if errorlevel 1 goto :fail

echo [3/4] Building HosnyPOS.exe with PyInstaller 4.10...
%PY_CMD% -m PyInstaller --clean --noconfirm "HosnyPOS-win7-x86.spec"
if errorlevel 1 goto :fail

echo Copying CA bundle beside HosnyPOS.exe...
%PY_CMD% -c "import certifi,shutil,pathlib; d=pathlib.Path('dist'); d.mkdir(exist_ok=True); shutil.copyfile(certifi.where(), d / 'cacert.pem')"
if errorlevel 1 goto :fail

echo [4/4] Verifying generated executable is 32-bit x86...
%PY_CMD% -c "from pathlib import Path; p=Path('dist/HosnyPOS.exe'); data=p.read_bytes(); pe=data.index(b'PE\0\0'); machine=int.from_bytes(data[pe+4:pe+6], 'little'); print('PE machine: 0x%04x' % machine); raise SystemExit(0 if machine == 0x014c else 1)"
if errorlevel 1 goto :bad_arch

echo.
echo Build completed successfully.
echo Output: "%~dp0dist\HosnyPOS.exe"
pause
exit /b 0

:missing_python
echo.
echo Python 3.7.3 32-bit was not found by the Windows py launcher.
echo Install the Windows x86 Python 3.7.3 release, then run this file again.
pause
exit /b 1

:bad_arch
echo.
echo Build finished, but the executable is not 32-bit x86.
echo Make sure the Windows x86 Python 3.7.3 installer is installed and selected by py -3.7-32.
pause
exit /b 1

:fail
echo.
echo Build failed.
pause
exit /b 1
