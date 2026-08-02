@echo off
setlocal EnableExtensions EnableDelayedExpansion

cd /d "%~dp0.."
title Publish Hosny POS Update

set "PY_CMD=py -3.10"
%PY_CMD% -V >nul 2>&1
if errorlevel 1 (
    set "PY_CMD=py -3"
    !PY_CMD! -V >nul 2>&1
    if errorlevel 1 (
        set "PY_CMD=python"
        !PY_CMD! -V >nul 2>&1
        if errorlevel 1 goto :missing_python
    )
)

set "NOTES=%~1"
if "%NOTES%"=="" set "NOTES=POS update"
set "HOSNY_POS_UPDATE_NOTES=%NOTES%"
set "REQUESTED_VERSION=%~2"
set "HOSNY_POS_UPDATE_VERSION=%REQUESTED_VERSION%"

echo Updating POS version and preparing local update package...
%PY_CMD% "Tools\publish_pos_update.py" --notes "%NOTES%" --version "%REQUESTED_VERSION%"
if errorlevel 1 goto :fail

echo.
echo Uploading POS update package to sync server...
%PY_CMD% "Tools\upload_pos_update.py"
if errorlevel 1 goto :upload_fail

echo.
echo POS update published successfully.
echo Manifest:
echo   "%CD%\sync_server\Hosny-sync-server\updates\pos\latest.json"
echo Package folder:
echo   "%CD%\sync_server\Hosny-sync-server\updates\pos"
echo.
pause
exit /b 0

:missing_python
echo.
echo No usable Python launcher was found.
echo Install Python or make sure py/python is available, then run this file again.
pause
exit /b 1

:fail
echo.
echo Failed to prepare POS update package.
pause
exit /b 1

:upload_fail
echo.
echo Failed to upload POS update package.
echo The local package was prepared, but it was not published to the sync server.
pause
exit /b 1
