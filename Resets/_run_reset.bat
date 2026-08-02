@echo off
setlocal

set "MODE=%~1"
set "SCRIPT_DIR=%~dp0"

if "%MODE%"=="" (
    echo Missing reset mode.
    echo Example: _run_reset.bat pos-zay
    pause
    exit /b 1
)

set "PYTHON_EXE="
for /f "delims=" %%I in ('where python 2^>nul') do (
    if not defined PYTHON_EXE set "PYTHON_EXE=%%I"
)
if not defined PYTHON_EXE (
    for /f "delims=" %%I in ('where py 2^>nul') do (
        if not defined PYTHON_EXE set "PYTHON_EXE=%%I -3"
    )
)
if not defined PYTHON_EXE (
    echo Python was not found.
    echo Install Python first, then run this file again.
    pause
    exit /b 1
)

call %PYTHON_EXE% "%SCRIPT_DIR%reset_sync_target.py" "%MODE%"
set "ERR=%ERRORLEVEL%"
echo.
if not "%ERR%"=="0" (
    echo Reset failed.
) else (
    echo Reset completed successfully.
)
echo.
pause
exit /b %ERR%
