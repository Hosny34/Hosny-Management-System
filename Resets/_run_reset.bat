@echo off
setlocal

set "MODE=%~1"
set "SCRIPT_DIR=%~dp0"
set "SERVER_DB=%SERVER_DB%"
if not defined SERVER_DB set "SERVER_DB=%SCRIPT_DIR%..\sync_server\Hosny-sync-server\sync_server.sqlite3"
set "PY_DEVICE="
set "PY_SCOPE="

if "%MODE%"=="" (
    echo Missing reset mode.
    echo Example: _run_reset.bat pos-zay
    pause
    exit /b 1
)

if not exist "%SERVER_DB%" (
    echo Could not find server DB:
    echo %SERVER_DB%
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

set "CONFIRM_TEXT="

if /I "%MODE%"=="pos-zay" (
    set "CONFIRM_TEXT=This will reset POS-ZAY server state and its targeted events."
    set "PY_DEVICE=POS-ZAY"
    set "PY_SCOPE=pos:POS-ZAY"
) else if /I "%MODE%"=="pos-oct" (
    set "CONFIRM_TEXT=This will reset POS-OCT server state and its targeted events."
    set "PY_DEVICE=POS-OCT"
    set "PY_SCOPE=pos:POS-OCT"
) else if /I "%MODE%"=="pos-obo" (
    set "CONFIRM_TEXT=This will reset POS-OBO server state and its targeted events."
    set "PY_DEVICE=POS-OBO"
    set "PY_SCOPE=pos:POS-OBO"
) else if /I "%MODE%"=="pos-gesr" (
    set "CONFIRM_TEXT=This will reset POS-GESR server state and its targeted events."
    set "PY_DEVICE=POS-GESR"
    set "PY_SCOPE=pos:POS-GESR"
) else if /I "%MODE%"=="pos-bah" (
    set "CONFIRM_TEXT=This will reset POS-BAH server state and its targeted events."
    set "PY_DEVICE=POS-BAH"
    set "PY_SCOPE=pos:POS-BAH"
) else if /I "%MODE%"=="pos-cen" (
    set "CONFIRM_TEXT=This will reset POS-CEN server state and its targeted events."
    set "PY_DEVICE=POS-CEN"
    set "PY_SCOPE=pos:POS-CEN"
) else if /I "%MODE%"=="warehouse" (
    set "CONFIRM_TEXT=This will reset WAREHOUSE server state and its device cursor/history."
    set "PY_DEVICE=WAREHOUSE"
    set "PY_SCOPE=warehouse"
) else if /I "%MODE%"=="all" (
    set "CONFIRM_TEXT=This will wipe ALL sync server devices, events, and cursors."
) else (
    echo Unknown mode: %MODE%
    echo.
    echo Allowed modes:
    echo   pos-zay
    echo   pos-oct
    echo   pos-obo
    echo   pos-gesr
    echo   pos-bah
    echo   pos-cen
    echo   warehouse
    echo   all
    pause
    exit /b 1
)

echo %CONFIRM_TEXT%
set /p "REPLY=Type Y to continue: "
if /I not "%REPLY%"=="Y" (
    echo Cancelled.
    pause
    exit /b 0
)

echo.
echo Running reset mode: %MODE%
echo DB: %SERVER_DB%
echo.
if /I "%MODE%"=="all" (
    call %PYTHON_EXE% -c "import sqlite3,sys; conn=sqlite3.connect(sys.argv[1]); conn.execute('PRAGMA foreign_keys = OFF'); conn.execute('BEGIN'); conn.execute('DELETE FROM events'); conn.execute('DELETE FROM device_cursors'); conn.execute('DELETE FROM devices'); conn.commit(); conn.execute('PRAGMA foreign_keys = ON'); conn.close()" "%SERVER_DB%"
) else (
    call %PYTHON_EXE% -c "import sqlite3,sys; conn=sqlite3.connect(sys.argv[1]); conn.execute('PRAGMA foreign_keys = OFF'); conn.execute('BEGIN'); conn.execute('DELETE FROM events WHERE source_device IN (SELECT device_uuid FROM devices WHERE device_name = ?) OR target_scope = ?', (sys.argv[2], sys.argv[3])); conn.execute('DELETE FROM device_cursors WHERE device_uuid IN (SELECT device_uuid FROM devices WHERE device_name = ?)', (sys.argv[2],)); conn.execute('DELETE FROM devices WHERE device_name = ?', (sys.argv[2],)); conn.commit(); conn.execute('PRAGMA foreign_keys = ON'); conn.close()" "%SERVER_DB%" "%PY_DEVICE%" "%PY_SCOPE%"
)
if errorlevel 1 (
    echo.
    echo Reset failed.
) else (
    echo.
    echo Reset completed successfully.
)
echo.
pause
exit /b 0
