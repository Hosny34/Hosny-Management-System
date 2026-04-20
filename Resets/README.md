# Reset Scripts Guide

This folder contains utilities for resetting sync-server state.

Current script:

- `reset_sync_server.sh`
- `_run_reset.bat`

One-click Windows launchers:

- `Reset POS-ZAY.bat`
- `Reset POS-OCT.bat`
- `Reset POS-OBO.bat`
- `Reset POS-GESR.bat`
- `Reset POS-BAH.bat`
- `Reset POS-CEN.bat`
- `Reset WAREHOUSE.bat`
- `Reset ALL Devices.bat`

## What It Does

The script works directly on the sync server SQLite database:

- `sync_server/Hosny-sync-server/sync_server.sqlite3`

It can delete:

- device registrations
- pulled cursor state
- stored sync events

## What "Stop The Server" Actually Means

- These reset tools work on the local SQLite file:
  `sync_server/Hosny-sync-server/sync_server.sqlite3`
- If you have a local sync server window/terminal currently using that file, close it first.
- If you are only using Railway and not a local SQLite server, these local reset files do not reset Railway by themselves.

Simple rule:

1. If local SQLite server is running, close it.
2. Run the reset file you want.
3. Start your local server again only if you actually use a local server.

## Before You Run It

Recommended:

1. Close any local sync server window if you are using one.
2. Close warehouse/POS apps if you want a clean restart flow.
3. Make sure Python is installed.
4. Git Bash is optional now. It is only needed if you want to run the `.sh` file manually.

## Easiest Way To Run It

On Windows, just double-click one of these files:

- `Reset POS-ZAY.bat`
- `Reset POS-OCT.bat`
- `Reset POS-OBO.bat`
- `Reset POS-GESR.bat`
- `Reset POS-BAH.bat`
- `Reset POS-CEN.bat`
- `Reset WAREHOUSE.bat`
- `Reset ALL Devices.bat`

Each file will:

- run the correct reset directly on Windows
- use Python's built-in `sqlite3`
- keep the window open so you can read the result

## Manual Way To Run It

Use:

- Git Bash
- MSYS2 Bash
- WSL Bash
- any Bash shell with `sqlite3`

Do not run the `.sh` file directly from plain PowerShell unless you are calling it through Bash.

## Main Commands

### 1. Reset POS-ZAY only

This removes:

- `POS-ZAY` device registration on server
- `POS-ZAY` cursor rows
- events created by `POS-ZAY`
- events targeted to `POS-ZAY`

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-zay
```

### 2. Reset POS-OCT only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-oct
```

### 3. Reset POS-OBO only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-obo
```

### 4. Reset POS-GESR only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-gesr
```

### 5. Reset POS-BAH only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-bah
```

### 6. Reset POS-CEN only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-cen
```

### 7. Reset warehouse only

This removes:

- `WAREHOUSE` device registration on server
- `WAREHOUSE` cursor rows
- events created by `WAREHOUSE`
- events targeted to `warehouse`

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" warehouse
```

### 8. Reset all devices completely

This removes:

- all devices
- all events
- all device cursors

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" all
```

## Help Command

You can print usage/help with:

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" --help
```

or

```bash
bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" help
```

## Optional DB Path Override

If you ever need to point the script at another server DB file, use:

```bash
SERVER_DB="/full/path/to/sync_server.sqlite3" bash "/c/Users/youssef.sherif/Downloads/ادارة المخازن/Resets/reset_sync_server.sh" pos-zay
```

## Typical Real Usage

### Clean reset for one POS branch

1. If you run a local SQLite sync server, close it first
2. Double-click the branch reset `.bat` file you want
3. Or run the matching Bash command manually
4. Delete local POS DB files if you want a fully fresh local start:

- `POS-ZAY/warehouse_data.sqlite3`
- `POS-ZAY/warehouse_data.sqlite3-shm`
- `POS-ZAY/warehouse_data.sqlite3-wal`

5. Start the app again
6. Configure sync if needed
7. Sync as a fresh device

### Full clean reset for everything

1. If you run a local SQLite sync server, close it first
2. Double-click `Reset ALL Devices.bat`
3. Delete local DB files on warehouse and POS if you want full clean restart locally too
4. Start apps again

## Important Warnings

- The script is destructive.
- `all` removes every device and every stored sync event from the server DB.
- If you reset server state but keep old local DBs, apps may still hold local history/state.
- If you reset local DBs but not the server DB, old events may be pulled again.
- These reset files are for the local SQLite sync DB path used by this project.

## Recommended Rule

Use:

- `pos-zay` when only one POS is problematic
- `warehouse` when only the warehouse device needs clean server identity
- `all` only when you want a full sync-system restart
