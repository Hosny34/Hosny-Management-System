# Reset Scripts Guide

This folder contains utilities for resetting sync-server state and, for POS
devices, wiping the local POS database too.

Current scripts:

- `reset_sync_server.sh`
- `_run_reset.bat`
- `reset_sync_target.py`

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

The reset flow now supports two server modes:

1. Remote Railway reset through the live sync server admin endpoint
2. Local SQLite reset through `sync_server/Hosny-sync-server/sync_server.sqlite3`

For POS modes, it also tries to delete the matching local POS DB:

- `warehouse_data.sqlite3`
- `warehouse_data.sqlite3-wal`
- `warehouse_data.sqlite3-shm`

Server-side it can delete:

- device registrations
- pulled cursor state
- stored sync events

## What "Stop The Server" Actually Means

- These reset tools work on the local SQLite file:
  `sync_server/Hosny-sync-server/sync_server.sqlite3`
- If you have a local sync server window/terminal currently using that file, close it first.
- If you are using Railway, set `SERVER_URL` so the script can call the
  live reset endpoint on the deployed server.

Simple rule:

1. If local SQLite server is running, close it.
2. Run the reset file you want.
3. Start your local server again only if you actually use a local server.

## Before You Run It

Recommended:

1. Close the target POS app before wiping its local DB.
2. Close any local sync server window if you are using one.
3. Make sure Python is installed.
4. If you want to reset Railway, set:

- `SERVER_URL`

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
- reset the live Railway server if `SERVER_URL` is set
- otherwise fall back to the local sync server SQLite DB
- wipe the local POS DB automatically for POS reset modes when it finds it
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
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-zay
```

### 2. Reset POS-OCT only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-oct
```

### 3. Reset POS-OBO only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-obo
```

### 4. Reset POS-GESR only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-gesr
```

### 5. Reset POS-BAH only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-bah
```

### 6. Reset POS-CEN only

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-cen
```

### 7. Reset warehouse only

This removes:

- `WAREHOUSE` device registration on server
- `WAREHOUSE` cursor rows
- events created by `WAREHOUSE`
- events targeted to `warehouse`

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" warehouse
```

### 8. Reset all devices completely

This removes:

- all devices
- all events
- all device cursors

Command:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" all
```

## Help Command

You can print usage/help with:

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" --help
```

or

```bash
bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" help
```

## Optional DB Path Override

If you ever need to point the script at another server DB file, use:

```bash
SERVER_DB="/full/path/to/sync_server.sqlite3" bash "/c/Users/youssef.sherif/Downloads/Hosny-Management-System/Resets/reset_sync_server.sh" pos-zay
```

## Typical Real Usage

### Clean reset for one POS branch

1. If you run a local SQLite sync server, close it first
2. Double-click the branch reset `.bat` file you want
3. Or run the matching Bash command manually
4. Start the app again
5. Configure sync if needed
6. Sync as a fresh device

### Full clean reset for everything

1. If you run a local SQLite sync server, close it first
2. Double-click `Reset ALL Devices.bat`
3. Start apps again

## Important Warnings

- The script is destructive.
- `all` removes every device and every stored sync event from the server DB.
- If you reset server state but keep old local DBs, apps may still hold local history/state.
- If you reset local DBs but not the server DB, old events may be pulled again.
- For Railway resets, the deployed sync server just needs the updated code
  deployed.

## Recommended Rule

Use:

- `pos-zay` when only one POS is problematic
- `warehouse` when only the warehouse device needs clean server identity
- `all` only when you want a full sync-system restart
