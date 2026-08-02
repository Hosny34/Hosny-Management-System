# MEMORY

## Overview

This workspace contains a warehouse management desktop app, a POS desktop app, and a FastAPI sync server that connects them.

The main goal of the system is:

- manage warehouse stock and billing locally on desktop
- manage branch/POS sales locally on desktop
- synchronize selected business events between devices through a central sync server
- keep local-first behavior so each app can continue to work with its own SQLite database

Current active app layout in this workspace:

- `Warehouse/` = main warehouse desktop app (nested folder name)
- `POS-ZAY/` = primary POS desktop app copy
- `POS-OCT/` = second POS copy when present (same codebase pattern as Zay, own `dist` build)
- `sync_server/` = FastAPI sync server project

Additional POS device folders (`POS-OBO`, `POS-GESR`, …) may exist or be recreated from `POS-ZAY` when rollout needs them; reset `.bat` files still name those devices for server-side cleanup.

## High-Level Architecture

The system is made of 3 major parts:

1. Warehouse desktop app
2. POS desktop app
3. Sync server

### Warehouse Desktop App

Main characteristics:

- Python desktop application
- Tkinter/ttk GUI
- SQLite local database
- packaged to Windows `.exe` with PyInstaller

Responsibilities:

- manage warehouse inventory
- create and review bills
- move stock
- audit stock
- manage admin settings
- view mirrored branch stock snapshots
- emit sync events for outbound changes
- pull/apply branch stock snapshots from POS devices

### POS Desktop App

Main characteristics:

- Python desktop application
- Tkinter/ttk GUI
- SQLite local database
- packaged to Windows `.exe` with PyInstaller

Responsibilities:

- manage branch/POS stock locally
- create sales, returns, exchanges, reservations
- manage shift lifecycle
- push branch-side events to the server
- receive warehouse shipment and pricing events
- apply inbound sync events into the local SQLite DB

### Sync Server

Main characteristics:

- FastAPI application
- Uvicorn runtime
- SQLite-backed in the current implemented phase
- JWT-based device auth
- deployed on Railway

Responsibilities:

- issue tokens for devices
- accept push events from clients
- serve pull events to clients
- keep event ordering through monotonic sequence IDs
- store registered/known devices

## Technical Stack

### Languages

- Python 3.10+ targeted
- production source files were re-checked for Python 3.10 compatibility
- current local environment may still build/test with Python 3.12, but the compatibility target should remain Python 3.10

### Desktop UI

- Tkinter
- ttk

### Data Storage

- SQLite for warehouse app
- SQLite for POS app
- SQLite for sync server in the currently wired phase

### Packaging

- PyInstaller (prefer `python -m PyInstaller` on Windows)
- checked-in **`.spec`** files per app: `HosnyWarehouse.spec` (warehouse inner folder), `HosnyPOS.spec` in each POS folder
- typical command from the app directory: `python -m PyInstaller --noconfirm HosnyWarehouse.spec` (or `HosnyPOS.spec`)
- legacy one-file CLI (`pyinstaller --onefile …`) still works but specs are the documented default
- project rule `.cursor/rules/pyinstaller-after-changes.mdc` (`alwaysApply`): after substantive Python fixes, rebuild the affected `dist\*.exe`

### Server

- FastAPI
- Uvicorn
- Pydantic
- PyJWT

### Networking

- HTTP/JSON
- bearer token auth
- `urllib` on client side

## Local Desktop App Design

Both desktop apps are local-first.

Important implication:

- each app has its own local SQLite DB
- UI reads from the local DB only
- sync is event-driven and asynchronous
- sync updates must also trigger UI refresh, not just DB updates

## Sync Model

The sync system is event-based.

Key concepts:

- `sync_outbox`: locally produced events waiting to be pushed
- `sync_inbox`: remotely pulled events waiting to be applied
- `device_identity`: stores local device sync identity/config
- `known_devices`: cache of devices known from server/device list sync
- monotonic `server_seq`: server ordering cursor
- idempotent appliers: event application logic should be safe to replay

### Typical Sync Flow

1. device loads sync configuration
2. device requests JWT token from `/v1/auth/token`
3. device pushes pending outbox events
4. device pulls newer events after last cursor
5. inbound events are stored in `sync_inbox`
6. matching appliers mutate domain tables
7. cursor and last sync timestamps are updated
8. some screens need explicit UI refresh to reflect new DB state

## Authentication Model

The current auth flow supports simple device-name JWT issuance.

### Token Endpoint

Endpoint:

- `POST /v1/auth/token`

Current accepted request body:

```json
{
  "device_name": "POS-ZAY"
}
```

Optional backward-compatible shape also supports old API-key-related flow.

Current response format:

```json
{
  "access_token": "TOKEN_HERE",
  "token_type": "bearer"
}
```

### JWT Details

- algorithm: `HS256`
- signed with shared server secret
- payload includes device information
- lifetime currently set to about 1 year

### Roles

Current device roles are inferred from device naming / server logic:

- warehouse
- pos

## Sync Event Types Seen in This Project

### Warehouse-originated or warehouse-relevant

- `PRICE_UPDATE`
- `STOCK_INCOME`
- `SALE_CREATED`
- `SALE_VOIDED`
- `SALE_RETURNED`
- `STOCK_TRANSFER`
- `STOCK_TRANSFER_OUT`
- `STOCK_AUDIT_APPLIED`
- `STOCK_ADJUST`

### POS-originated or POS-relevant

- `SALE_CREATED`
- `SALE_RETURNED`
- `SALE_EXCHANGED`
- `RESERVATION_CREATED`
- `RESERVATION_PAYMENT_UPDATED`
- `RESERVATION_COMPLETED`
- `RESERVATION_DELIVERED`
- `SHIFT_OPENED`
- `SHIFT_CLOSED`
- `INCOME_BILL_CREATED`
- `STOCK_ADJUST`

### Snapshot / mirror events

- `POS_STOCK_SNAPSHOT`

## Current Event Application Design

### POS-side appliers currently implemented

- `STOCK_TRANSFER_OUT`
- `PRICE_UPDATE`
- `CATALOG_UPSERT`

Meaning:

- the POS can apply warehouse shipments
- the POS can apply warehouse price changes
- the POS can seed catalog/spec metadata

### Warehouse-side appliers currently implemented

- `POS_STOCK_SNAPSHOT`

Meaning:

- warehouse receives and mirrors full POS stock snapshots
- warehouse can display branch stock in `pos_stocks_mirror`

## Important Database Concepts

### Warehouse DB

Contains domain tables such as:

- `stocks`
- `movements`
- `bills`
- `bill_items`
- sync-related tables
- branch mirror tables

Relevant branch mirror tables:

- `pos_stocks_mirror`
- `pos_stocks_snapshot_meta`

### POS DB

Contains domain tables such as:

- `stocks`
- `movements`
- `bills`
- `bill_items`
- `shifts`
- reservations-related tables
- sync-related tables

### Sync Server DB

Contains data such as:

- devices
- pushed events
- pull cursor data
- auth/device records

## Important User-Facing Behaviors

### Warehouse Branch Shipment Behavior

When the warehouse creates a bill targeting a branch POS, the system should emit a shipment event rather than a normal sale event.

Expected behavior:

- selecting a branch in customer list as `فرع: POS-ZAY`
- or entering raw known POS name like `POS-ZAY`
- should route the save path through branch shipment sync logic
- warehouse keeps its own audit/bill records
- POS receives shipment as inbound stock event

### POS Snapshot Behavior

The POS may push a stock snapshot to the warehouse.

This can make sync logs show uploads from POS even when the warehouse has just sent stock down.

That is expected behavior because:

- shipment download and snapshot upload can happen in the same sync cycle

## Major Work Already Done

This section records the important implementation and debugging work completed so far.

### 1. Added Missing Sync Token Endpoint

Problem:

- deployed server returned `404 Not Found` for `POST /v1/auth/token`

Work done:

- implemented token issuance endpoint under `/v1`
- added device-name-based JWT issuance
- kept compatibility for older API-key-style clients

Result:

- clients can request bearer tokens directly from device name

### 2. Fixed Railway Startup Crash

Problem:

- Railway injected a non-SQLite `DATABASE_URL`
- server crashed with:
  `Only SQLite is wired up in this phase`

Work done:

- changed config handling so non-SQLite `DATABASE_URL` falls back safely to SQLite
- updated README notes for Railway behavior

Result:

- server starts successfully on Railway in current SQLite-only phase

### 3. Added Lazy Device Provisioning for Simple JWT Flow

Problem:

- simple JWTs without DB-backed device records caused server-side failures later in sync flow
- downstream logic still needed real `device_uuid` / device row information

Work done:

- added `ensure_simple_device`
- added `upsert_device`
- updated JWT decode logic to provision/fill device records as needed

Result:

- stateless/simple device tokens now map to DB-backed device records
- downstream sync logic can continue to work

### 4. Fixed Warehouse 500 Error During Sync

Problem:

- warehouse-side sync reached auth stage successfully
- later sync failed with HTTP 500
- root cause was missing/provisioning mismatch for device identity behind simple token flow

Work done:

- same provisioning changes above ensured valid device info exists

Result:

- warehouse sync no longer fails there

### 5. Made API Key Optional in Desktop Clients

Problem:

- old clients required API key in UI and validation flow
- new server supports device-name-only token issuance

Work done:

- updated `sync_client.py` to allow missing `api_token`
- updated setup UI to make API key optional
- updated validation logic to only require device name

Result:

- warehouse and POS apps can sync with only server URL + device name

### 6. Fixed Windows/Arabic Editing Issues

Problem:

- path and encoding issues appeared because workspace path includes Arabic
- some automated text replacements caused mojibake / bad characters

Work done:

- used safer editing strategies
- normalized edits
- replaced some fragile UI text with safer ASCII/English where needed in sync setup

Result:

- code compiles correctly
- shell/path handling became more reliable during fixes

### 7. Fixed Warehouse-to-POS Shipment Routing Bug

Problem:

- warehouse only recognized shipment target when customer started with `فرع: `
- plain `POS-ZAY` became a normal `SALE_CREATED`
- POS downloaded but skipped applying it

Work done:

- updated warehouse bill finalize logic so raw known POS names are also treated as branch shipment targets

Result:

- branch shipment sync behavior now works even if user types the POS name directly

### 8. Fixed POS UI Refresh After Sync

Problem:

- POS DB updated after sync but selling point UI did not refresh immediately
- user had to change filters or restart shift/tab flow to see data

Work done:

- added post-sync host notification path
- added POS-specific `sync_refresh()` behavior
- refreshed filter values, favorites, and active POS selector view after sync

Result:

- POS screen updates immediately after sync

### 9. Removed Extra POS Workspace Copies

Problem:

- workspace had many cloned POS folders
- only one active POS copy was needed for current work

Work done:

- removed all extra POS copy folders
- kept only `POS-ZAY`

Result:

- cleaner workspace
- less confusion

### 10. Rebuilt Desktop Executables Multiple Times

Work done repeatedly after edits:

- rebuilt `POS-ZAY\dist\HosnyPOS.exe` (POS executable name)
- rebuilt `Warehouse\dist\HosnyWarehouse.exe`

Result:

- latest fixes were made available in `.exe` outputs

### 11. Cleaned Warehouse Duplicate Navigation

Problem:

- warehouse UI showed duplicate navigation surfaces
- same screens appeared in menu bar, custom header navigation, and visible notebook tabs

Work done:

- hid the redundant notebook tab strip
- kept cleaner top navigation behavior

Result:

- warehouse interface is less cluttered

### 12. Added Default Branch POS Names to Warehouse

Problem:

- after deleting extra POS folders, only synced/known devices appeared
- user still wanted the standard branch list to appear in shipment dropdown

Work done:

- added default branch POS name list in warehouse app

Current default branch names:

- `POS-ZAY`
- `POS-OCT`
- `POS-OBO`
- `POS-GESR`
- `POS-BAH`
- `POS-CEN`

### 13. Cleaned Branch Customer Dropdown Duplicates

Problem:

- dropdown showed duplicates like both `POS-ZAY` and `فرع: POS-ZAY`
- old branch names could mix with regular customers

Work done:

- normalized branch entries with one prefixed form
- removed duplicate customer values
- filtered branch names out of regular customer list

Result:

- cleaner customer dropdown

### 14. Added Filters to Branch Stock Window

Problem:

- branch stock window had only free-text search
- user wanted filters like normal stock window

Work done:

- added branch stock filter controls for:
  - item type
  - school
  - color
  - size
- kept free-text search
- added clear button
- filters are cascaded from selected branch mirror data

Result:

- branch stock tab behaves more like normal stock filtering

## Current UI Notes

### Warehouse App

Current notable windows/features:

- dashboard
- outcome
- income
- statistics
- inventory window
- bill history
- movements
- stock transfer
- stock audit
- admin settings
- branch stock mirror view
- sync dialog and setup

### POS App

Current notable windows/features:

- dashboard/home
- income
- POS selling point
- reservations
- statistics
- shifts summary
- sync dialog and setup

## Current Known Working Behavior

Based on completed work and user feedback:

- server deploy succeeded after crash fix
- warehouse sync can authenticate and complete
- POS sync can authenticate and complete
- warehouse shipments can reach POS
- POS stock can mirror back to warehouse
- POS selling point refresh issue has been addressed
- branch stock filtering has been improved

## Current Known Residual Caveats

These are not necessarily unresolved bugs, but important notes.

### 1. Branch list may still contain stale synced devices

If old device names are still in local cache tables, they may still appear until explicitly cleaned or filtered.

### 2. Branch stock window is similar to normal stock, not identical

It now has comparable filters, but not every advanced control from full inventory is necessarily copied one-to-one.

### 3. Sync is event-apply based

If a wrong event was previously synced before a code fix, that historic bad event does not auto-convert later. Sometimes a fresh shipment or new sync cycle is required after the fix.

### 4. SQLite-only server phase

Postgres is not fully implemented in server DB layer yet.

## Files Touched Significantly During This Work

### Sync Server

- `sync_server/Hosny-sync-server/main.py`
- `sync_server/Hosny-sync-server/auth.py`
- `sync_server/Hosny-sync-server/db.py`
- `sync_server/Hosny-sync-server/config.py`
- `sync_server/README.md`
- `sync_server/requirements.txt`

### Warehouse App

- `Warehouse/HosnyWarehouse.py`
- `Warehouse/sync_client.py`
- `Warehouse/sync_ui.py`
- `Warehouse/sync_appliers.py`
- `Warehouse/sync_core.py`

### POS App

- `POS-ZAY/HosnyPOS.py` (and parallel `POS-OCT/HosnyPOS.py` when that tree exists)
- `POS-ZAY/sync_client.py`
- `POS-ZAY/sync_ui.py`
- `POS-ZAY/sync_appliers.py`
- `POS-ZAY/sync_core.py`

## Important Runtime/Deployment URLs and Names

Known production server URL used in this work:

- `https://web-production-e022.up.railway.app`

Known warehouse device name:

- `WAREHOUSE`

Known branch device names:

- `POS-ZAY`
- `POS-OCT`
- `POS-OBO`
- `POS-GESR`
- `POS-BAH`
- `POS-CEN`

## Build Outputs

Expected Windows executables (after PyInstaller):

- `Warehouse/dist/HosnyWarehouse.exe`
- `POS-ZAY/dist/HosnyPOS.exe`
- `POS-OCT/dist/HosnyPOS.exe` (when `POS-OCT` is maintained in the workspace)

## Workflow Preference Learned From User

The user explicitly requested:

- after edits, run verification/build steps automatically
- when relevant to desktop apps, run PyInstaller after changes

Current preferred post-edit workflow:

1. edit code
2. run `py_compile` on touched modules
3. check lints
4. rebuild affected `.exe` with `python -m PyInstaller --noconfirm <App>.spec` from the correct app directory (warehouse inner folder, or each POS folder)

## Suggested Future Plans

These are the logical next improvements for the project.

### Sync / Server

- implement proper Postgres backend support if Railway DB usage becomes necessary
- add stronger token validation dependencies on sync endpoints if not already fully enforced
- add clearer device list management and stale-device cleanup
- add server-side observability/logging for event apply diagnostics

### Warehouse App

- make branch stock window even closer to regular inventory window if desired
- add stale branch removal / whitelist management in UI
- improve sync status visibility inside main app
- add explicit device management screen

### POS App

- continue validating post-sync UI refresh across all tabs
- ensure reservations/statistics also reflect inbound changes immediately if needed
- recreate other POS copies from `POS-ZAY` when rollout is ready

### General Product Direction

- keep local-first architecture
- keep sync idempotent and auditable
- reduce UI duplication
- improve consistency across warehouse/POS flows
- add focused regression tests where high-value

## Recreating Additional POS Copies Later

When needed, new POS folders can be recreated from the cleaned `POS-ZAY` copy.

Recommended future copies:

- `POS-OCT`
- `POS-OBO`
- `POS-GESR`
- `POS-BAH`
- `POS-CEN`

Each recreated copy should have:

- its own device name in sync setup
- its own local SQLite database
- rebuilt `.exe` if distributed separately

## Summary

This project is now a local-first warehouse/POS system with a functioning FastAPI sync server, simple device-name JWT auth, working warehouse-to-POS shipment flow, mirrored branch stock snapshots back to warehouse, reduced UI duplication, improved branch dropdown behavior, improved branch stock filtering, warehouse reservation mirror views (per-line and aggregated across branches, with robust POS device filtering), merged numeric size-range pickers when two presets overlap, and rebuilt desktop executables for the warehouse and maintained POS app folders.

This file is intended to be a long-term project memory document and should be updated whenever architecture, sync behavior, deployment assumptions, supported device list, or UI workflows change.

## Workflow Update: Additional Warehouse Numbers

The warehouse app originally exposed only 4 warehouse numbers:

- `1`
- `2`
- `3`
- `4`

This was extended to 7 numbers in the warehouse UI.

New warehouse numbers added:

- `5`
- `6`
- `7`

Current intended mapping recorded for operations:

- `5` = `POS-ZAY`
- `6` = `POS-OCT`
- `7` = `POS-OBO`

Implementation note:

- warehouse-side dropdowns and multi-select warehouse filters were updated to use one shared list instead of separate hardcoded `1..4` values
- warehouse selectors now display labels while still storing/processing numeric warehouse numbers internally

Current warehouse UI labels:

- `1` = `مخزن 1`
- `2` = `مخزن 2`
- `3` = `مخزن 3`
- `4` = `مخزن 4`
- `5` = `مخزن زايد`
- `6` = `مخزن اكتوبر`
- `7` = `مخزن العبور`

## Workflow Update: Branch Returns And Transfers

This section records the newer workflow requested after the earlier sync work.

### New Sync Events

- `STOCK_RETURN_TO_WAREHOUSE`
- `POS_TRANSFER_VIA_WAREHOUSE`

### New Warehouse Receiving Model

The earlier implementation auto-inserted POS returns directly into warehouse stock.

That behavior has now been replaced.

Current behavior:

- warehouse sync receives incoming POS returns and POS-to-POS transfer requests
- those lines are inserted into `branch_inventory_queue`
- they stay out of real warehouse stock until a warehouse user processes them

Queue statuses:

- `PENDING`
- `ASSIGNED`
- `DISCARDED`
- `REROUTED`

### Warehouse Processing Actions

Warehouse user can now process each queued line by:

- assigning it to a warehouse number and package number
- discarding it as defective
- rerouting it to another branch

Assign behavior:

- adds the item to real warehouse stock
- records normal inbound stock movement

Discard behavior:

- marks the item processed as defective/discarded
- item never enters warehouse stock

Reroute behavior:

- emits a warehouse shipment to a branch
- item does not enter warehouse stock first

### POS Return Naming Change

The user explicitly requested the POS return action be named:

- `الى المصنع`

This wording replaced the earlier warehouse-target wording in the POS UI flow.

### POS-To-POS Through Warehouse

The user explicitly requested that POS can create a bill and choose a branch target.

Implemented model:

- POS can choose a branch target
- POS does not deliver directly to that branch
- it emits `POS_TRANSFER_VIA_WAREHOUSE`
- warehouse receives that request into the same unhandled queue
- warehouse staff decides when/how to reroute it onward

### New Warehouse UI

Warehouse now includes a dedicated window for queued branch inventory items:

- unhandled branch returns
- unhandled branch transfer requests

That window allows:

- refresh
- assign to stock
- reroute to branch
- discard as defective

### Operational Meaning

After this workflow update:

- a branch return should not immediately appear in warehouse stock after sync
- it should first appear in the unhandled branch queue
- only after warehouse staff processes it should it affect stock or get rerouted

### Documentation Rule

The user explicitly asked that for important workflow changes:

- update `MEMORY.md`
- update `Read Me.txt`

This should be treated as a standing project preference for future sessions.

## Utility Update: Reset Script

A reusable Bash utility was added here:

- `Resets/reset_sync_server.sh`

Purpose:

- reset `POS-ZAY` only on the sync server
- reset `WAREHOUSE` only on the sync server
- wipe all devices/events/cursors on the sync server

Implementation behavior:

- operates directly on `sync_server/Hosny-sync-server/sync_server.sqlite3`
- deletes from `events`, `device_cursors`, and `devices`
- uses confirmation prompts before destructive actions

Later expansion:

- added support for `POS-OCT`, `POS-OBO`, `POS-GESR`, `POS-BAH`, and `POS-CEN`
- added Windows `.bat` launchers so resets can be triggered by double-click
- added `_run_reset.bat` helper
- later changed `_run_reset.bat` to use native Windows batch + `sqlite3.exe` directly so Git Bash is no longer required for the `.bat` reset launchers
- later changed `_run_reset.bat` again to use Python's built-in `sqlite3`, so the Windows reset launchers no longer require Git Bash or `sqlite3.exe`

Current one-click reset files in `Resets`:

- `Reset POS-ZAY.bat`
- `Reset POS-OCT.bat`
- `Reset POS-OBO.bat`
- `Reset POS-GESR.bat`
- `Reset POS-BAH.bat`
- `Reset POS-CEN.bat`
- `Reset WAREHOUSE.bat`
- `Reset ALL Devices.bat`

Clarified behavior:

- the reset tools target the local SQLite sync DB path used by this project
- "stop the server first" only means: if a local sync server process is currently using that same SQLite file, close it before reset

## Utility Update: Project Tools Setup Guide

A plain-text setup guide was added here:

- `Project Tools Setup.txt`

Purpose:

- list the tools needed to run the desktop apps
- list the packages needed to run the sync server
- list the tools needed to build `.exe` files
- list the tools needed to use reset scripts

It includes installation guidance for:

- Python
- desktop Python packages
- sync server requirements
- PyInstaller
- sqlite3
- Git Bash / Git for Windows

## Utility Update: Master Test Scenarios

A dedicated test checklist file was added here:

- `TEST SCENARIOS.md`

Purpose:

- provide a real end-to-end QA checklist for the current system
- give scenario IDs the user can report back quickly
- cover the full visible system feature set, including warehouse operations, POS operations, reset flows, sync flows, restrictions, manager permissions, labels, queue handling, and optional export/print cases

Expected usage:

- user runs scenarios and reports results in `PASS` / `FAIL` form by scenario ID
- this should speed up diagnosis in future sessions

## Utility Update: Staff Training Checklists

Two trainer-facing checklist files were added in the project root:

- `Warehouse Staff Training Checklist.txt`
- `POS Staff Training Checklist.txt`

Purpose:

- help the user train staff consistently
- make sure important workflows are demonstrated in both apps
- cover both daily operations and critical rules/restrictions

## Workflow Update: POS Cashier Lockdown

The user clarified the intended operating model:

- warehouse = source of truth
- POS = controlled execution terminal for sales activity

Current enforcement added in the POS app:

- manual incoming stock from POS is disabled
- POS inventory management entry points are hidden or blocked
- manual stock delete / spec edit / price edit from POS are blocked
- bulk price update from POS is blocked
- admin import / manual adjustment / reset actions are blocked on POS

Implementation note:

- old code paths were intentionally kept in source where practical and disabled with guards/comments instead of being removed completely

### POS transfer rule

Current allowed POS branch-transfer behavior is request-only:

- POS user can raise a transfer request to another branch
- the request goes to warehouse review first
- warehouse decides whether to ship onward

This is not treated as a direct POS-to-POS stock transfer.

### Price and discount policy

Current cashier policy:

- POS should not be able to edit price manually
- POS should not be able to bulk-edit price
- free discounting is not intended in POS

At the time of this update, the dangerous open price-editing paths were blocked.
There was no active free-discount UI exposed in the POS sale flow, so preset discount logic was not added yet.

### Finalized bill policy

Finalized POS bills are now treated as immutable.

New behavior:

- no delete-finalized-bill flow was added
- bill history now supports `VOID` with:
  - manager password
  - required reason
- `VOID` changes bill status instead of deleting it
- stock deducted by a sale is restored for stock-origin sale lines when a bill is voided
- a `SALE_VOIDED` sync event is emitted for audit/sync purposes

### POS naming

The active POS app entry file/executable was also renamed earlier:

- `HosnyPOS.py`
- `HosnyPOS.exe`

## Workflow Update: Manager-Controlled POS Permissions

The earlier cashier-lockdown implementation used hardcoded restrictions.

That was then changed to a manager-controlled model inside POS.

Current behavior:

- POS admin now includes a `صلاحيات المدير` tab
- manager enters password
- manager chooses which restricted POS features are allowed or blocked
- choices are saved locally in SQLite `app_settings`

Saved manager-controlled permissions currently cover:

- opening inventory window
- manual incoming stock
- bulk price update
- Excel import
- manual stock adjustment
- reset counts/history
- manual stock delete
- inventory price editing
- inventory specs editing
- size profile editing

Additional implementation note:

- admin password persistence is now stored in the same settings table instead of relying only on the old hardcoded constant fallback

Meaning:

- the POS can still be run in tight cashier mode
- but the manager can selectively re-enable individual powers without code edits

## Fix Update: Post-Test Corrections

After a later manual test round, several fixes were applied:

- sync server simple auth now rejects unknown device names instead of auto-accepting any typed name
- allowed simple device names are currently limited to `WAREHOUSE` and the approved POS names
- warehouse quick search now prefers item matches before school-name matches to reduce false jumps
- warehouse and POS size-profile lookup now uses trimmed/case-insensitive matching
- warehouse branch shipment events now carry matching size-profile metadata to POS so branch receipts keep the intended size ranges more reliably
- warehouse bill-history returns for branch-shipment bills now queue into `branch_inventory_queue` instead of trying to force stock back into a closed package
- warehouse bills history now labels branch shipments explicitly
- warehouse movements window now shows bill customer/branch and has context-aware linked filters for type/school/color/size

## Update (2026-04): Reservations mirror, financial device filter, merged size ranges

Warehouse UI / SQL (`Warehouse/HosnyWarehouse.py`):

- **`resolve_pos_mirror_device_sql_filter`**: POS filter for mirror + financial ledger expands a chosen label to all matching `source_device` values from `known_devices` (name and UUID), so filters work when sync stored only a UUID.
- **`list_pos_reservations_mirror_device_picklist`**: combobox values = known device names plus distinct `source_device` from `pos_reservations_mirror` and `pos_financial_ledger` (ledger-only devices still selectable for finance).
- **`display_name_for_sync_source`**: human-readable **الجهاز** column in the detail grid when a UUID maps to `known_devices`.
- **`PosReservationsMirrorWindow`**: `ttk.Notebook` — **تفصيلي** (per line) vs **مجمّع حسب المنتج** (`list_pos_reservations_mirror_aggregated`: `GROUP BY` normalized type/school/color/size, `SUM(qty)`, line count, distinct POS count); aggregate tab selected by default; hints on both tabs.
- **`list_pos_reservations_mirror_aggregated`**: `TRIM` + `CAST(size AS TEXT)` in grouping to avoid split duplicates from spacing.

POS + warehouse billing size lists (`merged_numeric_size_labels_from_profile`):

- When a size profile has **two** numeric presets from `ALLOWED_NUMERIC_RANGES`, the UI builds the **set union** of size labels sorted numerically — overlapping ranges (e.g. 6→22 and 14→28) show **6…28 once**, not two button blocks. Implemented in `HosnyWarehouse.py`, `POS-ZAY/HosnyPOS.py`, and `POS-OCT/HosnyPOS.py`; print-size-sheet helpers use the same merge.

Documentation and QA:

- `TEST SCENARIOS.md`: extended `BR-06`, `BR-07`, new `SYNC-05b`; coverage map lines updated.
- `Read Me.txt`, `Project Tools Setup.txt`, staff training checklists: aligned with the above.
- `.cursor/rules/pyinstaller-after-changes.mdc`: standing instruction to rebuild `dist` after fixes.

## Planned Next Session: Pricing Workflow Update

The user explicitly asked that the next continuation should start with pricing workflow improvements.

Requested plan for next session:

- make it easier to find and use warehouse-side price update actions
- when a bill is created and the item price is changed manually during that bill flow, that entered price should become the new default price going forward
- add targeted price update to one specific POS branch only
- add a clear choice between:
  - send price update to all POS branches
  - send price update to one selected POS branch
- add history / audit so it is visible which POS branch received which price update

Priority note:

- this should be treated as the first implementation task for the next session when the user asks to continue the plan

## Update (2026-04): Auth/scope smoke test + runbook alignment

Server smoke test automation:

- Added `sync_server/Hosny-sync-server/smoke_auth_scope.py`.
- Local run command: `python smoke_auth_scope.py` (from `sync_server/Hosny-sync-server`).
- Pass marker: `SMOKE_AUTH_SCOPE_PASSED`.
- Coverage includes:
  - allowlisted simple-device auth behavior (unknown names rejected),
  - simple JWT TTL wiring to `SIMPLE_DEVICE_JWT_TTL_SECONDS`,
  - `_normalize_target_scope` guardrails for POS and warehouse devices.

Runbook updates:

- `Read Me.txt` now includes an `AUTH / SCOPE SMOKE CHECK` section with command + expected output.
- Sync notes explicitly document staggered ~10-minute periodic sync across apps.
- Short practical rules now explicitly include `STOCK-MONITOR` as call-center read-only monitor identity.

## Update (2026-04): Warehouse diagnostics tab + movement monitor semantics

Warehouse app (`Warehouse/HosnyWarehouse.py`) now includes:

- A dedicated main tab **تشخيص المزامنة** (shortcut `F11`) in addition to dashboard widgets.
- Warehouse movements window switched from raw movement log rows to **aggregated per-product monitor** with columns:
  - `وارد`, `مباع`, `متبقي`, `حجوزات` (plus product specs columns).
- Business semantics enforced in monitor:
  - branch shipments are **not** counted as final sales,
  - warehouse sales are counted only for non-branch customers,
  - remaining is total warehouse stock + mirrored branch stock.

Docs updated:

- `Read Me.txt` latest-fix notes updated with diagnostics tab + movement monitor semantics.
- `TEST SCENARIOS.md` updated with:
  - `WH-05` aggregated warehouse movement monitor test,
  - `BR-08` dedicated diagnostics tab visibility/refresh test.

## Update (2026-04): Partial reservation delivery + shipment receipt verification + long-poll nudge

Reservation delivery (POS):

- Reservation bills now support **partial item-level delivery** instead of forcing full-bill delivery.
- Grouping is persisted with `reservations.reservation_group_uuid` so one bill can track multiple item rows coherently.
- Delivery/payment rule implemented:
  - if delivery is partial, cashier must collect full value of delivered items now,
  - any existing upfront paid amount stays allocated to still-pending reservation items.
- Bill is considered fully delivered only when all rows in the group are delivered.

Shipment receipt verification (POS + warehouse):

- POS no longer auto-applies incoming branch shipment quantities blindly.
- Incoming shipment lines are staged, then cashier confirms received quantity per line via checklist popup.
- POS emits `SHIPMENT_RECEIPT_REPORTED` back to warehouse when cashier confirms (including diffs).
- Warehouse receives review popup, can **accept/reject** diffs, and applies compensating stock adjustments only on acceptance.
- Warehouse now tracks per-branch mismatch frequency (`wrong bill` count) from receipt reviews.

Immediate receiver sync wake-up (server + clients):

- New server endpoint: `GET /v1/sync/wait` (long-poll by device scope and `since` cursor).
- Clients call wait in a background thread and trigger immediate sync when `has_updates=true`.
- Periodic sync remains enabled as fallback; long-poll acts as low-latency nudge.

## Update (2026-04): Unhandled branch queue warehouse/package UX

Warehouse queue window (`منتجات الفروع غير المعالجة`) in `Warehouse/HosnyWarehouse.py`:

- `مخزن` input changed from free text to a readonly dropdown using shared warehouse-number labels (`WAREHOUSE_NUMBER_DISPLAY_VALUES`).
- `عبوة` now auto-fills with the next available package number for the selected warehouse via `package_numbers_summary(...)`.
- Warehouse selection parsing in assign flow now uses `warehouse_numeric_value(...)` so labels and numeric IDs resolve consistently.

## Update (2026-04): POS visual parity rollout (Warehouse + Stock Monitor)

Warehouse app (`Warehouse/HosnyWarehouse.py`):

- Expanded from billing-only styling to app-wide parity with POS light theme.
- `ttk` base styles now follow POS spacing and typography: button/input padding, notebook tabs, tree headers, focus borders, and selection colors.
- Header/navigation visuals were shifted to the same light + blue POS style language for better readability consistency across screens.
- Classic Tk defaults were unified (`Segoe UI` defaults for dialogs/listboxes/menus), so popup windows match main UI better.

Stock Monitor app (`POS-STOCK-MONITOR/StockMonitor.py`):

- Added a centralized POS-like UI palette and a global `ttk.Style` theme pass.
- Unified button/input/table styling with Warehouse/POS look-and-feel (including row striping and selection colors).
- Added global Tk font defaults for dialog/menus so reservation-request popup and supporting UI controls keep consistent readability.

Builds:

- Rebuilt updated desktop executables after this visual pass using PyInstaller specs.

## Update (2026-04): UI polish follow-ups (sizing, split ratios, hover)

Warehouse (`Warehouse/HosnyWarehouse.py`):

- Billing split-pane startup ratio tuned to match the operator-adjusted warehouse layout (wider selector area on open).
- School/product tile hover behavior fixed so color returns to the correct base style after mouse leave (no sticky hover color state).

POS-ZAY (`POS-ZAY/HosnyPOS.py`):

- Billing split-pane startup ratio tuned separately to match the operator-adjusted POS layout.
- Partial reservation delivery dialog (`تسليم عناصر الحجز`) now opens larger with a minimum size so bottom action controls are visible immediately.

Stock Monitor (`POS-STOCK-MONITOR/StockMonitor.py`):

- Header layout adjusted so long snapshot/meta text no longer overlaps/hides the `طلب حجز` button; metadata moved to its own line.

Builds:

- Rebuilt affected executables after each UI polish fix (warehouse, POS-ZAY, stock monitor).

## Update (2026-04): Local time display across UI

Goal:

- Avoid user confusion from raw wire-format UTC timestamps (`...Z`) by converting them to **local device time in UI only**.
- Storage/sync payload timestamps remain unchanged (UTC on wire / DB where already used), preserving protocol behavior.

Applied in:

- `Warehouse/sync_ui.py`
- `POS-ZAY/sync_ui.py`
- `POS-OCT/sync_ui.py`
- `Warehouse/HosnyWarehouse.py`
- `POS-ZAY/HosnyPOS.py`
- `POS-OCT/HosnyPOS.py`
- `POS-STOCK-MONITOR/StockMonitor.py`

Behavior:

- ISO timestamps with `Z` are parsed as UTC and rendered in local time (`YYYY-mm-dd HH:MM:SS`) in visible labels/tables.
- Naive timestamps still render in human-readable local format without changing persisted values.
