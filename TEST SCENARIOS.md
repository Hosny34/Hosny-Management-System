# Test Scenarios

Use this file as the main QA checklist for the current system.

This version is intended to cover the full visible system feature set, not only the main sync flows.

Recommended reporting format back to me:

- `ENV-01 PASS`
- `SYNC-04 FAIL - POS did not refresh until filter changed`
- `RET-03 PASS`

Scenario types:

- required core scenarios: should be tested every serious round
- extended scenarios: test the rest of the visible app behavior
- optional scenarios: test only if you use that feature in real work

## 1. Test Preparation

Before starting:

1. Decide whether this is:
   - normal test on current data
   - clean test after reset
2. If you want a clean test:
   - run the correct reset `.bat` file from `Resets`
   - delete the local DB files for the device you want to reset
   - open the app again
3. Keep notes of:
   - device name used
   - exact steps
   - what you expected
   - what actually happened

### Automated subset (no GUI / no sync server)

From the repo root, with Python 3.10+ and `pytest` installed (`pip install pytest`):

```bash
python -m pytest tests/test_scenarios_automated.py -v
```

These tests use a temporary SQLite DB and `POS-ZAY/HosnyPOS.py` (`SqliteDatabase`). They map to checklist-style cases such as **ENV-03** (schema), **SALE-01**, **SALE-03**, **RET-01** (warehouse return bill row), **VOID-01/02**, **LOCK-01**, reservation **DELIVER_PAY** + `get_sales_stats`, and money parsing. They do **not** replace manual runs for **ENV-01**, **SYNC-***, UI, or print flows.

## 2. Feature Coverage Map

Use this as a quick master checklist to confirm that no area was skipped.

### Warehouse app coverage

- startup and DB recreation
- sync setup and connection
- stock entry / inbound stock
- stock search and filters
- warehouse-number labeled selectors
- package-number flows
- outbound shipment to branch
- branch target dropdown behavior
- branch stock mirror window
- branch stock mirror filters
- unhandled branch products queue
- assign / discard / reroute actions
- billing/history views
- branch bill log + sync queue log (shipments / inbound queue tabs)
- POS reservations mirror (device filter, detail vs aggregated tabs) and POS financial-by-day (after sync)
- size profiles with two numeric ranges: merged size picker (no duplicate size buttons)
- branch price sync audit log (POS fan-out decisions)
- multi-row price edit with mixed sizes + optional POS sync
- navigation cleanup / visible tabs
- export/print flows if used

### POS app coverage

- startup and DB recreation
- sync setup and connection
- selling screen refresh after sync
- normal sale
- multi-line sale
- customer selection / autocomplete
- `الى المصنع` return flow
- branch transfer request flow
- bill history (including bill type column and نوع الفاتورة filter)
- `VOID` with reason
- cashier lockdown restrictions
- manager permissions tab
- permission persistence after restart
- inventory window behavior
- admin/settings behavior
- export/print flows if used

### Sync/server/reset coverage

- auth and connection test
- warehouse to POS shipment
- POS stock mirror back to warehouse
- POS return to warehouse queue
- POS to POS via warehouse
- repeat sync safety (no duplicate stock or queue rows)
- periodic background sync (warehouse / POS) without long UI freeze
- manual sync summary and received-event details modal
- size profiles embedded in branch shipment payload applied on POS
- reset by branch
- reset warehouse
- reset all

## 3. Test Environments

### `ENV-01` App opens normally

Steps:

1. Open warehouse app
2. Open POS app

Expected:

- both apps open without crash
- main windows load fully
- no startup error dialog

### `ENV-02` Existing DB opens correctly

Steps:

1. Start app with current DB
2. Open main screens

Expected:

- old data loads
- stock grids open
- bills history opens

### `ENV-03` Fresh DB recreation works

Steps:

1. Close the app
2. Delete local DB files for the app under test
3. Reopen the app

Expected:

- app recreates a new DB automatically
- no schema crash
- app remains usable

### `ENV-04` Sync setup dialog opens

Steps:

1. Open sync settings from warehouse
2. Open sync settings from POS

Expected:

- dialog opens
- server URL field works
- device name field works
- token/API key guidance is visible

## 4. Reset Scenarios

### `RST-01` Single branch server reset

Steps:

1. Run the branch reset `.bat` file
2. Delete local DB files for that branch
3. Open the branch again

Expected:

- branch behaves like a fresh device
- old local state does not remain

### `RST-02` Warehouse-only reset

Steps:

1. Run `Reset WAREHOUSE.bat`
2. Delete warehouse local DB files
3. Reopen warehouse

Expected:

- warehouse starts clean
- no old local sync cursor remains

### `RST-03` Full system reset

Steps:

1. Run `Reset ALL Devices.bat`
2. Delete local DB files for warehouse and tested POS
3. Open apps again

Expected:

- system starts from zero state
- no old device history remains

## 5. Authentication And Connection

### `AUTH-01` Test connection succeeds

Steps:

1. Enter server URL in POS sync setup
2. Enter correct device name
3. Run test connection

Expected:

- connection succeeds
- no 401 or 500 error

### `AUTH-02` Wrong device name is rejected

Steps:

1. Enter invalid or mistyped device name
2. Run test connection or sync

Expected:

- clear error appears
- app does not silently continue

### `AUTH-03` Warehouse device authenticates

Steps:

1. Configure warehouse sync
2. Run test connection

Expected:

- warehouse authenticates successfully

## 6. Warehouse Core Operations

### `WHS-01` Warehouse main screens open

Steps:

1. Open warehouse app
2. Open main major screens/windows from the menus you normally use

Expected:

- no crash
- each window opens
- data grids render correctly

### `WHS-02` Warehouse stock filters work

Steps:

1. Open warehouse stock view
2. Filter by item type
3. Filter by school
4. Filter by color
5. Filter by size
6. Clear filters

Expected:

- filters narrow the result correctly
- clear returns the full set

### `WHS-03` Warehouse search works

Steps:

1. Search using known text from an item
2. Search using non-matching text

Expected:

- matching rows appear for valid search
- empty result behaves safely for invalid search

### `WHS-04` Inbound stock entry works

Steps:

1. Add a normal inbound stock line in warehouse
2. Save it

Expected:

- stock increases
- row appears in stock and history where appropriate

### `WHS-05` Warehouse numbers and package numbers save correctly

Steps:

1. Add stock using one of the labeled warehouse values
2. Enter package number
3. Save

Expected:

- save succeeds
- stock can be found later by the same warehouse/package values

### `WHS-06` Warehouse billing/history window loads

Steps:

1. Open warehouse billing/history views
2. Open an existing record if available

Expected:

- history opens
- old bills/rows are visible

### `WHS-07` Warehouse customer or branch dropdown behaves correctly

Steps:

1. Open warehouse billing target/customer dropdown
2. Type partial branch name
3. Select a branch

Expected:

- approved POS names appear
- no duplicates
- selection works normally

### `WHS-08` Warehouse export or print works if used

Steps:

1. Run your usual export/print action from warehouse

Expected:

- export file is created or print dialog opens
- no crash

### `WHS-09` Multi-select price change with different sizes (POS sync allowed)

Prerequisites: two or more inventory rows with the **same item type and school** but **different sizes** (and any colour/location you use in real data).

Steps:

1. Open warehouse **المخزون**, multi-select those rows (Ctrl/Shift).
2. Click **تعديل السعر…**, enter a new price.
3. Choose **كل فروع البيع** or **فروع محددة…** (not **لا**) for POS sync, then save.
4. Sync affected POS devices.

Expected:

- no warning that blocks you solely because sizes differ.
- warehouse rows update; POS matching stock lines show the new price after sync.

### `WHS-10` Branch price sync audit log opens

Steps:

1. From warehouse inventory toolbar, open **سجل أسعار الفروع…** (branch price sync audit).

Expected:

- dialog opens and table is visible.
- no error about mixing `grid` and `pack` geometry managers.

## 7. Warehouse To POS Shipment Flow

### `SYNC-01` Warehouse shipment reaches target POS

Steps:

1. In warehouse, create stock shipment to one branch
2. Sync warehouse
3. Sync that POS

Expected:

- warehouse upload succeeds
- POS download succeeds
- shipped stock appears in POS

### `SYNC-02` POS selling screen refreshes immediately after sync

Steps:

1. Add new items in warehouse shipment to POS
2. Sync warehouse
3. Sync POS
4. Go to POS selling point

Expected:

- items appear without needing shift restart
- items appear without forcing manual filter refresh

### `SYNC-03` Warehouse branch stock mirror updates

Steps:

1. Sync POS after receiving goods
2. In warehouse, sync again if needed
3. Open branch stock window

Expected:

- branch stock mirror shows updated rows

### `SYNC-04` Sync summary direction makes sense

Steps:

1. Perform shipment from warehouse to POS
2. Sync warehouse
3. Sync POS

Expected:

- warehouse mainly uploads shipment event
- POS mainly downloads incoming event
- summary is not misleading

### `SYNC-05` Branch shipment applies warehouse size profiles on POS

Prerequisites: in warehouse, configure **نطاقات المقاسات** for a specific `(item_type, school, color)` that you will ship to a POS.

Steps:

1. Create a branch shipment containing at least one line for that triple.
2. Sync warehouse, then sync the target POS.
3. On POS, use any screen that reads `size_profiles` for that triple (e.g. selling grid / size picker).

Expected:

- after the shipment event is applied, POS size ranges match the warehouse configuration for that triple.
- no TclError or SQL error during apply.

### `SYNC-05b` Two overlapping numeric ranges show one merged size list (warehouse + POS)

Prerequisites: in warehouse **تعديل نطاقات المقاسات…**, set **النطاق الأول** and **النطاق الثاني** to two presets that overlap (example: **6 → 22** and **14 → 28**). Use a triple you can open on POS selling flow.

Steps:

1. Save the profile; sync if the POS must receive the profile via shipment.
2. Open the size picker for that triple on **warehouse billing** and on **POS selling** (اختر المقاس).

Expected:

- each numeric size appears **once** (e.g. 6, 8, …, 28) — no second block repeating 14, 16, …
- counts still show per size from stock; ordering is ascending by size number.

## 8. POS Core Operations

### `POS-01` POS main screens open

Steps:

1. Open POS app
2. Open major screens/windows from the toolbar or menus you normally use

Expected:

- no crash
- key windows open successfully

### `POS-02` POS customer autocomplete works

Steps:

1. Open POS billing screen
2. Type known customer text
3. Type branch-related text

Expected:

- matching customer suggestions appear
- warehouse/branch special targets appear where expected

### `POS-03` POS inventory screen opens according to permissions

Steps:

1. Open inventory window in default restricted mode
2. Then enable permission and open it again if needed

Expected:

- behavior matches permission setting

### `POS-04` POS admin/settings screen opens

Steps:

1. Open POS admin/settings window

Expected:

- window opens
- manager permissions tab is available

### `POS-05` POS export or print works if used

Steps:

1. Run your usual POS export/print action

Expected:

- export file is created or print dialog opens
- no crash

### `POS-06` Bill history shows bill type and filter

Steps:

1. Open POS **سجل الفواتير**.
2. Confirm a **نوع** (or equivalent) column shows بيع / مرتجع / استبدال as appropriate.
3. Use **نوع الفاتورة** filter: الكل, بيع, مرتجع, استبدال.

Expected:

- list filters correctly.
- no crash when switching filter.

## 9. POS Sales Flow

### `SALE-01` Normal sale reduces POS stock

Steps:

1. Sell one item in POS
2. Finalize bill

Expected:

- bill is created
- POS stock decreases
- bill appears in history

### `SALE-02` Sale sync reaches warehouse visibility

Steps:

1. Create sale in POS
2. Sync POS
3. Sync warehouse

Expected:

- warehouse receives sale-related state update as designed
- branch stock mirror reflects change after branch sync cycle

### `SALE-03` Multi-line sale works

Steps:

1. Add multiple different items
2. Finalize one bill

Expected:

- total is correct
- all lines are saved
- stock decreases correctly for each line

## 10. POS Return To Warehouse Flow

### `RET-01` POS can create return to warehouse

Steps:

1. In POS, create a test bill
2. Use `الى المصنع`
3. Finalize

Expected:

- operation completes
- bill is saved
- sync event is prepared

### `RET-02` Return reaches warehouse queue

Steps:

1. Sync POS
2. Sync warehouse
3. Open unhandled branch products window

Expected:

- returned line appears in queue
- it does not go directly into warehouse stock

### `RET-03` Warehouse assign-to-stock works

Steps:

1. Select queued returned item
2. Assign warehouse number and package number

Expected:

- queue item becomes processed
- item enters warehouse stock
- warehouse number/package number are saved

### `RET-04` Warehouse discard works

Steps:

1. Select queued returned item
2. Discard as defective

Expected:

- item is marked discarded
- item does not enter warehouse stock

### `RET-05` Warehouse reroute to branch works

Steps:

1. Select queued returned item
2. Reroute it to another branch
3. Sync warehouse
4. Sync target POS

Expected:

- item leaves queue as processed
- target POS receives it after warehouse sync

## 11. POS To POS Via Warehouse Flow

### `XFER-01` POS creates branch transfer request

Steps:

1. In POS, create a bill
2. Choose target branch
3. Finalize

Expected:

- POS accepts request
- event is created as request flow
- this is not treated as direct branch delivery

### `XFER-02` Request reaches warehouse queue

Steps:

1. Sync source POS
2. Sync warehouse
3. Open unhandled branch products window

Expected:

- transfer request appears in queue

### `XFER-03` Warehouse reroute completes delivery

Steps:

1. Process queued transfer by rerouting to target branch
2. Sync warehouse
3. Sync destination POS

Expected:

- destination POS receives stock
- source POS does not directly control final delivery

## 12. Cashier Lockdown

### `LOCK-01` Manual incoming stock is blocked by default

Steps:

1. Open POS
2. Try to use manual incoming stock flow

Expected:

- feature is hidden or blocked
- restriction message appears if accessed

### `LOCK-02` Inventory editing is blocked by default

Steps:

1. Open POS inventory area
2. Try to edit specs, delete stock, or adjust stock

Expected:

- actions are blocked by default

### `LOCK-03` Price editing is blocked by default

Steps:

1. Open inventory or pricing tools in POS
2. Try to change prices manually

Expected:

- price editing is blocked by default

### `LOCK-03B` Bulk price change is blocked by default

Steps:

1. Try to open or use bulk price update tools in POS

Expected:

- bulk price changes are blocked by default

### `LOCK-03C` Excel import is blocked by default

Steps:

1. Try to run Excel import or similar admin import in POS

Expected:

- import is blocked by default

### `LOCK-03D` Reset counts/history action is blocked by default

Steps:

1. Try to run reset counts or similar destructive admin cleanup in POS

Expected:

- action is blocked by default

### `LOCK-04` Finalized bill is not deletable

Steps:

1. Create sale
2. Open bills history
3. Try to remove it

Expected:

- finalized bill is not deletable by old destructive flow
- only valid reversal path is available

### `LOCK-05` Direct POS-to-POS transfer does not bypass warehouse

Steps:

1. Create transfer request from source POS to destination POS
2. Sync source POS only

Expected:

- destination POS does not receive it until warehouse processes it

## 13. Manager Permissions

### `MGR-01` Permissions tab opens

Steps:

1. Open POS admin/settings
2. Open `صلاحيات المدير`

Expected:

- tab loads
- permission checkboxes are visible

### `MGR-02` Wrong password cannot save permissions

Steps:

1. Change a permission
2. Enter wrong password
3. Save

Expected:

- save is rejected
- old setting remains

### `MGR-03` Correct password saves permissions

Steps:

1. Enable one restricted feature
2. Enter correct password
3. Save

Expected:

- save succeeds
- feature becomes available

### `MGR-04` Disabled feature becomes blocked again

Steps:

1. Disable same feature
2. Save with correct password
3. Try to use the feature

Expected:

- feature is blocked again

### `MGR-05` Permissions survive app restart

Steps:

1. Change one permission
2. Close POS
3. Reopen POS

Expected:

- saved permission remains

### `MGR-06` Each individual permission controls its matching feature

Steps:

1. Toggle one permission at a time
2. Test the related screen or action each time

Expected:

- only the intended feature changes behavior
- unrelated features do not change unexpectedly

### `MGR-07` Manager password change works

Steps:

1. Change the manager password
2. Close and reopen POS
3. Try old password
4. Try new password

Expected:

- old password fails
- new password succeeds

## 14. Void Flow

### `VOID-01` Manager can void a sale with reason

Steps:

1. Create a POS sale
2. Open bill history
3. Use `VOID`
4. Enter manager password
5. Enter reason

Expected:

- bill status changes to voided
- bill remains in history

### `VOID-02` Void without reason is rejected

Steps:

1. Try to void a bill
2. Leave reason empty

Expected:

- action is rejected

### `VOID-03` Void restores stock when applicable

Steps:

1. Create sale from real stock item
2. Void it
3. Recheck stock count

Expected:

- stock is restored for that item

## 15. Warehouse Number And Label Tests

### `WH-01` New warehouse numbers exist

Steps:

1. Open warehouse forms that use warehouse numbers

Expected:

- options `1` to `7` exist internally

### `WH-02` UI shows labels instead of raw numbers

Steps:

1. Open warehouse-number dropdowns

Expected:

- labels show:
  - `مخزن 1`
  - `مخزن 2`
  - `مخزن 3`
  - `مخزن 4`
  - `مخزن زايد`
  - `مخزن اكتوبر`
  - `مخزن العبور`

### `WH-03` Internal storage still works with labeled values

Steps:

1. Save stock using labeled warehouse selection
2. Search/filter that stock later

Expected:

- saved records behave normally
- filters and stock movement still work

### `WH-04` New labeled warehouses work in transfer/selection dialogs

Steps:

1. Open any warehouse dialog that uses warehouse selection
2. Select `مخزن زايد`, `مخزن اكتوبر`, or `مخزن العبور`

Expected:

- selection is accepted
- downstream save/transfer works

## 16. Warehouse Branch UI Tests

### `BR-01` Branch list includes approved POS names

Steps:

1. Open warehouse branch-related customer/target dropdowns

Expected:

- list includes all approved POS names

### `BR-02` Branch list has no duplicates

Steps:

1. Inspect branch dropdown list carefully

Expected:

- no duplicate branch entries
- no raw/incorrect duplicate naming

### `BR-03` Branch stock window filters work

Steps:

1. Open branch stock window
2. Use item type, school, color, size, and text filters

Expected:

- each filter narrows results correctly
- clear button resets filters

### `BR-04` Unhandled branch queue window opens and filters work

Steps:

1. Open unhandled branch products window
2. Use available filters/search if present

Expected:

- window opens
- rows can be filtered safely

### `BR-05` Branch bill log and sync queue log window

Steps:

1. Warehouse: **المزامنة** → **سجل شحنات الفروع والوارد المتزامن…** (or **الفواتير** → **شحنات الفروع والوارد المتزامن…**).
2. Open tab **شحنات صادرة (فواتير فرع)** and tab **وارد متزامن (طابور الفروع)**.

Expected:

- both tabs load without error.
- outbound branch shipments appear on first tab when they exist.
- queue rows (returns / transfers) appear on second tab when they exist.

### `BR-06` POS reservations mirror (warehouse)

Prerequisites: a POS has created or updated reservations and both sides have synced so warehouse inbox applied `RESERVATION_*` events.

Steps:

1. Warehouse **المزامنة** → **حجوزات الفروع (مرآة)…**.
2. Confirm **نقطة البيع** lists friendly branch names **and** any raw device identifiers that appear only in mirror data (UUID-style entries are normal).
3. Choose a POS by **device name** (e.g. `POS-ZAY`); press **تحديث** (or change filters — the list refreshes on selection / pending toggle).
4. Tab **تفصيلي (كل حجز)**: each row is one mirrored reservation line (multiple lines can share the same product; **الجهاز** shows a readable name when known).
5. Tab **مجمّع حسب المنتج** (opens by default): same **نوع / مدرسة / لون / مقاس** is one row with **إجمالي الكمية**, **عدد الحجوزات** (mirror line count), **عدد الفروع** (distinct devices when not filtered to one POS).
6. Optionally enable **المعلقة فقط** and confirm both tabs respect it.

Expected:

- filtering by POS **name** shows rows whose mirror `source_device` was stored as UUID or name (no “empty filter” bug when names and DB values differed).
- aggregate tab sums quantities across lines and across branches when no POS is selected.
- empty tables are acceptable if no reservation traffic yet.

### `BR-07` POS financial summary by day (warehouse)

Prerequisites: POS sales or reservation payments exist and warehouse has applied ledger events after sync.

Steps:

1. Warehouse **المزامنة** → **التدفقات المالية للفروع (يومي)…**.
2. Optionally set from/to dates and **الجهاز** filter (same expanded picklist as reservations: known names plus identifiers seen in mirror/ledger); refresh.
3. Select a day in the upper summary grid.

Expected:

- lower detail panel fills with categories (sales, returns, voids, exchange, reservation buckets) for that day.
- filtering by POS **name** matches rows stored under that device’s UUID (or name), same resolution rules as reservation mirror.
- numbers are plausible for the test activity you created.

## 17. Sync Detail And Status Tests

### `SD-01` Sync settings are saved after reopening app

Steps:

1. Save sync setup
2. Close app
3. Reopen sync setup

Expected:

- saved server URL remains
- saved device name remains

### `SD-02` Test connection can be rerun multiple times

Steps:

1. Run test connection several times in a row

Expected:

- no crash
- result remains stable

### `SD-03` Sync log is readable and ends cleanly

Steps:

1. Run a normal sync cycle
2. Read the log text shown in UI

Expected:

- progress is understandable
- finish state is visible

### `SD-04` Known device or branch-related sync data remains usable

Steps:

1. Sync devices normally
2. Reopen relevant dropdowns/windows that depend on known branch names

Expected:

- names still appear correctly
- no corruption in dropdown content

### `SD-05` Periodic background sync does not freeze the UI

Prerequisites: sync configured on warehouse and/or POS; app left running.

Steps:

1. Leave the app open past the automatic sync interval (on the order of tens of minutes).
2. Continue light UI use (open menus, switch screens).

Expected:

- UI stays responsive; any background sync work does not lock the app for long stretches.
- if a cycle fails, user-visible feedback appears (toast / dialog) rather than silent failure only.

### `SD-06` Manual sync shows summary and received-event details

Steps:

1. Run **مزامنة الآن** after there is something to push or pull.
2. If a post-sync summary appears, open **تفاصيل** (or equivalent) for received / applied / skipped events when offered.

Expected:

- summary text is readable.
- details window opens without crash; applied vs skipped vs error rows are distinguishable.

### `SD-07` Warehouse mirror / ledger idempotency on repeat sync

Prerequisites: `BR-06` / `BR-07` already show data for a POS after one warehouse sync.

Steps:

1. Run **مزامنة الآن** on warehouse twice in a row with no new POS business between runs.
2. Re-check **التدفقات المالية للفروع** counts for the same day (and reservation mirror if applicable).

Expected:

- no duplicate financial ledger lines for the same underlying event UUID.
- reservation mirror rows remain consistent (no runaway duplicates from replay).

## 18. Negative And Edge Cases

### `EDGE-01` Sync with no new work is safe

Steps:

1. Run sync when nothing changed

Expected:

- no crash
- summary stays sensible

### `EDGE-02` Repeated sync does not duplicate data

Steps:

1. Create one known event
2. Sync multiple times

Expected:

- no duplicated stock
- no duplicated queue entries

### `EDGE-03` Bad network/server URL shows useful failure

Steps:

1. Enter invalid server URL
2. Test connection or sync

Expected:

- user sees clear error

### `EDGE-04` Fresh sync after reset does not revive deleted local state unexpectedly

Steps:

1. Reset server for target device
2. Delete local DB
3. Sync fresh

Expected:

- old deleted local state does not reappear unless intentionally re-sent by another device

### `EDGE-05` Wrong or unavailable permission path fails safely

Steps:

1. Attempt restricted feature without permission

Expected:

- user gets safe rejection
- app does not crash

### `EDGE-06` Empty queue states are safe

Steps:

1. Open unhandled queue when there is nothing pending

Expected:

- empty state is safe
- no crash on refresh

### `EDGE-07` Empty branch stock mirror state is safe

Steps:

1. Open branch stock for branch with no synced stock yet

Expected:

- empty view is safe
- no crash

### `EDGE-08` Reopening windows after sync is safe

Steps:

1. Open a relevant stock or billing window
2. Run sync
3. Reopen or refresh the window

Expected:

- no stale-window crash
- data refresh behaves safely

## 19. Optional Feature Scenarios

Run these only if you actively use these features in your daily workflow.

### `OPT-01` Excel export from warehouse

Expected:

- file is created
- opened file has correct headers and rows

### `OPT-02` Excel export from POS

Expected:

- file is created
- opened file has correct headers and rows

### `OPT-03` Windows print integration

Expected:

- print dialog or printer workflow opens without crashing

### `OPT-04` Manual local backup/restore workflow

Expected:

- copying/restoring SQLite files works as expected when app is closed

## 20. Recommended Real Test Order

If you want the fastest useful real-world test pass, run in this order:

1. `ENV-01`
2. `AUTH-01`
3. `SYNC-01`
4. `SYNC-02`
5. `SALE-01`
6. `RET-01`
7. `RET-02`
8. `RET-03`
9. `XFER-01`
10. `XFER-02`
11. `XFER-03`
12. `LOCK-01`
13. `LOCK-03`
14. `MGR-03`
15. `VOID-01`
16. `BR-03`
17. `WH-02`
18. `MGR-07`
19. `RST-01`
20. `EDGE-02`

If you want full system coverage, also run:

1. `WHS-01` to `WHS-10`
2. `POS-01` to `POS-06`
3. `LOCK-01` to `LOCK-05`
4. `MGR-01` to `MGR-07`
5. `SD-01` to `SD-07`
6. `SYNC-01` to `SYNC-05`
7. `BR-01` to `BR-07`
8. `EDGE-01` to `EDGE-08`
9. `OPT-01` to `OPT-04` if relevant

### Optional quick pass for new reporting / sync UX

If you already ran the list above but want a focused regression on recently added areas:

- `SYNC-05`, `WHS-09`, `WHS-10`, `POS-06`, `BR-05`, `BR-06`, `BR-07`, `SD-05`, `SD-06`, `SD-07`

## 21. What To Send Me After Testing

Send results in compact form like this:

```text
ENV-01 PASS
AUTH-01 PASS
SYNC-02 FAIL - items appeared only after changing filter
SYNC-05 PASS
WHS-10 FAIL - price audit dialog crashed
BR-07 PASS
RET-02 PASS
XFER-03 FAIL - destination POS did not receive rerouted item
VOID-01 PASS
```

If something fails, also send:

- screenshot if possible
- exact device name
- whether DB was fresh or old
- whether you ran reset before test
- the last sync messages shown in UI
