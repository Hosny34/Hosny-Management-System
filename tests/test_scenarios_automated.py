"""
Automated coverage for scenarios documented in TEST SCENARIOS.md.

These tests exercise SqliteDatabase and helpers from POS/HosnyPOS.py using a
temporary SQLite file (no GUI, no sync server). They map to checklist IDs where noted.

Run from repo root:
    python -m pytest tests/test_scenarios_automated.py -v
"""
from __future__ import annotations

import importlib.util
import json
import os
import sqlite3
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
POS_DIR = str(REPO / "POS")
POS_FILE = REPO / "POS" / "HosnyPOS.py"
POS_SYNC_APPLIERS_FILE = REPO / "POS" / "sync_appliers.py"
WAREHOUSE_DIR = str(REPO / "Warehouse")
WAREHOUSE_FILE = REPO / "Warehouse" / "HosnyWarehouse.py"
WAREHOUSE_SYNC_CORE_FILE = REPO / "Warehouse" / "sync_core.py"
WAREHOUSE_SYNC_APPLIERS_FILE = REPO / "Warehouse" / "sync_appliers.py"

# HosnyPOS imports sync_core from the POS folder; tests run with cwd = repo root.
if POS_DIR not in sys.path:
    sys.path.insert(0, POS_DIR)


def _load_pos_module():
    if not POS_FILE.is_file():
        raise FileNotFoundError(f"Missing POS module: {POS_FILE}")
    spec = importlib.util.spec_from_file_location("hosny_pos_autotest", POS_FILE)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load HosnyPOS spec")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


def _load_module(name: str, path: Path):
    if not path.is_file():
        raise FileNotFoundError(f"Missing module: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load module spec: {path}")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


_MOD = _load_pos_module()
sys.stdout = sys.__stdout__
sys.stderr = sys.__stderr__
SqliteDatabase = _MOD.SqliteDatabase
WAREHOUSE_RETURN_LABEL = _MOD.WAREHOUSE_RETURN_LABEL
WAREHOUSE_RETURN_BILL_TYPE = _MOD.WAREHOUSE_RETURN_BILL_TYPE
_parse_money_amount = _MOD._parse_money_amount


def _load_warehouse_module():
    if WAREHOUSE_DIR not in sys.path:
        sys.path.insert(0, WAREHOUSE_DIR)
    return _load_module("hosny_warehouse_autotest", WAREHOUSE_FILE)


def _db_path():
    fd, path = tempfile.mkstemp(prefix="pos_test_", suffix=".sqlite3")
    os.close(fd)
    return path


def _open_db(path: str) -> SqliteDatabase:
    db = SqliteDatabase(path=path, legacy_json=str(REPO / "nonexistent_legacy_xyz.json"))
    sid = db.start_shift()
    db.active_shift_id = sid
    return db


def _close_db(db: SqliteDatabase | None) -> None:
    if db is None:
        return
    try:
        db.conn.close()
    except Exception:
        pass


def _stock_sum(db: SqliteDatabase, item_type: str, school: str, color: str, size: str) -> int:
    cur = db.conn.execute(
        """
        SELECT COALESCE(SUM(count), 0) AS s FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
        """,
        (item_type, school, color, size),
    )
    row = cur.fetchone()
    return int(row["s"] if row and row["s"] is not None else 0)


class TestEnvAndSales(unittest.TestCase):
    """ENV-03 (DB init), SALE-01, SALE-03, WH return bill shape, stats."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        # LIFO: close DB before deleting the temp file.
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_ENV03_fresh_db_schema_usable(self):
        """ENV-03 (partial): new DB file initializes schema; core tables exist."""
        self._db = _open_db(self._path)
        cur = self._db.conn.execute(
            "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
        )
        names = {r["name"] for r in cur.fetchall()}
        for required in ("stocks", "bills", "bill_items", "movements", "shifts", "reservations"):
            self.assertIn(required, names, f"missing table {required}")

    def test_SALE01_normal_sale_reduces_stock(self):
        """SALE-01: finalize sale creates bill and reduces aggregate stock."""
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Tee", "IPS", "White", "M", 100.0, 5)
        self.assertEqual(_stock_sum(db, "Tee", "IPS", "White", "M"), 5)
        bid = db.create_bill(
            "Walk-in",
            [{"item_type": "Tee", "school": "IPS", "color": "White", "size": "M", "unit_price": 100.0, "qty": 2}],
        )
        self.assertGreater(bid, 0)
        self.assertEqual(_stock_sum(db, "Tee", "IPS", "White", "M"), 3)

    def test_SALE03_multiline_sale(self):
        """SALE-03: one bill with two specs; totals and both stock reductions."""
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("A", "S1", "C1", "1", 10.0, 3)
        db.add_stock("B", "S1", "C2", "2", 20.0, 4)
        bid = db.create_bill(
            "C",
            [
                {"item_type": "A", "school": "S1", "color": "C1", "size": "1", "unit_price": 10.0, "qty": 2},
                {"item_type": "B", "school": "S1", "color": "C2", "size": "2", "unit_price": 20.0, "qty": 1},
            ],
        )
        self.assertGreater(bid, 0)
        row = db.conn.execute("SELECT total FROM bills WHERE id=?", (bid,)).fetchone()
        self.assertAlmostEqual(float(row["total"]), 40.0, places=3)
        self.assertEqual(_stock_sum(db, "A", "S1", "C1", "1"), 1)
        self.assertEqual(_stock_sum(db, "B", "S1", "C2", "2"), 3)

    def test_admin_can_convert_wrong_exchange_return_bill_to_sale(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Wrong Mode Tee", "School X", "Red", "10", 100.0, 10)
        bid = db.create_exchange_bill(
            "Customer",
            [{"item_type": "Wrong Mode Tee", "school": "School X", "color": "Red", "size": "10", "unit_price": 100.0, "qty": 2}],
            [],
        )
        self.assertEqual(_stock_sum(db, "Wrong Mode Tee", "School X", "Red", "10"), 12)

        out = db.convert_exchange_bill_to_sale(bid, "wrong tab")
        self.assertEqual(out["corrected_return_qty"], 2)
        self.assertEqual(_stock_sum(db, "Wrong Mode Tee", "School X", "Red", "10"), 8)

        bill = next(b for b in db.list_bills() if int(b["id"]) == int(bid))
        self.assertEqual(bill["bill_type"], "SALE")
        self.assertAlmostEqual(float(bill["total"]), 200.0, places=2)
        items = db.list_bill_items(bid)
        self.assertEqual({it["origin"] for it in items}, {"STOCK"})
        moves = db.conn.execute(
            "SELECT direction, qty FROM movements WHERE bill_id=? ORDER BY id",
            (bid,),
        ).fetchall()
        self.assertIn(("OUT", 2), [(m["direction"], int(m["qty"])) for m in moves])
        event = db.conn.execute(
            "SELECT event_type, payload_json FROM sync_outbox WHERE event_type='SALE_BILL_TYPE_CORRECTED' ORDER BY local_seq DESC LIMIT 1"
        ).fetchone()
        self.assertIsNotNone(event)
        payload = json.loads(event["payload_json"])
        self.assertAlmostEqual(float(payload["amount_delta"]), 400.0, places=2)

    def test_RET01_warehouse_return_bill_recorded(self):
        """RET-01 (DB): bill to المصنع is saved as warehouse return with zero total."""
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("X", "S", "Col", "L", 50.0, 2)
        before = _stock_sum(db, "X", "S", "Col", "L")
        bid = db.create_bill(
            WAREHOUSE_RETURN_LABEL,
            [{"item_type": "X", "school": "S", "color": "Col", "size": "L", "unit_price": 50.0, "qty": 1}],
        )
        row = db.conn.execute(
            "SELECT total, bill_type FROM bills WHERE id=?", (bid,)
        ).fetchone()
        self.assertEqual(str(row["bill_type"]).upper(), WAREHOUSE_RETURN_BILL_TYPE)
        self.assertAlmostEqual(float(row["total"]), 0.0, places=3)
        self.assertEqual(_stock_sum(db, "X", "S", "Col", "L"), before - 1)

    def test_get_sales_stats_excludes_warehouse_customer(self):
        """Retail stats: warehouse return customer excluded from sales_total / count."""
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("P", "Sch", "Co", "9", 30.0, 10)
        db.create_bill("Retail", [{"item_type": "P", "school": "Sch", "color": "Co", "size": "9", "unit_price": 30.0, "qty": 1}])
        db.create_bill(
            WAREHOUSE_RETURN_LABEL,
            [{"item_type": "P", "school": "Sch", "color": "Co", "size": "9", "unit_price": 30.0, "qty": 1}],
        )
        from datetime import date

        today = date.today().isoformat()
        stats = db.get_sales_stats(date_from=today, date_to=today)
        self.assertEqual(stats["sales_count"], 1)
        self.assertAlmostEqual(stats["sales_total"], 30.0, places=2)


class TestVoidAndLock(unittest.TestCase):
    """VOID-01/02, LOCK-01 (income blocked)."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_VOID02_void_without_reason_rejected(self):
        """VOID-02: empty reason raises."""
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("V", "S", "C", "M", 15.0, 2)
        bid = db.create_bill(
            "Cust",
            [{"item_type": "V", "school": "S", "color": "C", "size": "M", "unit_price": 15.0, "qty": 1}],
        )
        with self.assertRaises(ValueError):
            db.void_bill(bid, "   ")

    def test_VOID01_void_restores_stock(self):
        """VOID-01 / VOID-03: void with reason restores STOCK-origin qty."""
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("V", "S", "C", "M", 15.0, 2)
        bid = db.create_bill(
            "Cust",
            [{"item_type": "V", "school": "S", "color": "C", "size": "M", "unit_price": 15.0, "qty": 1}],
        )
        self.assertEqual(_stock_sum(db, "V", "S", "C", "M"), 1)
        db.void_bill(bid, "test void")
        self.assertEqual(_stock_sum(db, "V", "S", "C", "M"), 2)
        st = db.conn.execute("SELECT status FROM bills WHERE id=?", (bid,)).fetchone()["status"]
        self.assertEqual(str(st).upper(), "VOID")

    def test_LOCK01_manual_incoming_blocked_under_lockdown(self):
        """LOCK-01: create_income_bill raises when cashier lockdown enabled."""
        self._db = _open_db(self._path)
        db = self._db
        with self.assertRaises((PermissionError,)):
            db.create_income_bill(
                "Sup",
                [{"item_type": "Z", "school": "S", "color": "C", "size": "1", "unit_price": 1.0, "qty": 1}],
            )


class TestPosStockAuditSync(unittest.TestCase):
    """POS stock audit corrections must reach the warehouse branch mirror."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_pos_audit_event_corrects_warehouse_mirror(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            UPDATE device_identity
               SET device_name = 'POS-TEST-AUDIT', device_role = 'pos'
             WHERE id = 1
            """
        )
        stock_id = db.add_stock("Audit Tee", "Audit School", "Navy", "10", 120.0, 5)
        row = {
            "stock_id": stock_id,
            "item_type": "Audit Tee",
            "school": "Audit School",
            "color": "Navy",
            "size": "10",
            "unit_price": 120.0,
            "expected": 5,
            "actual": 3,
            "diff": -2,
        }
        report_id = db.create_stock_audit_report([row], reason="manual")
        db.apply_stock_adjustments([row], note="POS stock audit manual equalization")
        db.record_pos_stock_audit_applied(report_id, [row], reason="manual")

        outbox = db.conn.execute(
            """
            SELECT event_uuid, event_type, payload_json, target_scope
              FROM sync_outbox
             WHERE event_type = 'POS_STOCK_AUDIT_APPLIED'
             ORDER BY local_seq DESC LIMIT 1
            """
        ).fetchone()
        self.assertIsNotNone(outbox)
        self.assertEqual(outbox["target_scope"], "warehouse")

        wh_core = _load_module("warehouse_sync_core_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.execute(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device,item_type,school,color,size,unit_price,count,snapshot_at)
                VALUES ('POS-TEST-AUDIT','Audit Tee','Audit School','Navy','10',120,5,'2026-07-28 10:00:00')
                """
            )
            wh.execute(
                """
                INSERT INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-TEST-AUDIT', '2026-07-28 10:00:00', 1, 600.0, '2026.07.28.1')
                """
            )
            wh.commit()
            payload = json.loads(outbox["payload_json"])
            result = wh_appliers.apply_pos_stock_audit_applied(wh, payload, outbox["event_uuid"])
            self.assertEqual(result["applied_rows"], 1)
            count = wh.execute(
                """
                SELECT COALESCE(SUM(count), 0)
                  FROM pos_stocks_mirror
                 WHERE source_device='POS-TEST-AUDIT'
                   AND item_type='Audit Tee'
                   AND school='Audit School'
                   AND color='Navy'
                   AND size='10'
                """
            ).fetchone()[0]
            self.assertEqual(int(count or 0), 3)
            meta = wh.execute(
                """
                SELECT snapshot_at, row_count, total_value
                  FROM pos_stocks_snapshot_meta
                 WHERE source_device='POS-TEST-AUDIT'
                """
            ).fetchone()
            self.assertEqual(meta[0], "2026-07-28 10:00:00")
            self.assertEqual(int(meta[1] or 0), 1)
            self.assertAlmostEqual(float(meta[2] or 0), 360.0, places=2)
            audit_rows = wh.execute("SELECT COUNT(*) FROM pos_stock_audit_items_mirror").fetchone()[0]
            self.assertEqual(int(audit_rows), 1)
        finally:
            wh.close()

    def test_repeated_audit_report_sync_uses_unique_audit_events(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            UPDATE device_identity
               SET device_name = 'POS-TEST-AUDIT', device_role = 'pos'
             WHERE id = 1
            """
        )
        stock_id = db.add_stock("Audit Tee", "Audit School", "Navy", "10", 315.0, 5)
        report_id = db.create_stock_audit_report([
            {
                "stock_id": stock_id,
                "item_type": "Audit Tee",
                "school": "Audit School",
                "color": "Navy",
                "size": "10",
                "unit_price": 315.0,
                "expected": 5,
                "actual": 4,
                "diff": -1,
            }
        ], reason="auto-equalization")
        first = {
            "stock_id": stock_id,
            "item_type": "Audit Tee",
            "school": "Audit School",
            "color": "Navy",
            "size": "10",
            "unit_price": 315.0,
            "expected": 5,
            "actual": 4,
            "diff": -1,
        }
        second = dict(first, expected=4, actual=5, diff=1)
        db.record_pos_stock_audit_applied(report_id, [first], reason="auto-equalization")
        db.record_pos_stock_audit_applied(report_id, [second], reason="auto-equalization")

        outbox_rows = db.conn.execute(
            """
            SELECT event_uuid, payload_json
              FROM sync_outbox
             WHERE event_type = 'POS_STOCK_AUDIT_APPLIED'
             ORDER BY local_seq ASC
            """
        ).fetchall()
        self.assertEqual(len(outbox_rows), 2)
        payloads = [json.loads(r["payload_json"]) for r in outbox_rows]
        self.assertNotEqual(payloads[0]["audit_uuid"], payloads[1]["audit_uuid"])
        self.assertTrue(str(payloads[0]["audit_uuid"]).startswith(f"POS-TEST-AUDIT:{report_id}:"))
        self.assertTrue(str(payloads[1]["audit_uuid"]).startswith(f"POS-TEST-AUDIT:{report_id}:"))

        wh_core = _load_module("warehouse_sync_core_repeat_audit_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_repeat_audit_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.execute(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device,item_type,school,color,size,unit_price,count,snapshot_at)
                VALUES ('POS-TEST-AUDIT','Audit Tee','Audit School','Navy','10',315,5,'2026-08-02 10:00:00')
                """
            )
            wh.execute(
                """
                INSERT INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-TEST-AUDIT', '2026-08-02 10:00:00', 1, 1575.0, '2026.08.02.5')
                """
            )
            wh.commit()
            for row, payload in zip(outbox_rows, payloads):
                wh_appliers.apply_pos_stock_audit_applied(wh, payload, row["event_uuid"])
            count = wh.execute(
                """
                SELECT COALESCE(SUM(count), 0)
                  FROM pos_stocks_mirror
                 WHERE source_device='POS-TEST-AUDIT'
                   AND item_type='Audit Tee'
                   AND school='Audit School'
                   AND color='Navy'
                   AND size='10'
                """
            ).fetchone()[0]
            self.assertEqual(int(count or 0), 5)
            total = wh.execute(
                """
                SELECT COALESCE(SUM(total_value), 0)
                  FROM pos_stock_audit_reports_mirror
                 WHERE source_device='POS-TEST-AUDIT'
                """
            ).fetchone()[0]
            self.assertAlmostEqual(float(total or 0), 0.0, places=2)
        finally:
            wh.close()

    def test_audit_snapshot_replaces_stale_partial_audit_mirror(self):
        wh_core = _load_module("warehouse_sync_core_audit_snapshot_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_audit_snapshot_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.execute(
                """
                INSERT INTO pos_stock_audit_reports_mirror
                    (audit_uuid, source_device, local_report_id, reason, created_at,
                     total_diff, total_value, event_uuid, received_at)
                VALUES ('POS-TEST-AUDIT:1', 'POS-TEST-AUDIT', 1, 'auto-equalization',
                        '2026-08-02T12:52:10', -1, -315, 'old-event', '2026-08-02T12:52:30')
                """
            )
            wh.execute(
                """
                INSERT INTO pos_stock_audit_items_mirror
                    (audit_uuid, source_device, item_type, school, color, size,
                     expected_qty, actual_qty, diff_qty, unit_price, diff_value)
                VALUES ('POS-TEST-AUDIT:1', 'POS-TEST-AUDIT', 'Audit Tee', 'Audit School',
                        'Navy', '12', 5, 4, -1, 315, -315)
                """
            )
            payload = {
                "source_device_name": "POS-TEST-AUDIT",
                "snapshot_at": "2026-08-02T14:30:00",
                "reports": [
                    {
                        "report_id": 1,
                        "created_at": "2026-08-02T12:52:15",
                        "reason": "auto-equalization",
                        "diff_count": 2,
                        "total_diff": 0,
                        "total_value": 0.0,
                        "lines": [
                            {
                                "item_type": "Audit Tee",
                                "school": "Audit School",
                                "color": "Navy",
                                "size": "12",
                                "expected": 5,
                                "actual": 4,
                                "diff": -1,
                                "unit_price": 315.0,
                                "diff_value": -315.0,
                            },
                            {
                                "item_type": "Audit Tee",
                                "school": "Audit School",
                                "color": "Navy",
                                "size": "12",
                                "expected": 4,
                                "actual": 5,
                                "diff": 1,
                                "unit_price": 315.0,
                                "diff_value": 315.0,
                            },
                        ],
                    }
                ],
            }
            result = wh_appliers.apply_pos_stock_audit_snapshot(wh, payload, "snapshot-event")
            self.assertEqual(result["reports"], 1)
            self.assertEqual(result["lines"], 2)
            row = wh.execute(
                """
                SELECT COUNT(*) AS c, COALESCE(SUM(total_diff), 0) AS qty,
                       COALESCE(SUM(total_value), 0) AS value
                  FROM pos_stock_audit_reports_mirror
                 WHERE source_device='POS-TEST-AUDIT'
                """
            ).fetchone()
            self.assertEqual(int(row[0] or 0), 1)
            self.assertEqual(int(row[1] or 0), 0)
            self.assertAlmostEqual(float(row[2] or 0), 0.0, places=2)
        finally:
            wh.close()


class TestPosStockSnapshotSync(unittest.TestCase):
    """Warehouse must compare POS stock snapshot timestamps by instant."""

    def test_utc_snapshot_after_local_snapshot_is_not_skipped(self):
        wh_core = _load_module("warehouse_sync_core_snapshot_time_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_snapshot_time_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.execute(
                """
                INSERT INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-OCT', '2026-08-02T12:52:10', 1, 1260.0, '2026.8.2.2')
                """
            )
            wh.execute(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device,item_type,school,color,size,unit_price,count,snapshot_at)
                VALUES ('POS-OCT','Audit Tee','Audit School','Navy','12',315,4,'2026-08-02T12:52:10')
                """
            )
            result = wh_appliers.apply_pos_stock_snapshot(
                wh,
                {
                    "source_device_name": "POS-OCT",
                    "snapshot_at": "2026-08-02T11:35:24.885389Z",
                    "app_version": "2026.8.2.10",
                    "rows": [
                        {
                            "item_type": "Audit Tee",
                            "school": "Audit School",
                            "color": "Navy",
                            "size": "12",
                            "unit_price": 315.0,
                            "count": 5,
                        }
                    ],
                },
                "snapshot-utc-newer",
            )
            self.assertFalse(result.get("skipped"))
            meta = wh.execute(
                "SELECT total_value, app_version FROM pos_stocks_snapshot_meta WHERE source_device='POS-OCT'"
            ).fetchone()
            self.assertAlmostEqual(float(meta[0] or 0), 1575.0, places=2)
            self.assertEqual(meta[1], "2026.8.2.10")
        finally:
            wh.close()


class TestSpecRenameSync(unittest.TestCase):
    """Spec rename events must repair leftover branch aliases generically."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_value_rename_updates_leftover_pos_stock_school(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("تيشيرت صيفي", "الالمانية", "احمر ف كحلي", "4", 260.0, 6)
        appliers = _load_module("pos_sync_appliers_value_rename_autotest", POS_SYNC_APPLIERS_FILE)
        payload = {
            "old_spec": {
                "item_type": "تيشيرت صيفي",
                "school": "الالمانية KG",
                "color": "احمر ف كحلي",
                "size": "4",
            },
            "new_spec": {
                "item_type": "تيشيرت صيفي",
                "school": "الالمانية KG",
                "color": "احمر ف كحلي",
                "size": "4",
            },
            "changed_fields": [],
            "value_renames": [
                {"field": "school", "old_value": "الالمانية", "new_value": "الالمانية KG"},
            ],
        }
        result = appliers.apply_spec_renamed(db.conn, payload, "event-value-rename")
        self.assertGreaterEqual(int(result["updated_rows"]), 1)
        old_count = _stock_sum(db, "تيشيرت صيفي", "الالمانية", "احمر ف كحلي", "4")
        new_count = _stock_sum(db, "تيشيرت صيفي", "الالمانية KG", "احمر ف كحلي", "4")
        self.assertEqual(old_count, 0)
        self.assertEqual(new_count, 6)


class TestWarehouseBranchStockViews(unittest.TestCase):
    """Branch monitor and cycle summary must use corrected POS mirror stock."""

    def setUp(self):
        self._path = _db_path()
        self._db = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(self._close)

    def _close(self):
        if self._db is not None:
            try:
                self._db.conn.close()
            except Exception:
                pass

    def test_monitor_and_cycle_use_corrected_pos_stock_mirror(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', '2026-07-28T10:00:00Z', '2026-07-28T10:00:00Z')
                """
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-ZAY', '2026-07-28T10:05:00Z', 1, 360.0, '2026.07.29.012056')
                """
            )
            db.conn.execute(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device, item_type, school, color, size, unit_price, count, snapshot_at)
                VALUES ('POS-ZAY', 'Audit Tee', 'Audit School', 'Navy', '10', 120.0, 3, '2026-07-28T10:05:00Z')
                """
            )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor("2026-07-28", "2026-07-28")}
        self.assertIn("POS-ZAY", monitor)
        self.assertEqual(int(monitor["POS-ZAY"]["stock_qty"]), 3)
        self.assertAlmostEqual(float(monitor["POS-ZAY"]["stock_value"]), 360.0, places=2)
        self.assertEqual(monitor["POS-ZAY"]["app_version"], "2026.07.29.012056")

        cycle = {r["branch_device"]: r for r in db.list_branch_cycle_reconciliation("2026-07-28", "2026-07-28")}
        self.assertIn("POS-ZAY", cycle)
        self.assertEqual(int(cycle["POS-ZAY"]["stock_qty"]), 3)
        self.assertAlmostEqual(float(cycle["POS-ZAY"]["stock_value"]), 360.0, places=2)

    def test_monitor_last_sync_uses_device_seen_time_when_snapshot_unchanged(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', '2026-07-30T18:30:00.000000Z', '2026-07-30T18:30:05.000000Z')
                """
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-ZAY', '2026-07-30T17:02:54.000000Z', 1, 100.0, '2026.07.30')
                """
            )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor("2026-07-30", "2026-07-30")}
        self.assertEqual(monitor["POS-ZAY"]["last_sync_at"], "2026-07-30T18:30:00.000000Z")
        self.assertEqual(monitor["POS-ZAY"]["snapshot_at"], "2026-07-30T17:02:54.000000Z")

    def test_monitor_app_version_uses_latest_sync_payload_not_stale_snapshot(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', '2026-08-02T11:20:00Z', '2026-08-02T11:20:05Z')
                """
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-ZAY', '2026-08-02T10:00:00Z', 1, 100.0, '2026.8.2.2')
                """
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO sync_inbox
                    (event_uuid, event_type, server_seq, source_device, payload_json,
                     applied_at, apply_status, apply_attempts, apply_at)
                VALUES ('version-event', 'POS_STOCK_AUDIT_SNAPSHOT', 50, 'POS-ZAY', ?,
                        '2026-08-02T11:21:00Z', 'ok', 1, '2026-08-02T11:21:01Z')
                """,
                (json.dumps({
                    "source_device_name": "POS-ZAY",
                    "app_version": "2026.8.2.5",
                    "reports": [],
                }, ensure_ascii=False),),
            )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor("2026-08-02", "2026-08-02")}
        self.assertEqual(monitor["POS-ZAY"]["app_version"], "2026.8.2.5")

    def test_monitor_app_version_uses_latest_apply_time_when_server_seq_resets(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', '2026-08-02T11:40:00Z', '2026-08-02T11:40:05Z')
                """
            )
            for event_uuid, server_seq, apply_at, version in (
                ("new-version-low-seq", 10, "2026-08-02T11:35:00Z", "2026.8.2.10"),
                ("old-version-high-seq", 170, "2026-08-02T10:37:00Z", "2026.08.02.3"),
            ):
                db.conn.execute(
                    """
                    INSERT OR REPLACE INTO sync_inbox
                        (event_uuid, event_type, server_seq, source_device, payload_json,
                         applied_at, apply_status, apply_attempts, apply_at)
                    VALUES (?, 'POS_STOCK_SNAPSHOT', ?, 'POS-ZAY', ?,
                            ?, 'ok', 1, ?)
                    """,
                    (
                        event_uuid,
                        server_seq,
                        json.dumps({"source_device_name": "POS-ZAY", "app_version": version, "rows": []}),
                        apply_at,
                        apply_at,
                    ),
                )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor("2026-08-02", "2026-08-02")}
        self.assertEqual(monitor["POS-ZAY"]["app_version"], "2026.8.2.10")

    def test_monitor_zero_net_audit_does_not_warn(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', '2026-08-02T11:40:00Z', '2026-08-02T11:40:05Z')
                """
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-ZAY', '2026-08-02T11:35:00Z', 1, 100.0, '2026.8.2.10')
                """
            )
            db.conn.execute(
                """
                INSERT INTO pos_stock_audit_reports_mirror
                    (audit_uuid, source_device, local_report_id, reason, created_at,
                     total_diff, total_value, event_uuid, received_at)
                VALUES ('POS-ZAY:1', 'POS-ZAY', 1, 'auto-equalization',
                        '2026-08-02T12:52:15', 0, 0, 'audit-snapshot', '2026-08-02T11:36:00Z')
                """
            )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor("2026-08-02", "2026-08-02")}
        self.assertEqual(monitor["POS-ZAY"]["audit_adjust_qty"], 0)
        self.assertAlmostEqual(float(monitor["POS-ZAY"]["audit_adjust_value"]), 0.0, places=2)
        self.assertEqual(monitor["POS-ZAY"]["status"], "جيد")


class TestWarehousePriceProfiles(unittest.TestCase):
    """Warehouse price-profile reports."""

    def setUp(self):
        self._path = _db_path()
        self._db = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(self._close)

    def _close(self):
        if self._db is not None:
            try:
                self._db.conn.close()
            except Exception:
                pass

    def test_missing_price_profile_report_lists_only_unassigned_stock_groups(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        db.add_or_update_stock_row("Assigned Item", "School A", "Blue", "10", 1, 1, 100.0, 3)
        db.add_or_update_stock_row("Missing Item", "School B", "Red", "12", 1, 1, 120.0, 4)
        profile_id = db.create_price_profile("Default Profile")
        db.assign_price_profile("Assigned Item", "School A", "Blue", profile_id)

        rows = db.list_items_missing_price_profile()
        keys = {(r["item_type"], r["school"], r["color"]) for r in rows}
        self.assertNotIn(("Assigned Item", "School A", "Blue"), keys)
        self.assertIn(("Missing Item", "School B", "Red"), keys)
        missing = next(r for r in rows if (r["item_type"], r["school"], r["color"]) == ("Missing Item", "School B", "Red"))
        self.assertEqual(int(missing["total_qty"]), 4)
        self.assertAlmostEqual(float(missing["total_value"]), 480.0, places=2)

    def test_price_profile_updates_all_group_sizes_and_catalog_rows(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        db.add_stock("Profile Tee", "School P", "Navy", "8", 1, 1, 111.0, 2)
        db.add_stock("Profile Tee", "School P", "Navy", "10", 1, 1, 222.0, 2)
        profile_id = db.create_price_profile("Profile Prices")
        db.replace_price_profile_item_prices(
            profile_id,
            "Profile Tee",
            [
                {"size": "8", "price": 180.0},
                {"size": "10", "price": 200.0},
                {"size": "12", "price": 220.0},
            ],
        )
        db.assign_price_profile("Profile Tee", "School P", "Navy", profile_id)
        result = db.apply_price_profile_to_stock(profile_id, [("Profile Tee", "School P", "Navy")])
        self.assertEqual(int(result["updated"]), 2)

        prices = {
            str(r["size"]): float(r["unit_price"])
            for r in db.current_inventory({"item_type": "Profile Tee", "school": "School P", "color": "Navy", "hide_zero": False})
        }
        self.assertAlmostEqual(prices["8"], 180.0, places=2)
        self.assertAlmostEqual(prices["10"], 200.0, places=2)
        self.assertAlmostEqual(float(db.get_effective_price("Profile Tee", "School P", "Navy", "12")), 220.0, places=2)

        catalog = db.price_profile_catalog_rows_for_targets(profile_id, [("Profile Tee", "School P", "Navy")])
        by_size = {row["size"]: float(row["unit_price"]) for row in catalog}
        self.assertEqual(by_size, {"8": 180.0, "10": 200.0, "12": 220.0})

    def test_size_profile_update_queues_catalog_upsert_for_all_pos(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', '', '')
                """
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-OBO', 'uuid-obo', 'pos', '', '')
                """
            )
        db.upsert_size_profile(
            "Profile Sync Tee",
            "School Sync",
            "Gray",
            r1_start=2,
            r1_end=10,
            r2_start=12,
            r2_end=16,
            has_alpha=True,
        )
        sent = db.send_size_profile_to_all_pos("Profile Sync Tee", "School Sync", "Gray")
        self.assertGreaterEqual(sent, 2)

        rows = db.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='CATALOG_UPSERT'
            ORDER BY local_seq
            """
        ).fetchall()
        self.assertEqual(len(rows), sent)
        targets = {r["target_scope"] for r in rows}
        self.assertIn("pos:POS-ZAY", targets)
        self.assertIn("pos:POS-OBO", targets)
        for row in rows:
            payload = json.loads(row["payload_json"])
            profile = payload["size_profiles"][0]
            self.assertEqual(profile["item_type"], "Profile Sync Tee")
            self.assertEqual(profile["school"], "School Sync")
            self.assertEqual(profile["color"], "Gray")
            self.assertEqual(profile["num_start_1"], 2)
            self.assertEqual(profile["num_end_1"], 10)
            self.assertEqual(profile["num_start_2"], 12)
            self.assertEqual(profile["num_end_2"], 16)
            self.assertEqual(profile["has_alpha"], 1)


class TestReservationDeliver(unittest.TestCase):
    """Reservation delivery cash in stats (related checklist: POS dashboard / shift)."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_deliver_reservation_inserts_deliver_pay(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("R", "S", "C", "L", 100.0, 1)
        ids = db.create_reservation(
            "RsvCust",
            [{"item_type": "R", "school": "S", "color": "C", "size": "L", "unit_price": 100.0, "qty": 1}],
            paid_amount=0.0,
        )
        rid = int(ids[0])
        db.deliver_reservation(rid, 40.0)
        row = db.conn.execute(
            "SELECT COUNT(*) AS n, COALESCE(SUM(unit_price),0) AS t FROM movements WHERE direction='DELIVER_PAY'"
        ).fetchone()
        self.assertGreaterEqual(int(row["n"]), 1)
        self.assertGreaterEqual(float(row["t"]), 40.0)
        from datetime import date

        today = date.today().isoformat()
        stats = db.get_sales_stats(date_from=today, date_to=today)
        self.assertGreaterEqual(stats.get("deliver_cash", 0), 40.0)

    def test_reservation_can_be_created_for_available_stock_without_deducting_until_delivery(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Reserved Tee", "School R", "Green", "12", 90.0, 4)
        before = _stock_sum(db, "Reserved Tee", "School R", "Green", "12")
        ids = db.create_reservation(
            "Reservation Customer",
            [{"item_type": "Reserved Tee", "school": "School R", "color": "Green", "size": "12", "unit_price": 90.0, "qty": 2}],
            paid_amount=50.0,
        )
        self.assertEqual(len(ids), 1)
        self.assertEqual(_stock_sum(db, "Reserved Tee", "School R", "Green", "12"), before)
        db.deliver_reservation(int(ids[0]), collected_amount=130.0)
        self.assertEqual(_stock_sum(db, "Reserved Tee", "School R", "Green", "12"), before - 2)

    def test_reservation_down_payment_is_distributed_with_5_rounding(self):
        self._db = _open_db(self._path)
        db = self._db
        for size in ("8", "10", "12"):
            db.add_stock("Rounding Tee", "School Round", "Black", size, 100.0, 1)
        ids = db.create_reservation(
            "Reservation Customer",
            [
                {"item_type": "Rounding Tee", "school": "School Round", "color": "Black", "size": "8", "unit_price": 100.0, "qty": 1},
                {"item_type": "Rounding Tee", "school": "School Round", "color": "Black", "size": "10", "unit_price": 100.0, "qty": 1},
                {"item_type": "Rounding Tee", "school": "School Round", "color": "Black", "size": "12", "unit_price": 100.0, "qty": 1},
            ],
            paid_amount=100.0,
        )
        rows = db.conn.execute(
            "SELECT id, paid_amount FROM reservations WHERE id IN (%s) ORDER BY id"
            % ",".join("?" * len(ids)),
            tuple(int(x) for x in ids),
        ).fetchall()
        self.assertEqual([float(r["paid_amount"]) for r in rows], [35.0, 35.0, 30.0])

        moves = db.conn.execute(
            "SELECT unit_price FROM movements WHERE direction='RESERVE_PAY' ORDER BY id"
        ).fetchall()
        self.assertEqual([float(r["unit_price"]) for r in moves], [35.0, 35.0, 30.0])

        event = db.conn.execute(
            "SELECT payload_json FROM sync_outbox WHERE event_type='RESERVATION_CREATED' ORDER BY local_seq DESC LIMIT 1"
        ).fetchone()
        payload = json.loads(event["payload_json"])
        self.assertEqual([float(r["paid_amount"]) for r in payload["reservations"]], [35.0, 35.0, 30.0])

    def test_legacy_reservation_down_payment_is_repaired_once(self):
        self._db = _open_db(self._path)
        db = self._db
        for size in ("8", "10", "12"):
            db.add_stock("Legacy Reserved Tee", "School Legacy", "Blue", size, 100.0, 1)
        ids = [int(x) for x in db.create_reservation(
            "Legacy Reservation Customer",
            [
                {"item_type": "Legacy Reserved Tee", "school": "School Legacy", "color": "Blue", "size": "8", "unit_price": 100.0, "qty": 1},
                {"item_type": "Legacy Reserved Tee", "school": "School Legacy", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1},
                {"item_type": "Legacy Reserved Tee", "school": "School Legacy", "color": "Blue", "size": "12", "unit_price": 100.0, "qty": 1},
            ],
            paid_amount=100.0,
        )]

        placeholders = ",".join("?" * len(ids))
        db.conn.execute(
            f"UPDATE reservations SET paid_amount=CASE WHEN id=? THEN 100 ELSE 0 END WHERE id IN ({placeholders})",
            (ids[0], *ids),
        )
        db.conn.execute("DELETE FROM movements WHERE direction='RESERVE_PAY'")
        first = db.conn.execute("SELECT * FROM reservations WHERE id=?", (ids[0],)).fetchone()
        db.conn.execute(
            """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                first["created_at"],
                "RESERVE_PAY",
                None,
                0,
                f"عربون حجز #{ids[0]}",
                None,
                first["item_type"],
                first["school"],
                first["color"],
                first["size"],
                100.0,
                first["shift_id"],
                first["payment_method"],
            ),
        )
        db.conn.commit()

        result = db._repair_legacy_reservation_down_payment_allocations()
        self.assertEqual(result["groups"], 1)

        rows = db.conn.execute(
            f"SELECT paid_amount FROM reservations WHERE id IN ({placeholders}) ORDER BY id",
            tuple(ids),
        ).fetchall()
        self.assertEqual([float(r["paid_amount"]) for r in rows], [35.0, 35.0, 30.0])

        moves = db.conn.execute(
            "SELECT note, unit_price FROM movements WHERE direction='RESERVE_PAY' ORDER BY id"
        ).fetchall()
        self.assertEqual([float(r["unit_price"]) for r in moves], [35.0, 35.0, 30.0])
        self.assertEqual([r["note"] for r in moves], [f"عربون حجز #{rid}" for rid in ids])

        second_result = db._repair_legacy_reservation_down_payment_allocations()
        self.assertEqual(second_result["groups"], 0)

    def test_partial_reservation_delivery_collects_full_selected_items(self):
        self._db = _open_db(self._path)
        db = self._db
        sizes = ("2", "4", "6", "8", "10")
        for size in sizes:
            db.add_stock("Partial Reserved Tee", "School Partial", "Red", size, 100.0, 2)
        ids = [int(x) for x in db.create_reservation(
            "Partial Customer",
            [
                {"item_type": "Partial Reserved Tee", "school": "School Partial", "color": "Red", "size": size, "unit_price": 100.0, "qty": 1}
                for size in sizes
            ],
            paid_amount=100.0,
        )]

        with self.assertRaisesRegex(ValueError, "300"):
            db.deliver_reservation_items(ids[:3], collected_amount=240.0)

        summary = db.deliver_reservation_items(ids[:3], collected_amount=300.0)
        self.assertEqual(summary["delivered_items"], 3)
        self.assertFalse(summary["group_completed"])

        placeholders = ",".join("?" * len(ids))
        rows = db.conn.execute(
            f"SELECT id, status, paid_amount FROM reservations WHERE id IN ({placeholders}) ORDER BY id",
            tuple(ids),
        ).fetchall()
        self.assertEqual([str(r["status"]) for r in rows[:3]], ["تم التسليم", "تم التسليم", "تم التسليم"])
        self.assertEqual([float(r["paid_amount"]) for r in rows[:3]], [100.0, 100.0, 100.0])
        self.assertEqual(sum(float(r["paid_amount"]) for r in rows[3:]), 100.0)

        paid_now = db.conn.execute(
            "SELECT COALESCE(SUM(unit_price),0) FROM movements WHERE direction='DELIVER_PAY'"
        ).fetchone()[0]
        self.assertEqual(float(paid_now), 300.0)

    def test_non_sale_bill_payment_methods_affect_shift_cash_and_visa(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Pay Tee", "School Pay", "Blue", "10", 100.0, 10)
        db.add_stock("Pay Hoodie", "School Pay", "Black", "12", 150.0, 3)

        db.create_bill(
            "Cash customer",
            [{"item_type": "Pay Tee", "school": "School Pay", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 2}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        db.create_return_bill(
            "Visa return",
            [{"item_type": "Pay Tee", "school": "School Pay", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )
        db.create_exchange_bill(
            "Visa exchange",
            [{"item_type": "Pay Tee", "school": "School Pay", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            [{"item_type": "Pay Hoodie", "school": "School Pay", "color": "Black", "size": "12", "unit_price": 150.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )
        ids = db.create_reservation(
            "Visa reservation",
            [{"item_type": "Pay Tee", "school": "School Pay", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            paid_amount=30.0,
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )
        db.deliver_reservation(int(ids[0]), 70.0, payment_method=_MOD.PAYMENT_METHOD_VISA)

        summary = db.get_shift_summary(db.active_shift_id)
        self.assertAlmostEqual(float(summary["cash_collected"]), 200.0, places=2)
        self.assertAlmostEqual(float(summary["visa_collected"]), 50.0, places=2)

    def test_shift_summary_recovers_missing_reservation_payment_movements(self):
        self._db = _open_db(self._path)
        db = self._db
        db.create_reservation(
            "Cash reservation",
            [{"item_type": "R", "school": "S", "color": "C", "size": "1", "unit_price": 370.0, "qty": 1}],
            paid_amount=370.0,
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        db.create_reservation(
            "Visa reservation",
            [{"item_type": "R", "school": "S", "color": "C", "size": "2", "unit_price": 480.0, "qty": 1}],
            paid_amount=480.0,
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )
        db.conn.execute("DELETE FROM movements WHERE direction='RESERVE_PAY'")
        db.conn.commit()

        summary = db.get_shift_summary(db.active_shift_id)
        self.assertAlmostEqual(float(summary["cash_collected"]), 370.0, places=2)
        self.assertAlmostEqual(float(summary["visa_collected"]), 480.0, places=2)
        self.assertAlmostEqual(float(summary["res_paid"]), 850.0, places=2)

    def test_closed_shift_list_uses_saved_summary_when_later_rows_change(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("T", "S", "C", "1", 100.0, 1)
        db.create_bill(
            "Customer",
            [{"item_type": "T", "school": "S", "color": "C", "size": "1", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        db.create_reservation(
            "Reservation",
            [{"item_type": "R", "school": "S", "color": "C", "size": "1", "unit_price": 400.0, "qty": 1}],
            paid_amount=400.0,
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        shift_id = int(db.active_shift_id)
        saved = db.get_shift_summary(shift_id)
        db.end_shift(shift_id, summary_json=json.dumps(saved, ensure_ascii=False, default=str))
        db.conn.execute("DELETE FROM movements WHERE direction='RESERVE_PAY'")
        db.conn.execute("UPDATE reservations SET paid_amount=0")
        db.conn.commit()

        detail = db.get_shift_summary(shift_id)
        listed = next(r for r in db.get_all_shifts() if int(r["id"]) == shift_id)
        self.assertAlmostEqual(float(detail["cash_collected"]), 500.0, places=2)
        self.assertAlmostEqual(float(listed["cash_collected"]), 500.0, places=2)
        self.assertAlmostEqual(float(listed["res_paid"]), 400.0, places=2)


class TestPosSizeGrid(unittest.TestCase):
    """Sales size grid must include real stock sizes outside configured ranges."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_size_profile_merges_inventory_sizes_outside_range(self):
        self._db = _open_db(self._path)
        db = self._db
        db.upsert_size_profile(
            "Range Tee",
            "School Size",
            "Black",
            r1_start=0,
            r1_end=10,
            r2_start=None,
            r2_end=None,
            has_alpha=False,
        )
        db.add_stock("Range Tee", "School Size", "Black", "12", 100.0, 3)
        pos_shell = object.__new__(_MOD.POSFrame)
        pos_shell.db = db
        sizes = pos_shell._get_sizes_for_bill("School Size", "Range Tee", "Black")
        self.assertIn("0", sizes)
        self.assertIn("10", sizes)
        self.assertIn("12", sizes)

    def test_inventory_show_zero_uses_size_profile_missing_rows(self):
        self._db = _open_db(self._path)
        db = self._db
        db.upsert_size_profile(
            "Zero Tee",
            "School Zero",
            "Gray",
            r1_start=None,
            r1_end=None,
            r2_start=None,
            r2_end=None,
            has_alpha=True,
        )
        db.add_stock("Zero Tee", "School Zero", "Gray", "M", 250.0, 1)
        db.add_stock("Zero Tee", "School Zero", "Gray", "XL", 250.0, 5)

        inv = object.__new__(_MOD.InventoryWindow)
        inv.db = db
        inv.show_zero_var = type("V", (), {"get": lambda self: True})()
        inv._filters = lambda: {
            "item_type": "Zero Tee",
            "school": "School Zero",
            "color": "Gray",
            "size": None,
            "hide_zero": False,
        }
        rows = inv._with_profile_zero_rows(db.current_inventory(inv._filters()))
        by_size = {r["size"]: int(r["count"]) for r in rows}
        self.assertEqual(by_size["M"], 1)
        self.assertEqual(by_size["XL"], 5)
        for size in ("S", "L", "2XL", "3XL", "4XL", "5XL"):
            self.assertIn(size, by_size)
            self.assertEqual(by_size[size], 0)

    def test_special_bill_finalize_returns_to_first_empty_sales_bill(self):
        pos_shell = object.__new__(_MOD.POSFrame)
        pos_shell.bills = {1: [], 2: [{"qty": 1}], 3: [], 4: [], 5: [], 6: []}
        self.assertEqual(pos_shell._sales_bill_for_next_entry(), 1)

        pos_shell.bills[1] = [{"qty": 1}]
        self.assertEqual(pos_shell._sales_bill_for_next_entry(), 3)

        pos_shell.bills[3] = [{"qty": 1}]
        self.assertEqual(pos_shell._sales_bill_for_next_entry(), 1)


class TestMovementSizeFilters(unittest.TestCase):
    def setUp(self):
        self._path = _db_path()
        self._db = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_pos_movement_log_filters_by_size(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Move Tee", "School M", "Navy", "10", 100.0, 2)
        db.add_stock("Move Tee", "School M", "Navy", "12", 100.0, 2)

        rows = db.list_movements({"item_type": "Move Tee", "school": "School M", "color": "Navy", "size": "10"})
        self.assertTrue(rows)
        self.assertEqual({str(r.get("size")) for r in rows}, {"10"})

    def test_warehouse_movement_log_filters_by_size(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        db.add_stock("Move Tee", "School M", "Navy", "10", 1, 1, 100.0, 2)
        db.add_stock("Move Tee", "School M", "Navy", "12", 1, 1, 100.0, 2)

        rows = db.list_movements({"item_type": "Move Tee", "school": "School M", "color": "Navy", "size": "10"})
        self.assertTrue(rows)
        self.assertEqual({str(r.get("size")) for r in rows}, {"10"})


class TestItemTypeOrdering(unittest.TestCase):
    def setUp(self):
        self._path = _db_path()
        self._db = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_pos_item_type_filters_use_preferred_order(self):
        self._db = _open_db(self._path)
        db = self._db
        for item in ("بنطلون", "تيشيرت شتوي", "تيشيرت صيفي", "شروال", "تيشيرت رياضي"):
            db.add_stock(item, "School O", "Black", "10", 100.0, 1)

        self.assertEqual(
            db.get_distinct_filtered("item_type", {}),
            ["تيشيرت صيفي", "تيشيرت شتوي", "شروال", "تيشيرت رياضي", "بنطلون"],
        )
        self.assertEqual(
            [item for item, _color in db.list_items_for_school("School O")],
            ["تيشيرت صيفي", "تيشيرت شتوي", "شروال", "تيشيرت رياضي", "بنطلون"],
        )

    def test_warehouse_item_type_filters_use_preferred_order(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        for item in ("بنطلون", "تيشيرت شتوي", "تيشيرت صيفي", "شروال", "تيشيرت رياضي"):
            db.add_stock(item, "School O", "Black", "10", 1, 1, 100.0, 1)

        expected = ["تيشيرت صيفي", "تيشيرت شتوي", "شروال", "تيشيرت رياضي", "بنطلون"]
        self.assertEqual(db.get_distinct_filtered("item_type", {}), expected)
        self.assertEqual(db.get_distinct("item_type"), expected)

        branch_shell = object.__new__(wh_mod.BranchStockWindow)
        branch_shell._all_rows = [(item, "School O", "Black", "10", 100.0, 1) for item in reversed(expected)]
        self.assertEqual(branch_shell._branch_distinct("item_type", {}), expected)

    def test_unknown_item_types_sort_after_preferred_items(self):
        wh_mod = _load_warehouse_module()
        hoodie = "\u0647\u0648\u062f\u064a"
        shirt = "\u062a\u064a\u0634\u064a\u0631\u062a \u0635\u064a\u0641\u064a"
        trouser = "\u0628\u0646\u0637\u0644\u0648\u0646"
        self.assertEqual(_MOD.sort_item_type_values([hoodie, trouser, shirt]), [shirt, trouser, hoodie])
        self.assertEqual(wh_mod.sort_warehouse_item_type_values([hoodie, trouser, shirt]), [shirt, trouser, hoodie])


class TestSchoolAccountsCashFlow(unittest.TestCase):
    """School accounts must track all cash movement, not sales only."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_school_accounts_include_reservation_return_void_exchange(self):
        self._db = _open_db(self._path)
        db = self._db
        school = "School Cash"
        db.add_stock("T", school, "Blue", "10", 100.0, 10)
        sale_id = db.create_bill(
            "Customer",
            [{"item_type": "T", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 2}],
        )
        db.void_bill(sale_id, "test")
        db.create_return_bill(
            "Customer",
            [{"item_type": "T", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
        )
        db.add_stock("P", school, "Black", "12", 150.0, 2)
        db.create_exchange_bill(
            "Customer",
            [{"item_type": "T", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            [{"item_type": "P", "school": school, "color": "Black", "size": "12", "unit_price": 150.0, "qty": 1}],
        )
        ids = db.create_reservation(
            "Customer",
            [{"item_type": "T", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            paid_amount=30.0,
        )
        db.deliver_reservation(int(ids[0]), 70.0)
        rows = db.get_school_accounts_report([school], date_from="2000-01-01", date_to="2099-12-31")
        self.assertTrue(rows)
        net = sum(float(r.get("net_total") or 0.0) for r in rows)
        cash_in = sum(float(r.get("cash_in") or 0.0) for r in rows)
        cash_out = sum(float(r.get("cash_out") or 0.0) for r in rows)
        self.assertAlmostEqual(cash_in, 350.0, places=2)
        self.assertAlmostEqual(cash_out, 300.0, places=2)
        self.assertAlmostEqual(net, 50.0, places=2)


class TestParseMoney(unittest.TestCase):
    def test_arabic_digits_and_comma(self):
        self.assertAlmostEqual(_parse_money_amount("٤٠٫٥٠"), 40.50, places=2)
        self.assertAlmostEqual(_parse_money_amount("40,5"), 40.5, places=2)

    def test_size_range_label_parser_tolerates_arrow_encoding(self):
        self.assertEqual(_MOD.parse_numeric_range_label("0 → 24"), (0, 24))
        self.assertEqual(_MOD.parse_numeric_range_label("0 â†’ 24"), (0, 24))
        wh_mod = _load_warehouse_module()
        self.assertEqual(wh_mod.parse_numeric_range_label("6 → 22"), (6, 22))
        self.assertEqual(wh_mod.parse_numeric_range_label("6 â†’ 22"), (6, 22))


    def test_warehouse_canceled_bill_status_excluded_from_totals(self):
        wh_mod = _load_warehouse_module()
        self.assertTrue(wh_mod.is_canceled_bill_status("VOID"))
        self.assertTrue(wh_mod.is_canceled_bill_status("ملغاة"))
        self.assertFalse(wh_mod.is_canceled_bill_status("CONFIRMED"))


class TestApplicationVersioning(unittest.TestCase):
    def test_app_versions_are_numeric_and_aligned(self):
        wh_mod = _load_warehouse_module()
        with (REPO / "sync_server" / "Hosny-sync-server" / "updates" / "pos" / "latest.json").open(
            "r", encoding="utf-8"
        ) as f:
            manifest = json.load(f)
        with (REPO / "POS" / "current_version.json").open("r", encoding="utf-8") as f:
            current = json.load(f)

        versions = {
            "pos_app": _MOD.APP_VERSION,
            "warehouse_app": wh_mod.APP_VERSION,
            "pos_current": current.get("version"),
            "pos_manifest": manifest.get("version"),
        }
        for name, version in versions.items():
            self.assertRegex(str(version), r"^\d+(?:\.\d+)+$", name)
        self.assertEqual(len(set(versions.values())), 1, versions)
        self.assertEqual(manifest.get("package_file"), "HosnyPOS-%s.zip" % _MOD.APP_VERSION)


if __name__ == "__main__":
    unittest.main(verbosity=2)
