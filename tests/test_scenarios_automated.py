"""
Automated coverage for scenarios documented in TEST SCENARIOS.md.

These tests exercise SqliteDatabase and helpers from POS-ZAY/HosnyPOS.py using a
temporary SQLite file (no GUI, no sync server). They map to checklist IDs where noted.

Run from repo root:
    python -m pytest tests/test_scenarios_automated.py -v
"""
from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
POS_ZAY_DIR = str(REPO / "POS-ZAY")
POS_ZAY = REPO / "POS-ZAY" / "HosnyPOS.py"

# HosnyPOS imports sync_core from the POS folder; tests run with cwd = repo root.
if POS_ZAY_DIR not in sys.path:
    sys.path.insert(0, POS_ZAY_DIR)


def _load_pos_zay_module():
    if not POS_ZAY.is_file():
        raise FileNotFoundError(f"Missing POS module: {POS_ZAY}")
    spec = importlib.util.spec_from_file_location("hosny_pos_zay_autotest", POS_ZAY)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load HosnyPOS spec")
    mod = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = mod
    spec.loader.exec_module(mod)
    return mod


_MOD = _load_pos_zay_module()
SqliteDatabase = _MOD.SqliteDatabase
WAREHOUSE_RETURN_LABEL = _MOD.WAREHOUSE_RETURN_LABEL
WAREHOUSE_RETURN_BILL_TYPE = _MOD.WAREHOUSE_RETURN_BILL_TYPE
_parse_money_amount = _MOD._parse_money_amount


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


class TestParseMoney(unittest.TestCase):
    def test_arabic_digits_and_comma(self):
        self.assertAlmostEqual(_parse_money_amount("٤٠٫٥٠"), 40.50, places=2)
        self.assertAlmostEqual(_parse_money_amount("40,5"), 40.5, places=2)


if __name__ == "__main__":
    unittest.main(verbosity=2)
