"""
Automated coverage for scenarios documented in TEST SCENARIOS.md.

These tests exercise SqliteDatabase and helpers from POS/HosnyPOS.py using a
temporary SQLite file (no GUI, no sync server). They map to checklist IDs where noted.

Run from repo root:
    python -m pytest tests/test_scenarios_automated.py -v
"""
from __future__ import annotations

import importlib.util
import hashlib
import json
import os
import sqlite3
import sys
import tempfile
import types
import unittest
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
POS_DIR = str(REPO / "POS")
POS_FILE = REPO / "POS" / "HosnyPOS.py"
POS_SYNC_APPLIERS_FILE = REPO / "POS" / "sync_appliers.py"
POS_SYNC_CLIENT_FILE = REPO / "POS" / "sync_client.py"
WAREHOUSE_DIR = str(REPO / "Warehouse")
WAREHOUSE_FILE = REPO / "Warehouse" / "HosnyWarehouse.py"
WAREHOUSE_SYNC_CORE_FILE = REPO / "Warehouse" / "sync_core.py"
WAREHOUSE_SYNC_APPLIERS_FILE = REPO / "Warehouse" / "sync_appliers.py"
SYNC_SERVER_DIR = str(REPO / "sync_server" / "Hosny-sync-server")
SYNC_SERVER_AUTH_FILE = REPO / "sync_server" / "Hosny-sync-server" / "auth.py"

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
now_iso = _MOD.now_iso


def _load_warehouse_module():
    if WAREHOUSE_DIR not in sys.path:
        sys.path.insert(0, WAREHOUSE_DIR)
    return _load_module("hosny_warehouse_autotest", WAREHOUSE_FILE)


def _load_sync_server_auth_module():
    if SYNC_SERVER_DIR not in sys.path:
        sys.path.insert(0, SYNC_SERVER_DIR)
    sys.modules.setdefault("jwt", types.SimpleNamespace())
    return _load_module("hosny_sync_server_auth_autotest", SYNC_SERVER_AUTH_FILE)


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


class TestTestingBranchConfiguration(unittest.TestCase):
    def test_pos_test_is_configured_as_normal_pos_branch(self):
        wh_mod = _load_warehouse_module()
        auth_mod = _load_sync_server_auth_module()
        ui_name = "\u0641\u0631\u0639 \u0627\u0644\u0627\u062e\u062a\u0628\u0627\u0631"

        self.assertIn("POS-TEST", wh_mod.DEFAULT_BRANCH_POS_NAMES)
        self.assertEqual(wh_mod.branch_display_name("POS-TEST"), ui_name)
        self.assertEqual(wh_mod.canonical_branch_device_name(ui_name), "POS-TEST")
        self.assertEqual(wh_mod.configured_branch_device_name("POS-TEST"), "POS-TEST")

        self.assertIn("POS-TEST", _MOD.DEFAULT_BRANCH_POS_NAMES)
        self.assertEqual(_MOD._branch_display_name("POS-TEST"), ui_name)
        self.assertEqual(_MOD._extract_branch_target(f"{_MOD.BRANCH_TARGET_PREFIX}{ui_name}"), "POS-TEST")

        self.assertEqual(auth_mod.validate_simple_device_name("pos-test"), "POS-TEST")
        self.assertEqual(auth_mod.infer_role_from_device_name("POS-TEST"), "pos")
        self.assertEqual(
            auth_mod.allowed_scopes_for_pull("POS-TEST", "pos"),
            ["pos:POS-TEST", "all-pos", "all"],
        )

    def test_pos_test_branch_shipment_targets_test_scope(self):
        wh_mod = _load_warehouse_module()
        path = _db_path()
        self.addCleanup(lambda p=path: os.path.isfile(p) and os.remove(p))
        db = wh_mod.SqliteDatabase(path=path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self.addCleanup(lambda: db.conn.close())

        bid = db.create_bill(
            wh_mod.branch_customer_label("POS-TEST"),
            [
                {
                    "item_type": "Testing Tee",
                    "school": "Testing School",
                    "color": "Gray",
                    "size": "12",
                    "unit_price": 275.0,
                    "qty": 4,
                    "allow_factory_fill": True,
                }
            ],
            target_pos="POS-TEST",
        )

        self.assertGreater(bid, 0)
        event = db.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='STOCK_TRANSFER_OUT'
             ORDER BY local_seq DESC
             LIMIT 1
            """
        ).fetchone()
        self.assertIsNotNone(event)
        self.assertEqual(event["target_scope"], "pos:POS-TEST")
        payload = json.loads(event["payload_json"])
        self.assertTrue(payload["shipment_uuid"])
        self.assertEqual(payload["items"][0]["qty"], 4)

    def test_branch_shipment_receipt_state_and_resend(self):
        wh_mod = _load_warehouse_module()
        path = _db_path()
        self.addCleanup(lambda p=path: os.path.isfile(p) and os.remove(p))
        db = wh_mod.SqliteDatabase(path=path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self.addCleanup(lambda: db.conn.close())

        bid = db.create_bill(
            wh_mod.branch_customer_label("POS-CEN"),
            [
                {
                    "item_type": "Center Tee",
                    "school": "Center School",
                    "color": "Red",
                    "size": "14",
                    "unit_price": 275.0,
                    "qty": 2,
                    "allow_factory_fill": True,
                }
            ],
            target_pos="POS-CEN",
        )
        bill = next(b for b in db.list_bills() if int(b["id"]) == int(bid))
        self.assertEqual(db.branch_shipment_receipt_state(bill), "بانتظار الإرسال")

        shipment_uuid = str(bill["uuid"])
        db.conn.execute(
            "UPDATE sync_outbox SET status='acked' WHERE payload_json LIKE ?",
            (f"%{shipment_uuid}%",),
        )
        self.assertEqual(db.branch_shipment_receipt_state(bill), "مرسل - لم يؤكد")

        event_uuid = db.resend_branch_shipment_bill(bid)
        event = db.conn.execute(
            """
            SELECT event_type, target_scope, status, payload_json
              FROM sync_outbox
             WHERE event_uuid=?
            """,
            (event_uuid,),
        ).fetchone()
        self.assertEqual(event["event_type"], "STOCK_TRANSFER_OUT")
        self.assertEqual(event["target_scope"], "pos:POS-CEN")
        self.assertEqual(event["status"], "pending")
        payload = json.loads(event["payload_json"])
        self.assertEqual(payload["shipment_uuid"], shipment_uuid)
        self.assertEqual(payload["items"][0]["qty"], 2)
        self.assertEqual(db.branch_shipment_receipt_state(bill), "بانتظار الإرسال")

        db.conn.execute(
            """
            INSERT INTO shipment_receipt_reviews(
                sync_event_uuid, shipment_uuid, source_device, payload_json,
                has_diff, note, created_at
            ) VALUES(?,?,?,?,?,?,?)
            """,
            ("receipt-event", shipment_uuid, "POS-CEN", "[]", 0, "", wh_mod.now_iso()),
        )
        self.assertEqual(db.branch_shipment_receipt_state(bill), "استلمه الفرع")

    def test_rerouted_branch_shipment_uses_bill_uuid_as_shipment_uuid(self):
        wh_mod = _load_warehouse_module()
        path = _db_path()
        self.addCleanup(lambda p=path: os.path.isfile(p) and os.remove(p))
        db = wh_mod.SqliteDatabase(path=path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self.addCleanup(lambda: db.conn.close())
        db.conn.execute(
            """
            INSERT INTO branch_inventory_queue(
                sync_event_uuid, queue_kind, source_device, requested_target_device,
                external_ref, line_index, created_at, item_type, school, color,
                size, unit_price, qty, has_badge, note, status
            ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                "source-return-event",
                "RETURN",
                "POS-OBO",
                "WAREHOUSE",
                "source-shipment",
                0,
                wh_mod.now_iso(),
                "Reroute Tee",
                "Reroute School",
                "Red",
                "14",
                400.0,
                2,
                0,
                "wrong branch",
                "PENDING",
            ),
        )

        bill_id = db.reroute_branch_inventory_queue_items([1], "POS-OCT")
        bill = next(b for b in db.list_bills() if int(b["id"]) == int(bill_id))
        event = db.conn.execute(
            """
            SELECT target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='STOCK_TRANSFER_OUT'
             ORDER BY local_seq DESC
             LIMIT 1
            """
        ).fetchone()
        payload = json.loads(event["payload_json"])

        self.assertEqual(event["target_scope"], "pos:POS-OCT")
        self.assertEqual(payload["shipment_uuid"], bill["uuid"])
        self.assertIn(f"bill #{bill_id}", payload["note"])
        self.assertEqual(db.branch_shipment_receipt_state(bill), "بانتظار الإرسال")


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

    def test_bill_history_includes_sale_and_reservation_notes(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Note Tee", "Note School", "Black", "10", 100.0, 3)
        bill_id = db.create_bill(
            "Note Customer",
            [{"item_type": "Note Tee", "school": "Note School", "color": "Black", "size": "10", "unit_price": 100.0, "qty": 1}],
            note="sale note visible",
        )
        reservation_ids = db.create_reservation(
            "Reservation Note Customer",
            [{"item_type": "Note Tee", "school": "Note School", "color": "Black", "size": "12", "unit_price": 120.0, "qty": 1, "note": "reservation note visible"}],
            paid_amount=120.0,
        )

        history = db.list_bill_history()
        sale = next(row for row in history if row.get("history_key") == f"bill:{bill_id}")
        reservation = next(row for row in history if row.get("history_key", "").startswith("reservation:") and int(row.get("id") or 0) == int(reservation_ids[0]))
        self.assertEqual(sale.get("note"), "sale note visible")
        self.assertEqual(reservation.get("note"), "reservation note visible")

    def test_pos_audit_log_records_sale_with_hash_chain(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Audit Tee", "Audit School", "Navy", "M", 125.0, 4)
        bid = db.create_bill(
            "Audit Customer",
            [{"item_type": "Audit Tee", "school": "Audit School", "color": "Navy", "size": "M", "unit_price": 125.0, "qty": 2}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
            customer_phone="01000000000",
        )
        self.assertGreater(bid, 0)

        audit_dir = Path(db.audit.log_dir)
        files = sorted(audit_dir.glob("pos-audit-*.jsonl"))
        self.assertTrue(files, "expected daily POS audit log file")
        entries = []
        for file_path in files:
            with file_path.open("r", encoding="utf-8") as f:
                entries.extend(json.loads(line) for line in f if line.strip())

        events = [entry["event"] for entry in entries]
        self.assertIn("shift_started", events)
        self.assertIn("sale_created", events)
        sale = next(entry for entry in entries if entry["event"] == "sale_created")
        self.assertEqual(sale["details"]["bill_id"], bid)
        self.assertEqual(sale["details"]["total"], 250.0)
        self.assertEqual(sale["details"]["qty_total"], 2)

        previous_hash = ""
        for entry in entries:
            expected_prev = previous_hash
            entry_hash = entry["entry_hash"]
            payload = dict(entry)
            payload.pop("entry_hash")
            canonical = json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
            self.assertEqual(hashlib.sha256(canonical.encode("utf-8")).hexdigest(), entry_hash)
            self.assertEqual(entry["prev_hash"], expected_prev)
            previous_hash = entry_hash

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

    def test_partial_reservation_delivery_collects_selected_remaining_only(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Jacket", "Public", "Green", "4", 240.0, 1)
        db.add_stock("Jacket", "Public", "Green", "8", 240.0, 1)
        reservation_ids = db.create_reservation(
            "Customer",
            [
                {"item_type": "Jacket", "school": "Public", "color": "Green", "size": "4", "unit_price": 240.0, "qty": 1},
                {"item_type": "Jacket", "school": "Public", "color": "Green", "size": "8", "unit_price": 240.0, "qty": 1},
            ],
            paid_amount=120.0,
        )

        with self.assertRaisesRegex(ValueError, "180"):
            db.deliver_reservation_items([reservation_ids[0]], collected_amount=240.0)

        summary = db.deliver_reservation_items([reservation_ids[0]], collected_amount=180.0)
        self.assertFalse(summary["group_completed"])
        delivered = db.conn.execute("SELECT status, paid_amount FROM reservations WHERE id=?", (reservation_ids[0],)).fetchone()
        pending = db.conn.execute("SELECT status, paid_amount FROM reservations WHERE id=?", (reservation_ids[1],)).fetchone()
        self.assertEqual(delivered["status"], "تم التسليم")
        self.assertAlmostEqual(float(delivered["paid_amount"]), 240.0, places=2)
        self.assertEqual(pending["status"], "معلق")
        self.assertAlmostEqual(float(pending["paid_amount"]), 60.0, places=2)
        movement = db.conn.execute(
            "SELECT unit_price FROM movements WHERE direction='DELIVER_PAY' ORDER BY id DESC LIMIT 1"
        ).fetchone()
        self.assertIsNotNone(movement)
        self.assertAlmostEqual(float(movement["unit_price"]), 180.0, places=2)

    def test_branch_shipment_direct_line_records_in_then_out(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        bid = db.create_bill(
            "فرع: POS-OCT",
            [
                {
                    "item_type": "Direct Tee",
                    "school": "Direct School",
                    "color": "Gray",
                    "size": "20",
                    "unit_price": 275.0,
                    "qty": 13,
                    "allow_factory_fill": True,
                }
            ],
            target_pos="POS-OCT",
        )
        self.assertGreater(bid, 0)
        self.assertEqual(_stock_sum(db, "Direct Tee", "Direct School", "Gray", "20"), 0)
        moves = db.conn.execute(
            """
            SELECT direction, qty, note
              FROM movements
             WHERE bill_id=?
             ORDER BY id
            """,
            (bid,),
        ).fetchall()
        self.assertEqual(
            [(m["direction"], int(m["qty"])) for m in moves],
            [("IN", 13), ("OUT", 13)],
        )
        self.assertIn("Direct branch shipment intake", moves[0]["note"])
        items = db.list_bill_items(bid)
        self.assertEqual([item["origin"] for item in items], ["STOCK"])
        event = db.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='STOCK_TRANSFER_OUT'
             ORDER BY local_seq DESC
             LIMIT 1
            """
        ).fetchone()
        self.assertIsNotNone(event)
        self.assertEqual(event["target_scope"], "pos:POS-OCT")
        payload = json.loads(event["payload_json"])
        self.assertEqual(payload["items"][0]["qty"], 13)

    def test_void_branch_shipment_emits_targeted_cancel_event(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        bid = db.create_bill(
            "فرع: POS-OCT",
            [
                {
                    "item_type": "Cancel Tee",
                    "school": "Cancel School",
                    "color": "Gray",
                    "size": "20",
                    "unit_price": 275.0,
                    "qty": 4,
                    "allow_factory_fill": True,
                }
            ],
            target_pos="POS-OCT",
        )
        bill_uuid = db.conn.execute("SELECT uuid FROM bills WHERE id=?", (bid,)).fetchone()[0]
        db.void_bill(bid)

        event = db.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             ORDER BY local_seq DESC
             LIMIT 1
            """
        ).fetchone()
        self.assertEqual(event["event_type"], "STOCK_TRANSFER_CANCELLED")
        self.assertEqual(event["target_scope"], "pos:POS-OCT")
        payload = json.loads(event["payload_json"])
        self.assertEqual(payload["shipment_uuid"], bill_uuid)
        self.assertEqual(payload["items"][0]["qty"], 4)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM sync_outbox WHERE event_type='SALE_VOIDED'").fetchone()[0],
            0,
        )

    def test_branch_stock_reclassification_emits_targeted_event_and_updates_mirror(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            db.conn.execute(
                """
                INSERT INTO pos_stocks_mirror(source_device, item_type, school, color, size, unit_price, count, snapshot_at)
                VALUES ('POS-OCT', 'Summer Tee', 'Sky KG', 'Lemon', '2', 290, 7, '2026-08-09T12:00:00Z')
                """
            )
            db.conn.execute(
                """
                INSERT INTO pos_stocks_snapshot_meta(source_device, snapshot_at, row_count, total_value)
                VALUES ('POS-OCT', '2026-08-09T12:00:00Z', 1, 2030)
                """
            )
        event_uuid = db.record_branch_stock_reclassification_event(
            "POS-OCT",
            "POS-OCT",
            {"item_type": "Summer Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 290, "count": 7},
            {"item_type": "Winter Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 310},
            4,
            "wrong season",
        )
        self.assertTrue(event_uuid)
        event = db.conn.execute(
            "SELECT event_type, target_scope, payload_json FROM sync_outbox ORDER BY local_seq DESC LIMIT 1"
        ).fetchone()
        self.assertEqual(event["event_type"], "BRANCH_STOCK_RECLASSIFIED")
        self.assertEqual(event["target_scope"], "pos:POS-OCT")
        payload = json.loads(event["payload_json"])
        self.assertEqual(payload["qty"], 4)
        self.assertEqual(payload["to_spec"]["item_type"], "Winter Tee")
        self.assertEqual(
            db.conn.execute("SELECT count FROM pos_stocks_mirror WHERE item_type='Summer Tee'").fetchone()[0],
            3,
        )
        self.assertEqual(
            db.conn.execute("SELECT count FROM pos_stocks_mirror WHERE item_type='Winter Tee'").fetchone()[0],
            4,
        )

    def test_branch_catalog_delete_emits_targeted_event(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db

        event_uuid = db.record_branch_catalog_delete_event(
            "POS-OCT",
            [
                {"item_type": "Delete Tee", "school": "Delete School", "color": "Red", "size": "10"},
                {"school": "Delete School"},
            ],
            note="delete wrong definition",
        )

        self.assertTrue(event_uuid)
        event = db.conn.execute(
            "SELECT event_type, target_scope, payload_json FROM sync_outbox ORDER BY local_seq DESC LIMIT 1"
        ).fetchone()
        self.assertEqual(event["event_type"], "BRANCH_CATALOG_DELETED")
        self.assertEqual(event["target_scope"], "pos:POS-OCT")
        payload = json.loads(event["payload_json"])
        self.assertEqual(payload["branch_device"], "POS-OCT")
        self.assertEqual(payload["filters"][0]["item_type"], "Delete Tee")
        self.assertEqual(payload["filters"][1], {"school": "Delete School"})

    def test_warehouse_audit_can_add_to_closed_package(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        db.ensure_package_open(7, 12)
        db.close_package(7, 12)

        count = db.apply_stock_adjustments([
            {
                "item_type": "Audit Add Tee",
                "school": "Audit School",
                "color": "Navy",
                "size": "20",
                "warehouse_no": 7,
                "package_no": 12,
                "unit_price": 275.0,
                "expected": 0,
                "actual": 5,
            }
        ])

        self.assertEqual(count, 1)
        self.assertEqual(_stock_sum(db, "Audit Add Tee", "Audit School", "Navy", "20"), 5)
        self.assertEqual(db.package_status(7, 12), "CLOSED")
        move = db.conn.execute(
            """
            SELECT direction, qty, warehouse_no, package_no
              FROM movements
             WHERE item_type='Audit Add Tee'
             ORDER BY id DESC
             LIMIT 1
            """
        ).fetchone()
        self.assertEqual(move["direction"], "ADJUST_IN")
        self.assertEqual(int(move["qty"]), 5)
        self.assertEqual(int(move["warehouse_no"]), 7)
        self.assertEqual(int(move["package_no"]), 12)

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
    """POS stock audits are local-only and must not load the warehouse sync."""

    def setUp(self):
        self._path = _db_path()
        self._db: SqliteDatabase | None = None
        self.addCleanup(lambda p=self._path: os.path.isfile(p) and os.remove(p))
        self.addCleanup(lambda: _close_db(self._db))

    def test_pos_audit_adjusts_local_stock_without_sync_event(self):
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
        event_uuid = db.record_pos_stock_audit_applied(report_id, [row], reason="manual")
        self.assertIsNone(event_uuid)

        outbox = db.conn.execute(
            """
            SELECT event_uuid, event_type, payload_json, target_scope
              FROM sync_outbox
             WHERE event_type = 'POS_STOCK_AUDIT_APPLIED'
             ORDER BY local_seq DESC LIMIT 1
            """
        ).fetchone()
        self.assertIsNone(outbox)
        count = db.conn.execute(
            """
            SELECT COALESCE(SUM(count), 0)
              FROM stocks
             WHERE item_type='Audit Tee'
               AND school='Audit School'
               AND color='Navy'
               AND size='10'
            """
        ).fetchone()[0]
        self.assertEqual(int(count or 0), 3)

    def test_repeated_audit_reports_remain_local_only(self):
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
        self.assertIsNone(db.record_pos_stock_audit_applied(report_id, [first], reason="auto-equalization"))
        self.assertIsNone(db.record_pos_stock_audit_applied(report_id, [second], reason="auto-equalization"))

        outbox_rows = db.conn.execute(
            """
            SELECT event_uuid, payload_json
              FROM sync_outbox
             WHERE event_type = 'POS_STOCK_AUDIT_APPLIED'
             ORDER BY local_seq ASC
            """
        ).fetchall()
        self.assertEqual(outbox_rows, [])
        self.assertIsNotNone(db.get_stock_audit_report(report_id)[0])

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

    def test_audit_snapshot_replaces_existing_uuid_even_if_source_name_changed(self):
        wh_core = _load_module("warehouse_sync_core_audit_snapshot_uuid_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_audit_snapshot_uuid_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.execute(
                """
                INSERT INTO pos_stock_audit_reports_mirror
                    (audit_uuid, source_device, local_report_id, reason, created_at,
                     total_diff, total_value, event_uuid, received_at)
                VALUES ('POS-OCT:1', 'old POS OCT name', 1, 'auto-equalization',
                        '2026-08-02T12:52:15', -1, -315, 'old-event', '2026-08-02T12:52:30')
                """
            )
            payload = {
                "source_device_name": "POS-OCT",
                "snapshot_at": "2026-08-02T14:30:00",
                "reports": [
                    {
                        "report_id": 1,
                        "created_at": "2026-08-02T12:52:15",
                        "reason": "auto-equalization",
                        "total_diff": 0,
                        "total_value": 0.0,
                        "lines": [],
                    }
                ],
            }
            result = wh_appliers.apply_pos_stock_audit_snapshot(wh, payload, "snapshot-event")
            self.assertEqual(result["reports"], 1)
            row = wh.execute(
                """
                SELECT COUNT(*) AS c, source_device, total_diff, total_value
                  FROM pos_stock_audit_reports_mirror
                 WHERE audit_uuid='POS-OCT:1'
                """
            ).fetchone()
            self.assertEqual(int(row[0] or 0), 1)
            self.assertEqual(row[1], "POS-OCT")
            self.assertEqual(int(row[2] or 0), 0)
            self.assertAlmostEqual(float(row[3] or 0), 0.0, places=2)
        finally:
            wh.close()


class TestPosStockSnapshotSync(unittest.TestCase):
    """Warehouse must compare POS stock snapshot timestamps by instant."""

    def test_unchanged_pending_stock_snapshot_is_reused(self):
        self._path = _db_path()
        self._db = None
        try:
            self._db = _open_db(self._path)
            db = self._db
            db.add_stock("Timeout Tee", "Center School", "Red", "14", 275.0, 3)
            sync_client = _load_module("pos_sync_client_snapshot_reuse_autotest", POS_SYNC_CLIENT_FILE)
            client = sync_client.SyncClient(db.conn)
            cfg = {"device_role": "pos", "device_name": "POS-CEN"}

            first_uuid = client.emit_stock_snapshot_event(cfg)
            second_uuid = client.emit_stock_snapshot_event(cfg)

            self.assertTrue(first_uuid)
            self.assertEqual(second_uuid, first_uuid)
            pending_rows = db.conn.execute(
                """
                SELECT COUNT(*) AS c
                  FROM sync_outbox
                 WHERE event_type='POS_STOCK_SNAPSHOT'
                   AND status='pending'
                """
            ).fetchone()["c"]
            self.assertEqual(int(pending_rows or 0), 1)

            db.conn.execute(
                "UPDATE sync_outbox SET status='acked' WHERE event_uuid=?",
                (first_uuid,),
            )
            third_uuid = client.emit_stock_snapshot_event(cfg)
            self.assertIsNone(third_uuid)
            total_rows = db.conn.execute(
                "SELECT COUNT(*) AS c FROM sync_outbox WHERE event_type='POS_STOCK_SNAPSHOT'"
            ).fetchone()["c"]
            self.assertEqual(int(total_rows or 0), 1)
        finally:
            _close_db(self._db)
            try:
                os.remove(self._path)
            except Exception:
                pass

    def test_duplicate_snapshot_rows_are_collapsed_before_mirror_insert(self):
        wh_core = _load_module("warehouse_sync_core_snapshot_dupes_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_snapshot_dupes_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            result = wh_appliers.apply_pos_stock_snapshot(
                wh,
                {
                    "source_device_name": "POS-ZAY",
                    "snapshot_at": "2026-08-24T11:50:00Z",
                    "rows": [
                        {
                            "item_type": "Summer Tee",
                            "school": "Raja",
                            "color": "Red",
                            "size": "10",
                            "unit_price": 295.0,
                            "count": 2,
                        },
                        {
                            "item_type": "Summer Tee",
                            "school": "Raja",
                            "color": "Red",
                            "size": "10",
                            "unit_price": 295,
                            "count": 3,
                        },
                    ],
                },
                "snapshot-dupe-rows",
            )
            self.assertEqual(result["mirrored_rows"], 1)
            row = wh.execute(
                """
                SELECT COUNT(*) AS rows, COALESCE(SUM(count), 0) AS total_count
                  FROM pos_stocks_mirror
                 WHERE source_device='POS-ZAY'
                """
            ).fetchone()
            self.assertEqual(int(row[0] or 0), 1)
            self.assertEqual(int(row[1] or 0), 5)
        finally:
            wh.close()

    def test_existing_duplicate_mirror_rows_are_repaired_before_snapshot_index(self):
        wh_core = _load_module("warehouse_sync_core_snapshot_existing_dupes_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_snapshot_existing_dupes_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.executemany(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device,item_type,school,color,size,unit_price,count,snapshot_at)
                VALUES ('POS-CEN','Polo','Raja','Red','10',295,?,?)
                """,
                [(4, "2026-08-24T10:00:00Z"), (4, "2026-08-24T10:05:00Z")],
            )
            result = wh_appliers.apply_pos_stock_snapshot(
                wh,
                {
                    "source_device_name": "POS-ZAY",
                    "snapshot_at": "2026-08-24T11:50:00Z",
                    "rows": [
                        {
                            "item_type": "Summer Tee",
                            "school": "Raja",
                            "color": "Red",
                            "size": "10",
                            "unit_price": 295,
                            "count": 3,
                        },
                    ],
                },
                "snapshot-existing-dupe-rows",
            )
            self.assertEqual(result["mirrored_rows"], 1)
            duplicate_groups = wh.execute(
                """
                SELECT COUNT(*)
                  FROM (
                    SELECT 1
                      FROM pos_stocks_mirror
                     GROUP BY source_device,item_type,school,color,size,unit_price
                    HAVING COUNT(*) > 1
                  )
                """
            ).fetchone()[0]
            self.assertEqual(int(duplicate_groups or 0), 0)
            indexes = [
                row[1]
                for row in wh.execute("PRAGMA index_list(pos_stocks_mirror)").fetchall()
            ]
            self.assertIn("idx_pos_stocks_mirror_unique_spec", indexes)
        finally:
            wh.close()

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


class TestWarehouseFinancialLedgerSync(unittest.TestCase):
    def test_sale_voided_applier_writes_signed_amount(self):
        wh_core = _load_module("warehouse_sync_core_voided_sale_autotest", WAREHOUSE_SYNC_CORE_FILE)
        wh_appliers = _load_module("warehouse_sync_appliers_voided_sale_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        wh = sqlite3.connect(":memory:")
        try:
            wh_core.apply_sync_migration(wh)
            wh.execute(
                """
                INSERT INTO sync_inbox(event_uuid, event_type, server_seq, source_device, payload_json, server_created_at, apply_status, applied_at)
                VALUES ('void-sale-event', 'SALE_VOIDED', 1, 'POS-ZAY', '{}', '2026-08-24T11:55:00Z', 'pending', '2026-08-24T11:55:00Z')
                """
            )
            result = wh_appliers.apply_wh_pos_ledger_sale_voided(
                wh,
                {
                    "bill_id": 559,
                    "total": 295,
                    "payment_method": "CASH",
                    "voided_at": "2026-08-24T11:55:00Z",
                    "reason": "mistake",
                },
                "void-sale-event",
            )
            self.assertEqual(result["amount"], -295)
            row = wh.execute(
                """
                SELECT amount, cash_amount, payment_method, category, day
                  FROM pos_financial_ledger
                 WHERE event_uuid='void-sale-event'
                """
            ).fetchone()
            self.assertEqual(float(row[0] or 0), -295)
            self.assertEqual(float(row[1] or 0), -295)
            self.assertEqual(row[2], "CASH")
            self.assertEqual(row[3], "void_bill")
            self.assertEqual(row[4], "2026-08-24")
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
            "allow_global_value_renames": True,
            "value_renames": [
                {"field": "school", "old_value": "الالمانية", "new_value": "الالمانية KG"},
            ],
        }
        result = appliers.apply_spec_renamed(db.conn, payload, "event-value-rename")
        self.assertTrue(result.get("skipped"))
        old_count = _stock_sum(db, "تيشيرت صيفي", "الالمانية", "احمر ف كحلي", "4")
        new_count = _stock_sum(db, "تيشيرت صيفي", "الالمانية KG", "احمر ف كحلي", "4")
        self.assertEqual(old_count, 6)
        self.assertEqual(new_count, 0)

    def test_spec_rename_does_not_apply_unsafe_global_value_rename(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Skirt", "General", "Plaid", "4", 100.0, 1)
        db.add_stock("Tee", "General", "Red", "10", 100.0, 5)
        appliers = _load_module("pos_sync_appliers_no_broad_rename_autotest", POS_SYNC_APPLIERS_FILE)
        payload = {
            "old_spec": {"item_type": "Skirt", "school": "General", "color": "Plaid", "size": "4"},
            "new_spec": {"item_type": "Skirt", "school": "Yahya KG", "color": "Plaid", "size": "4"},
            "changed_fields": ["school"],
            "value_renames": [
                {"field": "school", "old_value": "General", "new_value": "Yahya KG"},
            ],
        }

        result = appliers.apply_spec_renamed(db.conn, payload, "event-unsafe-value-rename")

        self.assertTrue(result.get("skipped"))
        self.assertEqual(_stock_sum(db, "Skirt", "General", "Plaid", "4"), 1)
        self.assertEqual(_stock_sum(db, "Skirt", "Yahya KG", "Plaid", "4"), 0)
        self.assertEqual(_stock_sum(db, "Tee", "General", "Red", "10"), 5)
        self.assertEqual(_stock_sum(db, "Tee", "Yahya KG", "Red", "10"), 0)

    def test_pos_repairs_old_unsafe_value_rename_damage_from_shipment_evidence(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Tee", "Wrong Yahya", "Red", "10", 100.0, 5)
        db.add_or_update_stock_row("Tee", "General", "Red", "10", 100.0, 0)
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "bad-rename-event",
                "SPEC_RENAMED",
                1,
                "WAREHOUSE",
                json.dumps(
                    {
                        "old_spec": {"item_type": "Skirt", "school": "General", "color": "Plaid", "size": "4"},
                        "new_spec": {"item_type": "Skirt", "school": "Wrong Yahya", "color": "Plaid", "size": "4"},
                        "changed_fields": ["school"],
                        "value_renames": [
                            {"field": "school", "old_value": "General", "new_value": "Wrong Yahya"},
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "shipment-general-tee",
                "STOCK_TRANSFER_OUT",
                2,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "ship-general-tee",
                        "items": [
                            {
                                "item_type": "Tee",
                                "school": "General",
                                "color": "Red",
                                "size": "10",
                                "unit_price": 100.0,
                                "qty": 5,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )

        result = db.repair_unsafe_spec_value_rename_damage()

        self.assertGreaterEqual(int(result["updated_rows"]), 1)
        self.assertEqual(_stock_sum(db, "Tee", "General", "Red", "10"), 5)
        self.assertEqual(_stock_sum(db, "Tee", "Wrong Yahya", "Red", "10"), 0)
        self.assertEqual(
            db.conn.execute(
                "SELECT COUNT(*) FROM stocks WHERE school='General' AND count=0"
            ).fetchone()[0],
            0,
        )

    def test_pos_spec_rename_updates_catalog_and_audit_rows(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_or_update_stock_row("Tee", "School A", "Red", "10", 100.0, 0)
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            ("Tee", "School A", "Red", "10", 100.0, "catalog-a", "تعريف حجز", now_iso()),
        )
        report_id = db.create_stock_audit_report([
            {
                "item_type": "Tee",
                "school": "School A",
                "color": "Red",
                "size": "10",
                "expected": 0,
                "actual": 1,
                "diff": 1,
                "unit_price": 100.0,
            }
        ])
        self.assertIsNotNone(report_id)
        appliers = _load_module("pos_sync_appliers_catalog_audit_rename_autotest", POS_SYNC_APPLIERS_FILE)
        payload = {
            "old_spec": {"item_type": "Tee", "school": "School A", "color": "Red", "size": "10"},
            "new_spec": {"item_type": "Tee", "school": "School B", "color": "Red", "size": "10"},
            "changed_fields": ["school"],
        }
        result = appliers.apply_spec_renamed(db.conn, payload, "event-rename-catalog-audit")
        self.assertGreaterEqual(int(result["updated_rows"]), 2)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM branch_catalog_definitions WHERE school='School B'").fetchone()[0],
            1,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stock_audit_report_lines WHERE school='School B'").fetchone()[0],
            1,
        )

    def test_warehouse_partial_spec_edit_emits_rename_even_if_old_spec_remains(self):
        wh_mod = _load_warehouse_module()
        wh = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self._db = wh
        first = wh.add_stock("Tee", "School A", "Red", "10", 1, 1, 100.0, 2)
        wh.add_stock("Tee", "School A", "Red", "10", 1, 2, 100.0, 3)

        updated = wh.update_specs_by_ids([first], school="School B")

        self.assertEqual(updated, 1)
        row = wh.conn.execute(
            "SELECT event_type, payload_json FROM sync_outbox WHERE event_type='SPEC_RENAMED' ORDER BY local_seq DESC LIMIT 1"
        ).fetchone()
        self.assertIsNotNone(row)
        payload = json.loads(row["payload_json"])
        self.assertEqual(payload["old_spec"]["school"], "School A")
        self.assertEqual(payload["new_spec"]["school"], "School B")
        self.assertNotIn("value_renames", payload)

    def test_warehouse_void_bill_returns_stock_to_closed_package(self):
        wh_mod = _load_warehouse_module()
        wh = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self._db = wh
        wh.add_stock("Jacket", "School G", "Navy", "16", 1, 127, 595.0, 5)
        bill_id = wh.create_bill(
            "Customer A",
            [{"item_type": "Jacket", "school": "School G", "color": "Navy", "size": "16", "unit_price": 595.0, "qty": 2}],
        )
        wh.close_package(1, 127)

        wh.void_bill(int(bill_id))

        row = wh.conn.execute(
            """
            SELECT count
              FROM stocks
             WHERE item_type='Jacket'
               AND school='School G'
               AND color='Navy'
               AND size='16'
               AND warehouse_no=1
               AND package_no=127
               AND unit_price=595
            """
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(int(row["count"]), 5)
        self.assertEqual(wh.package_status(1, 127), "CLOSED")

    def test_force_source_of_truth_replays_renames_and_targets_owned_prices(self):
        wh_mod = _load_warehouse_module()
        wh = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self._db = wh
        with wh.conn:
            for name, uuid in (("POS-ZAY", "uuid-zay"), ("POS-OBO", "uuid-obo")):
                wh.conn.execute(
                    """
                    INSERT OR REPLACE INTO known_devices
                        (device_name, device_uuid, role, last_seen_at, updated_at)
                    VALUES (?, ?, 'pos', '', '')
                    """,
                    (name, uuid),
                )
            wh._record_sync_event_or_raise(
                "SPEC_RENAMED",
                {
                    "old_spec": {"item_type": "Tee", "school": "Old School", "color": "Red", "size": "10"},
                    "new_spec": {"item_type": "Tee", "school": "New School", "color": "Red", "size": "10"},
                    "changed_fields": ["school"],
                    "value_renames": [
                        {"field": "school", "old_value": "Old School", "new_value": "New School"},
                    ],
                },
                target_scope="all-pos",
            )
        wh.add_stock("Tee", "Old School", "Red", "10", 1, 1, 100.0, 5)
        wh.create_bill(
            "POS-ZAY",
            [{"item_type": "Tee", "school": "Old School", "color": "Red", "size": "10", "unit_price": 100.0, "qty": 1}],
        )
        wh.add_stock("Tee", "New School", "Red", "10", 1, 1, 225.0, 5)
        wh.add_stock("Tee", "Other School", "Red", "10", 1, 2, 150.0, 5)

        result = wh.force_pos_source_of_truth_sync(["POS-ZAY"])

        self.assertGreaterEqual(int(result["rename_events"]), 1)
        self.assertEqual(int(result["price_events"]), 1)
        rows = wh.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             WHERE event_type IN ('SPEC_RENAMED', 'PRICE_UPDATE')
             ORDER BY local_seq DESC
             LIMIT 2
            """
        ).fetchall()
        scopes = [str(r["target_scope"]) for r in rows]
        self.assertEqual(scopes.count("pos:POS-ZAY"), 2)
        price_payloads = [
            json.loads(r["payload_json"])
            for r in rows
            if r["event_type"] == "PRICE_UPDATE"
        ]
        self.assertEqual(len(price_payloads), 1)
        self.assertEqual(price_payloads[0]["filters"]["school"], "New School")
        self.assertAlmostEqual(float(price_payloads[0]["new_price"]), 225.0, places=2)

    def test_force_source_of_truth_expands_owned_group_to_full_profiles(self):
        wh_mod = _load_warehouse_module()
        wh = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self._db = wh
        with wh.conn:
            wh.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-OBO', 'uuid-obo', 'pos', '', '')
                """
            )
        profile_id = wh.create_price_profile("Sky Profile")
        wh.replace_price_profile_item_prices(
            profile_id,
            "Summer Tee",
            [
                {"size": "10", "price": 310.0},
                {"size": "12", "price": 310.0},
                {"size": "14", "price": 320.0},
            ],
        )
        wh.assign_price_profile("Summer Tee", "Sky School", "Aqua", profile_id)
        wh.upsert_size_profile(
            "Summer Tee",
            "Sky School",
            "Aqua",
            r1_start=10,
            r1_end=14,
            r2_start=None,
            r2_end=None,
            has_alpha=False,
        )
        wh.add_stock("Summer Tee", "Sky School", "Aqua", "10", 1, 1, 330.0, 5)
        wh.create_bill(
            "POS-OBO",
            [{"item_type": "Summer Tee", "school": "Sky School", "color": "Aqua", "size": "10", "unit_price": 330.0, "qty": 1}],
        )

        result = wh.force_pos_source_of_truth_sync(["POS-OBO"])

        self.assertEqual(int(result["branches"]), 1)
        self.assertGreaterEqual(int(result["specs"]), 3)
        snapshot = wh.conn.execute(
            """
            SELECT payload_json
              FROM sync_outbox
             WHERE event_type='POS_OWNERSHIP_SNAPSHOT'
             ORDER BY local_seq DESC
             LIMIT 1
            """
        ).fetchone()
        self.assertIsNotNone(snapshot)
        payload = json.loads(snapshot["payload_json"])
        prices = {
            str(row["size"]): float(row["unit_price"])
            for row in payload["specs"]
            if row["item_type"] == "Summer Tee"
        }
        self.assertAlmostEqual(prices["10"], 310.0, places=2)
        self.assertAlmostEqual(prices["12"], 310.0, places=2)
        self.assertAlmostEqual(prices["14"], 320.0, places=2)

    def test_branch_stock_repair_multi_selection_builds_one_rename_per_row(self):
        wh_mod = _load_warehouse_module()
        rows = [
            {"item_type": "Tee", "school": "Old School", "color": "Red", "size": "10"},
            {"item_type": "Tee", "school": "Old School", "color": "Red", "size": "12"},
        ]

        payloads = wh_mod.BranchStockWindow._spec_rename_payloads_for_selected_rows(
            rows,
            new_school="New School",
            new_color="Blue",
        )

        self.assertEqual(len(payloads), 2)
        self.assertEqual({p["old_spec"]["size"] for p in payloads}, {"10", "12"})
        self.assertEqual({p["new_spec"]["size"] for p in payloads}, {"10", "12"})
        for payload in payloads:
            self.assertEqual(payload["old_spec"]["school"], "Old School")
            self.assertEqual(payload["new_spec"]["school"], "New School")
            self.assertEqual(payload["old_spec"]["color"], "Red")
            self.assertEqual(payload["new_spec"]["color"], "Blue")
            self.assertEqual(set(payload["changed_fields"]), {"school", "color"})

    def test_pos_price_update_creates_catalog_price_for_owned_zero_spec(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            ("Price Tee", "Owned School", "Blue", "12", 100.0, "catalog-price", "تعريف حجز", now_iso()),
        )
        appliers = _load_module("pos_sync_appliers_price_catalog_autotest", POS_SYNC_APPLIERS_FILE)

        result = appliers.apply_price_update(
            db.conn,
            {
                "new_price": 175.0,
                "filters": {
                    "item_type": "Price Tee",
                    "school": "Owned School",
                    "color": "Blue",
                    "size": "12",
                },
            },
            "event-price-owned-catalog",
        )

        self.assertEqual(int(result["catalog_rows"]), 1)
        row = db.conn.execute(
            """
            SELECT unit_price, count
              FROM stocks
             WHERE item_type='Price Tee'
               AND school='Owned School'
               AND color='Blue'
               AND size='12'
            """
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertAlmostEqual(float(row["unit_price"]), 175.0, places=2)
        self.assertEqual(int(row["count"]), 0)
        cat = db.conn.execute(
            "SELECT unit_price FROM branch_catalog_definitions WHERE item_type='Price Tee' AND school='Owned School'"
        ).fetchone()
        self.assertAlmostEqual(float(cat["unit_price"]), 175.0, places=2)

    def test_pos_price_update_skips_stock_without_exact_branch_definition(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Loose Tee", "Loose School", "Blue", "12", 100.0, 4)
        appliers = _load_module("pos_sync_appliers_price_unowned_autotest", POS_SYNC_APPLIERS_FILE)

        result = appliers.apply_price_update(
            db.conn,
            {
                "new_price": 175.0,
                "filters": {
                    "item_type": "Loose Tee",
                    "school": "Loose School",
                    "color": "Blue",
                    "size": "12",
                },
            },
            "event-price-unowned",
        )

        self.assertTrue(result.get("skipped"))
        row = db.conn.execute(
            "SELECT unit_price, count FROM stocks WHERE item_type='Loose Tee' AND school='Loose School'"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertAlmostEqual(float(row["unit_price"]), 100.0, places=2)
        self.assertEqual(int(row["count"]), 4)

    def test_pos_price_update_refuses_partial_filter_even_for_owned_definition(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            ("Partial Tee", "Partial School", "Blue", "12", 100.0, "catalog-partial", "تعريف حجز", now_iso()),
        )
        appliers = _load_module("pos_sync_appliers_price_partial_autotest", POS_SYNC_APPLIERS_FILE)

        result = appliers.apply_price_update(
            db.conn,
            {
                "new_price": 175.0,
                "filters": {
                    "item_type": "Partial Tee",
                    "school": "Partial School",
                    "color": "Blue",
                },
            },
            "event-price-partial",
        )

        self.assertTrue(result.get("skipped"))
        self.assertEqual(result.get("reason"), "refusing non-exact price update")

    def test_pos_price_update_refreshes_pending_shipment_price_anchor(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            ("Profile Tee", "Profile School", "Yellow", "16", 280.0, "old-catalog", "old catalog", now_iso()),
        )
        db.conn.execute(
            """
            INSERT INTO stocks(item_type,school,color,size,unit_price,count)
            VALUES(?,?,?,?,?,?)
            """,
            ("Profile Tee", "Profile School", "Yellow", "16", 280.0, 0),
        )
        db.conn.execute(
            """
            INSERT INTO incoming_shipment_items_pending(
                shipment_uuid,line_index,item_type,school,color,size,
                unit_price,expected_qty,received_qty,status
            )
            VALUES(?,?,?,?,?,?,?,?,?,?)
            """,
            ("ship-old-price", 0, "Profile Tee", "Profile School", "Yellow", "16", 280.0, 3, None, "PENDING"),
        )
        appliers = _load_module("pos_sync_appliers_pending_price_autotest", POS_SYNC_APPLIERS_FILE)

        result = appliers.apply_price_update(
            db.conn,
            {
                "new_price": 300.0,
                "filters": {
                    "item_type": "Profile Tee",
                    "school": "Profile School",
                    "color": "Yellow",
                    "size": "16",
                },
                "allow_catalog_definition": True,
            },
            "event-price-pending-anchor",
        )

        self.assertEqual(result.get("pending_rows"), 1)
        pending = db.conn.execute(
            "SELECT unit_price FROM incoming_shipment_items_pending WHERE shipment_uuid='ship-old-price'"
        ).fetchone()
        self.assertAlmostEqual(float(pending["unit_price"]), 300.0, places=2)

    def test_lightweight_visibility_repair_prefers_current_catalog_price_over_old_inbox_anchor(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            ("Profile Tee", "Profile School", "Yellow", "16", 300.0, "price-update-event", "Warehouse price profile sync", now_iso()),
        )
        payload = {
            "shipment_uuid": "old-inbox-price",
            "from_device": "WAREHOUSE",
            "note": "bill #old",
            "items": [
                {
                    "item_type": "Profile Tee",
                    "school": "Profile School",
                    "color": "Yellow",
                    "size": "16",
                    "unit_price": 280.0,
                    "qty": 2,
                }
            ],
        }
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            ("old-inbox-price-event", "STOCK_TRANSFER_OUT", 10, "WAREHOUSE", json.dumps(payload, ensure_ascii=False), now_iso(), "ok"),
        )

        result = db.ensure_branch_catalog_stock_rows_lightweight()

        self.assertEqual(result["created"], 1)
        row = db.conn.execute(
            """
            SELECT unit_price,count
              FROM stocks
             WHERE item_type='Profile Tee'
               AND school='Profile School'
               AND color='Yellow'
               AND size='16'
            """
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertAlmostEqual(float(row["unit_price"]), 300.0, places=2)
        self.assertEqual(int(row["count"]), 0)

    def test_price_profile_catalog_emits_related_branch_price_updates(self):
        wh_mod = _load_warehouse_module()
        wh = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        self._db = wh
        with wh.conn:
            wh.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-OBO', 'uuid-obo', 'pos', '', '')
                """
            )
        profile_id = wh.create_price_profile("Profile Sync")
        wh.replace_price_profile_item_prices(
            profile_id,
            "Profile Tee",
            [{"size": "8", "price": 180.0}],
        )
        wh.assign_price_profile("Profile Tee", "School P", "Navy", profile_id)
        wh.send_catalog_rows_to_pos(
            "POS-OBO",
            [{"item_type": "Profile Tee", "school": "School P", "color": "Navy", "size": "8", "unit_price": 100.0}],
            note="Catalog-only sync for POS reservations",
        )

        sent = wh.send_price_profile_catalog_to_all_pos(
            profile_id,
            [("Profile Tee", "School P", "Navy")],
            note="Price profile prices updated",
        )

        self.assertEqual(sent, 1)
        row = wh.conn.execute(
            "SELECT target_scope, payload_json FROM sync_outbox WHERE event_type='PRICE_UPDATE' ORDER BY local_seq DESC LIMIT 1"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(row["target_scope"], "pos:POS-OBO")
        payload = json.loads(row["payload_json"])
        self.assertEqual(payload["filters"]["school"], "School P")
        self.assertEqual(payload["filters"]["size"], "8")
        self.assertAlmostEqual(float(payload["new_price"]), 180.0, places=2)


    def test_pos_ignores_generic_catalog_upsert_from_other_branches(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_catalog_ignore_autotest", POS_SYNC_APPLIERS_FILE)
        result = appliers.apply_catalog_upsert(
            db.conn,
            {
                "items": [{"item_type": "Bahtim Tee", "default_price": 100.0}],
                "size_profiles": [
                    {
                        "item_type": "Bahtim Tee",
                        "school": "Bahtim School",
                        "color": "Red",
                        "num_start_1": 2,
                        "num_end_1": 10,
                        "num_start_2": None,
                        "num_end_2": None,
                        "has_alpha": 0,
                    }
                ],
                "spec_history": [
                    {"field": "school", "value": "Bahtim School"},
                    {"field": "item_type", "value": "Bahtim Tee"},
                ],
            },
            "event-catalog-bahtim",
        )
        self.assertTrue(result.get("skipped"))
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Bahtim School'").fetchone()[0],
            0,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM spec_history WHERE value='Bahtim School'").fetchone()[0],
            0,
        )
        self.assertIsNone(db.get_size_profile("Bahtim Tee", "Bahtim School", "Red"))

    def test_explicit_reservation_catalog_definition_survives_cleanup(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_catalog_definition_autotest", POS_SYNC_APPLIERS_FILE)
        result = appliers.apply_stock_transfer_out(
            db.conn,
            {
                "shipment_uuid": "catalog-explicit-1",
                "from_device": "WAREHOUSE-MAIN",
                "note": "Catalog-only sync for POS reservations",
                "items": [
                    {
                        "item_type": "Reservation Tee",
                        "school": "Reservation School",
                        "color": "Green",
                        "size": "8",
                        "unit_price": 150.0,
                        "qty": 0,
                        "catalog_only": True,
                    }
                ],
            },
            "event-explicit-catalog",
        )
        self.assertEqual(result["catalog_rows"], 1)
        db.cleanup_unowned_branch_catalog_rows()
        row = db.conn.execute(
            "SELECT count FROM stocks WHERE item_type='Reservation Tee' AND school='Reservation School'"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(int(row["count"]), 0)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM branch_catalog_definitions WHERE school='Reservation School'").fetchone()[0],
            1,
        )
        self.assertNotIn(
            "Reservation School",
            db.get_distinct_filtered("school", {}, available_only=True),
        )
        self.assertIn(
            "Reservation School",
            db.get_distinct_filtered("school", {}, available_only=False),
        )
        self.assertEqual(
            db.get_available_qty_for_reservation("Reservation Tee", "Reservation School", "Green", "8"),
            0,
        )

    def test_arabic_reservation_catalog_definition_survives_cleanup(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_arabic_catalog_definition_autotest", POS_SYNC_APPLIERS_FILE)
        result = appliers.apply_stock_transfer_out(
            db.conn,
            {
                "shipment_uuid": "catalog-arabic-reservation-1",
                "from_device": "WAREHOUSE-MAIN",
                "note": "\u062a\u0639\u0631\u064a\u0641 \u062d\u062c\u0632",
                "items": [
                    {
                        "item_type": "Arabic Reservation Tee",
                        "school": "Arabic Reservation School",
                        "color": "Green",
                        "size": "8",
                        "unit_price": 150.0,
                        "qty": 0,
                        "catalog_only": True,
                    }
                ],
            },
            "event-arabic-catalog",
        )
        self.assertEqual(result["catalog_rows"], 1)
        db.cleanup_unowned_branch_catalog_rows()
        self.assertEqual(
            db.conn.execute(
                "SELECT COUNT(*) FROM branch_catalog_definitions WHERE school='Arabic Reservation School'"
            ).fetchone()[0],
            1,
        )

    def test_positive_branch_shipment_seeds_visible_catalog_row_until_confirmed(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_positive_shipment_catalog_autotest", POS_SYNC_APPLIERS_FILE)
        result = appliers.apply_stock_transfer_out(
            db.conn,
            {
                "shipment_uuid": "positive-shipment-1",
                "from_device": "WAREHOUSE-MAIN",
                "note": "bill #999",
                "items": [
                    {
                        "item_type": "Shipment Tee",
                        "school": "Shipment School",
                        "color": "Black",
                        "size": "10",
                        "unit_price": 120.0,
                        "qty": 3,
                    }
                ],
            },
            "event-positive-shipment",
        )
        self.assertEqual(result["queued_rows"], 1)
        self.assertEqual(result["auto_received_rows"], 0)
        self.assertTrue(result["needs_verification"])
        self.assertEqual(result["catalog_rows"], 1)
        row = db.conn.execute(
            "SELECT count FROM stocks WHERE item_type='Shipment Tee' AND school='Shipment School'"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(int(row["count"]), 0)
        self.assertIn("Shipment School", db.get_distinct_filtered("school", {}))

        cleaned = db.cleanup_unowned_branch_catalog_rows()
        self.assertEqual(cleaned["stock_rows"], 0)
        self.assertEqual(
            db.conn.execute("SELECT COALESCE(SUM(count),0) FROM stocks WHERE school='Shipment School'").fetchone()[0],
            0,
        )
        pending = db.conn.execute(
            """
            SELECT expected_qty, received_qty, status
              FROM incoming_shipment_items_pending
             WHERE shipment_uuid='positive-shipment-1'
            """
        ).fetchone()
        self.assertIsNotNone(pending)
        self.assertEqual(int(pending["expected_qty"]), 3)
        self.assertIsNone(pending["received_qty"])
        self.assertEqual(pending["status"], "PENDING")
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM sync_outbox WHERE event_type='SHIPMENT_RECEIPT_REPORTED'").fetchone()[0],
            0,
        )

    def test_stock_transfer_cancelled_removes_pending_shipment_alert(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_cancel_pending_autotest", POS_SYNC_APPLIERS_FILE)
        shipment_payload = {
            "shipment_uuid": "cancel-pending-1",
            "from_device": "WAREHOUSE-MAIN",
            "note": "bill #1001",
            "items": [
                {
                    "item_type": "Cancel Pending Tee",
                    "school": "Cancel Pending School",
                    "color": "Black",
                    "size": "10",
                    "unit_price": 120.0,
                    "qty": 3,
                }
            ],
        }
        appliers.apply_stock_transfer_out(db.conn, shipment_payload, "event-cancel-pending-out")
        self.assertIsNotNone(db.get_next_incoming_shipment_alert())

        result = appliers.apply_stock_transfer_cancelled(
            db.conn,
            {
                "shipment_uuid": "cancel-pending-1",
                "note": "cancel bill #1001",
                "items": shipment_payload["items"],
            },
            "event-cancel-pending-cancel",
        )
        self.assertEqual(result["pending_rows_cancelled"], 1)
        self.assertEqual(result["removed_qty"], 0)
        self.assertIsNone(db.get_next_incoming_shipment_alert())
        status = db.conn.execute(
            "SELECT status, received_qty FROM incoming_shipment_items_pending WHERE shipment_uuid='cancel-pending-1'"
        ).fetchone()
        self.assertEqual(status["status"], "CANCELLED")
        self.assertEqual(int(status["received_qty"]), 0)

    def test_cancelled_shipment_anchor_does_not_resurrect_during_cleanup(self):
        self._db = _open_db(self._path)
        db = self._db
        shipment_payload = {
            "shipment_uuid": "cancel-cleanup-1",
            "from_device": "WAREHOUSE-MAIN",
            "note": "bill #1003",
            "items": [
                {
                    "item_type": "Cancel Cleanup Tee",
                    "school": "Cancel Cleanup School",
                    "color": "Black",
                    "size": "14",
                    "unit_price": 140.0,
                    "qty": 2,
                }
            ],
        }
        cancel_payload = {"shipment_uuid": "cancel-cleanup-1", "note": "cancel bill #1003"}
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid, event_type, server_seq, source_device, payload_json, applied_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                "event-cancel-cleanup-out",
                "STOCK_TRANSFER_OUT",
                10,
                "WAREHOUSE-MAIN",
                json.dumps(shipment_payload),
                now_iso(),
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid, event_type, server_seq, source_device, payload_json, applied_at)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                "event-cancel-cleanup-cancel",
                "STOCK_TRANSFER_CANCELLED",
                11,
                "WAREHOUSE-MAIN",
                json.dumps(cancel_payload),
                now_iso(),
            ),
        )
        db.conn.execute(
            """
            INSERT INTO stocks(item_type, school, color, size, unit_price, count)
            VALUES (?, ?, ?, ?, ?, 0)
            """,
            ("Cancel Cleanup Tee", "Cancel Cleanup School", "Black", "14", 140.0),
        )

        cleaned = db.cleanup_unowned_branch_catalog_rows()
        self.assertGreaterEqual(cleaned["stock_rows"], 1)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Cancel Cleanup School'").fetchone()[0],
            0,
        )

    def test_stock_transfer_cancelled_reverses_confirmed_received_stock(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_cancel_confirmed_autotest", POS_SYNC_APPLIERS_FILE)
        shipment_payload = {
            "shipment_uuid": "cancel-confirmed-1",
            "from_device": "WAREHOUSE-MAIN",
            "note": "bill #1002",
            "items": [
                {
                    "item_type": "Cancel Confirmed Tee",
                    "school": "Cancel Confirmed School",
                    "color": "Black",
                    "size": "12",
                    "unit_price": 130.0,
                    "qty": 5,
                }
            ],
        }
        appliers.apply_stock_transfer_out(db.conn, shipment_payload, "event-cancel-confirmed-out")
        db.confirm_incoming_shipment(
            "cancel-confirmed-1",
            [{"line_index": 0, "received_qty": 4}],
            note="received one short",
        )
        self.assertEqual(_stock_sum(db, "Cancel Confirmed Tee", "Cancel Confirmed School", "Black", "12"), 4)

        result = appliers.apply_stock_transfer_cancelled(
            db.conn,
            {
                "shipment_uuid": "cancel-confirmed-1",
                "note": "cancel bill #1002",
                "items": shipment_payload["items"],
            },
            "event-cancel-confirmed-cancel",
        )
        self.assertEqual(result["confirmed_rows_cancelled"], 1)
        self.assertEqual(result["removed_qty"], 4)
        self.assertEqual(result["shortage_qty"], 0)
        self.assertEqual(_stock_sum(db, "Cancel Confirmed Tee", "Cancel Confirmed School", "Black", "12"), 0)
        movement = db.conn.execute(
            """
            SELECT direction, qty
              FROM movements
             WHERE direction='SHIPMENT_CANCEL'
             ORDER BY id DESC
             LIMIT 1
            """
        ).fetchone()
        self.assertIsNotNone(movement)
        self.assertEqual(int(movement["qty"]), 4)

    def test_branch_stock_reclassification_moves_partial_quantity_only(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_reclass_partial_autotest", POS_SYNC_APPLIERS_FILE)
        db.add_stock("Summer Tee", "Sky KG", "Lemon", "2", 290.0, 7)
        payload = {
            "branch_device": "POS-OCT",
            "from_spec": {"item_type": "Summer Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 290.0},
            "to_spec": {"item_type": "Winter Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 310.0},
            "qty": 4,
            "note": "wrong season",
        }
        result = appliers.apply_branch_stock_reclassified(db.conn, payload, "event-reclass-partial-1")
        self.assertEqual(result["qty"], 4)
        self.assertEqual(_stock_sum(db, "Summer Tee", "Sky KG", "Lemon", "2"), 3)
        self.assertEqual(_stock_sum(db, "Winter Tee", "Sky KG", "Lemon", "2"), 4)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM branch_catalog_definitions WHERE item_type='Winter Tee'").fetchone()[0],
            1,
        )
        directions = [
            r[0] for r in db.conn.execute(
                "SELECT direction FROM movements WHERE direction LIKE 'RECLASS_%' ORDER BY id"
            ).fetchall()
        ]
        self.assertEqual(directions, ["RECLASS_OUT", "RECLASS_IN"])
        again = appliers.apply_branch_stock_reclassified(db.conn, payload, "event-reclass-partial-1")
        self.assertTrue(again["already_applied"])
        self.assertEqual(_stock_sum(db, "Summer Tee", "Sky KG", "Lemon", "2"), 3)
        self.assertEqual(_stock_sum(db, "Winter Tee", "Sky KG", "Lemon", "2"), 4)

    def test_branch_stock_reclassification_rejects_source_shortage(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_reclass_shortage_autotest", POS_SYNC_APPLIERS_FILE)
        db.add_stock("Summer Tee", "Sky KG", "Lemon", "2", 290.0, 2)
        with self.assertRaisesRegex(Exception, "source count is not enough"):
            appliers.apply_branch_stock_reclassified(
                db.conn,
                {
                    "from_spec": {"item_type": "Summer Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 290.0},
                    "to_spec": {"item_type": "Winter Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 310.0},
                    "qty": 4,
                },
                "event-reclass-shortage-1",
            )
        self.assertEqual(_stock_sum(db, "Summer Tee", "Sky KG", "Lemon", "2"), 2)
        self.assertEqual(_stock_sum(db, "Winter Tee", "Sky KG", "Lemon", "2"), 0)

    def test_branch_stock_reclassification_anchor_survives_cleanup(self):
        self._db = _open_db(self._path)
        db = self._db
        payload = {
            "to_spec": {"item_type": "Winter Cleanup Tee", "school": "Sky KG", "color": "Lemon", "size": "2", "unit_price": 310.0},
            "qty": 4,
        }
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid, event_type, server_seq, source_device, payload_json, applied_at)
            VALUES (?, 'BRANCH_STOCK_RECLASSIFIED', 20, 'WAREHOUSE-MAIN', ?, ?)
            """,
            ("event-reclass-cleanup-1", json.dumps(payload), now_iso()),
        )
        db.conn.execute(
            """
            INSERT INTO stocks(item_type, school, color, size, unit_price, count)
            VALUES ('Winter Cleanup Tee', 'Sky KG', 'Lemon', '2', 310, 0)
            """
        )
        cleaned = db.cleanup_unowned_branch_catalog_rows()
        self.assertEqual(cleaned["stock_rows"], 0)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE item_type='Winter Cleanup Tee'").fetchone()[0],
            1,
        )

    def test_duplicate_branch_shipment_heals_missing_visible_catalog_row(self):
        self._db = _open_db(self._path)
        db = self._db
        appliers = _load_module("pos_sync_appliers_duplicate_shipment_heal_autotest", POS_SYNC_APPLIERS_FILE)
        payload = {
            "shipment_uuid": "duplicate-shipment-heal",
            "from_device": "WAREHOUSE-MAIN",
            "note": "bill #93",
            "items": [
                {
                    "item_type": "شروال رياضي",
                    "school": "رجاك",
                    "color": "اسود",
                    "size": "10",
                    "unit_price": 360.0,
                    "qty": 2,
                }
            ],
        }
        appliers.apply_stock_transfer_out(db.conn, payload, "event-duplicate-heal-1")
        db.conn.execute(
            "DELETE FROM stocks WHERE item_type='شروال رياضي' AND school='رجاك' AND color='اسود' AND size='10'"
        )
        result = appliers.apply_stock_transfer_out(db.conn, payload, "event-duplicate-heal-1")
        self.assertTrue(result["duplicate"])
        row = db.conn.execute(
            "SELECT count, unit_price FROM stocks WHERE item_type='شروال رياضي' AND school='رجاك' AND color='اسود' AND size='10'"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(int(row["count"]), 0)
        self.assertEqual(float(row["unit_price"]), 360.0)

    def test_lightweight_visibility_repair_replays_applied_bill_shipment_anchor(self):
        self._db = _open_db(self._path)
        db = self._db
        payload = {
            "shipment_uuid": "331e5473-f9f8-46f5-ad9d-2ba142c65f98",
            "from_device": "WAREHOUSE",
            "note": "bill #93",
            "items": [
                {
                    "item_type": "شروال رياضي",
                    "school": "رجاك",
                    "color": "اسود",
                    "size": "\u200e10\u200e",
                    "unit_price": 360.0,
                    "qty": 2,
                }
            ],
        }
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid, event_type, server_seq, source_device, payload_json, applied_at, apply_status)
            VALUES(?, 'STOCK_TRANSFER_OUT', 930, 'WAREHOUSE', ?, ?, 'ok')
            """,
            ("event-bill-93", json.dumps(payload, ensure_ascii=False), now_iso()),
        )
        result = db.ensure_branch_catalog_stock_rows_lightweight()
        self.assertEqual(result["created"], 1)
        row = db.conn.execute(
            "SELECT count, unit_price FROM stocks WHERE item_type='شروال رياضي' AND school='رجاك' AND color='اسود' AND size='10'"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(int(row["count"]), 0)
        self.assertEqual(float(row["unit_price"]), 360.0)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM incoming_shipment_alerts WHERE shipment_uuid=?", (payload["shipment_uuid"],)).fetchone()[0],
            0,
        )


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

    def test_reservation_mirror_print_uses_selected_branch_stock_source(self):
        wh_mod = _load_warehouse_module()

        resolve = wh_mod.PosReservationsMirrorWindow._resolve_branch_stock_source
        self.assertEqual(resolve("POS-ZAY", "كل فروع POS"), "POS-ZAY")
        self.assertEqual(resolve("POS-ZAY", ""), "POS-ZAY")
        self.assertEqual(resolve("", "كل فروع POS"), "__ALL_POS__")
        self.assertEqual(resolve("POS-ZAY", "POS-OBO"), "POS-OBO")

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

    def test_monitor_prefers_latest_pos_financial_snapshot_for_today_totals(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        appliers = _load_module("warehouse_sync_appliers_financial_snapshot_autotest", WAREHOUSE_SYNC_APPLIERS_FILE)
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-OBO', 'uuid-obo-current', 'pos', '2026-08-08T14:00:00Z', '2026-08-08T14:00:00Z')
                """
            )
            db.conn.execute(
                """
                INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status,apply_at)
                VALUES('old-obo-snapshot-name','POS_STOCK_SNAPSHOT',1,'uuid-obo-old',?,'2026-08-08T10:00:00Z','ok','2026-08-08T10:00:00Z')
                """,
                (json.dumps({"source_device_name": "POS-OBO", "rows": []}),),
            )
            db.conn.execute(
                """
                INSERT INTO pos_financial_ledger
                    (source_device,event_uuid,event_type,category,amount,day,meta_json,created_at)
                VALUES('uuid-obo-old','sale-low','SALE_CREATED','sale',7285,'2026-08-08','{}','2026-08-08 16:00:00')
                """
            )
            appliers.apply_pos_financial_snapshot(
                db.conn,
                {
                    "__source_device__": "uuid-obo-old",
                    "source_device_name": "POS-OBO",
                    "day": "2026-08-08",
                    "cash_total": 6135.0,
                    "visa_total": 14500.0,
                    "total_collected": 20635.0,
                    "snapshot_at": "2026-08-08T14:05:00Z",
                },
                "financial-snapshot-obo",
            )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor("2026-08-08", "2026-08-08")}
        self.assertAlmostEqual(float(monitor["POS-OBO"]["cash_net"]), 6135.0, places=2)
        self.assertAlmostEqual(float(monitor["POS-OBO"]["visa_net"]), 14500.0, places=2)
        self.assertAlmostEqual(float(monitor["POS-OBO"]["total_collected"]), 20635.0, places=2)

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
        from datetime import datetime, timezone

        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        today = datetime.now(timezone.utc).date().isoformat()
        seen_at = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
        with db.conn:
            db.conn.execute(
                """
                INSERT OR REPLACE INTO known_devices
                    (device_name, device_uuid, role, last_seen_at, updated_at)
                VALUES ('POS-ZAY', 'uuid-zay', 'pos', ?, ?)
                """,
                (seen_at, seen_at),
            )
            db.conn.execute(
                """
                INSERT OR REPLACE INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES ('POS-ZAY', ?, 1, 100.0, '2026.8.2.10')
                """,
                (seen_at,),
            )
            db.conn.execute(
                """
                INSERT INTO pos_stock_audit_reports_mirror
                    (audit_uuid, source_device, local_report_id, reason, created_at,
                     total_diff, total_value, event_uuid, received_at)
                VALUES ('POS-ZAY:1', 'POS-ZAY', 1, 'auto-equalization',
                        ?, 0, 0, 'audit-snapshot', ?)
                """,
                (seen_at, seen_at),
            )

        monitor = {r["branch_device"]: r for r in db.list_pos_branch_monitor(today, today)}
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

    def test_branch_stock_profile_issue_keys_flags_missing_and_mismatched_profiles(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        profile_id = db.create_price_profile("Branch Profile")
        db.replace_price_profile_item_prices(
            profile_id,
            "Profile Tee",
            [
                {"size": "10", "price": 180.0},
                {"size": "12", "price": 200.0},
            ],
        )
        db.assign_price_profile("Profile Tee", "School P", "Navy", profile_id)
        window = object.__new__(wh_mod.BranchStockWindow)
        window.db = db
        window._current_branch_names = lambda: ("POS-ZAY", "POS-ZAY")

        issues = wh_mod.BranchStockWindow._price_profile_issue_keys(
            window,
            [
                ("Profile Tee", "School P", "Navy", "10", 180.0, 1),
                ("Profile Tee", "School P", "Navy", "12", 190.0, 1),
                ("No Profile Tee", "School Missing", "Red", "10", 100.0, 1),
            ],
        )

        self.assertNotIn(("profile tee", "school p", "navy", "10"), issues)
        self.assertIn(("profile tee", "school p", "navy", "12"), issues)
        self.assertIn(("no profile tee", "school missing", "red", "10"), issues)

    def test_manual_price_override_skips_only_that_branch_profile_catalog(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        db.add_stock("Manual Tee", "Manual School", "Olive", "14", 1, 1, 200.0, 2)
        db.add_stock("Manual Tee", "Manual School", "Olive", "16", 1, 1, 200.0, 2)
        profile_id = db.create_price_profile("School Profile")
        db.replace_price_profile_item_prices(
            profile_id,
            "Manual Tee",
            [
                {"size": "14", "price": 260.0},
                {"size": "16", "price": 280.0},
            ],
        )
        db.assign_price_profile("Manual Tee", "Manual School", "Olive", profile_id)
        db.set_price_profile_manual_overrides("POS-ZAY", [
            {
                "item_type": "Manual Tee",
                "school": "Manual School",
                "color": "Olive",
                "size": "14",
                "unit_price": 200.0,
            }
        ])

        unscoped_catalog = db.price_profile_catalog_rows_for_targets(
            profile_id,
            [("Manual Tee", "Manual School", "Olive")],
        )
        zayed_catalog = db.price_profile_catalog_rows_for_targets(
            profile_id,
            [("Manual Tee", "Manual School", "Olive")],
            pos_device="POS-ZAY",
        )
        obor_catalog = db.price_profile_catalog_rows_for_targets(
            profile_id,
            [("Manual Tee", "Manual School", "Olive")],
            pos_device="POS-OBO",
        )
        self.assertEqual([row["size"] for row in unscoped_catalog], ["14", "16"])
        self.assertEqual([row["size"] for row in zayed_catalog], ["16"])
        self.assertEqual([row["size"] for row in obor_catalog], ["14", "16"])

        db.related_pos_devices_for_price_update = lambda _filters: ["POS-ZAY", "POS-OBO"]
        sent = db.send_price_profile_catalog_to_all_pos(
            profile_id,
            [("Manual Tee", "Manual School", "Olive")],
        )
        self.assertEqual(sent, 3)
        events = db.conn.execute(
            """
            SELECT target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='PRICE_UPDATE'
             ORDER BY local_seq
            """
        ).fetchall()
        by_target_size = [
            (str(row["target_scope"]), str(json.loads(row["payload_json"])["filters"]["size"]))
            for row in events
        ]
        self.assertNotIn(("pos:POS-ZAY", "14"), by_target_size)
        self.assertIn(("pos:POS-OBO", "14"), by_target_size)
        self.assertIn(("pos:POS-ZAY", "16"), by_target_size)
        self.assertIn(("pos:POS-OBO", "16"), by_target_size)

        db.clear_price_profile_manual_overrides("POS-ZAY", [
            {
                "item_type": "Manual Tee",
                "school": "Manual School",
                "color": "Olive",
                "size": "14",
            }
        ])
        catalog_after_clear = db.price_profile_catalog_rows_for_targets(
            profile_id,
            [("Manual Tee", "Manual School", "Olive")],
            pos_device="POS-ZAY",
        )
        self.assertEqual([row["size"] for row in catalog_after_clear], ["14", "16"])

    def test_manual_price_override_is_not_reported_as_profile_issue(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        profile_id = db.create_price_profile("Profile With Exception")
        db.replace_price_profile_item_prices(
            profile_id,
            "Exception Tee",
            [
                {"size": "14", "price": 260.0},
                {"size": "16", "price": 280.0},
            ],
        )
        db.assign_price_profile("Exception Tee", "Exception School", "Olive", profile_id)
        db.set_price_profile_manual_overrides("POS-ZAY", [
            {
                "item_type": "Exception Tee",
                "school": "Exception School",
                "color": "Olive",
                "size": "14",
                "unit_price": 200.0,
            }
        ])
        window = object.__new__(wh_mod.BranchStockWindow)
        window.db = db
        window._current_branch_names = lambda: ("POS-ZAY", "POS-ZAY")

        issues = wh_mod.BranchStockWindow._price_profile_issue_keys(
            window,
            [
                ("Exception Tee", "Exception School", "Olive", "14", 200.0, 1),
                ("Exception Tee", "Exception School", "Olive", "16", 200.0, 1),
            ],
        )
        manual = wh_mod.BranchStockWindow._manual_price_override_keys(
            window,
            [
                ("Exception Tee", "Exception School", "Olive", "14", 200.0, 1),
                ("Exception Tee", "Exception School", "Olive", "16", 200.0, 1),
            ],
        )

        self.assertIn(("exception tee", "exception school", "olive", "14"), manual)
        self.assertNotIn(("exception tee", "exception school", "olive", "14"), issues)
        self.assertIn(("exception tee", "exception school", "olive", "16"), issues)

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

    def test_price_profile_assignment_does_not_sync_new_school_to_all_pos(self):
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
        db.add_stock("Profile Leak Tee", "New Profile School", "Black", "8", 1, 1, 100.0, 2)
        profile_id = db.create_price_profile("No POS Leak Profile")
        db.replace_price_profile_item_prices(
            profile_id,
            "Profile Leak Tee",
            [{"size": "8", "price": 175.0}],
        )
        db.assign_price_profile("Profile Leak Tee", "New Profile School", "Black", profile_id)

        result = db.apply_price_profile_to_stock(
            profile_id,
            [("Profile Leak Tee", "New Profile School", "Black")],
        )
        self.assertEqual(int(result["updated"]), 1)
        priced = db.conn.execute(
            """
            SELECT unit_price
              FROM stocks
             WHERE item_type='Profile Leak Tee'
               AND school='New Profile School'
               AND color='Black'
               AND size='8'
               AND COALESCE(count, 0) > 0
             ORDER BY id ASC
             LIMIT 1
            """
        ).fetchone()
        self.assertIsNotNone(priced)
        self.assertAlmostEqual(
            float(priced["unit_price"]),
            175.0,
            places=2,
        )
        rows = db.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             WHERE event_type IN ('PRICE_UPDATE', 'CATALOG_UPSERT')
            ORDER BY local_seq
            """
        ).fetchall()
        self.assertEqual(rows, [])

    def test_price_profile_syncs_only_to_related_pos_branches(self):
        wh_mod = _load_warehouse_module()
        self._db = wh_mod.SqliteDatabase(path=self._path, legacy_json=str(REPO / "nonexistent_warehouse_legacy.json"))
        db = self._db
        with db.conn:
            for name, uuid in (("POS-ZAY", "uuid-zay"), ("POS-OBO", "uuid-obo"), ("POS-OCT", "uuid-oct")):
                db.conn.execute(
                    """
                    INSERT OR REPLACE INTO known_devices
                        (device_name, device_uuid, role, last_seen_at, updated_at)
                    VALUES (?, ?, 'pos', '', '')
                    """,
                    (name, uuid),
                )
            db.conn.execute(
                """
                INSERT INTO pos_stocks_mirror
                    (source_device, item_type, school, color, size, unit_price, count, snapshot_at)
                VALUES ('POS-ZAY', 'Related Tee', 'Related School', 'Black', '8', 100.0, 0, '2026-08-03T10:00:00Z')
                """
            )
        db.add_stock("Related Tee", "Related School", "Black", "8", 1, 1, 100.0, 2)
        profile_id = db.create_price_profile("Related POS Profile")
        db.replace_price_profile_item_prices(
            profile_id,
            "Related Tee",
            [{"size": "8", "price": 175.0}],
        )
        db.assign_price_profile("Related Tee", "Related School", "Black", profile_id)
        db.send_catalog_rows_to_pos(
            "POS-OBO",
            [
                {
                    "item_type": "Related Tee",
                    "school": "Related School",
                    "color": "Black",
                    "size": "8",
                    "unit_price": 175.0,
                }
            ],
            note="Catalog-only sync for POS reservations",
        )

        result = db.apply_price_profile_to_stock(
            profile_id,
            [("Related Tee", "Related School", "Black")],
        )
        self.assertEqual(int(result["updated"]), 1)
        rows = db.conn.execute(
            """
            SELECT target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='PRICE_UPDATE'
             ORDER BY target_scope
            """
        ).fetchall()
        scopes = [str(r["target_scope"]) for r in rows]
        self.assertEqual(scopes, ["pos:POS-OBO"])
        for row in rows:
            payload = json.loads(row["payload_json"])
            self.assertEqual(payload["filters"]["item_type"], "Related Tee")
            self.assertEqual(payload["filters"]["school"], "Related School")
            self.assertEqual(payload["filters"]["color"], "Black")
            self.assertEqual(payload["filters"]["size"], "8")
        self.assertNotIn("pos:POS-OCT", scopes)
        self.assertNotIn("all-pos", scopes)

    def test_size_profile_update_does_not_seed_all_pos_catalogs(self):
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
        self.assertEqual(sent, 0)

        rows = db.conn.execute(
            """
            SELECT event_type, target_scope, payload_json
              FROM sync_outbox
             WHERE event_type='CATALOG_UPSERT'
            ORDER BY local_seq
            """
        ).fetchall()
        self.assertEqual(rows, [])


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

    def test_partial_reservation_delivery_collects_selected_remaining_items(self):
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

        with self.assertRaisesRegex(ValueError, "240"):
            db.deliver_reservation_items(ids[:3], collected_amount=300.0)

        summary = db.deliver_reservation_items(ids[:3], collected_amount=240.0)
        self.assertEqual(summary["delivered_items"], 3)
        self.assertFalse(summary["group_completed"])

        placeholders = ",".join("?" * len(ids))
        rows = db.conn.execute(
            f"SELECT id, status, paid_amount FROM reservations WHERE id IN ({placeholders}) ORDER BY id",
            tuple(ids),
        ).fetchall()
        self.assertEqual([str(r["status"]) for r in rows[:3]], ["تم التسليم", "تم التسليم", "تم التسليم"])
        self.assertEqual([float(r["paid_amount"]) for r in rows[:3]], [100.0, 100.0, 100.0])
        self.assertEqual(sum(float(r["paid_amount"]) for r in rows[3:]), 40.0)

        paid_now = db.conn.execute(
            "SELECT COALESCE(SUM(unit_price),0) FROM movements WHERE direction='DELIVER_PAY'"
        ).fetchone()[0]
        self.assertEqual(float(paid_now), 240.0)

    def test_reservation_paid_default_tracks_total_until_user_edits(self):
        class FakeVar:
            def __init__(self):
                self.value = ""

            def set(self, value):
                self.value = str(value)

            def get(self):
                return self.value

        frame = object.__new__(_MOD.POSFrame)
        frame.active_bill = 5
        frame.bills = {
            5: [
                {"unit_price": 100.0, "qty": 2},
                {"unit_price": 175.0, "qty": 1},
            ]
        }
        frame.reservation_paid_values = {5: "0", 7: "0"}
        frame.reservation_paid_manual = {5: False, 7: False}
        frame._updating_reservation_paid = False
        frame._paid_var = FakeVar()

        _MOD.POSFrame._refresh_reservation_paid_default(frame)
        self.assertEqual(frame._paid_var.get(), "375")
        self.assertFalse(frame.reservation_paid_manual[5])

        frame.reservation_paid_manual[5] = True
        frame._paid_var.set("50")
        frame.bills[5].append({"unit_price": 25.0, "qty": 1})
        _MOD.POSFrame._refresh_reservation_paid_default(frame)
        self.assertEqual(frame._paid_var.get(), "50")

    def test_reservation_delivery_can_change_size_and_reprice(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Change Tee", "School Change", "Red", "14", 275.0, 1)
        db.add_stock("Change Tee", "School Change", "Red", "16", 295.0, 1)
        before_old = _stock_sum(db, "Change Tee", "School Change", "Red", "14")
        before_new = _stock_sum(db, "Change Tee", "School Change", "Red", "16")
        rid = int(db.create_reservation(
            "Change Customer",
            [{"item_type": "Change Tee", "school": "School Change", "color": "Red", "size": "14", "unit_price": 275.0, "qty": 1}],
            paid_amount=100.0,
        )[0])

        summary = db.deliver_reservation_items(
            [rid],
            collected_amount=195.0,
            replacements={rid: {"item_type": "Change Tee", "size": "16"}},
        )
        self.assertEqual(summary["delivered_items"], 1)
        self.assertAlmostEqual(float(summary["collected_amount"]), 195.0, places=2)
        self.assertAlmostEqual(float(summary["refund_amount"]), 0.0, places=2)
        self.assertEqual(_stock_sum(db, "Change Tee", "School Change", "Red", "14"), before_old)
        self.assertEqual(_stock_sum(db, "Change Tee", "School Change", "Red", "16"), before_new - 1)

        row = db.conn.execute("SELECT item_type, size, unit_price, total_amount, paid_amount, status FROM reservations WHERE id=?", (rid,)).fetchone()
        self.assertEqual(str(row["item_type"]), "Change Tee")
        self.assertEqual(str(row["size"]), "16")
        self.assertAlmostEqual(float(row["unit_price"]), 295.0, places=2)
        self.assertAlmostEqual(float(row["total_amount"]), 295.0, places=2)
        self.assertAlmostEqual(float(row["paid_amount"]), 295.0, places=2)
        self.assertEqual(str(row["status"]), _MOD.RESERVATION_STATUS_DELIVERED)

        paid_now = db.conn.execute(
            "SELECT COALESCE(SUM(unit_price),0) FROM movements WHERE direction='DELIVER_PAY'"
        ).fetchone()[0]
        self.assertAlmostEqual(float(paid_now), 195.0, places=2)
        out_size = db.conn.execute(
            "SELECT size FROM movements WHERE direction='OUT' AND note LIKE '%Reservation delivered%' ORDER BY id DESC LIMIT 1"
        ).fetchone()["size"]
        self.assertEqual(str(out_size), "16")

    def test_reservation_delivery_can_change_item_type_and_refund_lower_price(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Reserved Shirt", "School Change", "Blue", "12", 300.0, 1)
        db.add_stock("Delivered Pants", "School Change", "Blue", "12", 200.0, 1)
        before_old = _stock_sum(db, "Reserved Shirt", "School Change", "Blue", "12")
        before_new = _stock_sum(db, "Delivered Pants", "School Change", "Blue", "12")
        rid = int(db.create_reservation(
            "Change Customer",
            [{"item_type": "Reserved Shirt", "school": "School Change", "color": "Blue", "size": "12", "unit_price": 300.0, "qty": 1}],
            paid_amount=250.0,
        )[0])

        summary = db.deliver_reservation_items(
            [rid],
            collected_amount=0.0,
            replacements={rid: {"item_type": "Delivered Pants", "size": "12"}},
        )
        self.assertEqual(summary["delivered_items"], 1)
        self.assertAlmostEqual(float(summary["collected_amount"]), 0.0, places=2)
        self.assertAlmostEqual(float(summary["refund_amount"]), 50.0, places=2)
        self.assertEqual(_stock_sum(db, "Reserved Shirt", "School Change", "Blue", "12"), before_old)
        self.assertEqual(_stock_sum(db, "Delivered Pants", "School Change", "Blue", "12"), before_new - 1)

        row = db.conn.execute("SELECT item_type, size, unit_price, total_amount, paid_amount, status FROM reservations WHERE id=?", (rid,)).fetchone()
        self.assertEqual(str(row["item_type"]), "Delivered Pants")
        self.assertEqual(str(row["size"]), "12")
        self.assertAlmostEqual(float(row["unit_price"]), 200.0, places=2)
        self.assertAlmostEqual(float(row["total_amount"]), 200.0, places=2)
        self.assertAlmostEqual(float(row["paid_amount"]), 200.0, places=2)
        self.assertEqual(str(row["status"]), _MOD.RESERVATION_STATUS_DELIVERED)

        refund_now = db.conn.execute(
            "SELECT COALESCE(SUM(unit_price),0) FROM movements WHERE direction='RESERVE_REFUND'"
        ).fetchone()[0]
        self.assertAlmostEqual(float(refund_now), 50.0, places=2)

    def test_reservation_delivery_can_change_color_and_reprice(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Color Tee", "School Color", "Red", "10", 270.0, 1)
        db.add_stock("Color Tee", "School Color", "Blue", "10", 290.0, 1)
        before_old = _stock_sum(db, "Color Tee", "School Color", "Red", "10")
        before_new = _stock_sum(db, "Color Tee", "School Color", "Blue", "10")
        rid = int(db.create_reservation(
            "Color Customer",
            [{"item_type": "Color Tee", "school": "School Color", "color": "Red", "size": "10", "unit_price": 270.0, "qty": 1}],
            paid_amount=100.0,
        )[0])

        summary = db.deliver_reservation_items(
            [rid],
            collected_amount=190.0,
            replacements={rid: {"item_type": "Color Tee", "color": "Blue", "size": "10"}},
        )

        self.assertEqual(summary["delivered_items"], 1)
        self.assertAlmostEqual(float(summary["collected_amount"]), 190.0, places=2)
        self.assertEqual(_stock_sum(db, "Color Tee", "School Color", "Red", "10"), before_old)
        self.assertEqual(_stock_sum(db, "Color Tee", "School Color", "Blue", "10"), before_new - 1)

        row = db.conn.execute("SELECT color, unit_price, total_amount, paid_amount, status FROM reservations WHERE id=?", (rid,)).fetchone()
        self.assertEqual(str(row["color"]), "Blue")
        self.assertAlmostEqual(float(row["unit_price"]), 290.0, places=2)
        self.assertAlmostEqual(float(row["total_amount"]), 290.0, places=2)
        self.assertAlmostEqual(float(row["paid_amount"]), 290.0, places=2)
        self.assertEqual(str(row["status"]), _MOD.RESERVATION_STATUS_DELIVERED)

        out_color = db.conn.execute(
            "SELECT color FROM movements WHERE direction='OUT' AND note LIKE '%Reservation delivered%' ORDER BY id DESC LIMIT 1"
        ).fetchone()["color"]
        self.assertEqual(str(out_color), "Blue")

    def test_cancel_reservation_items_redistributes_down_payment_to_remaining_items(self):
        self._db = _open_db(self._path)
        db = self._db
        ids = [int(x) for x in db.create_reservation(
            "Cancel Customer",
            [
                {"item_type": "Cancel Tee", "school": "School Cancel", "color": "Black", "size": "8", "unit_price": 100.0, "qty": 1},
                {"item_type": "Cancel Tee", "school": "School Cancel", "color": "Black", "size": "10", "unit_price": 100.0, "qty": 1},
                {"item_type": "Cancel Tee", "school": "School Cancel", "color": "Black", "size": "12", "unit_price": 100.0, "qty": 1},
            ],
            paid_amount=90.0,
        )]

        summary = db.cancel_reservation_items([ids[0]], reason="customer changed")
        self.assertEqual(summary["cancelled_items"], 1)
        self.assertAlmostEqual(float(summary["refund_amount"]), 0.0, places=2)

        rows = db.conn.execute(
            "SELECT id, status, paid_amount FROM reservations WHERE id IN (%s) ORDER BY id"
            % ",".join("?" * len(ids)),
            tuple(ids),
        ).fetchall()
        self.assertEqual(str(rows[0]["status"]), "ملغي")
        self.assertEqual(float(rows[0]["paid_amount"]), 0.0)
        self.assertEqual([float(r["paid_amount"]) for r in rows[1:]], [45.0, 45.0])

        refund = db.conn.execute(
            "SELECT COALESCE(SUM(unit_price),0) FROM movements WHERE direction='RESERVE_REFUND'"
        ).fetchone()[0]
        self.assertEqual(float(refund), 0.0)

    def test_cancel_fully_paid_reservation_item_refunds_overflow_payment(self):
        self._db = _open_db(self._path)
        db = self._db
        ids = [int(x) for x in db.create_reservation(
            "Fully Paid Cancel Customer",
            [
                {"item_type": "Full Cancel Tee", "school": "School Full", "color": "Black", "size": "1", "unit_price": 270.0, "qty": 1},
                {"item_type": "Full Cancel Tee", "school": "School Full", "color": "Black", "size": "2", "unit_price": 295.0, "qty": 1},
                {"item_type": "Full Cancel Tee", "school": "School Full", "color": "Black", "size": "3", "unit_price": 295.0, "qty": 1},
            ],
            paid_amount=860.0,
        )]

        summary = db.cancel_reservation_items([ids[2]], reason="customer changed")
        self.assertEqual(summary["cancelled_items"], 1)
        self.assertAlmostEqual(float(summary["redistributed_amount"]), 0.0, places=2)
        self.assertAlmostEqual(float(summary["refund_amount"]), 295.0, places=2)

        rows = db.conn.execute(
            "SELECT id, status, total_amount, paid_amount FROM reservations WHERE id IN (%s) ORDER BY id"
            % ",".join("?" * len(ids)),
            tuple(ids),
        ).fetchall()
        self.assertEqual([float(r["paid_amount"]) for r in rows[:2]], [270.0, 295.0])
        self.assertEqual(str(rows[2]["status"]), _MOD.RESERVATION_STATUS_CANCELLED)
        self.assertAlmostEqual(float(rows[2]["paid_amount"]), 0.0, places=2)

        refund = db.conn.execute(
            "SELECT COALESCE(SUM(unit_price),0) FROM movements WHERE direction='RESERVE_REFUND'"
        ).fetchone()[0]
        self.assertAlmostEqual(float(refund), 295.0, places=2)
        shift_summary = db.get_shift_summary(db.active_shift_id)
        self.assertAlmostEqual(float(shift_summary["cash_collected"]), 565.0, places=2)
        self.assertAlmostEqual(float(shift_summary["res_refund"]), 295.0, places=2)
        self.assertAlmostEqual(float(_MOD.ShiftsSummaryFrame._shift_cancel_total(shift_summary)), 295.0, places=2)

    def test_cancel_last_reservation_item_records_refund_and_reduces_shift_cash(self):
        self._db = _open_db(self._path)
        db = self._db
        ids = [int(x) for x in db.create_reservation(
            "Refund Customer",
            [
                {"item_type": "Refund Tee", "school": "School Refund", "color": "White", "size": "8", "unit_price": 100.0, "qty": 1},
            ],
            paid_amount=40.0,
        )]

        summary = db.cancel_reservation_items(ids, reason="customer cancelled")
        self.assertEqual(summary["cancelled_items"], 1)
        self.assertAlmostEqual(float(summary["refund_amount"]), 40.0, places=2)
        shift_summary = db.get_shift_summary(db.active_shift_id)
        self.assertAlmostEqual(float(shift_summary["cash_collected"]), 0.0, places=2)

        history = [r for r in db.list_bill_history() if str(r.get("history_kind")) == "reservation"]
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]["status"], "ملغي")
        self.assertAlmostEqual(float(history[0]["total"]), 0.0, places=2)

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

    def test_bill_history_includes_payment_method_for_cash_and_visa_bills(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("History Pay Tee", "School Pay", "Blue", "10", 100.0, 10)
        cash_id = db.create_bill(
            "Cash customer",
            [{"item_type": "History Pay Tee", "school": "School Pay", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        visa_id = db.create_bill(
            "Visa customer",
            [{"item_type": "History Pay Tee", "school": "School Pay", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )

        history = {int(r["id"]): r for r in db.list_bill_history() if str(r.get("history_kind")) == "bill"}
        self.assertEqual(history[int(cash_id)]["payment_method"], _MOD.PAYMENT_METHOD_CASH)
        self.assertEqual(history[int(visa_id)]["payment_method"], _MOD.PAYMENT_METHOD_VISA)
        self.assertEqual(_MOD._payment_method_label(history[int(cash_id)]["payment_method"]), "كاش")
        self.assertEqual(_MOD._payment_method_label(history[int(visa_id)]["payment_method"]), "فيزا")

    def test_voided_visa_sale_reduces_visa_not_cash(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Void Tee", "School Void", "Blue", "10", 100.0, 5)
        cash_bill = db.create_bill(
            "Cash customer",
            [{"item_type": "Void Tee", "school": "School Void", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        visa_bill = db.create_bill(
            "Visa customer",
            [{"item_type": "Void Tee", "school": "School Void", "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )
        db.void_bill(visa_bill, "customer cancelled")

        summary = db.get_shift_summary(db.active_shift_id)
        listed = next(r for r in db.get_all_shifts() if int(r["id"]) == int(db.active_shift_id))
        self.assertAlmostEqual(float(summary["cash_collected"]), 100.0, places=2)
        self.assertAlmostEqual(float(summary["visa_collected"]), 0.0, places=2)
        self.assertAlmostEqual(float(listed["cash_collected"]), 100.0, places=2)
        self.assertAlmostEqual(float(listed["visa_collected"]), 0.0, places=2)
        self.assertEqual(int(cash_bill), 1)

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

    def test_pos_available_filters_hide_zero_stock_branch_catalog_rows(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Tee", "Obor School", "Blue", "10", 100.0, 2)
        db.add_or_update_stock_row("Tee", "Bahtim School", "Red", "10", 100.0, 0)
        db.conn.execute(
            "INSERT OR IGNORE INTO spec_history(field,value) VALUES('school','Bahtim School')"
        )
        db.conn.execute(
            "INSERT OR IGNORE INTO spec_history(field,value) VALUES('item_type','Bahtim Tee')"
        )
        db.conn.execute(
            """
            INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)
            """,
            (now_iso(), "PRICE_UPDATE", None, 0, "old generic sync", None, "Bahtim Tee", "Bahtim School", "Red", "10", 100.0),
        )
        db.upsert_size_profile(
            "Bahtim Tee",
            "Bahtim School",
            "Red",
            r1_start=2,
            r1_end=10,
            r2_start=None,
            r2_end=None,
            has_alpha=False,
        )
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            (
                "Bahtim Tee",
                "Bahtim School",
                "Red",
                "10",
                100.0,
                "old-price-profile-event",
                "Price profile catalog applied",
                now_iso(),
            ),
        )

        self.assertIn("Bahtim School", db.get_distinct_filtered("school", {}))
        self.assertEqual(
            db.get_distinct_filtered("school", {}, available_only=True),
            ["Obor School"],
        )
        self.assertEqual(
            db.get_distinct_filtered("color", {"school": "Bahtim School"}, available_only=True),
            [],
        )
        cleaned = db.cleanup_unowned_branch_catalog_rows()
        self.assertGreaterEqual(cleaned["stock_rows"], 1)
        self.assertGreaterEqual(cleaned["catalog_definitions"], 1)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Bahtim School'").fetchone()[0],
            0,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM branch_catalog_definitions WHERE school='Bahtim School'").fetchone()[0],
            0,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM spec_history WHERE value='Bahtim School'").fetchone()[0],
            0,
        )
        self.assertIsNone(db.get_size_profile("Bahtim Tee", "Bahtim School", "Red"))

    def test_branch_catalog_delete_removes_zero_definitions_and_blocks_restore(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_or_update_stock_row("Delete Tee", "Delete School", "Red", "10", 100.0, 0)
        db.add_or_update_stock_row("Keep Tee", "Delete School", "Blue", "12", 150.0, 3)
        db.conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type,school,color,size,unit_price,source_event_uuid,note,created_at)
            VALUES(?,?,?,?,?,?,?,?)
            """,
            ("Delete Tee", "Delete School", "Red", "10", 100.0, "old-catalog", "Reservation definition", now_iso()),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "old-delete-school-anchor",
                "STOCK_TRANSFER_OUT",
                10,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "old-delete-school-anchor",
                        "note": "Reservation catalog definition",
                        "items": [
                            {
                                "item_type": "Delete Tee",
                                "school": "Delete School",
                                "color": "Red",
                                "size": "10",
                                "unit_price": 100.0,
                                "qty": 0,
                                "catalog_only": True,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "applied",
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "delete-school-event",
                "BRANCH_CATALOG_DELETED",
                20,
                "WAREHOUSE",
                json.dumps({"filters": [{"school": "Delete School"}]}, ensure_ascii=False),
                now_iso(),
                "",
            ),
        )
        appliers = _load_module("pos_sync_appliers_catalog_delete_autotest", POS_SYNC_APPLIERS_FILE)

        result = appliers.apply_branch_catalog_deleted(
            db.conn,
            {"filters": [{"school": "Delete School"}], "note": "manual delete"},
            "delete-school-event",
        )

        self.assertGreaterEqual(result["deleted_stock_rows"], 1)
        self.assertGreaterEqual(result["deleted_catalog_definitions"], 1)
        self.assertEqual(result["blocked_positive_filters"], 1)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE item_type='Delete Tee'").fetchone()[0],
            0,
        )
        self.assertEqual(
            db.conn.execute("SELECT count FROM stocks WHERE item_type='Keep Tee'").fetchone()[0],
            3,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM branch_catalog_delete_tombstones WHERE school='Delete School'").fetchone()[0],
            1,
        )

        db.ensure_branch_catalog_stock_rows_lightweight()
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE item_type='Delete Tee'").fetchone()[0],
            0,
        )

        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "new-delete-school-anchor",
                "STOCK_TRANSFER_OUT",
                30,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "new-delete-school-anchor",
                        "note": "Reservation catalog definition",
                        "items": [
                            {
                                "item_type": "Delete Tee",
                                "school": "Delete School",
                                "color": "Red",
                                "size": "10",
                                "unit_price": 100.0,
                                "qty": 0,
                                "catalog_only": True,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "applied",
            ),
        )
        db.ensure_branch_catalog_stock_rows_lightweight()
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE item_type='Delete Tee'").fetchone()[0],
            1,
        )

    def test_pos_cleanup_restores_old_positive_shipment_from_inbox_and_drops_price_catalog(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_or_update_stock_row("Bad Tee", "Unrelated School", "Red", "10", 100.0, 0)
        db.conn.execute(
            "INSERT OR IGNORE INTO spec_history(field,value) VALUES('school','Unrelated School')"
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "old-positive-shipment",
                "STOCK_TRANSFER_OUT",
                10,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "old-bill-shipment",
                        "note": "bill #112",
                        "items": [
                            {
                                "item_type": "Trouser",
                                "school": "\u0639\u0627\u0645",
                                "color": "Black",
                                "size": "10",
                                "unit_price": 275.0,
                                "qty": 4,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "applied",
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "old-price-catalog",
                "STOCK_TRANSFER_OUT",
                11,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "old-price-catalog",
                        "note": "Price profile catalog applied: wrong branch",
                        "items": [
                            {
                                "item_type": "Bad Tee",
                                "school": "Unrelated School",
                                "color": "Red",
                                "size": "10",
                                "unit_price": 100.0,
                                "qty": 0,
                                "catalog_only": True,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "applied",
            ),
        )
        db.conn.execute(
            "INSERT INTO hidden_definitions(field,value,hidden_at) VALUES(?,?,?)",
            ("school", "\u0639\u0627\u0645", now_iso()),
        )

        cleaned = db.cleanup_unowned_branch_catalog_rows()
        self.assertEqual(cleaned["restored_stock_rows"], 1)
        self.assertEqual(cleaned["hidden_definitions"], 1)
        self.assertEqual(
            db.conn.execute("SELECT count FROM stocks WHERE school=?", ("\u0639\u0627\u0645",)).fetchone()["count"],
            0,
        )
        self.assertIn("\u0639\u0627\u0645", db.get_distinct_filtered("school", {}))
        self.assertIn("Black", db.get_distinct_filtered("color", {"school": "\u0639\u0627\u0645"}))
        self.assertIn("10", db.get_distinct_filtered("size", {"school": "\u0639\u0627\u0645", "color": "Black"}))
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Unrelated School'").fetchone()[0],
            0,
        )

    def test_pos_cleanup_removes_positive_stock_even_with_manual_in_movement_without_branch_shipment_trail(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_or_update_stock_row("Wrong Tee", "Wrong School", "Red", "10", 100.0, 7)
        manual_id = db.add_stock("Manual Tee", "Manual School", "Blue", "12", 150.0, 5)
        db.conn.execute(
            """
            INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
            VALUES(?,?,?,?,?,?,?,?,?,?,?)
            """,
            (
                now_iso(),
                "IN",
                manual_id,
                5,
                "استلام شحنة مؤكد #owned",
                None,
                "Manual Tee",
                "Manual School",
                "Blue",
                "12",
                150.0,
            ),
        )
        db.conn.execute(
            "INSERT OR IGNORE INTO spec_history(field,value) VALUES('school','Wrong School')"
        )
        db.conn.execute(
            "INSERT OR IGNORE INTO spec_history(field,value) VALUES('school','Manual School')"
        )

        cleaned = db.cleanup_unowned_branch_catalog_rows()

        self.assertGreaterEqual(cleaned["stock_rows"], 1)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Wrong School'").fetchone()[0],
            0,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Manual School'").fetchone()[0],
            0,
        )
        schools = [r["school"] for r in db.current_inventory({"hide_zero": False})]
        self.assertNotIn("Wrong School", schools)
        self.assertNotIn("Manual School", schools)

    def test_pos_ownership_snapshot_preserves_allowed_counts_and_deletes_unowned(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_stock("Owned Tee", "Owned School", "Blue", "12", 150.0, 5)
        db.add_stock("Wrong Tee", "Wrong School", "Red", "10", 100.0, 7)
        appliers = _load_module("pos_sync_appliers_ownership_snapshot_autotest", POS_SYNC_APPLIERS_FILE)

        result = appliers.apply_pos_ownership_snapshot(
            db.conn,
            {
                "branch_device": "POS-ZAY",
                "mode": "replace",
                "specs": [
                    {
                        "item_type": "Owned Tee",
                        "school": "Owned School",
                        "color": "Blue",
                        "size": "12",
                        "unit_price": 175.0,
                    }
                ],
            },
            "ownership-event-1",
        )

        self.assertEqual(result["quantity_mode"], "preserved_for_allowed_specs")
        owned = db.conn.execute(
            "SELECT count, unit_price FROM stocks WHERE item_type='Owned Tee' AND school='Owned School'"
        ).fetchone()
        self.assertIsNotNone(owned)
        self.assertEqual(int(owned["count"]), 5)
        self.assertAlmostEqual(float(owned["unit_price"]), 175.0, places=2)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='Wrong School'").fetchone()[0],
            0,
        )

    def test_pos_cleanup_replays_local_rename_and_price_events_on_shipment_anchors(self):
        self._db = _open_db(self._path)
        db = self._db
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "shipment-before-rename",
                "STOCK_TRANSFER_OUT",
                10,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "bill-before-rename",
                        "note": "bill #500",
                        "items": [
                            {
                                "item_type": "Old Tee",
                                "school": "Old School",
                                "color": "Blue",
                                "size": "12",
                                "unit_price": 100.0,
                                "qty": 3,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "rename-after-shipment",
                "SPEC_RENAMED",
                11,
                "WAREHOUSE",
                json.dumps(
                    {
                        "old_spec": {
                            "item_type": "Old Tee",
                            "school": "Old School",
                            "color": "Blue",
                            "size": "12",
                        },
                        "new_spec": {
                            "item_type": "New Tee",
                            "school": "New School",
                            "color": "Blue",
                            "size": "12",
                        },
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "price-after-rename",
                "PRICE_UPDATE",
                12,
                "WAREHOUSE",
                json.dumps(
                    {
                        "filters": {
                            "item_type": "New Tee",
                            "school": "New School",
                            "color": "Blue",
                            "size": "12",
                        },
                        "new_price": 140.0,
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )

        cleaned = db.cleanup_unowned_branch_catalog_rows()

        self.assertEqual(cleaned["restored_stock_rows"], 1)
        row = db.conn.execute(
            "SELECT count, unit_price FROM stocks WHERE item_type='New Tee' AND school='New School'"
        ).fetchone()
        self.assertIsNotNone(row)
        self.assertEqual(int(row["count"]), 0)
        self.assertAlmostEqual(float(row["unit_price"]), 140.0, places=2)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE item_type='Old Tee' OR school='Old School'").fetchone()[0],
            0,
        )

    def test_pos_lightweight_catalog_restore_replays_spec_rename(self):
        self._db = _open_db(self._path)
        db = self._db
        db.add_or_update_stock_row("Tee", "عام", "Red", "10", 100.0, 0)
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "shipment-before-school-rename",
                "STOCK_TRANSFER_OUT",
                10,
                "WAREHOUSE",
                json.dumps(
                    {
                        "shipment_uuid": "shipment-before-school-rename",
                        "note": "bill #700",
                        "items": [
                            {
                                "item_type": "Tee",
                                "school": "عام",
                                "color": "Red",
                                "size": "10",
                                "unit_price": 100.0,
                                "qty": 1,
                            }
                        ],
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )
        db.conn.execute(
            """
            INSERT INTO sync_inbox(event_uuid,event_type,server_seq,source_device,payload_json,applied_at,apply_status)
            VALUES(?,?,?,?,?,?,?)
            """,
            (
                "rename-aam-october",
                "SPEC_RENAMED",
                11,
                "WAREHOUSE",
                json.dumps(
                    {
                        "old_spec": {"item_type": "Tee", "school": "عام", "color": "Red", "size": "10"},
                        "new_spec": {"item_type": "Tee", "school": "عام اكتوبر", "color": "Red", "size": "10"},
                    },
                    ensure_ascii=False,
                ),
                now_iso(),
                "ok",
            ),
        )

        result = db.ensure_branch_catalog_stock_rows_lightweight()

        self.assertGreaterEqual(int(result["stale_renamed_zero_rows"]), 1)
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='عام'").fetchone()[0],
            0,
        )
        self.assertEqual(
            db.conn.execute("SELECT COUNT(*) FROM stocks WHERE school='عام اكتوبر'").fetchone()[0],
            1,
        )


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

    def test_school_accounts_day_report_rows_reconcile_to_cash_visa_totals(self):
        self._db = _open_db(self._path)
        db = self._db
        school = "School Day"
        db.add_stock("Day Tee", school, "Blue", "10", 100.0, 10)
        db.create_bill(
            "Cash customer",
            [{"item_type": "Day Tee", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 2}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        db.create_bill(
            "Visa customer",
            [{"item_type": "Day Tee", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )
        db.create_return_bill(
            "Cash return",
            [{"item_type": "Day Tee", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            payment_method=_MOD.PAYMENT_METHOD_CASH,
        )
        db.create_reservation(
            "Visa reservation",
            [{"item_type": "Day Tee", "school": school, "color": "Blue", "size": "10", "unit_price": 100.0, "qty": 1}],
            paid_amount=30.0,
            payment_method=_MOD.PAYMENT_METHOD_VISA,
        )

        report = db.get_school_accounts_day_report([school], date_from="2000-01-01", date_to="2099-12-31")
        rows = report["rows"]
        self.assertGreaterEqual(len(rows), 4)
        cash_sum = sum(float(r["cash_total"]) for r in rows)
        visa_sum = sum(float(r["visa_total"]) for r in rows)
        total_sum = sum(float(r["total_paid"]) for r in rows)
        self.assertAlmostEqual(cash_sum, float(report["total_cash"]), places=2)
        self.assertAlmostEqual(visa_sum, float(report["total_visa"]), places=2)
        self.assertAlmostEqual(total_sum, float(report["total_day"]), places=2)
        self.assertAlmostEqual(float(report["total_cash"]), 100.0, places=2)
        self.assertAlmostEqual(float(report["total_visa"]), 130.0, places=2)
        self.assertAlmostEqual(float(report["total_day"]), 230.0, places=2)
        self.assertTrue(any(str(r["bill_type"]) == "RESERVATION" and float(r["visa_total"]) == 30.0 for r in rows))


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
    def test_app_versions_are_numeric_and_pos_manifest_matches_package(self):
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
        self.assertEqual(current.get("version"), _MOD.APP_VERSION)
        self.assertEqual(manifest.get("version"), _MOD.APP_VERSION)
        self.assertEqual(manifest.get("package_file"), "HosnyPOS-%s.zip" % _MOD.APP_VERSION)


if __name__ == "__main__":
    unittest.main(verbosity=2)
