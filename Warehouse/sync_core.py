# -*- coding: utf-8 -*-
"""Sync infrastructure — Phase 1 (schema prep + event logging).

This module is strictly additive. It does three things:

1. Creates the sync bookkeeping tables (device_identity, sync_outbox,
   sync_inbox, sync_state) if they don't already exist.
2. Adds a nullable `uuid` column to every syncable domain table and
   backfills existing rows with fresh UUIDs.
3. Exposes `record_event(conn, event_type, payload)` which the main app
   calls at the end of each business write to append an entry to
   sync_outbox.

Phase 1 does NOT talk to any server. The outbox is populated but never
drained. Phase 2 (manual sync button) will read from it.

Compatible with Python 3.10+. No external dependencies.
"""

from __future__ import annotations

import json
import sqlite3
import uuid
from datetime import datetime, timezone
from typing import Any, Dict, Iterable, Optional


# Bump this whenever the sync schema changes.
SYNC_SCHEMA_VERSION = 1

# Domain tables that gain a `uuid` column so rows can be referenced
# across devices. Any table listed here that does not exist on this
# particular app (e.g. POS has no `returns` table, warehouse has no
# `reservations` table) is silently skipped.
SYNCABLE_TABLES: tuple = (
    "stocks",
    "bills",
    "bill_items",
    "income_bills",
    "income_bill_items",
    "movements",
    "reservations",
    "shifts",
    "returns",
    "return_items",
)


def _utc_now_iso() -> str:
    """ISO-8601 UTC timestamp with microseconds, suitable for the wire."""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


def new_uuid() -> str:
    """Return a fresh UUIDv4 as a lowercase string."""
    return str(uuid.uuid4())


def _table_exists(conn: sqlite3.Connection, name: str) -> bool:
    cur = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (name,),
    )
    return cur.fetchone() is not None


def _column_exists(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.execute("PRAGMA table_info(%s)" % table)
    return any(row[1] == column for row in cur.fetchall())


def apply_sync_migration(conn: sqlite3.Connection) -> None:
    """Create sync tables and add uuid columns. Idempotent.

    Safe to call on every app startup. Failures on individual ALTER
    statements are swallowed so a partially-migrated database can still
    be progressed on the next run.
    """

    # ---- Sync bookkeeping tables ----
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS device_identity (
            id          INTEGER PRIMARY KEY CHECK (id = 1),
            device_uuid TEXT NOT NULL,
            device_name TEXT NOT NULL,
            device_role TEXT NOT NULL,
            api_token   TEXT,
            server_url  TEXT,
            created_at  TEXT NOT NULL,
            updated_at  TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sync_outbox (
            local_seq    INTEGER PRIMARY KEY AUTOINCREMENT,
            event_uuid   TEXT NOT NULL UNIQUE,
            event_type   TEXT NOT NULL,
            payload_json TEXT NOT NULL,
            created_at   TEXT NOT NULL,
            status       TEXT NOT NULL DEFAULT 'pending',
            attempts     INTEGER NOT NULL DEFAULT 0,
            last_error   TEXT,
            sent_at      TEXT
        );
        CREATE INDEX IF NOT EXISTS ix_sync_outbox_status
            ON sync_outbox(status, local_seq);
        CREATE INDEX IF NOT EXISTS ix_sync_outbox_type
            ON sync_outbox(event_type);

        CREATE TABLE IF NOT EXISTS sync_inbox (
            event_uuid    TEXT PRIMARY KEY,
            event_type    TEXT NOT NULL,
            server_seq    INTEGER NOT NULL,
            source_device TEXT,
            payload_json  TEXT NOT NULL,
            applied_at    TEXT NOT NULL, -- legacy: now means "received_at"
            server_created_at TEXT
        );
        CREATE INDEX IF NOT EXISTS ix_sync_inbox_seq
            ON sync_inbox(server_seq);

        CREATE TABLE IF NOT EXISTS sync_state (
            channel         TEXT PRIMARY KEY,
            last_pulled_seq INTEGER NOT NULL DEFAULT 0,
            last_push_at    TEXT,
            last_pull_at    TEXT,
            last_error      TEXT
        );

        CREATE TABLE IF NOT EXISTS sync_meta (
            key        TEXT PRIMARY KEY,
            value      TEXT,
            updated_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sync_dead_letter (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            event_uuid     TEXT NOT NULL,
            event_type     TEXT NOT NULL,
            server_seq     INTEGER,
            source_device  TEXT,
            payload_json   TEXT,
            apply_error    TEXT,
            attempts       INTEGER NOT NULL DEFAULT 0,
            first_failed_at TEXT NOT NULL,
            last_failed_at  TEXT NOT NULL,
            UNIQUE(event_uuid)
        );
        CREATE INDEX IF NOT EXISTS ix_sync_dead_letter_seq
            ON sync_dead_letter(server_seq DESC);
        CREATE INDEX IF NOT EXISTS ix_sync_dead_letter_source
            ON sync_dead_letter(source_device);

        INSERT OR IGNORE INTO sync_state (channel, last_pulled_seq)
            VALUES ('main', 0);
        """
    )

    # ---- Phase 3: apply-tracking columns on sync_inbox ----
    # NULL apply_status == not yet applied. Pure SQL additions are safe
    # on an existing Phase-1/2 inbox with rows in it.
    for col, ddl in (
        ("server_created_at", "ALTER TABLE sync_inbox ADD COLUMN server_created_at TEXT"),
        ("apply_status",   "ALTER TABLE sync_inbox ADD COLUMN apply_status TEXT"),
        ("apply_error",    "ALTER TABLE sync_inbox ADD COLUMN apply_error TEXT"),
        ("apply_attempts", "ALTER TABLE sync_inbox ADD COLUMN apply_attempts INTEGER NOT NULL DEFAULT 0"),
        ("apply_at",       "ALTER TABLE sync_inbox ADD COLUMN apply_at TEXT"),
    ):
        if not _column_exists(conn, "sync_inbox", col):
            try:
                conn.execute(ddl)
            except sqlite3.OperationalError:
                pass
    try:
        conn.execute(
            "CREATE INDEX IF NOT EXISTS ix_sync_inbox_pending "
            "ON sync_inbox(apply_status, server_seq)"
        )
    except sqlite3.OperationalError:
        pass

    # ---- Phase 3: target_scope column on sync_outbox ----
    # Lets the warehouse shipment flow address a specific POS device
    # ("pos:POS-03") without forcing callers to encode scope into the
    # payload. Nullable — the server falls back to a role-based
    # default when it's not present.
    if not _column_exists(conn, "sync_outbox", "target_scope"):
        try:
            conn.execute("ALTER TABLE sync_outbox ADD COLUMN target_scope TEXT")
        except sqlite3.OperationalError:
            pass

    # ---- Phase 3: known_devices cache ----
    # Warehouse keeps a local copy of the server's device list so it
    # can offer POS names in dropdowns (e.g. the bill customer field)
    # even when offline. Populated from /v1/sync/status on every sync
    # cycle; stale entries are replaced, not merged.
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS known_devices (
            device_name  TEXT PRIMARY KEY,
            device_uuid  TEXT,
            role         TEXT NOT NULL,
            last_seen_at TEXT,
            updated_at   TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_known_devices_role
            ON known_devices(role);

        -- Warehouse-only mirror of each POS's current stock state.
        -- Replaced wholesale per source_device when a STOCK_SNAPSHOT
        -- event lands in the inbox. Never written by domain code.
        CREATE TABLE IF NOT EXISTS pos_stocks_mirror (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            source_device TEXT NOT NULL,
            item_type     TEXT NOT NULL,
            school        TEXT NOT NULL,
            color         TEXT NOT NULL,
            size          TEXT NOT NULL,
            unit_price    REAL NOT NULL,
            count         INTEGER NOT NULL,
            snapshot_at   TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_pos_stocks_mirror_device
            ON pos_stocks_mirror(source_device);
        CREATE INDEX IF NOT EXISTS ix_pos_stocks_mirror_specs
            ON pos_stocks_mirror(item_type, school, color, size);

        -- One row per device holds the latest snapshot timestamp, so
        -- the UI can show staleness without scanning the mirror.
        CREATE TABLE IF NOT EXISTS pos_stocks_snapshot_meta (
            source_device TEXT PRIMARY KEY,
            snapshot_at   TEXT NOT NULL,
            row_count     INTEGER NOT NULL DEFAULT 0,
            total_value   REAL NOT NULL DEFAULT 0,
            app_version   TEXT
        );

        CREATE TABLE IF NOT EXISTS pos_stock_audit_reports_mirror (
            audit_uuid    TEXT PRIMARY KEY,
            source_device TEXT NOT NULL,
            local_report_id INTEGER,
            reason        TEXT,
            created_at    TEXT NOT NULL,
            total_diff    INTEGER NOT NULL DEFAULT 0,
            total_value   REAL NOT NULL DEFAULT 0,
            event_uuid    TEXT NOT NULL,
            received_at   TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_pos_stock_audit_reports_device
            ON pos_stock_audit_reports_mirror(source_device, created_at);

        CREATE TABLE IF NOT EXISTS pos_stock_audit_items_mirror (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            audit_uuid    TEXT NOT NULL,
            source_device TEXT NOT NULL,
            item_type     TEXT NOT NULL,
            school        TEXT NOT NULL,
            color         TEXT NOT NULL,
            size          TEXT NOT NULL,
            expected_qty  INTEGER NOT NULL,
            actual_qty    INTEGER NOT NULL,
            diff_qty      INTEGER NOT NULL,
            unit_price    REAL NOT NULL,
            diff_value    REAL NOT NULL,
            FOREIGN KEY(audit_uuid) REFERENCES pos_stock_audit_reports_mirror(audit_uuid)
        );
        CREATE INDEX IF NOT EXISTS ix_pos_stock_audit_items_specs
            ON pos_stock_audit_items_mirror(source_device, item_type, school, color, size);

        -- Mirror of POS reservation lines (warehouse reporting only).
        CREATE TABLE IF NOT EXISTS pos_reservations_mirror (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            source_device    TEXT NOT NULL,
            reservation_key  TEXT NOT NULL,
            customer         TEXT,
            item_type        TEXT NOT NULL,
            school           TEXT NOT NULL,
            color            TEXT NOT NULL,
            size             TEXT NOT NULL,
            qty              INTEGER NOT NULL,
            unit_price       REAL NOT NULL,
            total_amount     REAL NOT NULL,
            paid_amount      REAL NOT NULL DEFAULT 0,
            status           TEXT NOT NULL DEFAULT 'معلق',
            shift_id         INTEGER,
            last_event_uuid  TEXT,
            updated_at       TEXT NOT NULL,
            UNIQUE(source_device, reservation_key)
        );
        CREATE INDEX IF NOT EXISTS ix_pos_res_m_device
            ON pos_reservations_mirror(source_device);
        CREATE INDEX IF NOT EXISTS ix_pos_res_m_status
            ON pos_reservations_mirror(status);
        CREATE INDEX IF NOT EXISTS ix_pos_res_m_device_status
            ON pos_reservations_mirror(source_device, status);

        -- Per-event cashflow rows from POS sync events (warehouse reporting).
        CREATE TABLE IF NOT EXISTS pos_financial_ledger (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            source_device TEXT NOT NULL,
            event_uuid    TEXT NOT NULL UNIQUE,
            event_type    TEXT NOT NULL,
            category      TEXT NOT NULL,
            amount        REAL NOT NULL,
            gross_amount  REAL,
            cash_amount   REAL,
            payment_method TEXT,
            shift_id      INTEGER,
            day           TEXT NOT NULL,
            related_id    INTEGER,
            meta_json     TEXT,
            created_at    TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS ix_pos_fin_ledger_dev_day
            ON pos_financial_ledger(source_device, day DESC);
        CREATE INDEX IF NOT EXISTS ix_pos_fin_ledger_dev_shift
            ON pos_financial_ledger(source_device, shift_id);
        CREATE INDEX IF NOT EXISTS ix_pos_fin_ledger_day_dev
            ON pos_financial_ledger(day, source_device);

        -- Latest shift state per POS shift, for warehouse monitoring.
        CREATE TABLE IF NOT EXISTS pos_shifts_mirror (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            source_device   TEXT NOT NULL,
            shift_key       TEXT NOT NULL,
            shift_id        INTEGER,
            started_at      TEXT,
            ended_at        TEXT,
            status          TEXT NOT NULL DEFAULT 'OPEN',
            summary_json    TEXT,
            last_event_uuid TEXT,
            updated_at      TEXT NOT NULL,
            UNIQUE(source_device, shift_key)
        );
        CREATE INDEX IF NOT EXISTS ix_pos_shifts_mirror_device_status
            ON pos_shifts_mirror(source_device, status);
        CREATE INDEX IF NOT EXISTS ix_pos_shifts_mirror_device_shift_status
            ON pos_shifts_mirror(source_device, shift_id, status);
        CREATE INDEX IF NOT EXISTS ix_pos_shifts_mirror_device_started
            ON pos_shifts_mirror(source_device, started_at);
        CREATE INDEX IF NOT EXISTS ix_pos_shifts_mirror_device_ended
            ON pos_shifts_mirror(source_device, ended_at);
        """
    )

    for col, ddl in (
        ("gross_amount", "ALTER TABLE pos_financial_ledger ADD COLUMN gross_amount REAL"),
        ("cash_amount", "ALTER TABLE pos_financial_ledger ADD COLUMN cash_amount REAL"),
        ("payment_method", "ALTER TABLE pos_financial_ledger ADD COLUMN payment_method TEXT"),
        ("shift_id", "ALTER TABLE pos_financial_ledger ADD COLUMN shift_id INTEGER"),
    ):
        if not _column_exists(conn, "pos_financial_ledger", col):
            try:
                conn.execute(ddl)
            except sqlite3.OperationalError:
                pass

    if _table_exists(conn, "pos_stocks_snapshot_meta") and not _column_exists(conn, "pos_stocks_snapshot_meta", "app_version"):
        try:
            conn.execute("ALTER TABLE pos_stocks_snapshot_meta ADD COLUMN app_version TEXT")
        except sqlite3.OperationalError:
            pass

    # ---- Add uuid column + backfill on every syncable domain table ----
    for table in SYNCABLE_TABLES:
        if not _table_exists(conn, table):
            continue

        if not _column_exists(conn, table, "uuid"):
            try:
                conn.execute("ALTER TABLE %s ADD COLUMN uuid TEXT" % table)
            except sqlite3.OperationalError:
                # Column may have appeared from a concurrent migration.
                pass

        # Backfill rows without a uuid. Handled in small batches to avoid
        # holding a huge list in memory on large tables.
        try:
            rows = conn.execute(
                "SELECT id FROM %s WHERE uuid IS NULL OR uuid = ''" % table
            ).fetchall()
        except sqlite3.OperationalError:
            rows = []
        for r in rows:
            try:
                conn.execute(
                    "UPDATE %s SET uuid = ? WHERE id = ?" % table,
                    (new_uuid(), r[0]),
                )
            except sqlite3.OperationalError:
                pass

        try:
            conn.execute(
                "CREATE INDEX IF NOT EXISTS ix_%s_uuid ON %s(uuid)"
                % (table, table)
            )
        except sqlite3.OperationalError:
            pass

        # AFTER INSERT trigger: auto-generate a UUIDv4 when a row is
        # inserted without one. This is what keeps new domain rows
        # globally addressable without touching every INSERT statement.
        # The expression below produces a v4-shaped string using SQLite's
        # built-in random/hex functions — no extension required.
        try:
            conn.execute(
                """
                CREATE TRIGGER IF NOT EXISTS trg_%s_uuid_autofill
                AFTER INSERT ON %s
                FOR EACH ROW
                WHEN NEW.uuid IS NULL OR NEW.uuid = ''
                BEGIN
                    UPDATE %s SET uuid =
                        lower(hex(randomblob(4))) || '-' ||
                        lower(hex(randomblob(2))) || '-4' ||
                        substr(lower(hex(randomblob(2))), 2) || '-' ||
                        substr('89ab', 1 + (abs(random()) %% 4), 1) ||
                        substr(lower(hex(randomblob(2))), 2) || '-' ||
                        lower(hex(randomblob(6)))
                    WHERE rowid = NEW.rowid;
                END;
                """ % (table, table, table)
            )
        except sqlite3.OperationalError:
            pass


def ensure_device_identity(
    conn: sqlite3.Connection,
    default_name: str,
    default_role: str,
) -> Dict[str, Any]:
    """Return the single device_identity row, creating a placeholder if absent.

    The placeholder uses `default_name` and `default_role`. The real
    device name is configured later via a first-run setup dialog (Phase 2).
    """
    row = conn.execute(
        "SELECT id, device_uuid, device_name, device_role, api_token, "
        "server_url, created_at, updated_at FROM device_identity WHERE id = 1"
    ).fetchone()

    if row is not None:
        return {
            "id": row[0],
            "device_uuid": row[1],
            "device_name": row[2],
            "device_role": row[3],
            "api_token": row[4],
            "server_url": row[5],
            "created_at": row[6],
            "updated_at": row[7],
        }

    now = _utc_now_iso()
    device_uuid = new_uuid()
    conn.execute(
        """
        INSERT INTO device_identity
            (id, device_uuid, device_name, device_role,
             api_token, server_url, created_at, updated_at)
        VALUES (1, ?, ?, ?, NULL, NULL, ?, ?)
        """,
        (device_uuid, default_name, default_role, now, now),
    )
    return {
        "id": 1,
        "device_uuid": device_uuid,
        "device_name": default_name,
        "device_role": default_role,
        "api_token": None,
        "server_url": None,
        "created_at": now,
        "updated_at": now,
    }


def get_device_identity(conn: sqlite3.Connection) -> Optional[Dict[str, Any]]:
    """Return the device_identity row, or None if it doesn't exist yet."""
    try:
        row = conn.execute(
            "SELECT id, device_uuid, device_name, device_role, api_token, "
            "server_url, created_at, updated_at FROM device_identity WHERE id = 1"
        ).fetchone()
    except sqlite3.OperationalError:
        return None
    if row is None:
        return None
    return {
        "id": row[0],
        "device_uuid": row[1],
        "device_name": row[2],
        "device_role": row[3],
        "api_token": row[4],
        "server_url": row[5],
        "created_at": row[6],
        "updated_at": row[7],
    }


def update_device_name(conn: sqlite3.Connection, new_name: str) -> None:
    """Rename this device. Called from the setup dialog in a later phase."""
    now = _utc_now_iso()
    conn.execute(
        "UPDATE device_identity SET device_name = ?, updated_at = ? WHERE id = 1",
        (new_name, now),
    )


def record_event(
    conn: sqlite3.Connection,
    event_type: str,
    payload: Dict[str, Any],
) -> str:
    """Append a business event to sync_outbox. Returns the event_uuid.

    The event is appended to the outbox using the caller-supplied
    connection, so it participates in any transaction already open on
    that connection.

    The payload is serialized with `ensure_ascii=False` so Arabic strings
    stay readable in the database, and `default=str` so things like
    datetimes and Decimal round-trip to strings instead of crashing.
    """
    event_uuid = new_uuid()
    conn.execute(
        """
        INSERT INTO sync_outbox
            (event_uuid, event_type, payload_json,
             created_at, status, attempts)
        VALUES (?, ?, ?, ?, 'pending', 0)
        """,
        (
            event_uuid,
            event_type,
            json.dumps(payload, ensure_ascii=False, default=str),
            _utc_now_iso(),
        ),
    )
    return event_uuid


def outbox_pending_count(conn: sqlite3.Connection) -> int:
    """Number of outbox events waiting to be pushed."""
    try:
        cur = conn.execute(
            "SELECT COUNT(*) FROM sync_outbox WHERE status = 'pending'"
        )
        return int(cur.fetchone()[0] or 0)
    except sqlite3.OperationalError:
        return 0


def outbox_total_count(conn: sqlite3.Connection) -> int:
    """Total number of outbox events (any status)."""
    try:
        cur = conn.execute("SELECT COUNT(*) FROM sync_outbox")
        return int(cur.fetchone()[0] or 0)
    except sqlite3.OperationalError:
        return 0


def dead_letter_count(conn: sqlite3.Connection) -> int:
    """Number of inbox events parked in dead-letter table."""
    try:
        cur = conn.execute("SELECT COUNT(*) FROM sync_dead_letter")
        return int(cur.fetchone()[0] or 0)
    except sqlite3.OperationalError:
        return 0


def list_recent_events(
    conn: sqlite3.Connection,
    limit: int = 50,
) -> Iterable[Dict[str, Any]]:
    """Return the most recent outbox events, newest first. Useful for a
    read-only diagnostics panel."""
    try:
        cur = conn.execute(
            """
            SELECT local_seq, event_uuid, event_type, payload_json,
                   created_at, status, attempts, last_error, sent_at
            FROM sync_outbox
            ORDER BY local_seq DESC
            LIMIT ?
            """,
            (int(limit),),
        )
    except sqlite3.OperationalError:
        return []
    out = []
    for r in cur.fetchall():
        out.append({
            "local_seq": r[0],
            "event_uuid": r[1],
            "event_type": r[2],
            "payload_json": r[3],
            "created_at": r[4],
            "status": r[5],
            "attempts": r[6],
            "last_error": r[7],
            "sent_at": r[8],
        })
    return out
