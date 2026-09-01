# -*- coding: utf-8 -*-
"""Event appliers — Phase 3.

Each applier takes `(conn, payload, event_uuid)` and mutates the
domain tables. They are pure SQL, with no UI or business-layer
dependencies, and must be idempotent — the dispatcher in
sync_client.apply_inbox already guards against double-apply via
`apply_status='ok'`, but an applier that's called twice on the same
payload must still not corrupt state.

Registry
--------
`APPLIERS` is a dict keyed by event_type. The sync cycle dispatches
through it; unknown event types are recorded as `apply_status='skipped'`
and left in the inbox for a future phase to pick up.

Why per-app registries?
-----------------------
Warehouse and POS need to apply different sets of event types. Phase 3
only enables POS appliers; warehouse's registry is empty this phase
(Phase 4 will fill in the shadow-table mirror appliers for
SALE_CREATED, SHIFT_CLOSED, etc.). The `for_role()` helper returns
the right registry based on the device's role.

File layout
-----------
Both apps ship this file. The POS folder's copy enables the POS
handlers; the warehouse folder's copy has the same code but the
`for_role('warehouse')` path returns an empty dict for now. This keeps
the two trees symmetric and makes the Phase 4 warehouse-side additions
a pure registry edit.
"""

from __future__ import annotations

import json
import sqlite3
import uuid
from datetime import datetime, timezone, timedelta
from typing import Any, Callable, Dict, List, Optional, Set, Tuple


# ----------------------------- helpers -------------------------------- #

def _now_iso() -> str:
    """Local-time ISO timestamp matching existing `movements.ts` format."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


def _record_targeted_event(
    conn: sqlite3.Connection,
    event_type: str,
    target_scope: str,
    payload: Dict[str, Any],
) -> str:
    event_id = str(uuid.uuid4())
    if _has_column(conn, "sync_outbox", "target_scope"):
        conn.execute(
            """
            INSERT INTO sync_outbox
                (event_uuid, event_type, payload_json, created_at, status, attempts, target_scope)
            VALUES (?, ?, ?, ?, 'pending', 0, ?)
            """,
            (
                event_id,
                event_type,
                json.dumps(payload, ensure_ascii=False, default=str),
                _utc_now_iso(),
                target_scope,
            ),
        )
    else:
        payload = dict(payload)
        payload["__target_scope__"] = target_scope
        conn.execute(
            """
            INSERT INTO sync_outbox
                (event_uuid, event_type, payload_json, created_at, status, attempts)
            VALUES (?, ?, ?, ?, 'pending', 0)
            """,
            (
                event_id,
                event_type,
                json.dumps(payload, ensure_ascii=False, default=str),
                _utc_now_iso(),
            ),
        )
    return event_id


def _has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.execute("PRAGMA table_info(%s)" % table)
    return any(r[1] == column for r in cur.fetchall())


def _clean(value: Any) -> str:
    return str(value or "").replace("\u200e", "").replace("\u200f", "").strip()


def _timestamp_sort_value(value: Any) -> str:
    text = _clean(value)
    if not text:
        return ""
    raw = text.replace(" ", "T")
    try:
        if raw.endswith("Z"):
            dt = datetime.fromisoformat(raw[:-1] + "+00:00")
        else:
            dt = datetime.fromisoformat(raw)
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone(timedelta(hours=3)))
        return dt.astimezone(timezone.utc).isoformat()
    except Exception:
        return text


def _timestamp_is_older(candidate: Any, current: Any) -> bool:
    cand = _timestamp_sort_value(candidate)
    cur = _timestamp_sort_value(current)
    return bool(cand and cur and cand < cur)


class ApplyError(Exception):
    """Raised inside an applier to signal a structured failure.

    The dispatcher catches this and records the message in
    `sync_inbox.apply_error`. Bubbles out of the try/except so the
    transaction is rolled back.
    """


# ----------------------------- POS appliers --------------------------- #

def _upsert_spec_history(conn: sqlite3.Connection, specs: Dict[str, Any]) -> None:
    for field in ("item_type", "school", "color", "size"):
        v = _clean(specs.get(field))
        if v:
            conn.execute(
                "INSERT OR IGNORE INTO spec_history(field,value) VALUES(?,?)",
                (field, v),
            )


def _ensure_item_default_price(
    conn: sqlite3.Connection, item_type: str, price: float
) -> None:
    """Seed item_defaults only when the item_type is unknown.

    Never overwrites an existing default (that's a warehouse-side
    `PRICE_UPDATE` responsibility, applied via `apply_price_update`).
    """
    conn.execute(
        "INSERT OR IGNORE INTO item_defaults(item_type, default_price) VALUES(?, ?)",
        (item_type, float(price)),
    )


def _upsert_zero_stock_catalog_row(
    conn: sqlite3.Connection,
    item_type: str,
    school: str,
    color: str,
    size: str,
    unit_price: float,
    *,
    respect_tombstone: bool = True,
) -> bool:
    """Ensure a POS can select a catalog item even before stock arrives."""
    if respect_tombstone and _branch_catalog_spec_is_deleted(
        conn,
        {"item_type": item_type, "school": school, "color": color, "size": size},
    ):
        return False
    row = conn.execute(
        """
        SELECT id, count, unit_price
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
         ORDER BY
           CASE WHEN ABS(unit_price - ?) < 0.001 THEN 0 ELSE 1 END,
           count DESC,
           id ASC
         LIMIT 1
        """,
        (item_type, school, color, size, float(unit_price)),
    ).fetchone()
    if row:
        row_id = row["id"] if isinstance(row, sqlite3.Row) else row[0]
        row_count = row["count"] if isinstance(row, sqlite3.Row) else row[1]
        row_price = row["unit_price"] if isinstance(row, sqlite3.Row) else row[2]
        if int(row_count or 0) == 0 and abs(float(row_price or 0) - float(unit_price)) >= 0.001:
            conn.execute(
                "UPDATE stocks SET unit_price=? WHERE id=?",
                (float(unit_price), int(row_id)),
            )
        return False

    conn.execute(
        """
        INSERT INTO stocks(item_type, school, color, size, unit_price, count)
        VALUES (?, ?, ?, ?, ?, 0)
        """,
        (item_type, school, color, size, float(unit_price)),
    )
    return True


def _mark_branch_catalog_definition(
    conn: sqlite3.Connection,
    item_type: str,
    school: str,
    color: str,
    size: str,
    unit_price: float,
    event_uuid: str,
    note: str = "",
) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS branch_catalog_definitions(
            item_type TEXT NOT NULL,
            school TEXT NOT NULL,
            color TEXT NOT NULL,
            size TEXT NOT NULL,
            unit_price REAL NOT NULL DEFAULT 0,
            source_event_uuid TEXT,
            note TEXT,
            created_at TEXT NOT NULL,
            PRIMARY KEY(item_type, school, color, size)
        )
        """
    )
    now = _now_iso()
    conn.execute(
        """
        UPDATE branch_catalog_definitions
           SET unit_price = ?,
               source_event_uuid = ?,
               note = ?,
               created_at = ?
         WHERE item_type = ?
           AND school = ?
           AND color = ?
           AND size = ?
        """,
        (float(unit_price), str(event_uuid or ""), str(note or ""), now, item_type, school, color, size),
    )
    row = conn.execute(
        """
        SELECT 1 FROM branch_catalog_definitions
         WHERE item_type = ? AND school = ? AND color = ? AND size = ?
         LIMIT 1
        """,
        (item_type, school, color, size),
    ).fetchone()
    if row is None:
        conn.execute(
            """
            INSERT INTO branch_catalog_definitions
                (item_type, school, color, size, unit_price, source_event_uuid, note, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (item_type, school, color, size, float(unit_price), str(event_uuid or ""), str(note or ""), now),
        )


def _ensure_branch_catalog_delete_tombstones(conn: sqlite3.Connection) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS branch_catalog_delete_tombstones(
            item_type TEXT,
            school TEXT,
            color TEXT,
            size TEXT,
            delete_server_seq INTEGER,
            source_event_uuid TEXT,
            note TEXT,
            created_at TEXT NOT NULL
        )
        """
    )
    try:
        cols = {
            str(row[1] if not isinstance(row, sqlite3.Row) else row["name"])
            for row in conn.execute("PRAGMA table_info(branch_catalog_delete_tombstones)").fetchall()
        }
        if "delete_server_seq" not in cols:
            conn.execute("ALTER TABLE branch_catalog_delete_tombstones ADD COLUMN delete_server_seq INTEGER")
    except sqlite3.OperationalError:
        pass
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS idx_branch_catalog_delete_tombstones_specs
        ON branch_catalog_delete_tombstones(
            LOWER(TRIM(COALESCE(item_type, ''))),
            LOWER(TRIM(COALESCE(school, ''))),
            LOWER(TRIM(COALESCE(color, ''))),
            LOWER(TRIM(COALESCE(size, '')))
        )
        """
    )


def _branch_catalog_delete_filter_matches(spec: Dict[str, Any], filters: Dict[str, Any]) -> bool:
    for field in ("item_type", "school", "color", "size"):
        expected = _clean(filters.get(field))
        if not expected:
            continue
        actual = _clean(spec.get(field))
        if actual.casefold() != expected.casefold():
            return False
    return True


def _branch_catalog_spec_is_deleted(conn: sqlite3.Connection, spec: Dict[str, Any]) -> bool:
    try:
        _ensure_branch_catalog_delete_tombstones(conn)
        rows = conn.execute(
            """
            SELECT item_type, school, color, size
              FROM branch_catalog_delete_tombstones
            """
        ).fetchall()
    except sqlite3.OperationalError:
        return False
    for row in rows:
        filters = {
            "item_type": row["item_type"] if isinstance(row, sqlite3.Row) else row[0],
            "school": row["school"] if isinstance(row, sqlite3.Row) else row[1],
            "color": row["color"] if isinstance(row, sqlite3.Row) else row[2],
            "size": row["size"] if isinstance(row, sqlite3.Row) else row[3],
        }
        if _branch_catalog_delete_filter_matches(spec, filters):
            return True
    return False


def apply_branch_catalog_deleted(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Delete branch-visible catalog specs without touching positive stock."""
    raw_filters = payload.get("filters") or []
    if isinstance(raw_filters, dict):
        raw_filters = [raw_filters]
    if not isinstance(raw_filters, list) or not raw_filters:
        raise ApplyError("BRANCH_CATALOG_DELETED requires filters")

    cleaned_filters: List[Dict[str, str]] = []
    for raw in raw_filters:
        if not isinstance(raw, dict):
            continue
        filt = {field: _clean(raw.get(field)) for field in ("item_type", "school", "color", "size")}
        if not any(filt.values()):
            continue
        cleaned_filters.append(filt)
    if not cleaned_filters:
        raise ApplyError("BRANCH_CATALOG_DELETED has no usable filters")

    _ensure_branch_catalog_delete_tombstones(conn)
    now = _now_iso()
    seq_row = conn.execute(
        "SELECT server_seq FROM sync_inbox WHERE event_uuid = ? LIMIT 1",
        (str(event_uuid or ""),),
    ).fetchone()
    try:
        delete_server_seq = int(seq_row["server_seq"] if isinstance(seq_row, sqlite3.Row) else seq_row[0]) if seq_row else None
    except (TypeError, ValueError):
        delete_server_seq = None
    deleted_stock = 0
    deleted_defs = 0
    blocked_positive = 0
    tombstones = 0

    for filt in cleaned_filters:
        where = []
        args: List[Any] = []
        for field in ("item_type", "school", "color", "size"):
            value = filt.get(field) or ""
            if not value:
                continue
            where.append(f"LOWER(TRIM({field})) = LOWER(TRIM(?))")
            args.append(value)
        where_sql = " AND ".join(where) if where else "1=0"

        pos_row = conn.execute(
            f"SELECT 1 FROM stocks WHERE {where_sql} AND COALESCE(count, 0) > 0 LIMIT 1",
            args,
        ).fetchone()
        if pos_row is not None:
            blocked_positive += 1
            continue

        cur = conn.execute(
            f"DELETE FROM stocks WHERE {where_sql} AND COALESCE(count, 0) = 0",
            args,
        )
        deleted_stock += int(cur.rowcount or 0)

        try:
            cur = conn.execute(
                f"""
                DELETE FROM branch_catalog_definitions
                 WHERE {where_sql}
                   AND NOT EXISTS (
                        SELECT 1
                          FROM stocks
                         WHERE LOWER(TRIM(COALESCE(stocks.item_type, ''))) =
                               LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                           AND LOWER(TRIM(COALESCE(stocks.school, ''))) =
                               LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                           AND LOWER(TRIM(COALESCE(stocks.color, ''))) =
                               LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                           AND LOWER(TRIM(COALESCE(stocks.size, ''))) =
                               LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
                           AND COALESCE(stocks.count, 0) > 0
                   )
                """,
                args,
            )
            deleted_defs += int(cur.rowcount or 0)
        except sqlite3.OperationalError:
            pass

        conn.execute(
            """
            INSERT INTO branch_catalog_delete_tombstones(
                item_type, school, color, size, delete_server_seq, source_event_uuid, note, created_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                filt.get("item_type") or "",
                filt.get("school") or "",
                filt.get("color") or "",
                filt.get("size") or "",
                delete_server_seq,
                str(event_uuid or ""),
                _clean(payload.get("note")) or "Warehouse branch catalog delete",
                now,
            ),
        )
        tombstones += 1

    return {
        "deleted_stock_rows": deleted_stock,
        "deleted_catalog_definitions": deleted_defs,
        "blocked_positive_filters": blocked_positive,
        "tombstones": tombstones,
    }


def _is_reservation_catalog_note(note: Any) -> bool:
    text = _clean(note).casefold()
    return "reservation" in text or "حجز" in text


def _apply_size_profile_rows(conn: sqlite3.Connection, profiles: Any) -> int:
    """Upsert warehouse-originated size profile rows on the POS DB.

    Payload rows match `create_branch_shipment` / `CATALOG_UPSERT`:
        {"item_type", "school", "color",
         "num_start_1", "num_end_1", "num_start_2", "num_end_2", "has_alpha"}

    Returns how many profile rows were written.
    """
    if not isinstance(profiles, list) or not profiles:
        return 0
    n = 0
    for p in profiles:
        if not isinstance(p, dict):
            continue
        it = _clean(p.get("item_type"))
        sc = _clean(p.get("school"))
        cl = _clean(p.get("color"))
        if not (it and sc and cl):
            continue
        values = (
            p.get("num_start_1"), p.get("num_end_1"),
            p.get("num_start_2"), p.get("num_end_2"),
            int(p.get("has_alpha") or 0),
        )
        cur = conn.execute(
            """
            UPDATE size_profiles
               SET num_start_1 = ?,
                   num_end_1   = ?,
                   num_start_2 = ?,
                   num_end_2   = ?,
                   has_alpha   = ?,
                   updated_at  = datetime('now')
             WHERE item_type = ?
               AND school = ?
               AND color = ?
            """,
            (*values, it, sc, cl),
        )
        if cur.rowcount == 0:
            conn.execute(
                """
                INSERT OR IGNORE INTO size_profiles
                    (item_type, school, color,
                     num_start_1, num_end_1, num_start_2, num_end_2,
                     has_alpha, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
                """,
                (it, sc, cl, *values),
            )
        n += 1
    return n


def _apply_branch_catalog_rows(conn: sqlite3.Connection, rows: Any, event_uuid: str, note: str = "") -> int:
    if not isinstance(rows, list) or not rows:
        return 0
    n = 0
    for raw in rows:
        if not isinstance(raw, dict):
            continue
        it = _clean(raw.get("item_type"))
        sc = _clean(raw.get("school"))
        cl = _clean(raw.get("color"))
        sz = _clean(raw.get("size"))
        if not (it and sc and cl and sz):
            continue
        try:
            price = float(raw.get("unit_price") or 0)
        except (TypeError, ValueError):
            price = 0.0
        if _upsert_zero_stock_catalog_row(conn, it, sc, cl, sz, price):
            n += 1
        _mark_branch_catalog_definition(
            conn,
            it,
            sc,
            cl,
            sz,
            price,
            event_uuid,
            note=note or "Branch catalog from reclassification",
        )
        _upsert_spec_history(conn, {"item_type": it, "school": sc, "color": cl, "size": sz})
        _ensure_item_default_price(conn, it, price)
    return n


def _update_item_default_price(
    conn: sqlite3.Connection,
    item_type: str,
    price: float,
) -> None:
    cur = conn.execute(
        "UPDATE item_defaults SET default_price = ? WHERE item_type = ?",
        (float(price), item_type),
    )
    if cur.rowcount == 0:
        conn.execute(
            "INSERT OR IGNORE INTO item_defaults(item_type, default_price) VALUES(?, ?)",
            (item_type, float(price)),
        )


def _update_pending_shipment_spec_price(
    conn: sqlite3.Connection,
    filters: Dict[str, str],
    price: float,
) -> int:
    try:
        cur = conn.execute(
            """
            UPDATE incoming_shipment_items_pending
               SET unit_price = ?
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
            """,
            (
                float(price),
                filters["item_type"],
                filters["school"],
                filters["color"],
                filters["size"],
            ),
        )
        return int(cur.rowcount or 0)
    except sqlite3.OperationalError:
        return 0


def _spec_is_branch_defined(conn: sqlite3.Connection, specs: Dict[str, str]) -> bool:
    """True only when this exact POS spec is defined for the branch.

    Branch definition comes from warehouse anchors and from exact positive
    stock already present on the POS. The positive-stock fallback prevents an
    incomplete ownership repair from making real quantities unsellable.
    """
    args = (specs["item_type"], specs["school"], specs["color"], specs["size"])
    checks = (
        (
            """
            SELECT 1
              FROM stocks
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND COALESCE(count, 0) > 0
             LIMIT 1
            """,
            args,
        ),
        (
            """
            SELECT 1
              FROM branch_catalog_definitions
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
             LIMIT 1
            """,
            args,
        ),
        (
            """
            SELECT 1
              FROM incoming_shipment_items_pending
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND (COALESCE(expected_qty, 0) > 0 OR COALESCE(received_qty, 0) > 0)
               AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
             LIMIT 1
            """,
            args,
        ),
        (
            """
            SELECT 1
              FROM stock_audit_report_lines
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
             LIMIT 1
            """,
            args,
        ),
    )
    for sql, sql_args in checks:
        try:
            if conn.execute(sql, sql_args).fetchone():
                return True
        except sqlite3.OperationalError:
            pass

    try:
        inbox_rows = conn.execute(
            """
            SELECT event_type, payload_json
              FROM sync_inbox
             WHERE event_type IN ('STOCK_TRANSFER_OUT', 'BRANCH_STOCK_RECLASSIFIED')
            """
        ).fetchall()
    except sqlite3.OperationalError:
        inbox_rows = []
    cancelled_shipments: Set[str] = set()
    try:
        cancel_rows = conn.execute(
            """
            SELECT payload_json
              FROM sync_inbox
             WHERE event_type = 'STOCK_TRANSFER_CANCELLED'
            """
        ).fetchall()
    except sqlite3.OperationalError:
        cancel_rows = []
    for row in cancel_rows:
        try:
            payload = json.loads(row[0] or "{}")
        except Exception:
            continue
        shipment_uuid = _clean(payload.get("shipment_uuid") or payload.get("bill_uuid"))
        if shipment_uuid:
            cancelled_shipments.add(shipment_uuid.casefold())
    for row in inbox_rows:
        try:
            event_type = str(row["event_type"] if isinstance(row, sqlite3.Row) else row[0] or "")
            payload_json = row["payload_json"] if isinstance(row, sqlite3.Row) else row[1]
            payload = json.loads(payload_json or "{}")
        except Exception:
            continue
        if event_type == "BRANCH_STOCK_RECLASSIFIED":
            to_spec = payload.get("to_spec") or {}
            if isinstance(to_spec, dict) and all(
                _clean(to_spec.get(k)).casefold() == specs[k].casefold()
                for k in ("item_type", "school", "color", "size")
            ):
                return True
            continue
        shipment_uuid = _clean(payload.get("shipment_uuid"))
        if shipment_uuid and shipment_uuid.casefold() in cancelled_shipments:
            continue
        note_text = _clean(payload.get("note")).casefold()
        is_reservation_definition = "reservation" in note_text or "حجز" in note_text
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            try:
                qty = int(float(item.get("qty") or 0))
            except (TypeError, ValueError):
                qty = 0
            catalog_only = bool(item.get("catalog_only"))
            owns = qty > 0 or (catalog_only and is_reservation_definition)
            if not owns:
                continue
            if all(_clean(item.get(k)).casefold() == specs[k].casefold() for k in ("item_type", "school", "color", "size")):
                return True
    return False


def _exact_specs_for_filter(conn: sqlite3.Connection, filt: Dict[str, str]) -> List[Dict[str, str]]:
    where = []
    args: List[Any] = []
    for field in ("item_type", "school", "color", "size"):
        value = _clean(filt.get(field))
        if value:
            where.append(f"LOWER(TRIM(COALESCE({field}, ''))) = LOWER(TRIM(?))")
            args.append(value)
    if not where:
        return []
    specs: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}
    for table in ("stocks", "branch_catalog_definitions", "incoming_shipment_items_pending", "stock_audit_report_lines"):
        try:
            rows = conn.execute(
                f"""
                SELECT item_type, school, color, size
                  FROM {table}
                 WHERE {" AND ".join(where)}
                   AND COALESCE(TRIM(item_type), '') <> ''
                   AND COALESCE(TRIM(school), '') <> ''
                   AND COALESCE(TRIM(color), '') <> ''
                   AND COALESCE(TRIM(size), '') <> ''
                """,
                args,
            ).fetchall()
        except sqlite3.OperationalError:
            rows = []
        for row in rows:
            spec = {
                "item_type": _clean(row["item_type"] if isinstance(row, sqlite3.Row) else row[0]),
                "school": _clean(row["school"] if isinstance(row, sqlite3.Row) else row[1]),
                "color": _clean(row["color"] if isinstance(row, sqlite3.Row) else row[2]),
                "size": _clean(row["size"] if isinstance(row, sqlite3.Row) else row[3]),
            }
            key = tuple(spec[k].casefold() for k in ("item_type", "school", "color", "size"))
            specs[key] = spec
    if all(_clean(filt.get(k)) for k in ("item_type", "school", "color", "size")):
        spec = {k: _clean(filt.get(k)) for k in ("item_type", "school", "color", "size")}
        key = tuple(spec[k].casefold() for k in ("item_type", "school", "color", "size"))
        specs.setdefault(key, spec)
    return list(specs.values())


def _filter_has_protected_branch_anchor(conn: sqlite3.Connection, filt: Dict[str, str]) -> bool:
    return any(_spec_is_branch_defined(conn, spec) for spec in _exact_specs_for_filter(conn, filt))


def _spec_has_positive_stock(conn: sqlite3.Connection, specs: Dict[str, str]) -> bool:
    try:
        return bool(conn.execute(
            """
            SELECT 1
              FROM stocks
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND COALESCE(count, 0) > 0
             LIMIT 1
            """,
            (specs["item_type"], specs["school"], specs["color"], specs["size"]),
        ).fetchone())
    except sqlite3.OperationalError:
        return False


def _spec_has_hard_branch_anchor(conn: sqlite3.Connection, specs: Dict[str, str]) -> bool:
    """True for anchors that manual cleanup/delete-definition must not hide."""
    args = (specs["item_type"], specs["school"], specs["color"], specs["size"])
    checks = (
        (
            """
            SELECT 1
              FROM stocks
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND COALESCE(count, 0) > 0
             LIMIT 1
            """,
            args,
        ),
        (
            """
            SELECT 1
              FROM incoming_shipment_items_pending
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND (COALESCE(expected_qty, 0) > 0 OR COALESCE(received_qty, 0) > 0)
               AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
             LIMIT 1
            """,
            args,
        ),
        (
            """
            SELECT 1
              FROM stock_audit_report_lines
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
             LIMIT 1
            """,
            args,
        ),
        (
            """
            SELECT 1
              FROM branch_catalog_definitions
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND (
                    LOWER(COALESCE(note, '')) LIKE '%reservation%'
                    OR COALESCE(note, '') LIKE '%حجز%'
                    OR LOWER(COALESCE(note, '')) LIKE '%reclassification%'
               )
             LIMIT 1
            """,
            args,
        ),
    )
    for sql, sql_args in checks:
        try:
            if conn.execute(sql, sql_args).fetchone():
                return True
        except sqlite3.OperationalError:
            pass
    try:
        rows = conn.execute(
            """
            SELECT event_type, payload_json
              FROM sync_inbox
             WHERE event_type IN ('STOCK_TRANSFER_OUT', 'BRANCH_STOCK_RECLASSIFIED')
            """
        ).fetchall()
    except sqlite3.OperationalError:
        rows = []
    for row in rows:
        try:
            event_type = str(row["event_type"] if isinstance(row, sqlite3.Row) else row[0] or "")
            payload_json = row["payload_json"] if isinstance(row, sqlite3.Row) else row[1]
            payload = json.loads(payload_json or "{}")
        except Exception:
            continue
        candidates: List[Dict[str, Any]] = []
        if event_type == "STOCK_TRANSFER_OUT":
            note_text = _clean(payload.get("note")).casefold()
            is_reservation_definition = "reservation" in note_text or "حجز" in note_text
            for item in payload.get("items") or []:
                if not isinstance(item, dict):
                    continue
                try:
                    qty = int(float(item.get("qty") or 0))
                except (TypeError, ValueError):
                    qty = 0
                if qty > 0 or (bool(item.get("catalog_only")) and is_reservation_definition):
                    candidates.append(item)
        elif event_type == "BRANCH_STOCK_RECLASSIFIED":
            to_spec = payload.get("to_spec") or {}
            if isinstance(to_spec, dict):
                candidates.append(to_spec)
            catalog_rows = payload.get("catalog_rows") or []
            if isinstance(catalog_rows, list):
                candidates.extend([r for r in catalog_rows if isinstance(r, dict)])
        for raw in candidates:
            if all(_clean(raw.get(k)).casefold() == specs[k].casefold() for k in ("item_type", "school", "color", "size")):
                return True
    return False


def _price_update_spec_is_branch_owned(conn: sqlite3.Connection, specs: Dict[str, str]) -> bool:
    return _spec_is_branch_defined(conn, specs)


def _upsert_owned_price_catalog_row(
    conn: sqlite3.Connection,
    specs: Dict[str, str],
    unit_price: float,
) -> bool:
    if not _price_update_spec_is_branch_owned(conn, specs):
        return False
    _upsert_spec_history(conn, specs)
    _update_item_default_price(conn, specs["item_type"], float(unit_price))
    conn.execute(
        """
        UPDATE branch_catalog_definitions
           SET unit_price = ?
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
        """,
        (
            float(unit_price),
            specs["item_type"],
            specs["school"],
            specs["color"],
            specs["size"],
        ),
    )
    existing = conn.execute(
        """
        SELECT id
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
         ORDER BY id ASC
         LIMIT 1
        """,
        (specs["item_type"], specs["school"], specs["color"], specs["size"]),
    ).fetchone()
    if existing:
        return False
    conn.execute(
        """
        INSERT INTO stocks(item_type, school, color, size, unit_price, count)
        VALUES (?, ?, ?, ?, ?, 0)
        """,
        (
            specs["item_type"],
            specs["school"],
            specs["color"],
            specs["size"],
            float(unit_price),
        ),
    )
    return True


def _upsert_pos_stock_snapshot_meta(
    conn: sqlite3.Connection,
    source_name: str,
    snapshot_at: str,
    inserted: int,
    total_value: float,
    app_version: str = "",
) -> None:
    if _has_column(conn, "pos_stocks_snapshot_meta", "app_version"):
        cur = conn.execute(
            """
            UPDATE pos_stocks_snapshot_meta
               SET snapshot_at = ?,
                   row_count   = ?,
                   total_value = ?,
                   app_version = ?
             WHERE source_device = ?
            """,
            (snapshot_at, inserted, total_value, app_version, source_name),
        )
        if cur.rowcount == 0:
            conn.execute(
                """
                INSERT OR IGNORE INTO pos_stocks_snapshot_meta
                    (source_device, snapshot_at, row_count, total_value, app_version)
                VALUES (?, ?, ?, ?, ?)
                """,
                (source_name, snapshot_at, inserted, total_value, app_version),
            )
        return
    cur = conn.execute(
        """
        UPDATE pos_stocks_snapshot_meta
           SET snapshot_at = ?,
               row_count   = ?,
               total_value = ?
         WHERE source_device = ?
        """,
        (snapshot_at, inserted, total_value, source_name),
    )
    if cur.rowcount == 0:
        conn.execute(
            """
            INSERT OR IGNORE INTO pos_stocks_snapshot_meta
                (source_device, snapshot_at, row_count, total_value)
            VALUES (?, ?, ?, ?)
            """,
            (source_name, snapshot_at, inserted, total_value),
        )


def _delete_old_spec_history_if_unused(
    conn: sqlite3.Connection,
    field: str,
    old_value: str,
) -> None:
    if field not in ("item_type", "school", "color", "size"):
        return
    value = _clean(old_value)
    if not value:
        return
    row = conn.execute(
        f"""
        SELECT 1
          FROM stocks
         WHERE LOWER(TRIM({field})) = LOWER(TRIM(?))
         LIMIT 1
        """,
        (value,),
    ).fetchone()
    if not row:
        row = conn.execute(
            f"""
            SELECT 1
              FROM bill_items
             WHERE LOWER(TRIM({field})) = LOWER(TRIM(?))
             LIMIT 1
            """,
            (value,),
        ).fetchone()
    if not row:
        conn.execute(
            "DELETE FROM spec_history WHERE field = ? AND LOWER(TRIM(value)) = LOWER(TRIM(?))",
            (field, value),
        )


def _apply_spec_value_renames(conn: sqlite3.Connection, payload: Dict[str, Any]) -> int:
    """Apply broad value-level spec renames carried with SPEC_RENAMED events."""
    if not bool(payload.get("allow_global_value_renames")):
        return 0
    fields = {"item_type", "school", "color", "size"}
    renames = payload.get("value_renames") or []
    if not isinstance(renames, list):
        return 0
    updated = 0
    for raw in renames:
        if not isinstance(raw, dict):
            continue
        field = _clean(raw.get("field"))
        old_value = _clean(raw.get("old_value"))
        new_value = _clean(raw.get("new_value"))
        if field not in fields or not old_value or not new_value or old_value == new_value:
            continue
        for table in (
            "stocks",
            "movements",
            "bill_items",
            "reservations",
            "reservation_alerts",
            "incoming_shipment_items_pending",
            "branch_catalog_definitions",
            "stock_audit_report_lines",
            "pos_stocks_mirror",
        ):
            try:
                cur = conn.execute(
                    f"UPDATE {table} SET {field} = ? WHERE LOWER(TRIM({field})) = LOWER(TRIM(?))",
                    (new_value, old_value),
                )
                updated += int(cur.rowcount or 0)
            except sqlite3.OperationalError:
                pass
        try:
            conn.execute(
                "INSERT OR IGNORE INTO spec_history(field,value) VALUES(?,?)",
                (field, new_value),
            )
            _delete_old_spec_history_if_unused(conn, field, old_value)
        except sqlite3.OperationalError:
            pass
    return updated


def apply_spec_renamed(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    old_spec_raw = payload.get("old_spec") or {}
    new_spec_raw = payload.get("new_spec") or {}
    if not isinstance(old_spec_raw, dict) or not isinstance(new_spec_raw, dict):
        raise ApplyError("SPEC_RENAMED payload must include old_spec and new_spec")

    fields = ("item_type", "school", "color", "size")
    old_spec = {fld: _clean(old_spec_raw.get(fld)) for fld in fields}
    new_spec = {fld: _clean(new_spec_raw.get(fld)) for fld in fields}
    if not all(old_spec.values()) or not all(new_spec.values()):
        raise ApplyError("SPEC_RENAMED requires complete old/new specs")

    if not _spec_is_branch_defined(conn, old_spec):
        return {
            "skipped": True,
            "reason": "spec is not defined for this branch",
            "old_spec": old_spec,
            "new_spec": new_spec,
            "event_uuid": str(event_uuid),
        }

    if old_spec == new_spec:
        return {"skipped": True, "reason": "old and new specs are identical"}

    set_sql = ", ".join(f"{fld}=?" for fld in fields)
    set_args = tuple(new_spec[fld] for fld in fields)
    where_sql = " AND ".join(f"LOWER(TRIM({fld})) = LOWER(TRIM(?))" for fld in fields)
    where_args = tuple(old_spec[fld] for fld in fields)

    updated_tables = 0
    for table in (
        "stocks",
        "movements",
        "bill_items",
        "reservations",
        "reservation_alerts",
        "incoming_shipment_items_pending",
        "branch_catalog_definitions",
        "stock_audit_report_lines",
    ):
        try:
            cur = conn.execute(
                f"UPDATE {table} SET {set_sql} WHERE {where_sql}",
                (*set_args, *where_args),
            )
            updated_tables += int(cur.rowcount or 0)
        except sqlite3.OperationalError:
            pass

    old_item = old_spec["item_type"]
    new_item = new_spec["item_type"]
    if old_item != new_item:
        try:
            row = conn.execute(
                "SELECT default_price FROM item_defaults WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))",
                (old_item,),
            ).fetchone()
            if row and row[0] is not None:
                conn.execute(
                    "INSERT OR IGNORE INTO item_defaults(item_type, default_price) VALUES(?, ?)",
                    (new_item, float(row[0])),
                )
                conn.execute(
                    "DELETE FROM item_defaults WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))",
                    (old_item,),
                )
        except sqlite3.OperationalError:
            pass

    old_prof_key = (old_spec["item_type"], old_spec["school"], old_spec["color"])
    new_prof_key = (new_spec["item_type"], new_spec["school"], new_spec["color"])
    if old_prof_key != new_prof_key:
        try:
            conn.execute(
                """
                INSERT OR IGNORE INTO size_profiles
                    (item_type, school, color,
                     num_start_1, num_end_1, num_start_2, num_end_2,
                     has_alpha, updated_at)
                SELECT ?, ?, ?,
                       num_start_1, num_end_1, num_start_2, num_end_2,
                       has_alpha, datetime('now')
                  FROM size_profiles
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                """,
                (*new_prof_key, *old_prof_key),
            )
            conn.execute(
                """
                DELETE FROM size_profiles
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                """,
                old_prof_key,
            )
        except sqlite3.OperationalError:
            pass

    _upsert_spec_history(conn, new_spec)
    for fld in fields:
        if old_spec[fld] != new_spec[fld]:
            _delete_old_spec_history_if_unused(conn, fld, old_spec[fld])

    return {
        "updated_rows": int(updated_tables),
        "old_spec": old_spec,
        "new_spec": new_spec,
        "event_uuid": str(event_uuid),
    }


def apply_stock_transfer_out(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Queue a warehouse → POS shipment for cashier verification.

    Payload shape (see create_branch_shipment on the warehouse side):
        {
          "shipment_uuid": "...",
          "from_device":   "WAREHOUSE-MAIN",
          "note":          "...",
          "items": [
            {"item_type":"...", "school":"...", "color":"...",
             "size":"...", "unit_price": 150.0, "qty": 5}
          ],
          "size_profiles": [
            {"item_type":"...", "school":"...", "color":"...",
             "num_start_1": ..., "num_end_1": ..., ...}
          ]
        }

    Behaviour:
        - Stores shipment lines into `incoming_shipment_items_pending`
          (no stock mutation yet).
        - Creates one `incoming_shipment_alerts` row so the POS UI can
          open an item-by-item checklist for the cashier.
        - Auto-heals `spec_history` and `item_defaults` and applies
          `size_profiles` immediately (safe metadata updates).
    """
    items = payload.get("items") or []
    if not isinstance(items, list) or not items:
        raise ApplyError("shipment has no items")

    shipment_uuid = _clean(payload.get("shipment_uuid"))
    from_dev = _clean(payload.get("from_device")) or "warehouse"
    short_ship = (shipment_uuid[:8] if shipment_uuid else event_uuid[:8])
    note_base = f"شحنة من {from_dev} #{short_ship}"

    duplicate = False
    # Idempotency guard by shipment UUID. Even when the shipment was seen by
    # an older build, keep walking the payload so missing pending/catalog rows
    # can be healed from the bill-shipment anchor.
    if shipment_uuid:
        dup = conn.execute(
            """
            SELECT 1
              FROM incoming_shipment_alerts
             WHERE shipment_uuid = ?
            UNION ALL
            SELECT 1
              FROM incoming_shipment_items_pending
             WHERE shipment_uuid = ?
            LIMIT 1
            """,
            (shipment_uuid, shipment_uuid),
        ).fetchone()
        if dup:
            duplicate = True

    queued_rows = 0
    total_qty = 0
    catalog_rows = 0
    pending_line_index = 0

    for item in items:
        it = _clean(item.get("item_type"))
        sc = _clean(item.get("school"))
        cl = _clean(item.get("color"))
        sz = _clean(item.get("size"))
        try:
            qty = int(item.get("qty") or 0)
        except (TypeError, ValueError):
            raise ApplyError(f"invalid qty in shipment line: {item}")
        try:
            price = float(item.get("unit_price") or 0)
        except (TypeError, ValueError):
            raise ApplyError(f"invalid unit_price in shipment line: {item}")

        if not (it and sc and cl and sz):
            raise ApplyError(f"incomplete shipment line specs: {item}")
        _upsert_spec_history(conn, {"item_type": it, "school": sc, "color": cl, "size": sz})
        _ensure_item_default_price(conn, it, price)
        catalog_note = _clean(payload.get("note")) or "Catalog-only sync for POS reservations"
        is_reservation_definition = bool(item.get("catalog_only")) and _is_reservation_catalog_note(catalog_note)

        if qty <= 0:
            if _upsert_zero_stock_catalog_row(
                conn,
                it,
                sc,
                cl,
                sz,
                price,
                respect_tombstone=True,
            ):
                catalog_rows += 1
            if is_reservation_definition and not _branch_catalog_spec_is_deleted(
                conn,
                {"item_type": it, "school": sc, "color": cl, "size": sz},
            ):
                _mark_branch_catalog_definition(
                    conn,
                    it,
                    sc,
                    cl,
                    sz,
                    price,
                    event_uuid,
                    note=catalog_note,
                )
            continue

        if _upsert_zero_stock_catalog_row(conn, it, sc, cl, sz, price, respect_tombstone=False):
            catalog_rows += 1

        conn.execute(
            """
            INSERT OR IGNORE INTO incoming_shipment_items_pending(
                shipment_uuid, line_index, item_type, school, color, size,
                unit_price, expected_qty, received_qty, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, 'PENDING')
            """,
            (shipment_uuid or event_uuid, pending_line_index, it, sc, cl, sz, price, qty),
        )
        pending_line_index += 1
        queued_rows += 1
        total_qty += qty

    profile_rows = _apply_size_profile_rows(conn, payload.get("size_profiles"))
    if queued_rows > 0:
        conn.execute(
            """
            INSERT OR IGNORE INTO incoming_shipment_alerts(
                sync_event_uuid, shipment_uuid, from_device, note, total_qty, created_at, shown_at
            ) VALUES (?, ?, ?, ?, ?, ?, NULL)
            """,
            (
                event_uuid,
                shipment_uuid or event_uuid,
                from_dev,
                note_base,
                total_qty,
                _now_iso(),
            ),
        )
        return {
            "queued_rows": queued_rows,
            "auto_received_rows": 0,
            "catalog_rows": catalog_rows,
            "total_qty": total_qty,
            "shipment": shipment_uuid or None,
            "needs_verification": True,
            "size_profiles_applied": profile_rows,
            "duplicate": duplicate,
        }

    return {
        "queued_rows": queued_rows,
        "auto_received_rows": 0,
        "catalog_rows": catalog_rows,
        "total_qty":  total_qty,
        "shipment":   shipment_uuid or None,
        "needs_verification": queued_rows > 0,
        "size_profiles_applied": profile_rows,
        "duplicate": duplicate,
    }


def _consume_stock_for_shipment_cancel(
    conn: sqlite3.Connection,
    item_type: str,
    school: str,
    color: str,
    size: str,
    unit_price: float,
    qty: int,
    note: str,
    direction: str = "SHIPMENT_CANCEL",
) -> Tuple[int, int]:
    remaining = max(0, int(qty or 0))
    removed = 0
    if remaining <= 0:
        return 0, 0
    rows = conn.execute(
        """
        SELECT id, COALESCE(count, 0) AS count, unit_price
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
           AND COALESCE(count, 0) > 0
         ORDER BY id DESC
        """,
        (item_type, school, color, size),
    ).fetchall()
    for row in rows:
        if remaining <= 0:
            break
        stock_id = int(row["id"])
        available = int(row["count"] or 0)
        take = min(available, remaining)
        if take <= 0:
            continue
        conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?", (take, stock_id))
        conn.execute(
            """
            INSERT INTO movements
                (ts, direction, stock_id, qty, note, bill_id,
                 item_type, school, color, size, unit_price)
            VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?, ?, ?)
            """,
            (
                _now_iso(),
                str(direction or "SHIPMENT_CANCEL"),
                stock_id,
                take,
                note,
                item_type,
                school,
                color,
                size,
                float(row["unit_price"] if row["unit_price"] is not None else unit_price),
            ),
        )
        removed += take
        remaining -= take
    return removed, remaining


def _delete_unanchored_cancelled_zero_rows(
    conn: sqlite3.Connection,
    specs: Set[Tuple[str, str, str, str]],
) -> int:
    deleted = 0
    for it, sc, cl, sz in specs:
        has_catalog = False
        try:
            has_catalog = bool(conn.execute(
                """
                SELECT 1
                  FROM branch_catalog_definitions
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                 LIMIT 1
                """,
                (it, sc, cl, sz),
            ).fetchone())
        except sqlite3.OperationalError:
            has_catalog = False
        if has_catalog:
            continue
        has_active_shipment = bool(conn.execute(
            """
            SELECT 1
              FROM incoming_shipment_items_pending
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
               AND (COALESCE(expected_qty, 0) > 0 OR COALESCE(received_qty, 0) > 0)
             LIMIT 1
            """,
            (it, sc, cl, sz),
        ).fetchone())
        if has_active_shipment:
            continue
        cur = conn.execute(
            """
            DELETE FROM stocks
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
               AND COALESCE(count, 0) <= 0
            """,
            (it, sc, cl, sz),
        )
        deleted += int(cur.rowcount or 0)
    return deleted


def apply_stock_transfer_cancelled(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Cancel a warehouse -> POS shipment without corrupting counts.

    Pending shipments are removed from the cashier checklist. Confirmed
    shipments are reversed only when this POS still has rows for that exact
    shipment UUID. If the original shipment is unknown locally, do not fall
    back to subtracting by item spec because that can consume a later resent
    bill for the same items.
    """
    shipment_uuid = _clean(payload.get("shipment_uuid") or payload.get("bill_uuid"))
    if not shipment_uuid:
        raise ApplyError("STOCK_TRANSFER_CANCELLED requires shipment_uuid")

    rows = conn.execute(
        """
        SELECT *
          FROM incoming_shipment_items_pending
         WHERE shipment_uuid = ?
         ORDER BY line_index ASC
        """,
        (shipment_uuid,),
    ).fetchall()
    if rows and all(str(r["status"] or "").upper() == "CANCELLED" for r in rows):
        return {"shipment": shipment_uuid, "already_cancelled": True}

    removed_qty = 0
    shortage_qty = 0
    pending_rows = 0
    confirmed_rows = 0
    note = _clean(payload.get("note")) or f"Shipment cancelled #{shipment_uuid[:8]}"
    cancel_specs: Set[Tuple[str, str, str, str]] = set()

    for row in rows:
        it = _clean(row["item_type"])
        sc = _clean(row["school"])
        cl = _clean(row["color"])
        sz = _clean(row["size"])
        if it and sc and cl and sz:
            cancel_specs.add((it, sc, cl, sz))
        status = str(row["status"] or "").upper()
        if status == "PENDING":
            pending_rows += 1
            continue
        received = row["received_qty"]
        try:
            qty_to_reverse = int(received if received is not None else row["expected_qty"] or 0)
        except (TypeError, ValueError):
            qty_to_reverse = 0
        if qty_to_reverse <= 0:
            continue
        confirmed_rows += 1
        removed, shortage = _consume_stock_for_shipment_cancel(
            conn,
            it,
            sc,
            cl,
            sz,
            float(row["unit_price"] or 0.0),
            qty_to_reverse,
            note,
        )
        removed_qty += removed
        shortage_qty += shortage

    if not rows:
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            it = _clean(item.get("item_type"))
            sc = _clean(item.get("school"))
            cl = _clean(item.get("color"))
            sz = _clean(item.get("size"))
            if not (it and sc and cl and sz):
                continue
            cancel_specs.add((it, sc, cl, sz))
            try:
                qty = int(item.get("received_qty") if item.get("received_qty") is not None else item.get("qty") or 0)
            except (TypeError, ValueError):
                qty = 0
            # Unknown local shipment: never subtract by spec-only payload.
            # A resent branch bill may already own the same specs/quantities.
            if qty > 0:
                shortage_qty += qty

    conn.execute(
        """
        UPDATE incoming_shipment_items_pending
           SET status = 'CANCELLED',
               received_qty = COALESCE(received_qty, 0)
         WHERE shipment_uuid = ?
        """,
        (shipment_uuid,),
    )
    conn.execute(
        """
        UPDATE incoming_shipment_alerts
           SET shown_at = COALESCE(shown_at, ?),
               note = CASE
                    WHEN COALESCE(note, '') = '' THEN ?
                    ELSE note || ' | ' || ?
               END
         WHERE shipment_uuid = ?
        """,
        (_now_iso(), note, note, shipment_uuid),
    )
    deleted_zero_rows = _delete_unanchored_cancelled_zero_rows(conn, cancel_specs)
    return {
        "shipment": shipment_uuid,
        "pending_rows_cancelled": pending_rows,
        "confirmed_rows_cancelled": confirmed_rows,
        "removed_qty": removed_qty,
        "shortage_qty": shortage_qty,
        "quantity_mode": "exact_shipment_only",
        "deleted_zero_rows": deleted_zero_rows,
    }


def _stock_sum_for_spec(conn: sqlite3.Connection, spec: Dict[str, str]) -> int:
    row = conn.execute(
        """
        SELECT COALESCE(SUM(count), 0)
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
        """,
        (spec["item_type"], spec["school"], spec["color"], spec["size"]),
    ).fetchone()
    return int((row[0] if row else 0) or 0)


def _add_reclassified_stock(
    conn: sqlite3.Connection,
    spec: Dict[str, str],
    qty: int,
    unit_price: float,
    note: str,
) -> int:
    row = conn.execute(
        """
        SELECT id
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
           AND ABS(COALESCE(unit_price, 0) - ?) < 0.001
         ORDER BY id ASC
         LIMIT 1
        """,
        (spec["item_type"], spec["school"], spec["color"], spec["size"], float(unit_price)),
    ).fetchone()
    if row:
        stock_id = int(row["id"] if isinstance(row, sqlite3.Row) else row[0])
        conn.execute("UPDATE stocks SET count = COALESCE(count, 0) + ?, unit_price = ? WHERE id = ?", (int(qty), float(unit_price), stock_id))
    else:
        cur = conn.execute(
            """
            INSERT INTO stocks(item_type, school, color, size, unit_price, count)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (spec["item_type"], spec["school"], spec["color"], spec["size"], float(unit_price), int(qty)),
        )
        stock_id = int(cur.lastrowid)
    conn.execute(
        """
        INSERT INTO movements
            (ts, direction, stock_id, qty, note, bill_id,
             item_type, school, color, size, unit_price)
        VALUES (?, 'RECLASS_IN', ?, ?, ?, NULL, ?, ?, ?, ?, ?)
        """,
        (
            _now_iso(),
            stock_id,
            int(qty),
            note,
            spec["item_type"],
            spec["school"],
            spec["color"],
            spec["size"],
            float(unit_price),
        ),
    )
    return stock_id


def apply_branch_stock_reclassified(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    from_raw = payload.get("from_spec") or {}
    to_raw = payload.get("to_spec") or {}
    if not isinstance(from_raw, dict) or not isinstance(to_raw, dict):
        raise ApplyError("BRANCH_STOCK_RECLASSIFIED requires from_spec/to_spec")

    def _spec(raw: Dict[str, Any]) -> Dict[str, str]:
        spec = {
            "item_type": _clean(raw.get("item_type")),
            "school": _clean(raw.get("school")),
            "color": _clean(raw.get("color")),
            "size": _clean(raw.get("size")),
        }
        if not all(spec.values()):
            raise ApplyError("BRANCH_STOCK_RECLASSIFIED has incomplete spec")
        return spec

    src = _spec(from_raw)
    dst = _spec(to_raw)
    if src == dst:
        return {"skipped": True, "reason": "source and target specs are identical"}
    try:
        qty = int(payload.get("qty") or 0)
    except (TypeError, ValueError):
        raise ApplyError("BRANCH_STOCK_RECLASSIFIED invalid qty")
    if qty <= 0:
        raise ApplyError("BRANCH_STOCK_RECLASSIFIED qty must be positive")
    try:
        target_price = float(to_raw.get("unit_price") or from_raw.get("unit_price") or 0.0)
    except (TypeError, ValueError):
        target_price = 0.0
    profile_rows = _apply_size_profile_rows(conn, payload.get("size_profiles"))
    catalog_rows = _apply_branch_catalog_rows(
        conn,
        payload.get("catalog_rows"),
        event_uuid,
        note=_clean(payload.get("note")) or "Branch stock reclassification",
    )
    already = conn.execute(
        """
        SELECT 1
          FROM movements
         WHERE direction = 'RECLASS_IN'
           AND note LIKE ?
         LIMIT 1
        """,
        ("%" + str(event_uuid) + "%",),
    ).fetchone()
    if already:
        return {
            "already_applied": True,
            "event_uuid": str(event_uuid),
            "size_profiles_applied": profile_rows,
            "catalog_rows": catalog_rows,
        }
    available = _stock_sum_for_spec(conn, src)
    if available < qty:
        target_existing = _stock_sum_for_spec(conn, dst)
        missing_qty = max(0, qty - target_existing)
        if missing_qty <= 0:
            return {
                "repaired": True,
                "reason": "source missing but target quantity already exists",
                "available": available,
                "requested": qty,
                "target_existing": target_existing,
                "size_profiles_applied": profile_rows,
                "catalog_rows": catalog_rows,
            }
        note = _clean(payload.get("note")) or "Branch stock reclassified"
        note = f"{note} repair #{event_uuid}"
        _upsert_spec_history(conn, dst)
        _ensure_item_default_price(conn, dst["item_type"], target_price)
        stock_id = _add_reclassified_stock(conn, dst, missing_qty, target_price, note)
        _mark_branch_catalog_definition(
            conn,
            dst["item_type"],
            dst["school"],
            dst["color"],
            dst["size"],
            target_price,
            event_uuid,
            note="Branch stock reclassification",
        )
        return {
            "repaired": True,
            "reason": "source quantity missing; restored target quantity",
            "from_spec": src,
            "to_spec": dst,
            "qty": missing_qty,
            "requested": qty,
            "available": available,
            "target_existing_before": target_existing,
            "target_stock_id": stock_id,
            "target_became_branch_defined": True,
            "size_profiles_applied": profile_rows,
            "catalog_rows": catalog_rows,
        }

    note = _clean(payload.get("note")) or "Branch stock reclassified"
    note = f"{note} #{event_uuid}"
    removed, shortage = _consume_stock_for_shipment_cancel(
        conn,
        src["item_type"],
        src["school"],
        src["color"],
        src["size"],
        float(from_raw.get("unit_price") or 0.0),
        qty,
        note,
        direction="RECLASS_OUT",
    )
    if shortage or removed != qty:
        raise ApplyError("BRANCH_STOCK_RECLASSIFIED could not subtract full source quantity")

    _upsert_spec_history(conn, dst)
    _ensure_item_default_price(conn, dst["item_type"], target_price)
    stock_id = _add_reclassified_stock(conn, dst, qty, target_price, note)
    _mark_branch_catalog_definition(
        conn,
        dst["item_type"],
        dst["school"],
        dst["color"],
        dst["size"],
        target_price,
        event_uuid,
        note="Branch stock reclassification",
    )
    return {
        "from_spec": src,
        "to_spec": dst,
        "qty": qty,
        "target_stock_id": stock_id,
        "target_became_branch_defined": True,
        "size_profiles_applied": profile_rows,
        "catalog_rows": catalog_rows,
    }


def apply_price_update(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Apply a warehouse-issued price update to this POS's stock.

    Only the `{new_price, filters}` shape is applied. The older
    `{mode, value, constraints}` bulk-percentage shape is recorded as
    skipped (no determinism without knowing the warehouse's starting
    prices). The warehouse app now emits the filter shape, so this
    covers real traffic.

    Rules:
        - `filters` must include at least one non-empty spec field
          (item_type / school / color / size). A naked update that
          would rewrite the entire catalogue is rejected on purpose.
        - Applies via UPDATE + one PRICE_UPDATE movement per affected
          row, so the audit trail matches a locally-issued update.
    """
    new_price = payload.get("new_price")
    if new_price is None:
        return {"skipped": True, "reason": "unsupported payload shape (no new_price)"}
    try:
        new_price = float(new_price)
    except (TypeError, ValueError):
        raise ApplyError(f"invalid new_price: {new_price}")
    if new_price < 0:
        raise ApplyError("new_price must be >= 0")

    filters = payload.get("filters") or {}
    if not isinstance(filters, dict):
        raise ApplyError("filters must be an object")

    cleaned_filters = {k: _clean(filters.get(k)) for k in ("item_type", "school", "color", "size")}
    if not all(cleaned_filters.values()):
        return {"skipped": True, "reason": "refusing non-exact price update"}

    allow_catalog_definition = bool(
        payload.get("allow_catalog_definition")
        or payload.get("catalog_definition")
        or payload.get("catalog_only")
    )
    note_text = _clean(payload.get("note")) or "Warehouse price profile sync"
    if (
        allow_catalog_definition
        and _branch_catalog_spec_is_deleted(conn, cleaned_filters)
        and not _spec_has_positive_stock(conn, cleaned_filters)
    ):
        return {
            "skipped": True,
            "reason": "manual delete-definition tombstone",
            "updated": 0,
            "catalog_rows": 0,
        }
    if allow_catalog_definition:
        _mark_branch_catalog_definition(
            conn,
            cleaned_filters["item_type"],
            cleaned_filters["school"],
            cleaned_filters["color"],
            cleaned_filters["size"],
            float(new_price),
            event_uuid,
            note=note_text,
        )
        _upsert_spec_history(conn, cleaned_filters)

    if not allow_catalog_definition and not _price_update_spec_is_branch_owned(conn, cleaned_filters):
        return {
            "skipped": True,
            "reason": "spec is not defined for this branch",
            "updated": 0,
            "catalog_rows": 0,
        }

    where_parts = [f"LOWER(TRIM({k})) = LOWER(?)" for k in ("item_type", "school", "color", "size")]
    args = [cleaned_filters[k] for k in ("item_type", "school", "color", "size")]
    pending_updated = _update_pending_shipment_spec_price(conn, cleaned_filters, float(new_price))

    rows = conn.execute(
        "SELECT id, item_type, school, color, size FROM stocks WHERE "
        + " AND ".join(where_parts),
        args,
    ).fetchall()
    if not rows:
        catalog_inserted = False
        if all(cleaned_filters.values()):
            if allow_catalog_definition:
                catalog_inserted = _upsert_zero_stock_catalog_row(
                    conn,
                    cleaned_filters["item_type"],
                    cleaned_filters["school"],
                    cleaned_filters["color"],
                    cleaned_filters["size"],
                    float(new_price),
                )
                _update_item_default_price(conn, cleaned_filters["item_type"], float(new_price))
            else:
                catalog_inserted = _upsert_owned_price_catalog_row(conn, cleaned_filters, float(new_price))
        return {
            "updated": 0,
            "catalog_rows": 1 if catalog_inserted else 0,
            "pending_rows": pending_updated,
        }

    conn.execute(
        "UPDATE stocks SET unit_price = ? WHERE "
        + " AND ".join(where_parts),
        (new_price, *args),
    )
    try:
        conn.execute(
            "UPDATE branch_catalog_definitions SET unit_price = ? WHERE "
            + " AND ".join(where_parts),
            (new_price, *args),
        )
    except sqlite3.OperationalError:
        pass

    ts = _now_iso()
    note = f"تعديل سعر من المخزن (#{event_uuid[:8]})"
    for r in rows:
        conn.execute(
            """INSERT INTO movements
               (ts, direction, stock_id, qty, note, bill_id,
                item_type, school, color, size, unit_price)
               VALUES(?, 'PRICE_UPDATE', ?, 0, ?, NULL, ?, ?, ?, ?, ?)""",
            (
                ts, int(r[0]), note,
                r[1], r[2], r[3], r[4], new_price,
            ),
        )

    # Also refresh item_defaults so fresh catalog lookups get the new
    # price when the item has no stock left later.
    item_types = {r[1] for r in rows}
    if cleaned_filters.get("item_type"):
        item_types.add(cleaned_filters["item_type"])
    for it in item_types:
        _update_item_default_price(conn, it, new_price)

    return {"updated": len(rows), "catalog_rows": 0, "pending_rows": pending_updated}


def apply_pos_stock_snapshot(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Warehouse-side applier for a POS's full stock snapshot.

    Replaces all rows in `pos_stocks_mirror` for the snapshot's source
    device with the new set, and refreshes `pos_stocks_snapshot_meta`.
    Designed to be idempotent and safely replayable — apply the same
    snapshot twice and the mirror ends up in the same state.

    Payload shape (see sync_client.emit_stock_snapshot_event):
        {
          "source_device_name": "POS-01",
          "snapshot_at":        "2026-04-15T…",
          "rows": [
            {"item_type":"…","school":"…","color":"…","size":"…",
             "unit_price": 150.0, "count": 5},
            ...
          ]
        }
    """
    source_name = _clean(payload.get("source_device_name"))
    if not source_name:
        raise ApplyError("snapshot missing source_device_name")
    snapshot_at = _clean(payload.get("snapshot_at")) or _now_iso()
    app_version = _clean(payload.get("app_version"))
    rows = payload.get("rows") or []
    if not isinstance(rows, list):
        raise ApplyError("snapshot rows must be a list")

    inbound_seq = 0
    try:
        seq_row = conn.execute(
            "SELECT COALESCE(server_seq, 0) FROM sync_inbox WHERE event_uuid = ?",
            (event_uuid,),
        ).fetchone()
        inbound_seq = int((seq_row[0] if seq_row else 0) or 0)
    except (sqlite3.OperationalError, TypeError, ValueError):
        inbound_seq = 0
    try:
        if not _has_column(conn, "pos_stocks_snapshot_meta", "last_server_seq"):
            conn.execute("ALTER TABLE pos_stocks_snapshot_meta ADD COLUMN last_server_seq INTEGER NOT NULL DEFAULT 0")
    except sqlite3.OperationalError:
        pass

    current = conn.execute(
        "SELECT snapshot_at, COALESCE(last_server_seq, 0) AS last_server_seq FROM pos_stocks_snapshot_meta WHERE source_device = ?",
        (source_name,),
    ).fetchone()
    current_snapshot_at = _clean(current[0] if current else "")
    current_seq = 0
    try:
        current_seq = int((current["last_server_seq"] if isinstance(current, sqlite3.Row) else current[1]) if current else 0)
    except (TypeError, ValueError, IndexError):
        current_seq = 0
    if (
        inbound_seq
        and current_seq
        and inbound_seq < current_seq
        and _timestamp_is_older(snapshot_at, current_snapshot_at)
    ):
        return {
            "skipped": True,
            "reason": "stale stock snapshot",
            "current_server_seq": current_seq,
            "server_seq": inbound_seq,
            "current_snapshot_at": current_snapshot_at,
            "snapshot_at": snapshot_at,
        }
    if not inbound_seq and _timestamp_is_older(snapshot_at, current_snapshot_at):
        return {
            "skipped": True,
            "reason": "stale stock snapshot",
            "current_snapshot_at": current_snapshot_at,
            "snapshot_at": snapshot_at,
        }

    # Replace the whole mirror for this source.
    conn.execute(
        "DELETE FROM pos_stocks_mirror WHERE source_device = ?",
        (source_name,),
    )

    collapsed_rows: Dict[Tuple[str, str, str, str, float], Dict[str, Any]] = {}
    for r in rows:
        it = _clean(r.get("item_type"))
        sc = _clean(r.get("school"))
        cl = _clean(r.get("color"))
        sz = _clean(r.get("size"))
        try:
            price = float(r.get("unit_price") or 0)
            count = int(r.get("count") or 0)
        except (TypeError, ValueError):
            continue
        if not (it and sc and cl and sz) or count < 0:
            continue
        key = (it.casefold(), sc.casefold(), cl.casefold(), sz.casefold(), price)
        prev = collapsed_rows.get(key)
        if prev is None:
            collapsed_rows[key] = {
                "item_type": it,
                "school": sc,
                "color": cl,
                "size": sz,
                "unit_price": price,
                "count": count,
            }
        else:
            prev["count"] = int(prev.get("count") or 0) + count

    total_value = 0.0
    insert_rows = []
    for r in collapsed_rows.values():
        price = float(r["unit_price"] or 0)
        count = int(r["count"] or 0)
        insert_rows.append(
            (source_name, r["item_type"], r["school"], r["color"], r["size"], price, count, snapshot_at)
        )
        total_value += price * count
    if insert_rows:
        try:
            conn.executemany(
                """INSERT INTO pos_stocks_mirror
                   (source_device, item_type, school, color, size,
                    unit_price, count, snapshot_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                   ON CONFLICT(source_device, item_type, school, color, size, unit_price)
                   DO UPDATE SET
                       count = COALESCE(pos_stocks_mirror.count, 0) + COALESCE(excluded.count, 0),
                       snapshot_at = excluded.snapshot_at""",
                insert_rows,
            )
        except sqlite3.OperationalError as exc:
            if "ON CONFLICT clause does not match" not in str(exc):
                raise
            conn.executemany(
                """INSERT INTO pos_stocks_mirror
                   (source_device, item_type, school, color, size,
                    unit_price, count, snapshot_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                insert_rows,
            )
    inserted = len(insert_rows)

    _upsert_pos_stock_snapshot_meta(conn, source_name, snapshot_at, inserted, total_value, app_version)
    try:
        conn.execute(
            "UPDATE pos_stocks_snapshot_meta SET last_server_seq = ? WHERE source_device = ?",
            (inbound_seq, source_name),
        )
    except sqlite3.OperationalError:
        pass

    return {"mirrored_rows": inserted, "total_value": total_value}


def _refresh_pos_stock_snapshot_meta(
    conn: sqlite3.Connection,
    source_name: str,
    snapshot_at: str,
    *,
    update_snapshot_at: bool = True,
) -> None:
    row = conn.execute(
        """
        SELECT COUNT(*) AS row_count, COALESCE(SUM(count * unit_price), 0) AS total_value
          FROM pos_stocks_mirror
         WHERE source_device = ?
        """,
        (source_name,),
    ).fetchone()
    row_count = int((row[0] if row else 0) or 0)
    total_value = float((row[1] if row else 0.0) or 0.0)
    if update_snapshot_at:
        conn.execute(
            """
            UPDATE pos_stocks_snapshot_meta
               SET snapshot_at = ?,
                   row_count = ?,
                   total_value = ?
             WHERE source_device = ?
            """,
            (snapshot_at, row_count, total_value, source_name),
        )
        row = conn.execute(
            "SELECT 1 FROM pos_stocks_snapshot_meta WHERE source_device = ? LIMIT 1",
            (source_name,),
        ).fetchone()
        if row is None:
            conn.execute(
                """INSERT INTO pos_stocks_snapshot_meta
                       (source_device, snapshot_at, row_count, total_value)
                   VALUES (?, ?, ?, ?)""",
                (source_name, snapshot_at, row_count, total_value),
            )
    else:
        current = conn.execute(
            "SELECT snapshot_at FROM pos_stocks_snapshot_meta WHERE source_device = ?",
            (source_name,),
        ).fetchone()
        preserved_snapshot_at = _clean(current[0] if current else "") or snapshot_at
        conn.execute(
            """
            UPDATE pos_stocks_snapshot_meta
               SET row_count = ?,
                   total_value = ?
             WHERE source_device = ?
            """,
            (row_count, total_value, source_name),
        )
        row = conn.execute(
            "SELECT 1 FROM pos_stocks_snapshot_meta WHERE source_device = ? LIMIT 1",
            (source_name,),
        ).fetchone()
        if row is None:
            conn.execute(
                """INSERT INTO pos_stocks_snapshot_meta
                       (source_device, snapshot_at, row_count, total_value)
                   VALUES (?, ?, ?, ?)""",
                (source_name, preserved_snapshot_at, row_count, total_value),
            )


def apply_pos_stock_audit_applied(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Warehouse-side applier for a POS physical-count adjustment.

    The POS has already changed its local `stocks` rows. The warehouse
    mirror must therefore set the affected branch-stock rows to the
    counted `actual` quantity, not add the diff as a warehouse sale or
    warehouse shipment.
    """
    source_name = _clean(payload.get("source_device_name")) or _src_device(conn, payload, event_uuid)
    if not source_name:
        raise ApplyError("POS_STOCK_AUDIT_APPLIED: missing source device")
    audit_uuid = _clean(payload.get("audit_uuid")) or event_uuid
    if conn.execute(
        "SELECT 1 FROM pos_stock_audit_reports_mirror WHERE audit_uuid = ? LIMIT 1",
        (audit_uuid,),
    ).fetchone():
        return {"already_applied": True}

    lines = payload.get("lines") or []
    if not isinstance(lines, list) or not lines:
        raise ApplyError("POS_STOCK_AUDIT_APPLIED: missing lines")

    created_at = _clean(payload.get("created_at")) or _now_iso()
    reason = _clean(payload.get("reason"))
    try:
        report_id = int(payload.get("report_id")) if payload.get("report_id") is not None else None
    except (TypeError, ValueError):
        report_id = None
    total_diff = int(payload.get("total_diff") or 0)
    total_value = float(payload.get("total_value") or 0)

    conn.execute(
        """
        INSERT INTO pos_stock_audit_reports_mirror
            (audit_uuid, source_device, local_report_id, reason, created_at,
             total_diff, total_value, event_uuid, received_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            audit_uuid,
            source_name,
            report_id,
            reason,
            created_at,
            total_diff,
            total_value,
            event_uuid,
            _now_iso(),
        ),
    )

    applied_rows = 0
    for raw in lines:
        if not isinstance(raw, dict):
            continue
        it = _clean(raw.get("item_type"))
        sc = _clean(raw.get("school"))
        cl = _clean(raw.get("color"))
        sz = _clean(raw.get("size"))
        if not (it and sc and cl and sz):
            raise ApplyError("POS_STOCK_AUDIT_APPLIED: incomplete line specs")
        try:
            expected = int(raw.get("expected") or 0)
            actual = int(raw.get("actual") or 0)
            diff = int(raw.get("diff", actual - expected) or 0)
            price = float(raw.get("unit_price") or 0)
            diff_value = float(raw.get("diff_value", diff * price) or 0)
        except (TypeError, ValueError):
            raise ApplyError("POS_STOCK_AUDIT_APPLIED: invalid line numbers")
        if actual < 0:
            raise ApplyError("POS_STOCK_AUDIT_APPLIED: actual count cannot be negative")

        conn.execute(
            """
            INSERT INTO pos_stock_audit_items_mirror
                (audit_uuid, source_device, item_type, school, color, size,
                 expected_qty, actual_qty, diff_qty, unit_price, diff_value)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (audit_uuid, source_name, it, sc, cl, sz, expected, actual, diff, price, diff_value),
        )

        conn.execute(
            """
            DELETE FROM pos_stocks_mirror
             WHERE source_device = ?
               AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
            """,
            (source_name, it, sc, cl, sz),
        )
        if actual > 0:
            conn.execute(
                """INSERT INTO pos_stocks_mirror
                   (source_device, item_type, school, color, size,
                    unit_price, count, snapshot_at)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (source_name, it, sc, cl, sz, price, actual, created_at),
            )
        applied_rows += 1

    _refresh_pos_stock_snapshot_meta(conn, source_name, created_at, update_snapshot_at=False)
    return {"audit_uuid": audit_uuid, "applied_rows": applied_rows}


def apply_stock_return_to_warehouse(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Queue a POS -> warehouse stock return for warehouse review."""
    return _queue_branch_inventory_items(conn, payload, event_uuid, queue_kind="RETURN")


def apply_pos_transfer_via_warehouse(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Queue a POS -> warehouse transfer request for warehouse rerouting."""
    return _queue_branch_inventory_items(conn, payload, event_uuid, queue_kind="TRANSFER")


def _queue_branch_inventory_items(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
    queue_kind: str,
) -> Dict[str, Any]:
    items = payload.get("items") or []
    if not isinstance(items, list) or not items:
        raise ApplyError("branch inventory request has no items")

    from_dev = _clean(payload.get("from_device")) or "pos"
    external_ref = _clean(payload.get("return_uuid") or payload.get("request_uuid"))
    requested_target = _clean(payload.get("target_device"))
    if conn.execute(
        "SELECT 1 FROM branch_inventory_queue WHERE sync_event_uuid = ? LIMIT 1",
        (event_uuid,),
    ).fetchone():
        return {"already_applied": True}

    queued = 0
    total_qty = 0
    for idx, item in enumerate(items):
        it = _clean(item.get("item_type"))
        sc = _clean(item.get("school"))
        cl = _clean(item.get("color"))
        sz = _clean(item.get("size"))
        try:
            qty = int(item.get("qty") or 0)
        except (TypeError, ValueError):
            raise ApplyError(f"invalid qty in return line: {item}")
        try:
            price = float(item.get("unit_price") or 0)
        except (TypeError, ValueError):
            raise ApplyError(f"invalid unit_price in return line: {item}")
        if not (it and sc and cl and sz):
            raise ApplyError(f"incomplete queue line specs: {item}")
        if qty <= 0:
            continue

        conn.execute(
            """
            INSERT INTO branch_inventory_queue
                (sync_event_uuid, queue_kind, source_device, requested_target_device,
                 external_ref, line_index, created_at, item_type, school, color,
                 size, unit_price, qty, note, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING')
            """,
            (
                event_uuid,
                queue_kind,
                from_dev,
                requested_target or None,
                external_ref or None,
                idx,
                _now_iso(),
                it,
                sc,
                cl,
                sz,
                price,
                qty,
                _clean(payload.get("note")),
            ),
        )
        queued += 1
        total_qty += qty

    return {
        "queued_rows": queued,
        "total_qty": total_qty,
        "requested_target_device": requested_target or None,
        "external_ref": external_ref or None,
    }


def apply_catalog_upsert(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Ignore generic catalog broadcasts on POS branches.

    Branches should learn sellable/reservable definitions only from stock
    they actually receive, or from the explicit targeted reservation-definition
    flow (`send_catalog_rows_to_pos`, applied as STOCK_TRANSFER_OUT with qty 0).
    Older all-POS CATALOG_UPSERT broadcasts are intentionally harmless here.
    """
    return {
        "skipped": True,
        "reason": "generic catalog broadcasts are ignored by POS branches",
    }


def apply_pos_ownership_snapshot(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Replace this POS branch's allowed stock/catalog specs from Warehouse.

    Warehouse is the source of truth for which specs belong to a branch:
    confirmed bill shipments plus explicit reservation-definition rows. This
    removes bad schools/specs that entered a branch through old broad syncs.

    Important: the snapshot is an ownership/catalog snapshot, not a quantity
    snapshot. Existing positive `stocks.count` values are always preserved,
    even when Warehouse sends an incomplete ownership list. Only zero-stock
    unowned catalog rows are removed, and missing allowed specs are created as
    zero-stock catalog rows.
    """
    rows = payload.get("specs") or []
    if not isinstance(rows, list):
        raise ApplyError("POS_OWNERSHIP_SNAPSHOT specs must be a list")

    allowed: Set[Tuple[str, str, str, str]] = set()
    kept = 0
    skipped = 0
    prices_updated = 0
    catalog_created = 0
    tombstone_skipped = 0

    conn.execute("DROP TABLE IF EXISTS _pos_allowed_specs")
    conn.execute(
        """
        CREATE TEMP TABLE _pos_allowed_specs(
            item_type TEXT NOT NULL,
            school TEXT NOT NULL,
            color TEXT NOT NULL,
            size TEXT NOT NULL,
            PRIMARY KEY(item_type, school, color, size)
        )
        """
    )
    conn.execute("DROP TABLE IF EXISTS _pos_protected_specs")
    conn.execute(
        """
        CREATE TEMP TABLE _pos_protected_specs(
            item_type TEXT NOT NULL,
            school TEXT NOT NULL,
            color TEXT NOT NULL,
            size TEXT NOT NULL,
            PRIMARY KEY(item_type, school, color, size)
        )
        """
    )

    for raw in rows:
        if not isinstance(raw, dict):
            skipped += 1
            continue
        it = _clean(raw.get("item_type"))
        sc = _clean(raw.get("school"))
        cl = _clean(raw.get("color"))
        sz = _clean(raw.get("size"))
        if not (it and sc and cl and sz):
            skipped += 1
            continue
        key = (it.casefold(), sc.casefold(), cl.casefold(), sz.casefold())
        if key in allowed:
            continue
        allowed.add(key)
        try:
            price = float(raw.get("unit_price") or 0)
        except (TypeError, ValueError):
            price = 0.0
        spec = {"item_type": it, "school": sc, "color": cl, "size": sz}
        if _branch_catalog_spec_is_deleted(conn, spec) and not _spec_has_positive_stock(conn, spec):
            tombstone_skipped += 1
            continue

        conn.execute(
            """
            INSERT OR IGNORE INTO _pos_allowed_specs(item_type, school, color, size)
            VALUES (LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)))
            """,
            (it, sc, cl, sz),
        )
        _upsert_spec_history(conn, spec)
        _ensure_item_default_price(conn, it, price)
        _mark_branch_catalog_definition(
            conn,
            it,
            sc,
            cl,
            sz,
            price,
            event_uuid,
            note="POS ownership source-of-truth",
        )
        if _upsert_zero_stock_catalog_row(conn, it, sc, cl, sz, price):
            catalog_created += 1
        cur = conn.execute(
            """
            UPDATE stocks
               SET unit_price = ?
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
            """,
            (price, it, sc, cl, sz),
        )
        prices_updated += int(cur.rowcount or 0)
        kept += 1

    for table_name, where_sql in (
        ("stocks", "COALESCE(count, 0) > 0"),
        ("incoming_shipment_items_pending", "(COALESCE(expected_qty, 0) > 0 OR COALESCE(received_qty, 0) > 0) AND UPPER(COALESCE(status, '')) <> 'CANCELLED'"),
        ("stock_audit_report_lines", "1=1"),
        ("branch_catalog_definitions", "LOWER(COALESCE(note, '')) LIKE '%reservation%' OR COALESCE(note, '') LIKE '%حجز%' OR LOWER(COALESCE(note, '')) LIKE '%reclassification%'"),
    ):
        try:
            conn.execute(
                f"""
                INSERT OR IGNORE INTO _pos_protected_specs(item_type, school, color, size)
                SELECT LOWER(TRIM(COALESCE(item_type, ''))),
                       LOWER(TRIM(COALESCE(school, ''))),
                       LOWER(TRIM(COALESCE(color, ''))),
                       LOWER(TRIM(COALESCE(size, '')))
                  FROM {table_name}
                 WHERE {where_sql}
                   AND COALESCE(TRIM(item_type), '') <> ''
                   AND COALESCE(TRIM(school), '') <> ''
                   AND COALESCE(TRIM(color), '') <> ''
                   AND COALESCE(TRIM(size), '') <> ''
                """
            )
        except sqlite3.OperationalError:
            pass

    try:
        inbox_rows = conn.execute(
            """
            SELECT event_type, payload_json
              FROM sync_inbox
             WHERE event_type IN ('STOCK_TRANSFER_OUT', 'BRANCH_STOCK_RECLASSIFIED')
            """
        ).fetchall()
    except sqlite3.OperationalError:
        inbox_rows = []
    for row in inbox_rows:
        try:
            event_type = str(row["event_type"] if isinstance(row, sqlite3.Row) else row[0] or "")
            payload_json = row["payload_json"] if isinstance(row, sqlite3.Row) else row[1]
            payload_obj = json.loads(payload_json or "{}")
        except Exception:
            continue
        protected_rows: List[Dict[str, Any]] = []
        if event_type == "STOCK_TRANSFER_OUT":
            note_text = _clean(payload_obj.get("note")).casefold()
            is_reservation_definition = "reservation" in note_text or "حجز" in note_text
            for item in payload_obj.get("items") or []:
                if not isinstance(item, dict):
                    continue
                try:
                    qty = int(float(item.get("qty") or 0))
                except (TypeError, ValueError):
                    qty = 0
                if qty > 0 or (bool(item.get("catalog_only")) and is_reservation_definition):
                    protected_rows.append(item)
        elif event_type == "BRANCH_STOCK_RECLASSIFIED":
            to_spec = payload_obj.get("to_spec") or {}
            if isinstance(to_spec, dict):
                protected_rows.append(to_spec)
            catalog_rows = payload_obj.get("catalog_rows") or []
            if isinstance(catalog_rows, list):
                protected_rows.extend([r for r in catalog_rows if isinstance(r, dict)])
        for raw in protected_rows:
            it = _clean(raw.get("item_type"))
            sc = _clean(raw.get("school"))
            cl = _clean(raw.get("color"))
            sz = _clean(raw.get("size"))
            if not (it and sc and cl and sz):
                continue
            conn.execute(
                """
                INSERT OR IGNORE INTO _pos_protected_specs(item_type, school, color, size)
                VALUES (LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)))
                """,
                (it, sc, cl, sz),
            )

    preserved_positive = 0
    for row in conn.execute(
        """
        SELECT item_type, school, color, size, unit_price, SUM(COALESCE(count, 0)) AS qty
          FROM stocks
         WHERE COALESCE(count, 0) > 0
           AND NOT EXISTS (
                SELECT 1
                  FROM _pos_allowed_specs
                 WHERE _pos_allowed_specs.item_type = LOWER(TRIM(COALESCE(stocks.item_type, '')))
                   AND _pos_allowed_specs.school = LOWER(TRIM(COALESCE(stocks.school, '')))
                   AND _pos_allowed_specs.color = LOWER(TRIM(COALESCE(stocks.color, '')))
                   AND _pos_allowed_specs.size = LOWER(TRIM(COALESCE(stocks.size, '')))
           )
         GROUP BY item_type, school, color, size, unit_price
        """
    ).fetchall():
        it = _clean(row["item_type"])
        sc = _clean(row["school"])
        cl = _clean(row["color"])
        sz = _clean(row["size"])
        if not (it and sc and cl and sz):
            continue
        try:
            price = float(row["unit_price"] or 0)
        except (TypeError, ValueError):
            price = 0.0
        _upsert_spec_history(conn, {"item_type": it, "school": sc, "color": cl, "size": sz})
        _ensure_item_default_price(conn, it, price)
        _mark_branch_catalog_definition(
            conn,
            it,
            sc,
            cl,
            sz,
            price,
            event_uuid,
            note="Preserved positive stock outside ownership snapshot",
        )
        preserved_positive += int(row["qty"] or 0)

    deleted_stock = conn.execute(
        """
        DELETE FROM stocks
         WHERE NOT EXISTS (
                SELECT 1
                  FROM _pos_allowed_specs
                 WHERE _pos_allowed_specs.item_type = LOWER(TRIM(COALESCE(stocks.item_type, '')))
                   AND _pos_allowed_specs.school = LOWER(TRIM(COALESCE(stocks.school, '')))
                   AND _pos_allowed_specs.color = LOWER(TRIM(COALESCE(stocks.color, '')))
                   AND _pos_allowed_specs.size = LOWER(TRIM(COALESCE(stocks.size, '')))
           )
           AND COALESCE(count, 0) <= 0
           AND NOT EXISTS (
                SELECT 1
                  FROM _pos_protected_specs
                 WHERE _pos_protected_specs.item_type = LOWER(TRIM(COALESCE(stocks.item_type, '')))
                   AND _pos_protected_specs.school = LOWER(TRIM(COALESCE(stocks.school, '')))
                   AND _pos_protected_specs.color = LOWER(TRIM(COALESCE(stocks.color, '')))
                   AND _pos_protected_specs.size = LOWER(TRIM(COALESCE(stocks.size, '')))
           )
        """
    ).rowcount
    deleted_profiles = conn.execute(
        """
        DELETE FROM size_profiles
         WHERE NOT EXISTS (
                SELECT 1
                  FROM _pos_allowed_specs
                 WHERE _pos_allowed_specs.item_type = LOWER(TRIM(COALESCE(size_profiles.item_type, '')))
                   AND _pos_allowed_specs.school = LOWER(TRIM(COALESCE(size_profiles.school, '')))
                   AND _pos_allowed_specs.color = LOWER(TRIM(COALESCE(size_profiles.color, '')))
           )
           AND NOT EXISTS (
                SELECT 1
                  FROM stocks
                 WHERE LOWER(TRIM(COALESCE(stocks.item_type, ''))) = LOWER(TRIM(COALESCE(size_profiles.item_type, '')))
                   AND LOWER(TRIM(COALESCE(stocks.school, ''))) = LOWER(TRIM(COALESCE(size_profiles.school, '')))
                   AND LOWER(TRIM(COALESCE(stocks.color, ''))) = LOWER(TRIM(COALESCE(size_profiles.color, '')))
                   AND COALESCE(stocks.count, 0) > 0
           )
           AND NOT EXISTS (
                SELECT 1
                  FROM _pos_protected_specs
                 WHERE _pos_protected_specs.item_type = LOWER(TRIM(COALESCE(size_profiles.item_type, '')))
                   AND _pos_protected_specs.school = LOWER(TRIM(COALESCE(size_profiles.school, '')))
                   AND _pos_protected_specs.color = LOWER(TRIM(COALESCE(size_profiles.color, '')))
           )
        """
    ).rowcount
    try:
        deleted_catalog = conn.execute(
            """
            DELETE FROM branch_catalog_definitions
             WHERE NOT EXISTS (
                    SELECT 1
                      FROM _pos_allowed_specs
                     WHERE _pos_allowed_specs.item_type = LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                       AND _pos_allowed_specs.school = LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                       AND _pos_allowed_specs.color = LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                       AND _pos_allowed_specs.size = LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
               )
               AND NOT EXISTS (
                    SELECT 1
                      FROM stocks
                     WHERE LOWER(TRIM(COALESCE(stocks.item_type, ''))) = LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                       AND LOWER(TRIM(COALESCE(stocks.school, ''))) = LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                       AND LOWER(TRIM(COALESCE(stocks.color, ''))) = LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                       AND LOWER(TRIM(COALESCE(stocks.size, ''))) = LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
                       AND COALESCE(stocks.count, 0) > 0
               )
               AND NOT EXISTS (
                    SELECT 1
                      FROM _pos_protected_specs
                     WHERE _pos_protected_specs.item_type = LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                       AND _pos_protected_specs.school = LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                       AND _pos_protected_specs.color = LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                       AND _pos_protected_specs.size = LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
               )
            """
        ).rowcount
    except sqlite3.OperationalError:
        deleted_catalog = 0

    profile_rows = _apply_size_profile_rows(conn, payload.get("size_profiles"))
    conn.execute("DROP TABLE IF EXISTS _pos_allowed_specs")
    conn.execute("DROP TABLE IF EXISTS _pos_protected_specs")

    return {
        "kept_specs": kept,
        "skipped_specs": skipped,
        "tombstone_skipped_specs": tombstone_skipped,
        "catalog_created": catalog_created,
        "prices_updated": prices_updated,
        "deleted_stock": int(deleted_stock or 0),
        "deleted_size_profiles": int(deleted_profiles or 0),
        "deleted_catalog_definitions": int(deleted_catalog or 0),
        "preserved_positive_qty": preserved_positive,
        "size_profiles_applied": profile_rows,
        "quantity_mode": "preserve_all_positive_counts",
    }


def apply_remote_reservation_request(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Apply a call-center reservation request on the POS.

    Rules:
    - Creates a pending reservation line and a RESERVE movement.
    - Rejects the request when resulting available quantity would be < 4.
    - Stores a one-shot alert so the POS UI can pop a reservation notice.
    """
    if conn.execute(
        "SELECT 1 FROM reservation_alerts WHERE sync_event_uuid=? LIMIT 1",
        (event_uuid,),
    ).fetchone():
        return {"already_applied": True}

    line = payload.get("line") or {}
    if not isinstance(line, dict):
        raise ApplyError("REMOTE_RESERVATION_REQUEST: missing line payload")
    customer = _clean(payload.get("customer"))
    if not customer:
        raise ApplyError("REMOTE_RESERVATION_REQUEST: customer is required")
    it = _clean(line.get("item_type"))
    sc = _clean(line.get("school"))
    cl = _clean(line.get("color"))
    sz = _clean(line.get("size"))
    if not (it and sc and cl and sz):
        raise ApplyError("REMOTE_RESERVATION_REQUEST: incomplete item specs")
    try:
        qty = int(line.get("qty") or 0)
    except (TypeError, ValueError):
        raise ApplyError("REMOTE_RESERVATION_REQUEST: invalid qty")
    if qty <= 0:
        raise ApplyError("REMOTE_RESERVATION_REQUEST: qty must be > 0")

    row = conn.execute(
        """
        SELECT COALESCE(SUM(count),0)
        FROM stocks
        WHERE LOWER(TRIM(item_type))=LOWER(TRIM(?))
          AND LOWER(TRIM(school))=LOWER(TRIM(?))
          AND LOWER(TRIM(color))=LOWER(TRIM(?))
          AND LOWER(TRIM(size))=LOWER(TRIM(?))
        """,
        (it, sc, cl, sz),
    ).fetchone()
    on_hand = int((row[0] if row else 0) or 0)
    row = conn.execute(
        """
        SELECT COALESCE(SUM(qty),0)
        FROM reservations
        WHERE status='معلق'
          AND LOWER(TRIM(item_type))=LOWER(TRIM(?))
          AND LOWER(TRIM(school))=LOWER(TRIM(?))
          AND LOWER(TRIM(color))=LOWER(TRIM(?))
          AND LOWER(TRIM(size))=LOWER(TRIM(?))
        """,
        (it, sc, cl, sz),
    ).fetchone()
    reserved_pending = int((row[0] if row else 0) or 0)
    available = max(0, on_hand - reserved_pending)
    if qty > available:
        raise ApplyError(f"REMOTE_RESERVATION_REQUEST: requested qty {qty} exceeds available {available}")
    if (available - qty) < 4:
        raise ApplyError("REMOTE_RESERVATION_REQUEST: cannot reserve because remaining stock would be < 4")

    try:
        unit_price = float(line.get("unit_price") or 0)
    except (TypeError, ValueError):
        unit_price = 0.0
    if unit_price <= 0:
        prow = conn.execute(
            """
            SELECT unit_price
            FROM stocks
            WHERE LOWER(TRIM(item_type))=LOWER(TRIM(?))
              AND LOWER(TRIM(school))=LOWER(TRIM(?))
              AND LOWER(TRIM(color))=LOWER(TRIM(?))
              AND LOWER(TRIM(size))=LOWER(TRIM(?))
            ORDER BY id DESC
            LIMIT 1
            """,
            (it, sc, cl, sz),
        ).fetchone()
        if prow and prow[0] is not None:
            unit_price = float(prow[0] or 0)

    try:
        hold_days = int(payload.get("hold_days") or 2)
    except (TypeError, ValueError):
        hold_days = 2
    if hold_days <= 0:
        hold_days = 2
    hold_until = conn.execute(
        "SELECT datetime('now', ?)",
        (f"+{hold_days} days",),
    ).fetchone()[0]
    note = f"طلب حجز من الكول سنتر (ينتهي {hold_until})"
    total = float(unit_price) * int(qty)
    ts = _now_iso()
    cur = conn.execute(
        """
        INSERT INTO reservations(created_at, customer, item_type, school, color, size, qty, unit_price, total_amount, paid_amount, status, note)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 'معلق', ?)
        """,
        (ts, customer, it, sc, cl, sz, qty, float(unit_price), total, note),
    )
    reservation_id = int(cur.lastrowid)
    conn.execute(
        """
        INSERT INTO movements(ts, direction, stock_id, qty, note, bill_id, item_type, school, color, size, unit_price)
        VALUES (?, 'RESERVE', NULL, ?, ?, NULL, ?, ?, ?, ?, ?)
        """,
        (ts, qty, f"حجز - {customer}", it, sc, cl, sz, float(unit_price)),
    )
    conn.execute(
        """
        INSERT INTO reservation_alerts
            (sync_event_uuid, request_uuid, customer, branch_device, item_type, school, color, size, qty, hold_until, note, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            event_uuid,
            _clean(payload.get("request_uuid")) or None,
            customer,
            _clean(payload.get("__target_device__")) or None,
            it, sc, cl, sz, qty,
            hold_until,
            note,
            ts,
        ),
    )
    return {"reservation_id": reservation_id, "qty": qty, "hold_until": hold_until}


# ---------------- warehouse: POS mirror + financial ledger appliers ---------------- #


def _src_device(conn: sqlite3.Connection, payload: Dict[str, Any], event_uuid: str) -> str:
    s = _clean(payload.get("__source_device__"))
    if s:
        return s
    try:
        row = conn.execute(
            "SELECT source_device FROM sync_inbox WHERE event_uuid = ?",
            (event_uuid,),
        ).fetchone()
    except sqlite3.OperationalError:
        return ""
    return _clean(row[0] if row else "")


def _ledger_day(conn: sqlite3.Connection, event_uuid: str) -> str:
    try:
        row = conn.execute(
            "SELECT substr(applied_at, 1, 10) FROM sync_inbox WHERE event_uuid = ?",
            (event_uuid,),
        ).fetchone()
    except sqlite3.OperationalError:
        return _now_iso()[:10]
    if row and row[0] and len(str(row[0])) >= 10:
        return str(row[0])[:10]
    return _now_iso()[:10]


def _ledger_append(
    conn: sqlite3.Connection,
    *,
    src: str,
    event_uuid: str,
    event_type: str,
    category: str,
    amount: float,
    day: str,
    related_id: Optional[int] = None,
    meta: Optional[Dict[str, Any]] = None,
) -> None:
    if not src:
        raise ApplyError("ledger row missing source_device")
    conn.execute(
        """
        INSERT OR IGNORE INTO pos_financial_ledger
            (source_device, event_uuid, event_type, category, amount, day,
             related_id, meta_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            src,
            event_uuid,
            event_type,
            category,
            float(amount),
            day[:10],
            related_id,
            json.dumps(meta or {}, ensure_ascii=False, default=str),
            _now_iso(),
        ),
    )


def _reservation_key_from_line(line: Dict[str, Any]) -> str:
    ru = _clean(line.get("reservation_uuid"))
    if ru:
        return ru
    rid = line.get("reservation_id")
    try:
        return f"id:{int(rid)}" if rid is not None else ""
    except (TypeError, ValueError):
        return ""


def _reservation_key_from_payload(payload: Dict[str, Any]) -> str:
    ru = _clean(payload.get("reservation_uuid"))
    if ru:
        return ru
    rid = payload.get("reservation_id")
    try:
        return f"id:{int(rid)}" if rid is not None else ""
    except (TypeError, ValueError):
        return ""


def apply_wh_pos_reservation_created(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("RESERVATION_CREATED: missing source_device")
    lines = payload.get("reservations") or []
    if not isinstance(lines, list) or not lines:
        return {"rows": 0}
    now = _now_iso()
    day = _ledger_day(conn, event_uuid)
    customer = _clean(payload.get("customer"))
    shift_id = payload.get("shift_id")
    try:
        shift_id = int(shift_id) if shift_id is not None else None
    except (TypeError, ValueError):
        shift_id = None
    paid_batch = float(payload.get("paid_amount") or 0)
    n = 0
    for idx, line in enumerate(lines):
        if not isinstance(line, dict):
            continue
        key = _reservation_key_from_line(line)
        if not key:
            continue
        it = _clean(line.get("item_type"))
        sc = _clean(line.get("school"))
        cl = _clean(line.get("color"))
        sz = _clean(line.get("size"))
        if not (it and sc and cl and sz):
            raise ApplyError("RESERVATION_CREATED: incomplete line specs")
        try:
            qty = int(line.get("qty") or 0)
            price = float(line.get("unit_price") or 0)
        except (TypeError, ValueError):
            raise ApplyError("RESERVATION_CREATED: invalid qty or unit_price")
        total_amount = price * qty
        alloc_paid = paid_batch if idx == 0 else 0.0
        conn.execute(
            """
            INSERT OR REPLACE INTO pos_reservations_mirror
                (source_device, reservation_key, customer, item_type, school, color, size,
                 qty, unit_price, total_amount, paid_amount, status, shift_id, last_event_uuid, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'معلق', ?, ?, ?)
            """,
            (
                src, key, customer, it, sc, cl, sz,
                qty, price, total_amount, alloc_paid,
                shift_id, event_uuid, now,
            ),
        )
        n += 1
    if paid_batch:
        _ledger_append(
            conn,
            src=src,
            event_uuid=event_uuid,
            event_type="RESERVATION_CREATED",
            category="reservation_downpayment",
            amount=paid_batch,
            day=day,
            related_id=None,
            meta={"customer": customer},
        )
    return {"rows": n}


def apply_wh_pos_reservation_payment_updated(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("RESERVATION_PAYMENT_UPDATED: missing source_device")
    key = _reservation_key_from_payload(payload)
    if not key:
        raise ApplyError("RESERVATION_PAYMENT_UPDATED: missing reservation id")
    new_paid = float(payload.get("paid_amount") or 0)
    row = conn.execute(
        "SELECT paid_amount FROM pos_reservations_mirror WHERE source_device=? AND reservation_key=?",
        (src, key),
    ).fetchone()
    old = float(row[0]) if row else 0.0
    delta = new_paid - old
    now = _now_iso()
    day = _ledger_day(conn, event_uuid)
    rid = payload.get("reservation_id")
    try:
        rid_int = int(rid) if rid is not None else None
    except (TypeError, ValueError):
        rid_int = None
    cur = conn.execute(
        """
        UPDATE pos_reservations_mirror
           SET paid_amount = ?, last_event_uuid = ?, updated_at = ?
         WHERE source_device = ? AND reservation_key = ?
        """,
        (new_paid, event_uuid, now, src, key),
    )
    if cur.rowcount == 0:
        conn.execute(
            """
            INSERT INTO pos_reservations_mirror
                (source_device, reservation_key, customer,
                 item_type, school, color, size, qty, unit_price, total_amount,
                 paid_amount, status, shift_id, last_event_uuid, updated_at)
            VALUES (?, ?, '', '', '', '', '', 0, 0, 0, ?, 'معلق', NULL, ?, ?)
            """,
            (src, key, new_paid, event_uuid, now),
        )
    if abs(delta) > 1e-9:
        _ledger_append(
            conn,
            src=src,
            event_uuid=event_uuid,
            event_type="RESERVATION_PAYMENT_UPDATED",
            category="reservation_payment",
            amount=delta,
            day=day,
            related_id=rid_int,
            meta={"new_paid": new_paid},
        )
    return {"delta": delta, "new_paid": new_paid}


def apply_wh_pos_reservation_completed(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("RESERVATION_COMPLETED: missing source_device")
    key = _reservation_key_from_payload(payload)
    if not key:
        raise ApplyError("RESERVATION_COMPLETED: missing reservation id")
    now = _now_iso()
    conn.execute(
        """
        UPDATE pos_reservations_mirror
           SET status = 'تم التسليم', last_event_uuid = ?, updated_at = ?
         WHERE source_device = ? AND reservation_key = ?
        """,
        (event_uuid, now, src, key),
    )
    return {"updated": 1}


def apply_wh_pos_reservation_delivered(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("RESERVATION_DELIVERED: missing source_device")
    key = _reservation_key_from_payload(payload)
    if not key:
        raise ApplyError("RESERVATION_DELIVERED: missing reservation id")
    new_total = float(payload.get("paid_amount_total") or 0)
    coll = float(payload.get("collected_amount") or 0)
    now = _now_iso()
    day = _ledger_day(conn, event_uuid)
    rid = payload.get("reservation_id")
    try:
        rid_int = int(rid) if rid is not None else None
    except (TypeError, ValueError):
        rid_int = None
    cur = conn.execute(
        """
        UPDATE pos_reservations_mirror
           SET status = 'تم التسليم', paid_amount = ?, last_event_uuid = ?, updated_at = ?
         WHERE source_device = ? AND reservation_key = ?
        """,
        (new_total, event_uuid, now, src, key),
    )
    if cur.rowcount == 0:
        conn.execute(
            """
            INSERT INTO pos_reservations_mirror
                (source_device, reservation_key, customer,
                 item_type, school, color, size, qty, unit_price, total_amount,
                 paid_amount, status, shift_id, last_event_uuid, updated_at)
            VALUES (?, ?, '', '', '', '', '', 0, 0, 0, ?, 'تم التسليم', NULL, ?, ?)
            """,
            (src, key, new_total, event_uuid, now),
        )
    if coll:
        leg_uid = "%s:deliver_collect" % (event_uuid,)
        _ledger_append(
            conn,
            src=src,
            event_uuid=leg_uid,
            event_type="RESERVATION_DELIVERED",
            category="reservation_collect",
            amount=coll,
            day=day,
            related_id=rid_int,
            meta={"paid_amount_total": new_total},
        )
    return {"collected": coll, "paid_total": new_total}


def apply_wh_pos_ledger_sale_created(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SALE_CREATED: missing source_device")
    total = float(payload.get("total") or 0)
    day = _ledger_day(conn, event_uuid)
    bid = payload.get("bill_id")
    try:
        bid_int = int(bid) if bid is not None else None
    except (TypeError, ValueError):
        bid_int = None
    _ledger_append(
        conn,
        src=src,
        event_uuid=event_uuid,
        event_type="SALE_CREATED",
        category="sale",
        amount=total,
        day=day,
        related_id=bid_int,
        meta={"customer": _clean(payload.get("customer"))},
    )
    return {"amount": total}


def apply_wh_pos_ledger_sale_returned(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SALE_RETURNED: missing source_device")
    total = float(payload.get("total") or 0)
    day = _ledger_day(conn, event_uuid)
    bid = payload.get("bill_id")
    try:
        bid_int = int(bid) if bid is not None else None
    except (TypeError, ValueError):
        bid_int = None
    _ledger_append(
        conn,
        src=src,
        event_uuid=event_uuid,
        event_type="SALE_RETURNED",
        category="return_bill",
        amount=-abs(total),
        day=day,
        related_id=bid_int,
        meta={},
    )
    return {"amount": -abs(total)}


def apply_wh_pos_ledger_sale_voided(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SALE_VOIDED: missing source_device")
    total = float(payload.get("total") or 0)
    day = _ledger_day(conn, event_uuid)
    bid = payload.get("bill_id")
    try:
        bid_int = int(bid) if bid is not None else None
    except (TypeError, ValueError):
        bid_int = None
    _ledger_append(
        conn,
        src=src,
        event_uuid=event_uuid,
        event_type="SALE_VOIDED",
        category="void_bill",
        amount=-abs(total),
        day=day,
        related_id=bid_int,
        meta={"reason": _clean(payload.get("reason"))},
    )
    return {"amount": -abs(total)}


def apply_wh_pos_ledger_sale_exchanged(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SALE_EXCHANGED: missing source_device")
    diff = float(payload.get("diff") or 0)
    day = _ledger_day(conn, event_uuid)
    bid = payload.get("bill_id")
    try:
        bid_int = int(bid) if bid is not None else None
    except (TypeError, ValueError):
        bid_int = None
    _ledger_append(
        conn,
        src=src,
        event_uuid=event_uuid,
        event_type="SALE_EXCHANGED",
        category="exchange_net",
        amount=diff,
        day=day,
        related_id=bid_int,
        meta={
            "return_total": float(payload.get("return_total") or 0),
            "take_total": float(payload.get("take_total") or 0),
        },
    )
    return {"amount": diff}


def apply_wh_pos_ledger_sale_bill_type_corrected(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SALE_BILL_TYPE_CORRECTED: missing source_device")
    amount_delta = float(payload.get("amount_delta") or 0)
    day = _ledger_day(conn, event_uuid)
    bid = payload.get("bill_id")
    try:
        bid_int = int(bid) if bid is not None else None
    except (TypeError, ValueError):
        bid_int = None
    _ledger_append(
        conn,
        src=src,
        event_uuid=event_uuid,
        event_type="SALE_BILL_TYPE_CORRECTED",
        category="bill_correction",
        amount=amount_delta,
        day=day,
        related_id=bid_int,
        meta={
            "from_bill_type": _clean(payload.get("from_bill_type")),
            "to_bill_type": _clean(payload.get("to_bill_type")),
            "reason": _clean(payload.get("reason")),
        },
    )
    return {"amount": amount_delta}


# ----------------------------- registry ------------------------------- #

ApplierFn = Callable[[sqlite3.Connection, Dict[str, Any], str], Dict[str, Any]]


_POS_REGISTRY: Dict[str, ApplierFn] = {
    "STOCK_TRANSFER_OUT": apply_stock_transfer_out,
    "STOCK_TRANSFER_CANCELLED": apply_stock_transfer_cancelled,
    "BRANCH_STOCK_RECLASSIFIED": apply_branch_stock_reclassified,
    "BRANCH_CATALOG_DELETED": apply_branch_catalog_deleted,
    "PRICE_UPDATE":       apply_price_update,
    "CATALOG_UPSERT":     apply_catalog_upsert,
    "POS_OWNERSHIP_SNAPSHOT": apply_pos_ownership_snapshot,
    "SPEC_RENAMED":       apply_spec_renamed,
    "REMOTE_RESERVATION_REQUEST": apply_remote_reservation_request,
}


_WAREHOUSE_REGISTRY: Dict[str, ApplierFn] = {
    "POS_STOCK_SNAPSHOT": apply_pos_stock_snapshot,
    "STOCK_RETURN_TO_WAREHOUSE": apply_stock_return_to_warehouse,
    "POS_TRANSFER_VIA_WAREHOUSE": apply_pos_transfer_via_warehouse,
    "RESERVATION_CREATED": apply_wh_pos_reservation_created,
    "RESERVATION_PAYMENT_UPDATED": apply_wh_pos_reservation_payment_updated,
    "RESERVATION_COMPLETED": apply_wh_pos_reservation_completed,
    "RESERVATION_DELIVERED": apply_wh_pos_reservation_delivered,
    "SALE_CREATED": apply_wh_pos_ledger_sale_created,
    "SALE_RETURNED": apply_wh_pos_ledger_sale_returned,
    "SALE_VOIDED": apply_wh_pos_ledger_sale_voided,
    "SALE_EXCHANGED": apply_wh_pos_ledger_sale_exchanged,
    "SALE_BILL_TYPE_CORRECTED": apply_wh_pos_ledger_sale_bill_type_corrected,
}


def for_role(role: Optional[str]) -> Dict[str, ApplierFn]:
    """Return the applier dict appropriate for a device role."""
    if role == "pos":
        return _POS_REGISTRY
    if role == "warehouse":
        return _WAREHOUSE_REGISTRY
    return {}
