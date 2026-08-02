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
from datetime import datetime, timezone
from typing import Any, Callable, Dict, Optional


# ----------------------------- helpers -------------------------------- #

def _now_iso() -> str:
    """Local-time ISO timestamp matching existing `movements.ts` format."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.execute("PRAGMA table_info(%s)" % table)
    return any(r[1] == column for r in cur.fetchall())


def _clean(value: Any) -> str:
    return str(value or "").strip()


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
) -> bool:
    """Ensure a POS can select a catalog item even before stock arrives."""
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

    # Idempotency guard by shipment UUID.
    if shipment_uuid:
        dup = conn.execute(
            "SELECT id FROM incoming_shipment_alerts WHERE shipment_uuid=? LIMIT 1",
            (shipment_uuid,),
        ).fetchone()
        if dup:
            return {"queued": True, "duplicate": True, "shipment": shipment_uuid}

    queued_rows = 0
    total_qty = 0
    catalog_rows = 0

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

        if qty <= 0:
            if _upsert_zero_stock_catalog_row(conn, it, sc, cl, sz, price):
                catalog_rows += 1
            continue

        conn.execute(
            """
            INSERT OR IGNORE INTO incoming_shipment_items_pending(
                shipment_uuid, line_index, item_type, school, color, size,
                unit_price, expected_qty, received_qty, status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULL, 'PENDING')
            """,
            (shipment_uuid or event_uuid, queued_rows, it, sc, cl, sz, price, qty),
        )
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
        "catalog_rows": catalog_rows,
        "total_qty":  total_qty,
        "shipment":   shipment_uuid or None,
        "needs_verification": queued_rows > 0,
        "size_profiles_applied": profile_rows,
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

    where_parts = []
    args = []
    for k in ("item_type", "school", "color", "size"):
        v = _clean(filters.get(k))
        if v:
            where_parts.append(f"LOWER(TRIM({k})) = LOWER(?)")
            args.append(v)

    if not where_parts:
        return {"skipped": True, "reason": "refusing unfiltered price update"}

    rows = conn.execute(
        "SELECT id, item_type, school, color, size FROM stocks WHERE "
        + " AND ".join(where_parts),
        args,
    ).fetchall()
    if not rows:
        return {"updated": 0}

    conn.execute(
        "UPDATE stocks SET unit_price = ? WHERE "
        + " AND ".join(where_parts),
        (new_price, *args),
    )

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
    for it in item_types:
        _update_item_default_price(conn, it, new_price)

    return {"updated": len(rows)}


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

    current = conn.execute(
        "SELECT snapshot_at FROM pos_stocks_snapshot_meta WHERE source_device = ?",
        (source_name,),
    ).fetchone()
    current_snapshot_at = _clean(current[0] if current else "")
    if current_snapshot_at and snapshot_at < current_snapshot_at:
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

    total_value = 0.0
    inserted = 0
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
        if not (it and sc and cl and sz) or count <= 0:
            continue
        conn.execute(
            """INSERT INTO pos_stocks_mirror
               (source_device, item_type, school, color, size,
                unit_price, count, snapshot_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (source_name, it, sc, cl, sz, price, count, snapshot_at),
        )
        inserted += 1
        total_value += price * count

    _upsert_pos_stock_snapshot_meta(conn, source_name, snapshot_at, inserted, total_value, app_version)

    return {"mirrored_rows": inserted, "total_value": total_value}


def _refresh_pos_stock_snapshot_meta(conn: sqlite3.Connection, source_name: str, snapshot_at: str) -> None:
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
    conn.execute(
        """INSERT INTO pos_stocks_snapshot_meta
               (source_device, snapshot_at, row_count, total_value)
           VALUES (?, ?, ?, ?)
           ON CONFLICT(source_device) DO UPDATE SET
               snapshot_at = excluded.snapshot_at,
               row_count   = excluded.row_count,
               total_value = excluded.total_value""",
        (source_name, snapshot_at, row_count, total_value),
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

    _refresh_pos_stock_snapshot_meta(conn, source_name, created_at)
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
    """Seed catalog entries (item_defaults + spec_history + size_profile)
    from a warehouse broadcast.

    Phase 3 uses this as a safety net so new catalog items pre-seed a
    POS device before any sale touches them. Not strictly required:
    `apply_stock_transfer_out` also self-heals the same tables. Kept
    separate so the warehouse can push catalog changes without
    attaching them to a shipment.

    Payload shape:
        {
          "items": [
            {"item_type":"...", "default_price": 150.0},
            ...
          ],
          "size_profiles": [
            {"item_type":"...", "school":"...", "color":"...",
             "num_start_1": 0, "num_end_1": 9, "has_alpha": 0}
          ],
          "spec_history": [
            {"field":"item_type", "value":"..."}
          ]
        }
    """
    items = payload.get("items") or []
    profiles = payload.get("size_profiles") or []
    history = payload.get("spec_history") or []

    seeded_items = 0
    for item in items:
        it = _clean(item.get("item_type"))
        if not it:
            continue
        price = item.get("default_price")
        if price is None:
            continue
        try:
            price = float(price)
        except (TypeError, ValueError):
            continue
        conn.execute(
            "INSERT OR IGNORE INTO item_defaults(item_type, default_price) VALUES(?, ?)",
            (it, price),
        )
        seeded_items += 1

    seeded_profiles = _apply_size_profile_rows(conn, profiles)

    seeded_history = 0
    for h in history:
        field = _clean(h.get("field"))
        value = _clean(h.get("value"))
        if field in ("item_type", "school", "color", "size") and value:
            conn.execute(
                "INSERT OR IGNORE INTO spec_history(field, value) VALUES(?, ?)",
                (field, value),
            )
            seeded_history += 1

    return {
        "items":   seeded_items,
        "profiles": seeded_profiles,
        "history":  seeded_history,
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
    "PRICE_UPDATE":       apply_price_update,
    "CATALOG_UPSERT":     apply_catalog_upsert,
    "SPEC_RENAMED":       apply_spec_renamed,
    "REMOTE_RESERVATION_REQUEST": apply_remote_reservation_request,
}


_WAREHOUSE_REGISTRY: Dict[str, ApplierFn] = {
    "POS_STOCK_SNAPSHOT": apply_pos_stock_snapshot,
    "POS_STOCK_AUDIT_APPLIED": apply_pos_stock_audit_applied,
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
