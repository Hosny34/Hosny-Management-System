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
import math
import sqlite3
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Sequence


# ----------------------------- helpers -------------------------------- #

def _now_iso() -> str:
    """Local-time ISO timestamp matching existing `movements.ts` format."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.execute("PRAGMA table_info(%s)" % table)
    return any(r[1] == column for r in cur.fetchall())


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _round_up_to_step(value: float, step: int = 5) -> float:
    step_f = float(step or 1)
    if value <= 0:
        return 0.0
    return float(math.ceil(float(value) / step_f) * step_f)


def _allocate_reservation_down_payments(
    totals: Sequence[float],
    paid_amount: float,
    *,
    round_step: int = 5,
) -> List[float]:
    paid = max(0.0, float(paid_amount or 0.0))
    clean_totals = [max(0.0, float(t or 0.0)) for t in totals]
    if not clean_totals:
        return []
    if paid - sum(clean_totals) > 1e-6:
        raise ValueError("reservation down payment exceeds total")

    allocations: List[float] = []
    remaining = paid
    for idx, line_total in enumerate(clean_totals):
        if remaining <= 1e-9:
            allocations.append(0.0)
            continue
        if idx == len(clean_totals) - 1:
            alloc = remaining
        else:
            remaining_count = len(clean_totals) - idx
            alloc = _round_up_to_step(remaining / remaining_count, round_step)
            alloc = min(alloc, line_total, remaining)
        allocations.append(round(float(alloc), 2))
        remaining = round(remaining - float(alloc), 2)

    if allocations:
        allocations[-1] = round(allocations[-1] + remaining, 2)
    return allocations


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
        conn.execute(
            """
            INSERT INTO size_profiles
                (item_type, school, color,
                 num_start_1, num_end_1, num_start_2, num_end_2,
                 has_alpha, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(item_type, school, color) DO UPDATE SET
                num_start_1 = excluded.num_start_1,
                num_end_1   = excluded.num_end_1,
                num_start_2 = excluded.num_start_2,
                num_end_2   = excluded.num_end_2,
                has_alpha   = excluded.has_alpha,
                updated_at  = datetime('now')
            """,
            (
                it, sc, cl,
                p.get("num_start_1"), p.get("num_end_1"),
                p.get("num_start_2"), p.get("num_end_2"),
                int(p.get("has_alpha") or 0),
            ),
        )
        n += 1
    return n


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
                INSERT INTO size_profiles
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
                ON CONFLICT(item_type, school, color) DO NOTHING
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
    """Apply a warehouse → this-POS stock shipment.

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
        - Adds one `stocks` row per item (matching existing POS
          `add_stock` convention of row-per-batch, not merged).
        - Inserts a `movements` row with direction='IN' and a note
          referencing the shipment.
        - Auto-heals `spec_history` and `item_defaults` so the new
          specs show up in dropdowns next time the user opens a
          picker.
        - Applies `size_profiles` from the warehouse payload so POS
          size pickers match the warehouse configuration for shipped
          (item_type, school, color) keys.
    """
    items = payload.get("items") or []
    if not isinstance(items, list) or not items:
        raise ApplyError("shipment has no items")

    shipment_uuid = _clean(payload.get("shipment_uuid"))
    from_dev = _clean(payload.get("from_device")) or "warehouse"
    short_ship = (shipment_uuid[:8] if shipment_uuid else event_uuid[:8])
    note_base = f"شحنة من {from_dev} #{short_ship}"

    added_rows = 0
    total_qty = 0

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
            _upsert_zero_stock_catalog_row(conn, it, sc, cl, sz, price)
            continue

        cur = conn.execute(
            """INSERT INTO stocks(item_type, school, color, size, unit_price, count)
               VALUES(?,?,?,?,?,?)""",
            (it, sc, cl, sz, price, qty),
        )
        stock_id = int(cur.lastrowid)

        conn.execute(
            """INSERT INTO movements
               (ts, direction, stock_id, qty, note, bill_id,
                item_type, school, color, size, unit_price)
               VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
            (
                _now_iso(), "IN", stock_id, qty, note_base, None,
                it, sc, cl, sz, price,
            ),
        )
        added_rows += 1
        total_qty += qty

    profile_rows = _apply_size_profile_rows(conn, payload.get("size_profiles"))

    return {
        "added_rows": added_rows,
        "total_qty":  total_qty,
        "shipment":   shipment_uuid or None,
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
        conn.execute(
            "INSERT INTO item_defaults(item_type, default_price) VALUES(?, ?) "
            "ON CONFLICT(item_type) DO UPDATE SET default_price = excluded.default_price",
            (it, new_price),
        )

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

    if _has_column(conn, "pos_stocks_snapshot_meta", "app_version"):
        conn.execute(
            """INSERT INTO pos_stocks_snapshot_meta
                   (source_device, snapshot_at, row_count, total_value, app_version)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(source_device) DO UPDATE SET
                   snapshot_at = excluded.snapshot_at,
                   row_count   = excluded.row_count,
                   total_value = excluded.total_value,
                   app_version = excluded.app_version""",
            (source_name, snapshot_at, inserted, total_value, app_version),
        )
    else:
        conn.execute(
            """INSERT INTO pos_stocks_snapshot_meta
                   (source_device, snapshot_at, row_count, total_value)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(source_device) DO UPDATE SET
                   snapshot_at = excluded.snapshot_at,
                   row_count   = excluded.row_count,
                   total_value = excluded.total_value""",
            (source_name, snapshot_at, inserted, total_value),
        )

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
    gross_amount: Optional[float] = None,
    cash_amount: Optional[float] = None,
) -> None:
    if not src:
        raise ApplyError("ledger row missing source_device")
    conn.execute(
        """
        INSERT OR IGNORE INTO pos_financial_ledger
            (source_device, event_uuid, event_type, category, amount, gross_amount, cash_amount, day,
             related_id, meta_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            src,
            event_uuid,
            event_type,
            category,
            float(amount),
            None if gross_amount is None else float(gross_amount),
            None if cash_amount is None else float(cash_amount),
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


def _shift_key_from_payload(payload: Dict[str, Any]) -> str:
    su = _clean(payload.get("shift_uuid"))
    if su:
        return su
    sid = payload.get("shift_id")
    try:
        return f"id:{int(sid)}" if sid is not None else ""
    except (TypeError, ValueError):
        return ""


def apply_wh_pos_shift_opened(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SHIFT_OPENED: missing source_device")
    key = _shift_key_from_payload(payload)
    if not key:
        raise ApplyError("SHIFT_OPENED: missing shift id")
    sid = payload.get("shift_id")
    try:
        sid_int = int(sid) if sid is not None else None
    except (TypeError, ValueError):
        sid_int = None
    started_at = _clean(payload.get("started_at")) or _now_iso()
    now = _now_iso()
    conn.execute(
        """
        INSERT INTO pos_shifts_mirror
            (source_device, shift_key, shift_id, started_at, ended_at, status,
             summary_json, last_event_uuid, updated_at)
        VALUES (?, ?, ?, ?, NULL, 'OPEN', '', ?, ?)
        ON CONFLICT(source_device, shift_key) DO UPDATE SET
            shift_id = COALESCE(excluded.shift_id, pos_shifts_mirror.shift_id),
            started_at = COALESCE(NULLIF(excluded.started_at, ''), pos_shifts_mirror.started_at),
            status = CASE
                WHEN pos_shifts_mirror.status = 'CLOSED' THEN pos_shifts_mirror.status
                ELSE 'OPEN'
            END,
            last_event_uuid = excluded.last_event_uuid,
            updated_at = excluded.updated_at
        """,
        (src, key, sid_int, started_at, event_uuid, now),
    )
    return {"status": "OPEN"}


def apply_wh_pos_shift_closed(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("SHIFT_CLOSED: missing source_device")
    key = _shift_key_from_payload(payload)
    if not key:
        raise ApplyError("SHIFT_CLOSED: missing shift id")
    sid = payload.get("shift_id")
    try:
        sid_int = int(sid) if sid is not None else None
    except (TypeError, ValueError):
        sid_int = None
    ended_at = _clean(payload.get("ended_at")) or _now_iso()
    summary_json = payload.get("summary_json")
    now = _now_iso()
    conn.execute(
        """
        INSERT INTO pos_shifts_mirror
            (source_device, shift_key, shift_id, started_at, ended_at, status,
             summary_json, last_event_uuid, updated_at)
        VALUES (?, ?, ?, NULL, ?, 'CLOSED', ?, ?, ?)
        ON CONFLICT(source_device, shift_key) DO UPDATE SET
            shift_id = COALESCE(excluded.shift_id, pos_shifts_mirror.shift_id),
            ended_at = excluded.ended_at,
            status = 'CLOSED',
            summary_json = excluded.summary_json,
            last_event_uuid = excluded.last_event_uuid,
            updated_at = excluded.updated_at
        """,
        (src, key, sid_int, ended_at, str(summary_json or ""), event_uuid, now),
    )
    return {"status": "CLOSED"}


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
    line_totals: List[float] = []
    for line in lines:
        if not isinstance(line, dict):
            line_totals.append(0.0)
            continue
        try:
            line_totals.append(float(line.get("unit_price") or 0) * int(line.get("qty") or 0))
        except (TypeError, ValueError):
            line_totals.append(0.0)
    try:
        fallback_allocations = _allocate_reservation_down_payments(line_totals, paid_batch)
    except Exception:
        fallback_allocations = [paid_batch if idx == 0 else 0.0 for idx, _ in enumerate(lines)]
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    cash_amount = paid_batch if payment_method == "CASH" else 0.0
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
        if "paid_amount" in line:
            try:
                alloc_paid = float(line.get("paid_amount") or 0)
            except (TypeError, ValueError):
                alloc_paid = 0.0
        else:
            alloc_paid = float(fallback_allocations[idx] if idx < len(fallback_allocations) else 0.0)
        conn.execute(
            """
            INSERT INTO pos_reservations_mirror
                (source_device, reservation_key, customer, item_type, school, color, size,
                 qty, unit_price, total_amount, paid_amount, status, shift_id, last_event_uuid, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'معلق', ?, ?, ?)
            ON CONFLICT(source_device, reservation_key) DO UPDATE SET
                customer     = excluded.customer,
                item_type    = excluded.item_type,
                school       = excluded.school,
                color        = excluded.color,
                size         = excluded.size,
                qty          = excluded.qty,
                unit_price   = excluded.unit_price,
                total_amount = excluded.total_amount,
                paid_amount  = excluded.paid_amount,
                status       = 'معلق',
                shift_id     = excluded.shift_id,
                last_event_uuid = excluded.last_event_uuid,
                updated_at   = excluded.updated_at
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
            meta={"customer": customer, "payment_method": payment_method},
            gross_amount=0.0,
            cash_amount=cash_amount,
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
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    cash_amount = delta if payment_method == "CASH" else 0.0
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
            meta={"new_paid": new_paid, "payment_method": payment_method},
            gross_amount=0.0,
            cash_amount=cash_amount,
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
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    cash_amount = coll if payment_method == "CASH" else 0.0
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
            meta={"paid_amount_total": new_total, "payment_method": payment_method},
            gross_amount=0.0,
            cash_amount=cash_amount,
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
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    cash_amount = -abs(total) if payment_method == "CASH" else 0.0
    day = _ledger_day(conn, event_uuid)
    bid = payload.get("bill_id")
    try:
        bid_int = int(bid) if bid is not None else None
    except (TypeError, ValueError):
        bid_int = None
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    cash_amount = total if payment_method == "CASH" else 0.0
    _ledger_append(
        conn,
        src=src,
        event_uuid=event_uuid,
        event_type="SALE_CREATED",
        category="sale",
        amount=total,
        day=day,
        related_id=bid_int,
        meta={"customer": _clean(payload.get("customer")), "payment_method": payment_method},
        gross_amount=total,
        cash_amount=cash_amount,
    )
    return {"amount": total, "cash_amount": cash_amount, "payment_method": payment_method}


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
        meta={"payment_method": payment_method},
        gross_amount=-abs(total),
        cash_amount=cash_amount,
    )
    return {"amount": -abs(total), "cash_amount": cash_amount, "payment_method": payment_method}


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
        gross_amount=-abs(total),
        cash_amount=-abs(total),
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
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    cash_amount = diff if payment_method == "CASH" else 0.0
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
            "payment_method": payment_method,
        },
        gross_amount=diff,
        cash_amount=cash_amount,
    )
    return {"amount": diff, "cash_amount": cash_amount, "payment_method": payment_method}


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
    from_type = str(payload.get("from_bill_type") or "").upper()
    to_type = str(payload.get("to_bill_type") or "").upper()
    if from_type == "EXCHANGE" and to_type == "SALE":
        old_diff = float(payload.get("old_diff") or 0)
        new_total = float(payload.get("new_total") or 0)
        _ledger_append(
            conn,
            src=src,
            event_uuid=f"{event_uuid}:reverse_exchange",
            event_type="SALE_BILL_TYPE_CORRECTED",
            category="exchange_net",
            amount=-old_diff,
            day=day,
            related_id=bid_int,
            meta={
                "from_bill_type": _clean(payload.get("from_bill_type")),
                "to_bill_type": _clean(payload.get("to_bill_type")),
                "reason": _clean(payload.get("reason")),
                "correction_event_uuid": event_uuid,
            },
            gross_amount=-old_diff,
            cash_amount=-old_diff,
        )
        _ledger_append(
            conn,
            src=src,
            event_uuid=f"{event_uuid}:sale",
            event_type="SALE_BILL_TYPE_CORRECTED",
            category="sale",
            amount=new_total,
            day=day,
            related_id=bid_int,
            meta={
                "from_bill_type": _clean(payload.get("from_bill_type")),
                "to_bill_type": _clean(payload.get("to_bill_type")),
                "reason": _clean(payload.get("reason")),
                "correction_event_uuid": event_uuid,
            },
            gross_amount=new_total,
            cash_amount=new_total,
        )
        return {"amount": amount_delta, "reclassified": "EXCHANGE_TO_SALE"}
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
        gross_amount=amount_delta,
        cash_amount=amount_delta,
    )
    return {"amount": amount_delta}


def apply_shipment_receipt_reported(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Queue POS shipment verification report for warehouse decision."""
    shipment_uuid = _clean(payload.get("shipment_uuid")) or event_uuid
    src = _src_device(conn, payload, event_uuid) or _clean(payload.get("source_device")) or "POS"
    note = _clean(payload.get("note"))
    has_diff = int(bool(payload.get("has_diff")))
    try:
        payload_json = json.dumps(payload.get("lines") or [], ensure_ascii=False, default=str)
    except Exception:
        payload_json = "[]"
    conn.execute(
        """
        INSERT OR IGNORE INTO shipment_receipt_reviews(
            sync_event_uuid, shipment_uuid, source_device,
            payload_json, has_diff, note, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (event_uuid, shipment_uuid, src, payload_json, has_diff, note, _now_iso()),
    )
    return {"queued": True, "shipment_uuid": shipment_uuid, "has_diff": bool(has_diff)}


# ----------------------------- registry ------------------------------- #

ApplierFn = Callable[[sqlite3.Connection, Dict[str, Any], str], Dict[str, Any]]


_POS_REGISTRY: Dict[str, ApplierFn] = {
    "STOCK_TRANSFER_OUT": apply_stock_transfer_out,
    "PRICE_UPDATE":       apply_price_update,
    "CATALOG_UPSERT":     apply_catalog_upsert,
    "SPEC_RENAMED":       apply_spec_renamed,
}


_WAREHOUSE_REGISTRY: Dict[str, ApplierFn] = {
    "POS_STOCK_SNAPSHOT": apply_pos_stock_snapshot,
    "POS_STOCK_AUDIT_APPLIED": apply_pos_stock_audit_applied,
    "STOCK_RETURN_TO_WAREHOUSE": apply_stock_return_to_warehouse,
    "POS_TRANSFER_VIA_WAREHOUSE": apply_pos_transfer_via_warehouse,
    "SHIFT_OPENED": apply_wh_pos_shift_opened,
    "SHIFT_CLOSED": apply_wh_pos_shift_closed,
    "RESERVATION_CREATED": apply_wh_pos_reservation_created,
    "RESERVATION_PAYMENT_UPDATED": apply_wh_pos_reservation_payment_updated,
    "RESERVATION_COMPLETED": apply_wh_pos_reservation_completed,
    "RESERVATION_DELIVERED": apply_wh_pos_reservation_delivered,
    "SALE_CREATED": apply_wh_pos_ledger_sale_created,
    "SALE_RETURNED": apply_wh_pos_ledger_sale_returned,
    "SALE_VOIDED": apply_wh_pos_ledger_sale_voided,
    "SALE_EXCHANGED": apply_wh_pos_ledger_sale_exchanged,
    "SALE_BILL_TYPE_CORRECTED": apply_wh_pos_ledger_sale_bill_type_corrected,
    "SHIPMENT_RECEIPT_REPORTED": apply_shipment_receipt_reported,
}


def for_role(role: Optional[str]) -> Dict[str, ApplierFn]:
    """Return the applier dict appropriate for a device role."""
    if role == "pos":
        return _POS_REGISTRY
    if role == "warehouse":
        return _WAREHOUSE_REGISTRY
    return {}
