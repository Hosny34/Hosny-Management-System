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
from datetime import datetime, timezone, timedelta
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

TEST_POS_DEVICE_NAMES = {"POS-TEST"}


# ----------------------------- helpers -------------------------------- #

def _now_iso() -> str:
    """Local-time ISO timestamp matching existing `movements.ts` format."""
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _has_column(conn: sqlite3.Connection, table: str, column: str) -> bool:
    cur = conn.execute("PRAGMA table_info(%s)" % table)
    return any(r[1] == column for r in cur.fetchall())


def _dedupe_pos_stocks_mirror_exact_prices(conn: sqlite3.Connection) -> Dict[str, int]:
    """Collapse old duplicate mirror rows without merging different prices."""
    result = {"groups": 0, "deleted_rows": 0}
    try:
        groups = conn.execute(
            """
            SELECT source_device, item_type, school, color, size, unit_price, COUNT(*) AS c
              FROM pos_stocks_mirror
             GROUP BY source_device, item_type, school, color, size, unit_price
            HAVING COUNT(*) > 1
            """
        ).fetchall()
    except sqlite3.OperationalError:
        return result

    for g in groups:
        source_device, item_type, school, color, size, unit_price = g[:6]
        rows = conn.execute(
            """
            SELECT id, count, snapshot_at
              FROM pos_stocks_mirror
             WHERE source_device = ?
               AND item_type = ?
               AND school = ?
               AND color = ?
               AND size = ?
               AND unit_price = ?
             ORDER BY snapshot_at DESC, count DESC, id DESC
            """,
            (source_device, item_type, school, color, size, unit_price),
        ).fetchall()
        if not rows:
            continue
        keep = rows[0]
        keep_id = int(keep[0])
        keep_count = max(int((r[1] if len(r) > 1 else 0) or 0) for r in rows)
        keep_snapshot = str((keep[2] if len(keep) > 2 else "") or "")
        conn.execute(
            """
            UPDATE pos_stocks_mirror
               SET count = ?, snapshot_at = ?
             WHERE id = ?
            """,
            (keep_count, keep_snapshot, keep_id),
        )
        delete_ids = [int(r[0]) for r in rows if int(r[0]) != keep_id]
        if delete_ids:
            placeholders = ",".join("?" for _ in delete_ids)
            conn.execute(
                f"DELETE FROM pos_stocks_mirror WHERE id IN ({placeholders})",
                delete_ids,
            )
        result["groups"] += 1
        result["deleted_rows"] += len(delete_ids)
    return result


def _ensure_pos_stocks_mirror_price_unique_index(conn: sqlite3.Connection) -> None:
    """Allow one POS mirror spec to exist at multiple prices."""
    try:
        _dedupe_pos_stocks_mirror_exact_prices(conn)
        conn.execute("DROP INDEX IF EXISTS idx_pos_stocks_mirror_unique_spec")
        conn.execute(
            """
            CREATE UNIQUE INDEX IF NOT EXISTS idx_pos_stocks_mirror_unique_spec
            ON pos_stocks_mirror(source_device, item_type, school, color, size, unit_price)
            """
        )
    except (sqlite3.OperationalError, sqlite3.IntegrityError):
        pass


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _is_test_pos_device(value: Any) -> bool:
    return _clean(value).casefold() in {d.casefold() for d in TEST_POS_DEVICE_NAMES}


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
    total_due = sum(clean_totals)
    if abs(paid - total_due) <= 0.01:
        return [round(t, 2) for t in clean_totals]

    allocations: List[float] = [0.0 for _ in clean_totals]
    remaining = paid
    open_indexes = {idx for idx, total in enumerate(clean_totals) if total > 1e-9}
    while remaining > 0.01 and open_indexes:
        share = _round_up_to_step(remaining / len(open_indexes), round_step)
        progressed = False
        for idx in list(sorted(open_indexes)):
            capacity = max(0.0, clean_totals[idx] - allocations[idx])
            if capacity <= 0.01:
                open_indexes.discard(idx)
                continue
            alloc = min(share, capacity, remaining)
            allocations[idx] = round(allocations[idx] + alloc, 2)
            remaining = round(remaining - alloc, 2)
            progressed = True
            if clean_totals[idx] - allocations[idx] <= 0.01:
                open_indexes.discard(idx)
            if remaining <= 0.01:
                break
        if not progressed:
            break

    if remaining > 0.01:
        for idx in sorted(open_indexes):
            capacity = max(0.0, clean_totals[idx] - allocations[idx])
            alloc = min(capacity, remaining)
            allocations[idx] = round(allocations[idx] + alloc, 2)
            remaining = round(remaining - alloc, 2)
            if remaining <= 0.01:
                break
    return [min(round(a, 2), round(clean_totals[idx], 2)) for idx, a in enumerate(allocations)]


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


_MIRROR_ALPHA_SIZES = ["S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL"]
_MIRROR_ALLOWED_NUMERIC_RANGES: Dict[Tuple[int, int], List[str]] = {
    (0, 24): [str(i) for i in range(0, 26, 2)],
    (0, 16): [str(i) for i in range(0, 18, 2)],
    (6, 22): [str(i) for i in range(6, 24, 2)],
    (14, 28): [str(i) for i in range(14, 30, 2)],
    (18, 30): [str(i) for i in range(18, 32, 2)],
    (0, 9): [str(i) for i in range(0, 10, 1)],
    (30, 62): [str(i) for i in range(30, 63, 2)],
}


def _has_table(conn: sqlite3.Connection, table: str) -> bool:
    try:
        row = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type='table' AND name=? LIMIT 1",
            (table,),
        ).fetchone()
        return row is not None
    except sqlite3.OperationalError:
        return False


def _mirror_size_sort_key(size: Any) -> Tuple[Any, ...]:
    text = _clean(size).upper()
    if text.isdigit():
        return (0, int(text), "")
    rank = {label: idx for idx, label in enumerate(_MIRROR_ALPHA_SIZES, start=1)}
    return (1, rank.get(text, 999), text.casefold())


def _mirror_profile_labels(profile: Tuple[Optional[int], Optional[int], Optional[int], Optional[int], int]) -> List[str]:
    r1s, r1e, r2s, r2e, has_alpha = profile
    labels: Dict[str, str] = {}
    for start, end in ((r1s, r1e), (r2s, r2e)):
        if start is None or end is None:
            continue
        for label in _MIRROR_ALLOWED_NUMERIC_RANGES.get((int(start), int(end)), []):
            labels.setdefault(label.casefold(), label)
    if int(has_alpha or 0):
        for label in _MIRROR_ALPHA_SIZES:
            labels.setdefault(label.casefold(), label)
    values = list(labels.values())
    values.sort(key=_mirror_size_sort_key)
    return values


def _infer_mirror_size_profile(sizes: Sequence[str]) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], int]:
    numeric = sorted({int(s) for s in sizes if _clean(s).isdigit()})
    alpha = {_clean(s).upper() for s in sizes if _clean(s).upper() in _MIRROR_ALPHA_SIZES}
    if not numeric:
        return (None, None, None, None, int(bool(alpha)))

    observed = {str(n) for n in numeric}
    ranges = list(_MIRROR_ALLOWED_NUMERIC_RANGES.items())
    covering = [
        (rng, labels)
        for rng, labels in ranges
        if observed.issubset(set(labels))
    ]
    if covering:
        rng, _labels = min(covering, key=lambda pair: (len(pair[1]) - len(observed), pair[0][1] - pair[0][0]))
        return (rng[0], rng[1], None, None, int(bool(alpha)))

    best_pair: Optional[Tuple[Tuple[int, int], Tuple[int, int], int]] = None
    for idx, (r1, labels1) in enumerate(ranges):
        set1 = set(labels1)
        for r2, labels2 in ranges[idx + 1:]:
            union = set1 | set(labels2)
            if observed.issubset(union):
                extra = len(union) - len(observed)
                candidate = (r1, r2, extra)
                if best_pair is None or candidate[2] < best_pair[2]:
                    best_pair = candidate
    if best_pair:
        r1, r2, _extra = best_pair
        return (r1[0], r1[1], r2[0], r2[1], int(bool(alpha)))

    return (0, 24, None, None, int(bool(alpha)))


def _warehouse_catalog_location(conn: sqlite3.Connection) -> Tuple[int, int]:
    try:
        row = conn.execute(
            """
            SELECT warehouse_no, package_no
              FROM stocks
             WHERE warehouse_no IS NOT NULL AND package_no IS NOT NULL
             GROUP BY warehouse_no, package_no
             ORDER BY COUNT(*) DESC, MIN(id) ASC
             LIMIT 1
            """
        ).fetchone()
        if row:
            return int(row[0] or 1), int(row[1] or 1)
    except sqlite3.OperationalError:
        pass
    return 1, 1


def _warehouse_catalog_spec_exists(
    conn: sqlite3.Connection,
    item_type: str,
    school: str,
    color: str,
    size: str,
) -> bool:
    row = conn.execute(
        """
        SELECT 1
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
         LIMIT 1
        """,
        (item_type, school, color, size),
    ).fetchone()
    return row is not None


def _ensure_warehouse_zero_stock_catalog_row(
    conn: sqlite3.Connection,
    item_type: str,
    school: str,
    color: str,
    size: str,
    unit_price: float,
) -> bool:
    if not _has_column(conn, "stocks", "warehouse_no"):
        return _upsert_zero_stock_catalog_row(conn, item_type, school, color, size, unit_price)

    row = conn.execute(
        """
        SELECT id, count, unit_price
          FROM stocks
         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
         ORDER BY COALESCE(count, 0) DESC, id ASC
         LIMIT 1
        """,
        (item_type, school, color, size),
    ).fetchone()
    if row:
        row_id = row["id"] if isinstance(row, sqlite3.Row) else row[0]
        row_count = row["count"] if isinstance(row, sqlite3.Row) else row[1]
        row_price = row["unit_price"] if isinstance(row, sqlite3.Row) else row[2]
        if (
            int(row_count or 0) == 0
            and float(unit_price or 0) > 0
            and abs(float(row_price or 0) - float(unit_price)) >= 0.001
        ):
            conn.execute("UPDATE stocks SET unit_price=? WHERE id=?", (float(unit_price), int(row_id)))
        return False

    warehouse_no, package_no = _warehouse_catalog_location(conn)
    conn.execute(
        """
        INSERT INTO stocks(
            item_type, school, color, size,
            warehouse_no, package_no, unit_price, count
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, 0)
        """,
        (item_type, school, color, size, int(warehouse_no), int(package_no), float(unit_price or 0)),
    )
    return True


def _ensure_mirror_import_price_profile(
    conn: sqlite3.Connection,
    source_device: str,
    item_type: str,
    school: str,
    color: str,
    prices_by_size: Dict[str, float],
) -> int:
    if not all(_has_table(conn, t) for t in ("price_profiles", "price_profile_lines", "price_profile_assignments")):
        return 0

    row = conn.execute(
        """
        SELECT profile_id
          FROM price_profile_assignments
         WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
         ORDER BY
             CASE
                 WHEN COALESCE(TRIM(item_type), '') = ''
                  AND COALESCE(TRIM(color), '') = '' THEN 0
                 ELSE 1
             END,
             updated_at DESC,
             id DESC
         LIMIT 1
        """,
        (school,),
    ).fetchone()
    profile_id: Optional[int] = None
    created_assignment = False
    if row:
        profile_id = int((row["profile_id"] if isinstance(row, sqlite3.Row) else row[0]) or 0) or None
    if profile_id is None:
        base_name = f"Mirror import - {source_device} - {school}"
        name = base_name[:140]
        conn.execute(
            """
            INSERT OR IGNORE INTO price_profiles(name, notes, created_at, updated_at)
            VALUES (?, ?, datetime('now'), datetime('now'))
            """,
            (name, "Created automatically from POS mirror definitions."),
        )
        row = conn.execute("SELECT id FROM price_profiles WHERE name=?", (name,)).fetchone()
        if not row:
            return 0
        profile_id = int((row["id"] if isinstance(row, sqlite3.Row) else row[0]) or 0)
        conn.execute(
            """
            INSERT INTO price_profile_assignments(item_type, school, color, profile_id, updated_at)
            VALUES (?, ?, ?, ?, datetime('now'))
            ON CONFLICT(item_type, school, color) DO NOTHING
            """,
            ("", school, "", profile_id),
        )
        created_assignment = True

    written = 0
    for size, price in prices_by_size.items():
        if not size:
            continue
        existing = conn.execute(
            """
            SELECT id, price
              FROM price_profile_lines
             WHERE profile_id=?
               AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
             LIMIT 1
            """,
            (int(profile_id), item_type, size),
        ).fetchone()
        if existing:
            line_id = existing["id"] if isinstance(existing, sqlite3.Row) else existing[0]
            old_price = float((existing["price"] if isinstance(existing, sqlite3.Row) else existing[1]) or 0)
            if created_assignment and float(price or 0) > 0 and abs(old_price - float(price)) >= 0.001:
                conn.execute(
                    "UPDATE price_profile_lines SET price=?, updated_at=datetime('now') WHERE id=?",
                    (float(price), int(line_id)),
                )
                written += 1
            continue
        conn.execute(
            """
            INSERT INTO price_profile_lines(profile_id, item_type, size, price, created_at, updated_at)
            VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
            """,
            (int(profile_id), item_type, size, float(price or 0)),
        )
        written += 1
    return written


def ensure_warehouse_catalog_from_pos_mirror(
    conn: sqlite3.Connection,
    source_device: Optional[str] = None,
) -> Dict[str, int]:
    """Import missing Warehouse definitions from POS mirror rows.

    This repairs catalog metadata only. It never copies branch quantities into
    warehouse stock and never writes movements.
    """
    if not _has_table(conn, "pos_stocks_mirror") or not _has_table(conn, "stocks"):
        return {"groups": 0, "stock_rows": 0, "size_profiles": 0, "price_profile_lines": 0}

    where = [
        "COALESCE(TRIM(item_type), '') <> ''",
        "COALESCE(TRIM(school), '') <> ''",
        "COALESCE(TRIM(color), '') <> ''",
        "COALESCE(TRIM(size), '') <> ''",
    ]
    args: List[Any] = []
    if source_device:
        if _is_test_pos_device(source_device):
            return {"groups": 0, "stock_rows": 0, "size_profiles": 0, "price_profile_lines": 0}
        where.append("source_device = ?")
        args.append(source_device)
    else:
        where.append("UPPER(TRIM(source_device)) NOT IN (%s)" % ",".join("?" for _ in TEST_POS_DEVICE_NAMES))
        args.extend(sorted(TEST_POS_DEVICE_NAMES))

    rows = conn.execute(
        f"""
        SELECT source_device, item_type, school, color, size, unit_price
          FROM pos_stocks_mirror
         WHERE {' AND '.join(where)}
         ORDER BY source_device, item_type, school, color, size, unit_price
        """,
        tuple(args),
    ).fetchall()

    groups: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}
    for row in rows:
        src = _clean(row["source_device"] if isinstance(row, sqlite3.Row) else row[0])
        if _is_test_pos_device(src):
            continue
        it = _clean(row["item_type"] if isinstance(row, sqlite3.Row) else row[1])
        sc = _clean(row["school"] if isinstance(row, sqlite3.Row) else row[2])
        cl = _clean(row["color"] if isinstance(row, sqlite3.Row) else row[3])
        sz = _clean(row["size"] if isinstance(row, sqlite3.Row) else row[4])
        try:
            price = float((row["unit_price"] if isinstance(row, sqlite3.Row) else row[5]) or 0)
        except (TypeError, ValueError):
            price = 0.0
        key = (src.casefold(), it.casefold(), sc.casefold(), cl.casefold())
        group = groups.setdefault(
            key,
            {
                "source_device": src,
                "item_type": it,
                "school": sc,
                "color": cl,
                "prices_by_size": {},
                "sizes": set(),
            },
        )
        group["sizes"].add(sz)
        current_price = group["prices_by_size"].get(sz)
        if current_price is None or (float(current_price or 0) <= 0 < price):
            group["prices_by_size"][sz] = price

    stock_rows = 0
    size_profiles = 0
    price_lines = 0
    for group in groups.values():
        it = group["item_type"]
        sc = group["school"]
        cl = group["color"]
        src = group["source_device"]
        observed_sizes = sorted(group["sizes"], key=_mirror_size_sort_key)
        prices_by_size = dict(group["prices_by_size"])

        profile_exists = conn.execute(
            """
            SELECT 1
              FROM size_profiles
             WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
             LIMIT 1
            """,
            (it, sc, cl),
        ).fetchone()
        if not profile_exists:
            profile = _infer_mirror_size_profile(observed_sizes)
            conn.execute(
                """
                INSERT INTO size_profiles(
                    item_type, school, color,
                    num_start_1, num_end_1, num_start_2, num_end_2,
                    has_alpha, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
                ON CONFLICT(item_type, school, color) DO NOTHING
                """,
                (it, sc, cl, profile[0], profile[1], profile[2], profile[3], int(profile[4] or 0)),
            )
            size_profiles += 1
            for label in _mirror_profile_labels(profile):
                prices_by_size.setdefault(label, 0.0)

        price_lines += _ensure_mirror_import_price_profile(conn, src, it, sc, cl, prices_by_size)

        for size in sorted(prices_by_size.keys(), key=_mirror_size_sort_key):
            if _ensure_warehouse_zero_stock_catalog_row(conn, it, sc, cl, size, float(prices_by_size.get(size) or 0)):
                stock_rows += 1
            _upsert_spec_history(conn, {"item_type": it, "school": sc, "color": cl, "size": size})

    return {
        "groups": len(groups),
        "stock_rows": int(stock_rows),
        "size_profiles": int(size_profiles),
        "price_profile_lines": int(price_lines),
    }


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
            "pos_stocks_mirror",
            "pos_reservations_mirror",
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
    if old_spec == new_spec:
        value_updates = _apply_spec_value_renames(conn, payload)
        if value_updates:
            return {"updated_rows": int(value_updates), "old_spec": old_spec, "new_spec": new_spec, "event_uuid": str(event_uuid)}
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

    updated_tables += _apply_spec_value_renames(conn, payload)

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
    _ensure_pos_stocks_mirror_price_unique_index(conn)

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
            # Some live DBs/exes have a stale mirror unique index.  The snapshot
            # path deletes this source_device first, and collapsed_rows is
            # already unique by source/spec/price, so a plain insert preserves
            # the intended full-snapshot mirror even if ON CONFLICT is unusable.
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
    imported_catalog = ensure_warehouse_catalog_from_pos_mirror(conn, source_name)

    if _has_column(conn, "pos_stocks_snapshot_meta", "app_version"):
        conn.execute(
            """INSERT INTO pos_stocks_snapshot_meta
                   (source_device, snapshot_at, row_count, total_value, app_version, last_server_seq)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(source_device) DO UPDATE SET
                   snapshot_at = excluded.snapshot_at,
                   row_count   = excluded.row_count,
                   total_value = excluded.total_value,
                   app_version = excluded.app_version,
                   last_server_seq = excluded.last_server_seq""",
            (source_name, snapshot_at, inserted, total_value, app_version, inbound_seq),
        )
    else:
        conn.execute(
            """INSERT INTO pos_stocks_snapshot_meta
                   (source_device, snapshot_at, row_count, total_value, last_server_seq)
               VALUES (?, ?, ?, ?, ?)
               ON CONFLICT(source_device) DO UPDATE SET
                   snapshot_at = excluded.snapshot_at,
                   row_count   = excluded.row_count,
                   total_value = excluded.total_value,
                   last_server_seq = excluded.last_server_seq""",
            (source_name, snapshot_at, inserted, total_value, inbound_seq),
        )

    return {"mirrored_rows": inserted, "total_value": total_value, "catalog_import": imported_catalog}


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
            """INSERT INTO pos_stocks_snapshot_meta
                   (source_device, snapshot_at, row_count, total_value)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(source_device) DO UPDATE SET
                   snapshot_at = excluded.snapshot_at,
                   row_count   = excluded.row_count,
                   total_value = excluded.total_value""",
            (source_name, snapshot_at, row_count, total_value),
        )
    else:
        current = conn.execute(
            "SELECT snapshot_at FROM pos_stocks_snapshot_meta WHERE source_device = ?",
            (source_name,),
        ).fetchone()
        preserved_snapshot_at = _clean(current[0] if current else "") or snapshot_at
        conn.execute(
            """INSERT INTO pos_stocks_snapshot_meta
                   (source_device, snapshot_at, row_count, total_value)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(source_device) DO UPDATE SET
                   row_count   = excluded.row_count,
                   total_value = excluded.total_value""",
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


def apply_pos_stock_audit_snapshot(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Warehouse-side applier for a POS audit-report snapshot."""
    source_name = _clean(payload.get("source_device_name")) or _src_device(conn, payload, event_uuid)
    if not source_name:
        raise ApplyError("POS_STOCK_AUDIT_SNAPSHOT: missing source device")
    reports = payload.get("reports") or []
    if not isinstance(reports, list):
        raise ApplyError("POS_STOCK_AUDIT_SNAPSHOT: reports must be a list")

    conn.execute(
        "DELETE FROM pos_stock_audit_items_mirror WHERE source_device = ?",
        (source_name,),
    )
    conn.execute(
        "DELETE FROM pos_stock_audit_reports_mirror WHERE source_device = ?",
        (source_name,),
    )

    report_count = 0
    line_count = 0
    received_at = _now_iso()
    for report in reports:
        if not isinstance(report, dict):
            continue
        try:
            report_id = int(report.get("report_id"))
        except (TypeError, ValueError):
            continue
        audit_uuid = _clean(report.get("audit_uuid")) or f"{source_name}:{report_id}"
        created_at = _clean(report.get("created_at")) or received_at
        reason = _clean(report.get("reason"))
        try:
            total_diff = int(report.get("total_diff") or 0)
            total_value = float(report.get("total_value") or 0)
        except (TypeError, ValueError):
            total_diff = 0
            total_value = 0.0
        conn.execute(
            "DELETE FROM pos_stock_audit_items_mirror WHERE audit_uuid = ?",
            (audit_uuid,),
        )
        conn.execute(
            "DELETE FROM pos_stock_audit_reports_mirror WHERE audit_uuid = ?",
            (audit_uuid,),
        )
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
                received_at,
            ),
        )
        report_count += 1
        lines = report.get("lines") or []
        if not isinstance(lines, list):
            continue
        for raw in lines:
            if not isinstance(raw, dict):
                continue
            it = _clean(raw.get("item_type"))
            sc = _clean(raw.get("school"))
            cl = _clean(raw.get("color"))
            sz = _clean(raw.get("size"))
            if not (it and sc and cl and sz):
                continue
            try:
                expected = int(raw.get("expected") or 0)
                actual = int(raw.get("actual") or 0)
                diff = int(raw.get("diff", actual - expected) or 0)
                price = float(raw.get("unit_price") or 0)
                diff_value = float(raw.get("diff_value", diff * price) or 0)
            except (TypeError, ValueError):
                continue
            conn.execute(
                """
                INSERT INTO pos_stock_audit_items_mirror
                    (audit_uuid, source_device, item_type, school, color, size,
                     expected_qty, actual_qty, diff_qty, unit_price, diff_value)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (audit_uuid, source_name, it, sc, cl, sz, expected, actual, diff, price, diff_value),
            )
            line_count += 1
    return {"source_device": source_name, "reports": report_count, "lines": line_count}


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


def _payload_day(payload: Dict[str, Any]) -> str:
    for key in ("business_day", "created_at", "voided_at", "ended_at", "started_at"):
        value = _clean(payload.get(key))
        if len(value) >= 10:
            return value[:10]
    return ""


def _ledger_day(conn: sqlite3.Connection, event_uuid: str, payload: Optional[Dict[str, Any]] = None) -> str:
    if payload:
        day = _payload_day(payload)
        if day:
            return day
    try:
        row = conn.execute(
            """
            SELECT substr(COALESCE(NULLIF(server_created_at, ''), applied_at), 1, 10)
              FROM sync_inbox
             WHERE event_uuid = ?
            """,
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
    payment_method: Optional[str] = None,
    shift_id: Optional[int] = None,
) -> None:
    if not src:
        raise ApplyError("ledger row missing source_device")
    conn.execute(
        """
        INSERT OR IGNORE INTO pos_financial_ledger
            (source_device, event_uuid, event_type, category, amount, gross_amount, cash_amount,
             payment_method, shift_id, day,
             related_id, meta_json, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            src,
            event_uuid,
            event_type,
            category,
            float(amount),
            None if gross_amount is None else float(gross_amount),
            None if cash_amount is None else float(cash_amount),
            _clean(payment_method).upper() if payment_method else None,
            None if shift_id is None else int(shift_id),
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


def _shift_id_from_payload(payload: Dict[str, Any]) -> Optional[int]:
    sid = payload.get("shift_id")
    try:
        return int(sid) if sid is not None else None
    except (TypeError, ValueError):
        return None


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
    started_at = _clean(payload.get("started_at"))
    ended_at = _clean(payload.get("ended_at")) or _now_iso()
    summary_json = payload.get("summary_json")
    now = _now_iso()
    conn.execute(
        """
        INSERT INTO pos_shifts_mirror
            (source_device, shift_key, shift_id, started_at, ended_at, status,
             summary_json, last_event_uuid, updated_at)
        VALUES (?, ?, ?, ?, ?, 'CLOSED', ?, ?, ?)
        ON CONFLICT(source_device, shift_key) DO UPDATE SET
            shift_id = COALESCE(excluded.shift_id, pos_shifts_mirror.shift_id),
            started_at = COALESCE(NULLIF(excluded.started_at, ''), pos_shifts_mirror.started_at),
            ended_at = excluded.ended_at,
            status = 'CLOSED',
            summary_json = excluded.summary_json,
            last_event_uuid = excluded.last_event_uuid,
            updated_at = excluded.updated_at
        """,
        (src, key, sid_int, started_at or None, ended_at, str(summary_json or ""), event_uuid, now),
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
    day = _ledger_day(conn, event_uuid, payload)
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
            payment_method=payment_method,
            shift_id=shift_id,
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
    day = _ledger_day(conn, event_uuid, payload)
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
            payment_method=payment_method,
            shift_id=_shift_id_from_payload(payload),
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
    day = _ledger_day(conn, event_uuid, payload)
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
            payment_method=payment_method,
            shift_id=_shift_id_from_payload(payload),
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
    day = _ledger_day(conn, event_uuid, payload)
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
        payment_method=payment_method,
        shift_id=_shift_id_from_payload(payload),
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
    bill_type = _clean(payload.get("bill_type")).upper() or "SALE"
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    signed_amount = abs(total) if bill_type == "RETURN" else -abs(total)
    cash_amount = signed_amount if payment_method == "CASH" else 0.0
    day = _ledger_day(conn, event_uuid, payload)
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
        payment_method=payment_method,
        shift_id=_shift_id_from_payload(payload),
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
    bill_type = _clean(payload.get("bill_type")).upper() or "SALE"
    payment_method = _clean(payload.get("payment_method")).upper() or "CASH"
    signed_amount = abs(total) if bill_type == "RETURN" else -abs(total)
    cash_amount = signed_amount if payment_method == "CASH" else 0.0
    day = _ledger_day(conn, event_uuid, payload)
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
        amount=signed_amount,
        day=day,
        related_id=bid_int,
        meta={"reason": _clean(payload.get("reason")), "payment_method": payment_method, "bill_type": bill_type},
        gross_amount=signed_amount,
        cash_amount=cash_amount,
        payment_method=payment_method,
        shift_id=_shift_id_from_payload(payload),
    )
    return {"amount": signed_amount, "cash_amount": cash_amount, "payment_method": payment_method}


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
    day = _ledger_day(conn, event_uuid, payload)
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
        payment_method=payment_method,
        shift_id=_shift_id_from_payload(payload),
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
    day = _ledger_day(conn, event_uuid, payload)
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


def apply_pos_financial_snapshot(
    conn: sqlite3.Connection,
    payload: Dict[str, Any],
    event_uuid: str,
) -> Dict[str, Any]:
    """Store POS-computed daily cash/Visa totals for the branch monitor."""
    src = _src_device(conn, payload, event_uuid)
    if not src:
        raise ApplyError("POS_FINANCIAL_SNAPSHOT: missing source_device")
    day = _clean(payload.get("day"))[:10]
    if not day:
        raise ApplyError("POS_FINANCIAL_SNAPSHOT: missing day")
    try:
        cash_total = float(payload.get("cash_total") or 0.0)
        visa_total = float(payload.get("visa_total") or 0.0)
        total_collected = float(payload.get("total_collected") or (cash_total + visa_total))
    except (TypeError, ValueError):
        raise ApplyError("POS_FINANCIAL_SNAPSHOT: invalid totals")

    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS pos_financial_daily_snapshot (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source_device TEXT NOT NULL,
            source_device_name TEXT,
            day TEXT NOT NULL,
            cash_total REAL NOT NULL DEFAULT 0,
            visa_total REAL NOT NULL DEFAULT 0,
            total_collected REAL NOT NULL DEFAULT 0,
            snapshot_at TEXT NOT NULL,
            app_version TEXT,
            event_uuid TEXT NOT NULL UNIQUE,
            received_at TEXT NOT NULL
        )
        """
    )
    conn.execute(
        """
        CREATE INDEX IF NOT EXISTS ix_pos_fin_daily_snapshot_source_day
            ON pos_financial_daily_snapshot(source_device, day, snapshot_at)
        """
    )
    conn.execute(
        """
        INSERT OR IGNORE INTO pos_financial_daily_snapshot
            (source_device, source_device_name, day, cash_total, visa_total,
             total_collected, snapshot_at, app_version, event_uuid, received_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            src,
            _clean(payload.get("source_device_name")) or None,
            day,
            cash_total,
            visa_total,
            total_collected,
            _clean(payload.get("snapshot_at")) or _now_iso(),
            _clean(payload.get("app_version")) or None,
            event_uuid,
            _now_iso(),
        ),
    )
    return {
        "source_device": src,
        "day": day,
        "cash_total": cash_total,
        "visa_total": visa_total,
        "total_collected": total_collected,
    }


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
    "POS_STOCK_AUDIT_SNAPSHOT": apply_pos_stock_audit_snapshot,
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
    "POS_FINANCIAL_SNAPSHOT": apply_pos_financial_snapshot,
    "SHIPMENT_RECEIPT_REPORTED": apply_shipment_receipt_reported,
}


def for_role(role: Optional[str]) -> Dict[str, ApplierFn]:
    """Return the applier dict appropriate for a device role."""
    if role == "pos":
        return _POS_REGISTRY
    if role == "warehouse":
        return _WAREHOUSE_REGISTRY
    return {}
