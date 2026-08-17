from __future__ import annotations

import argparse
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any


def _connect(path: Path, readonly: bool = False) -> sqlite3.Connection:
    if readonly:
        conn = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    else:
        conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


def _backup_database(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_path = path.with_name(f"{path.stem}-before-owned-catalog-repair-{stamp}{path.suffix}")
    src = _connect(path, readonly=True)
    try:
        dst = sqlite3.connect(str(backup_path))
        try:
            src.backup(dst)
        finally:
            dst.close()
    finally:
        src.close()
    return backup_path


def _summary(conn: sqlite3.Connection) -> dict[str, Any]:
    rows = conn.execute(
        """
        SELECT school,
               COUNT(*) AS rows,
               SUM(CASE WHEN COALESCE(count, 0) = 0 THEN 1 ELSE 0 END) AS zero_rows,
               SUM(CASE WHEN COALESCE(count, 0) > 0 THEN 1 ELSE 0 END) AS positive_rows,
               COALESCE(SUM(count), 0) AS qty,
               COALESCE(SUM(count * unit_price), 0) AS value
          FROM stocks
         GROUP BY school
         ORDER BY rows DESC
        """
    ).fetchall()
    return {
        "schools": [
            {
                "school": row["school"],
                "rows": int(row["rows"] or 0),
                "zero_rows": int(row["zero_rows"] or 0),
                "positive_rows": int(row["positive_rows"] or 0),
                "qty": int(row["qty"] or 0),
                "value": float(row["value"] or 0),
            }
            for row in rows
        ],
    }


def _prepare_owned_specs(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        DROP TABLE IF EXISTS temp._owned_branch_specs;
        CREATE TEMP TABLE _owned_branch_specs(
            item_type TEXT NOT NULL,
            school TEXT NOT NULL,
            color TEXT NOT NULL,
            size TEXT NOT NULL,
            PRIMARY KEY(item_type, school, color, size)
        );
        """
    )
    for row in conn.execute("SELECT event_uuid, payload_json FROM sync_inbox WHERE event_type='STOCK_TRANSFER_OUT'"):
        try:
            payload = json.loads(row["payload_json"] or "{}")
        except Exception:
            continue
        note_text = str(payload.get("note") or "")
        is_reservation_definition = "reservation" in note_text.casefold() or "حجز" in note_text
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            try:
                qty = int(float(item.get("qty") or 0))
            except (TypeError, ValueError):
                qty = 0
            owns_spec = qty > 0 or (bool(item.get("catalog_only")) and is_reservation_definition)
            if not owns_spec:
                continue
            values = [str(item.get(k) or "").strip() for k in ("item_type", "school", "color", "size")]
            if all(values):
                conn.execute(
                    """
                    INSERT OR IGNORE INTO _owned_branch_specs(item_type, school, color, size)
                    VALUES (LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)))
                    """,
                    values,
                )

    for table_name, where_sql in (
        ("incoming_shipment_items_pending", "(COALESCE(expected_qty, 0) > 0 OR COALESCE(received_qty, 0) > 0)"),
        ("movements", "UPPER(COALESCE(direction, '')) = 'IN' AND COALESCE(qty, 0) > 0"),
        (
            "branch_catalog_definitions",
            "LOWER(COALESCE(note, '')) LIKE '%reservation%' OR COALESCE(note, '') LIKE '%حجز%'",
        ),
    ):
        try:
            conn.execute(
                f"""
                INSERT OR IGNORE INTO _owned_branch_specs(item_type, school, color, size)
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
            continue


def repair(path: Path, apply: bool) -> dict[str, Any]:
    conn = _connect(path, readonly=not apply)
    try:
        _prepare_owned_specs(conn)
        before = _summary(conn)
        unowned = conn.execute(
            """
            SELECT COUNT(*) AS rows,
                   SUM(CASE WHEN COALESCE(count, 0) = 0 THEN 1 ELSE 0 END) AS zero_rows,
                   SUM(CASE WHEN COALESCE(count, 0) > 0 THEN 1 ELSE 0 END) AS positive_rows
              FROM stocks AS s
             WHERE NOT EXISTS (
                    SELECT 1 FROM _owned_branch_specs AS o
                     WHERE o.item_type = LOWER(TRIM(COALESCE(s.item_type, '')))
                       AND o.school = LOWER(TRIM(COALESCE(s.school, '')))
                       AND o.color = LOWER(TRIM(COALESCE(s.color, '')))
                       AND o.size = LOWER(TRIM(COALESCE(s.size, '')))
               )
            """
        ).fetchone()
        result: dict[str, Any] = {
            "database": str(path),
            "applied": bool(apply),
            "backup": None,
            "unowned_rows": int(unowned["rows"] or 0),
            "unowned_zero_rows": int(unowned["zero_rows"] or 0),
            "unowned_positive_rows": int(unowned["positive_rows"] or 0),
            "deleted_stock_rows": 0,
            "deleted_size_profiles": 0,
            "deleted_spec_history": 0,
            "before": before,
            "after": None,
        }
        if not apply:
            return result
        if result["unowned_positive_rows"]:
            raise RuntimeError(
                "Refusing automatic repair because unowned positive stock rows were found. "
                "Inspect the DB before deleting stock."
            )

        result["backup"] = str(_backup_database(path))
        with conn:
            result["deleted_stock_rows"] = int(
                conn.execute(
                    """
                    DELETE FROM stocks AS s
                     WHERE COALESCE(count, 0) = 0
                       AND NOT EXISTS (
                            SELECT 1 FROM _owned_branch_specs AS o
                             WHERE o.item_type = LOWER(TRIM(COALESCE(s.item_type, '')))
                               AND o.school = LOWER(TRIM(COALESCE(s.school, '')))
                               AND o.color = LOWER(TRIM(COALESCE(s.color, '')))
                               AND o.size = LOWER(TRIM(COALESCE(s.size, '')))
                       )
                    """
                ).rowcount
                or 0
            )
            result["deleted_size_profiles"] = int(
                conn.execute(
                    """
                    DELETE FROM size_profiles AS p
                     WHERE NOT EXISTS (
                            SELECT 1 FROM _owned_branch_specs AS o
                             WHERE o.item_type = LOWER(TRIM(COALESCE(p.item_type, '')))
                               AND o.school = LOWER(TRIM(COALESCE(p.school, '')))
                               AND o.color = LOWER(TRIM(COALESCE(p.color, '')))
                       )
                    """
                ).rowcount
                or 0
            )
            cur = conn.execute(
                """
                DELETE FROM spec_history
                 WHERE field='school'
                   AND value NOT IN (SELECT DISTINCT school FROM stocks)
                """
            )
            result["deleted_spec_history"] = int(cur.rowcount or 0)
        result["after"] = _summary(conn)
        return result
    finally:
        try:
            conn.execute("DROP TABLE IF EXISTS temp._owned_branch_specs")
        except Exception:
            pass
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Remove zero POS catalog rows that have no shipment or reservation evidence.")
    parser.add_argument("db", type=Path)
    parser.add_argument("--apply", action="store_true", help="Apply the repair. Without this, prints a dry-run preview.")
    args = parser.parse_args()
    print(json.dumps(repair(args.db, bool(args.apply)), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
