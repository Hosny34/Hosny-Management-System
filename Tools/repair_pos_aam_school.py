from __future__ import annotations

import argparse
import json
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable


AAM_SCHOOL = "\u0639\u0627\u0645"
DEFAULT_WRONG_SCHOOL = "\u064a\u062d\u064a \u0627\u0644\u0645\u0634\u062f / \u0642\u0627\u0633\u0645 \u0627\u0645\u064a\u0646 KG"


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
    backup_path = path.with_name(f"{path.stem}-before-aam-repair-{stamp}{path.suffix}")
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


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _affected_specs(conn: sqlite3.Connection, wrong_school: str) -> set[tuple[str, str, str]]:
    specs: set[tuple[str, str, str]] = set()
    rows = conn.execute(
        """
        SELECT payload_json
          FROM sync_inbox
         WHERE event_type='STOCK_TRANSFER_OUT'
           AND payload_json LIKE ?
        """,
        (f"%{AAM_SCHOOL}%",),
    ).fetchall()
    for row in rows:
        try:
            payload = json.loads(row["payload_json"] or "{}")
        except Exception:
            continue
        for item in payload.get("items") or []:
            if not isinstance(item, dict):
                continue
            if _clean(item.get("school")) != AAM_SCHOOL:
                continue
            try:
                qty = int(float(item.get("qty") or 0))
            except (TypeError, ValueError):
                qty = 0
            if qty <= 0:
                continue
            item_type = _clean(item.get("item_type"))
            color = _clean(item.get("color"))
            size = _clean(item.get("size"))
            if item_type and color and size:
                specs.add((item_type, color, size))

    # Only keep specs that are currently present under the wrong school.
    present: set[tuple[str, str, str]] = set()
    for item_type, color, size in specs:
        row = conn.execute(
            """
            SELECT 1
              FROM stocks
             WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
               AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
               AND LOWER(TRIM(color)) = LOWER(TRIM(?))
               AND LOWER(TRIM(size)) = LOWER(TRIM(?))
             LIMIT 1
            """,
            (wrong_school, item_type, color, size),
        ).fetchone()
        if row:
            present.add((item_type, color, size))
    return present


def _summary(conn: sqlite3.Connection, schools: Iterable[str]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for school in schools:
        row = conn.execute(
            """
            SELECT COUNT(*) AS rows,
                   COALESCE(SUM(count), 0) AS qty,
                   COALESCE(SUM(count * unit_price), 0) AS value
              FROM stocks
             WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
            """,
            (school,),
        ).fetchone()
        out.append({
            "school": school,
            "rows": int(row["rows"] or 0),
            "qty": int(row["qty"] or 0),
            "value": float(row["value"] or 0),
        })
    return out


def _update_table(
    conn: sqlite3.Connection,
    table: str,
    specs: set[tuple[str, str, str]],
    wrong_school: str,
) -> int:
    total = 0
    for item_type, color, size in specs:
        try:
            cur = conn.execute(
                f"""
                UPDATE {table}
                   SET school = ?
                 WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                """,
                (AAM_SCHOOL, wrong_school, item_type, color, size),
            )
            total += int(cur.rowcount or 0)
        except sqlite3.OperationalError:
            continue
    return total


def _delete_duplicate_zero_stock_rows(conn: sqlite3.Connection) -> int:
    try:
        cur = conn.execute(
            """
            DELETE FROM stocks
             WHERE school = ?
               AND COALESCE(count, 0) = 0
               AND EXISTS (
                    SELECT 1
                      FROM stocks AS s2
                     WHERE s2.id <> stocks.id
                       AND s2.school = stocks.school
                       AND LOWER(TRIM(s2.item_type)) = LOWER(TRIM(stocks.item_type))
                       AND LOWER(TRIM(s2.color)) = LOWER(TRIM(stocks.color))
                       AND LOWER(TRIM(s2.size)) = LOWER(TRIM(stocks.size))
                       AND COALESCE(s2.count, 0) > 0
               )
            """,
            (AAM_SCHOOL,),
        )
        return int(cur.rowcount or 0)
    except sqlite3.OperationalError:
        return 0


def repair(path: Path, wrong_school: str, apply: bool) -> dict[str, Any]:
    conn = _connect(path, readonly=not apply)
    try:
        specs = _affected_specs(conn, wrong_school)
        before = _summary(conn, [AAM_SCHOOL, wrong_school])
        result: dict[str, Any] = {
            "database": str(path),
            "wrong_school": wrong_school,
            "affected_specs": len(specs),
            "before": before,
            "applied": bool(apply),
            "backup": None,
            "updates": {},
            "deleted_duplicate_zero_stock_rows": 0,
            "after": None,
        }
        if not apply:
            return result

        backup_path = _backup_database(path)
        result["backup"] = str(backup_path)

        tables = (
            "stocks",
            "movements",
            "bill_items",
            "reservations",
            "reservation_alerts",
            "incoming_shipment_items_pending",
            "branch_catalog_definitions",
            "stock_audit_report_lines",
            "size_profiles",
        )
        with conn:
            for table in tables:
                result["updates"][table] = _update_table(conn, table, specs, wrong_school)
            result["deleted_duplicate_zero_stock_rows"] = _delete_duplicate_zero_stock_rows(conn)
            conn.execute(
                "INSERT OR IGNORE INTO spec_history(field, value) VALUES('school', ?)",
                (AAM_SCHOOL,),
            )
            row = conn.execute(
                """
                SELECT 1
                  FROM stocks
                 WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
                 LIMIT 1
                """,
                (wrong_school,),
            ).fetchone()
            if row is None:
                conn.execute(
                    "DELETE FROM spec_history WHERE field='school' AND LOWER(TRIM(value)) = LOWER(TRIM(?))",
                    (wrong_school,),
                )
        result["after"] = _summary(conn, [AAM_SCHOOL, wrong_school])
        return result
    finally:
        conn.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Repair POS rows where عام was wrongly moved to Yahya/Qasem KG.")
    parser.add_argument("db", type=Path)
    parser.add_argument("--wrong-school", default=DEFAULT_WRONG_SCHOOL)
    parser.add_argument("--apply", action="store_true", help="Apply the repair. Without this, prints a dry-run preview.")
    args = parser.parse_args()

    result = repair(args.db, args.wrong_school, bool(args.apply))
    print(json.dumps(result, ensure_ascii=True, indent=2, default=str))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
