from __future__ import annotations

import argparse
import json
import os
import sqlite3
import sys
from contextlib import closing
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_DB = ROOT / "Warehouse" / "warehouse_data.sqlite3"
DEFAULT_SERVER_URL = "https://web-production-e022.up.railway.app"


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")


def _clean(value: Any) -> str:
    return str(value or "").strip()


def _open_readonly(path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(f"file:{path.as_posix()}?mode=ro", uri=True, timeout=10.0)
    conn.row_factory = sqlite3.Row
    return conn


def load_customer_stock_rows(db_path: Path) -> List[Dict[str, Any]]:
    with closing(_open_readonly(db_path)) as conn:
        rows = conn.execute(
            """
            SELECT
                pm.source_device,
                pm.item_type,
                pm.school,
                pm.color,
                pm.size,
                pm.unit_price,
                SUM(COALESCE(pm.count, 0)) AS count,
                COALESCE(meta.snapshot_at, MAX(pm.snapshot_at)) AS snapshot_at
              FROM pos_stocks_mirror pm
              LEFT JOIN pos_stocks_snapshot_meta meta
                ON meta.source_device = pm.source_device
             WHERE COALESCE(TRIM(pm.source_device), '') <> ''
               AND COALESCE(TRIM(pm.item_type), '') <> ''
               AND COALESCE(TRIM(pm.school), '') <> ''
               AND COALESCE(TRIM(pm.color), '') <> ''
               AND COALESCE(TRIM(pm.size), '') <> ''
             GROUP BY
                pm.source_device,
                pm.item_type,
                pm.school,
                pm.color,
                pm.size,
                pm.unit_price
             ORDER BY pm.source_device, pm.school, pm.item_type, pm.color, pm.size, pm.unit_price
            """
        ).fetchall()

    out: List[Dict[str, Any]] = []
    for r in rows:
        try:
            unit_price = float(r["unit_price"] or 0)
            count = int(r["count"] or 0)
        except (TypeError, ValueError):
            continue
        out.append(
            {
                "source_device": _clean(r["source_device"]),
                "item_type": _clean(r["item_type"]),
                "school": _clean(r["school"]),
                "color": _clean(r["color"]),
                "size": _clean(r["size"]),
                "unit_price": unit_price,
                "count": max(0, count),
                "snapshot_at": _clean(r["snapshot_at"]) or _utc_now_iso(),
            }
        )
    return out


def upload_rows(server_url: str, token: str, rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    if not token:
        raise RuntimeError("CUSTOMER_STOCK_UPLOAD_TOKEN is required")
    endpoint = server_url.rstrip("/") + "/v1/customer/stock/upload"
    body = json.dumps({"uploaded_at": _utc_now_iso(), "rows": rows}, ensure_ascii=False).encode("utf-8")
    req = Request(
        endpoint,
        data=body,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json",
            "X-Customer-Stock-Token": token,
        },
        method="POST",
    )
    try:
        with urlopen(req, timeout=60) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except HTTPError as ex:
        detail = ex.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"upload failed: HTTP {ex.code} {detail}") from ex
    except URLError as ex:
        raise RuntimeError(f"upload failed: {ex.reason}") from ex


def main() -> int:
    parser = argparse.ArgumentParser(description="Upload customer stock cache from Warehouse DB.")
    parser.add_argument("--db", default=os.environ.get("WAREHOUSE_DB_PATH") or str(DEFAULT_DB))
    parser.add_argument("--server-url", default=os.environ.get("CUSTOMER_STOCK_SERVER_URL") or DEFAULT_SERVER_URL)
    parser.add_argument("--token", default=os.environ.get("CUSTOMER_STOCK_UPLOAD_TOKEN") or "")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    db_path = Path(args.db)
    if not db_path.is_absolute():
        db_path = ROOT / db_path
    if not db_path.is_file():
        print(f"Warehouse DB not found: {db_path}", file=sys.stderr)
        return 2

    rows = load_customer_stock_rows(db_path)
    total_count = sum(int(r["count"] or 0) for r in rows)
    total_value = sum(float(r["unit_price"] or 0) * int(r["count"] or 0) for r in rows)
    print(f"Loaded {len(rows)} stock rows, total_count={total_count}, total_value={round(total_value, 2)}")
    if args.dry_run:
        return 0

    result = upload_rows(args.server_url, args.token, rows)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
