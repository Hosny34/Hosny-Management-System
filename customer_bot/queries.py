from __future__ import annotations

import sqlite3
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

try:
    from .config import BotConfig
except ImportError:
    from config import BotConfig


BRANCH_NAMES = {
    "POS-ZAY": "فرع زايد",
    "POS-OCT": "فرع اكتوبر",
    "POS-OBO": "فرع العبور",
    "POS-GESR": "فرع جسر السويس",
    "POS-BAH": "فرع بهتيم",
    "POS-CEN": "فرع السنتر",
}
BRANCH_BY_NAME = {v: k for k, v in BRANCH_NAMES.items()}


def clean(value: Any) -> str:
    return str(value or "").strip()


def branch_display_name(device: Any) -> str:
    text = clean(device)
    return BRANCH_NAMES.get(text, text)


def canonical_branch(value: Any) -> Optional[str]:
    text = clean(value)
    if not text:
        return None
    if text in BRANCH_NAMES:
        return text
    if text.startswith("فرع:"):
        text = text.split(":", 1)[1].strip()
    return BRANCH_BY_NAME.get(text)


def _parse_dt(value: Any) -> Optional[datetime]:
    text = clean(value)
    if not text:
        return None
    try:
        if text.endswith("Z"):
            text = text[:-1] + "+00:00"
        dt = datetime.fromisoformat(text)
        if dt.tzinfo is not None:
            return dt.astimezone().replace(tzinfo=None)
        return dt
    except ValueError:
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
    return None


def _open_readonly(path: Path) -> sqlite3.Connection:
    uri = f"file:{path.as_posix()}?mode=ro"
    conn = sqlite3.connect(uri, uri=True, timeout=5.0)
    conn.row_factory = sqlite3.Row
    return conn


def _where_like(field: str, value: Any, where: List[str], args: List[Any]) -> None:
    text = clean(value)
    if not text:
        return
    where.append(f"LOWER(TRIM({field})) LIKE LOWER(?)")
    args.append(f"%{text}%")


def _allowed_devices(config: BotConfig) -> List[str]:
    return [clean(b.get("device")) for b in config.branches if clean(b.get("device"))]


@dataclass
class WarehouseCustomerQueries:
    config: BotConfig

    def _connect(self) -> sqlite3.Connection:
        return _open_readonly(self.config.warehouse_db_path)

    def branch_info(self, branch: str | None = None) -> List[Dict[str, Any]]:
        wanted = canonical_branch(branch) if branch else None
        out: List[Dict[str, Any]] = []
        for b in self.config.branches:
            device = clean(b.get("device"))
            if wanted and device != wanted:
                continue
            out.append(
                {
                    "device": device,
                    "name": clean(b.get("name")) or branch_display_name(device),
                    "address": clean(b.get("address")),
                    "phone": clean(b.get("phone")),
                    "maps_url": clean(b.get("maps_url")),
                    "hours": clean(b.get("hours")),
                }
            )
        return out

    def known_values(self, field: str, limit: int = 2000) -> List[str]:
        if field not in {"item_type", "school", "color", "size"}:
            return []
        devices = _allowed_devices(self.config)
        device_filter = ""
        args: List[Any] = []
        if devices:
            device_filter = f"AND source_device IN ({','.join('?' for _ in devices)})"
            args.extend(devices)
        try:
            with closing(self._connect()) as conn:
                rows = conn.execute(
                    f"""
                    SELECT DISTINCT TRIM({field}) AS value
                      FROM pos_stocks_mirror
                     WHERE COALESCE(TRIM({field}), '') <> ''
                       {device_filter}
                     ORDER BY LENGTH(TRIM({field})) DESC, TRIM({field})
                     LIMIT ?
                    """,
                    (*args, int(limit)),
                ).fetchall()
        except sqlite3.Error:
            return []
        return [clean(r["value"]) for r in rows if clean(r["value"])]

    def distinct_values(
        self,
        field: str,
        *,
        source_device: str = "",
        school: str = "",
        item_type: str = "",
        color: str = "",
        min_count: int = 1,
        limit: int = 200,
    ) -> List[str]:
        if field not in {"item_type", "school", "color", "size"}:
            return []
        where = ["COALESCE(count, 0) >= ?", f"COALESCE(TRIM({field}), '') <> ''"]
        args: List[Any] = [int(min_count)]
        devices = _allowed_devices(self.config)
        if devices:
            where.append(f"source_device IN ({','.join('?' for _ in devices)})")
            args.extend(devices)
        if source_device:
            where.append("source_device = ?")
            args.append(clean(source_device))
        _where_like("school", school, where, args)
        _where_like("item_type", item_type, where, args)
        _where_like("color", color, where, args)
        try:
            with closing(self._connect()) as conn:
                rows = conn.execute(
                    f"""
                    SELECT DISTINCT TRIM({field}) AS value
                      FROM pos_stocks_mirror
                     WHERE {' AND '.join(where)}
                     ORDER BY TRIM({field})
                     LIMIT ?
                    """,
                    (*args, int(limit)),
                ).fetchall()
        except sqlite3.Error:
            return []
        return [clean(r["value"]) for r in rows if clean(r["value"])]

    def search_stock(
        self,
        *,
        item_type: str = "",
        school: str = "",
        color: str = "",
        size: str = "",
        source_device: str = "",
        min_count: int = 1,
        limit: int = 30,
    ) -> List[Dict[str, Any]]:
        where = ["COALESCE(pm.count, 0) >= ?"]
        args: List[Any] = [int(min_count)]
        devices = _allowed_devices(self.config)
        if devices:
            where.append(f"pm.source_device IN ({','.join('?' for _ in devices)})")
            args.extend(devices)
        if source_device:
            where.append("pm.source_device = ?")
            args.append(clean(source_device))
        _where_like("pm.item_type", item_type, where, args)
        _where_like("pm.school", school, where, args)
        _where_like("pm.color", color, where, args)
        _where_like("pm.size", size, where, args)
        try:
            with closing(self._connect()) as conn:
                rows = conn.execute(
                    f"""
                    SELECT
                        pm.source_device,
                        pm.item_type,
                        pm.school,
                        pm.color,
                        pm.size,
                        pm.unit_price,
                        SUM(pm.count) AS count,
                        COALESCE(meta.snapshot_at, MAX(pm.snapshot_at)) AS snapshot_at
                      FROM pos_stocks_mirror pm
                      LEFT JOIN pos_stocks_snapshot_meta meta
                        ON meta.source_device = pm.source_device
                     WHERE {' AND '.join(where)}
                     GROUP BY pm.source_device, pm.item_type, pm.school, pm.color, pm.size, pm.unit_price
                     ORDER BY pm.school, pm.item_type, pm.color, pm.size, pm.unit_price, pm.source_device
                     LIMIT ?
                    """,
                    (*args, int(limit)),
                ).fetchall()
        except sqlite3.Error:
            return []
        now = datetime.now()
        stale_after = timedelta(minutes=max(1, int(self.config.stock_stale_minutes)))
        out: List[Dict[str, Any]] = []
        for r in rows:
            snap = clean(r["snapshot_at"])
            snap_dt = _parse_dt(snap)
            is_stale = True if snap_dt is None else (now - snap_dt) > stale_after
            out.append(
                {
                    "branch_device": clean(r["source_device"]),
                    "branch": branch_display_name(r["source_device"]),
                    "item_type": clean(r["item_type"]),
                    "school": clean(r["school"]),
                    "color": clean(r["color"]),
                    "size": clean(r["size"]),
                    "unit_price": float(r["unit_price"] or 0.0),
                    "count": int(r["count"] or 0),
                    "last_sync": snap,
                    "stale": bool(is_stale),
                }
            )
        return out

    def search_prices(self, **filters: Any) -> List[Dict[str, Any]]:
        limit = int(filters.pop("limit", 30) or 30)
        rows = self.search_stock(**filters, min_count=0, limit=limit)
        seen = set()
        prices: List[Dict[str, Any]] = []
        for row in rows:
            key = (
                row["branch_device"],
                row["item_type"].casefold(),
                row["school"].casefold(),
                row["color"].casefold(),
                row["size"].casefold(),
                row["unit_price"],
            )
            if key in seen:
                continue
            seen.add(key)
            prices.append(row)
        return prices

    def reservation_status(self, *, branch: str, bill_number: str) -> Dict[str, Any] | None:
        device = canonical_branch(branch)
        bill = clean(bill_number)
        if not device or not bill:
            return None
        candidates = [bill, f"id:{bill}"]
        try:
            with closing(self._connect()) as conn:
                rows = conn.execute(
                    f"""
                    SELECT *
                      FROM pos_reservations_mirror
                     WHERE source_device = ?
                       AND reservation_key IN ({','.join('?' for _ in candidates)})
                     ORDER BY updated_at DESC
                    """,
                    (device, *candidates),
                ).fetchall()
        except sqlite3.Error:
            return None
        if not rows:
            return None
        total = sum(float(r["total_amount"] or 0.0) for r in rows)
        paid = sum(float(r["paid_amount"] or 0.0) for r in rows)
        pending = [r for r in rows if clean(r["status"]) == "معلق"]
        status = "معلق" if pending else clean(rows[0]["status"])
        return {
            "branch": branch_display_name(device),
            "bill_number": bill,
            "status": status,
            "total": total,
            "paid": paid,
            "remaining": max(0.0, total - paid),
            "items": [
                {
                    "item_type": clean(r["item_type"]),
                    "school": clean(r["school"]),
                    "color": clean(r["color"]),
                    "size": clean(r["size"]),
                    "qty": int(r["qty"] or 0),
                    "status": clean(r["status"]),
                }
                for r in rows
            ],
            "updated_at": clean(rows[0]["updated_at"]),
        }
