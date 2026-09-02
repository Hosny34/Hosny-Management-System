from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Any, Dict, List
from urllib.parse import urlencode
from urllib.request import urlopen

try:
    from .config import BotConfig
    from .queries import branch_display_name, canonical_branch, clean
except ImportError:
    from config import BotConfig
    from queries import branch_display_name, canonical_branch, clean


def _parse_dt(value: Any) -> datetime | None:
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
        return None


@dataclass
class CloudCustomerQueries:
    config: BotConfig

    def _allowed_devices(self) -> set[str]:
        return {clean(b.get("device")) for b in self.config.branches if clean(b.get("device"))}

    def _get_json(self, path: str, params: Dict[str, Any]) -> Dict[str, Any]:
        base = self.config.customer_stock_api_url.rstrip("/")
        if not base:
            return {}
        query = urlencode({k: v for k, v in params.items() if v not in ("", None)})
        url = base + path + (("?" + query) if query else "")
        with urlopen(url, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8"))

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
        try:
            payload = self._get_json("/v1/customer/known-values", {"field": field, "limit": limit})
        except Exception:
            return []
        return [clean(x) for x in payload.get("values") or [] if clean(x)]

    def search_stock(
        self,
        *,
        item_type: str = "",
        school: str = "",
        color: str = "",
        size: str = "",
        min_count: int = 1,
        limit: int = 30,
    ) -> List[Dict[str, Any]]:
        try:
            payload = self._get_json(
                "/v1/customer/stock",
                {
                    "item_type": item_type,
                    "school": school,
                    "color": color,
                    "size": size,
                    "min_count": min_count,
                    "limit": limit,
                },
            )
        except Exception:
            return []
        now = datetime.now()
        stale_after = timedelta(minutes=max(1, int(self.config.stock_stale_minutes)))
        allowed_devices = self._allowed_devices()
        out: List[Dict[str, Any]] = []
        for r in payload.get("rows") or []:
            device = clean(r.get("source_device"))
            if allowed_devices and device not in allowed_devices:
                continue
            uploaded_dt = _parse_dt(r.get("uploaded_at") or r.get("snapshot_at"))
            stale = True if uploaded_dt is None else (now - uploaded_dt) > stale_after
            out.append(
                {
                    "branch_device": device,
                    "branch": branch_display_name(device),
                    "item_type": clean(r.get("item_type")),
                    "school": clean(r.get("school")),
                    "color": clean(r.get("color")),
                    "size": clean(r.get("size")),
                    "unit_price": float(r.get("unit_price") or 0),
                    "count": int(r.get("count") or 0),
                    "last_sync": clean(r.get("uploaded_at") or r.get("snapshot_at")),
                    "stale": stale,
                }
            )
        return out

    def search_prices(self, **filters: Any) -> List[Dict[str, Any]]:
        limit = int(filters.pop("limit", 30) or 30)
        return self.search_stock(**filters, min_count=0, limit=limit)

    def reservation_status(self, *, branch: str, bill_number: str) -> Dict[str, Any] | None:
        return None
