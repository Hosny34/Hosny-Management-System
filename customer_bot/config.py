from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG_PATH = Path(__file__).with_name("config.json")


@dataclass(frozen=True)
class BotConfig:
    warehouse_db_path: Path
    stock_stale_minutes: int
    branches: List[Dict[str, str]]


def load_config(path: str | os.PathLike[str] | None = None) -> BotConfig:
    cfg_path = Path(path) if path else DEFAULT_CONFIG_PATH
    data: Dict[str, Any] = json.loads(cfg_path.read_text(encoding="utf-8"))
    db_raw = os.environ.get("WAREHOUSE_DB_PATH") or data.get("warehouse_db_path") or "Warehouse/warehouse_data.sqlite3"
    db_path = Path(db_raw)
    if not db_path.is_absolute():
        db_path = ROOT / db_path
    stale_raw = os.environ.get("BOT_STOCK_STALE_MINUTES") or data.get("stock_stale_minutes") or 30
    try:
        stale_minutes = int(stale_raw)
    except (TypeError, ValueError):
        stale_minutes = 30
    branches = [dict(b) for b in data.get("branches") or []]
    return BotConfig(
        warehouse_db_path=db_path,
        stock_stale_minutes=stale_minutes,
        branches=branches,
    )

