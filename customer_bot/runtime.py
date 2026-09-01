from __future__ import annotations

from typing import Any

try:
    from .cloud_queries import CloudCustomerQueries
    from .config import BotConfig, load_config
    from .queries import WarehouseCustomerQueries
except ImportError:
    from cloud_queries import CloudCustomerQueries
    from config import BotConfig, load_config
    from queries import WarehouseCustomerQueries


def make_queries(config: BotConfig | None = None) -> Any:
    cfg = config or load_config()
    if cfg.customer_stock_api_url:
        return CloudCustomerQueries(cfg)
    return WarehouseCustomerQueries(cfg)
