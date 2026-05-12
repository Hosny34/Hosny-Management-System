"""HosnyFactory - Local-first Factory Production Manager (V1).

Single-PC desktop app matching the HosnyWarehouse / HosnyPOS family:
  - Python + Tkinter/ttk
  - SQLite local DB (factory_data.sqlite3 next to this file)
  - PyInstaller-packaged (HosnyFactory.spec)

V1 covers: products, stages, workers+allowed-stages, order wizard with
default/optional stages, pipeline board, rule-based scheduler with
shop-hours ETA, pause/resume/block/skip/insert (manager-gated),
append-only event log, reports, and settings.
"""

from __future__ import annotations

import json
import logging
import os
import sqlite3
import sys
import threading
import traceback
import csv
from dataclasses import dataclass
from datetime import date, datetime, time as dtime, timedelta
from typing import Any, Dict, Iterable, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk


# =============================================================================
# Paths, logging
# =============================================================================

def _app_dir() -> str:
    """Directory containing this script (or the exe when frozen)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _app_dir()
DB_PATH = os.path.join(APP_DIR, "factory_data.sqlite3")
LOG_DIR = os.path.join(APP_DIR, "logs")
LOG_PATH = os.path.join(LOG_DIR, "HosnyFactory.log")
ANALYTICS_DIR = os.path.join(APP_DIR, "analytics_archive")

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(ANALYTICS_DIR, exist_ok=True)

logging.basicConfig(
    filename=LOG_PATH,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8",
)
log = logging.getLogger("hosny.factory")


# =============================================================================
# Theme / fonts (POS-style light palette, matches HosnyWarehouse)
# =============================================================================

UI = {
    "BG":       "#f8fafc",
    "SURFACE":  "#ffffff",
    "SURFACE2": "#f1f5f9",
    "BORDER":   "#cbd5e1",
    "ACCENT":   "#2563eb",
    "ACCENT_H": "#1d4ed8",
    "BRAND":    "#2563eb",
    "BRAND_L":  "#bfdbfe",
    "TEXT":     "#0f172a",
    "TEXT_SEC": "#475569",
    "TEXT_DIM": "#94a3b8",
    "OK":       "#16a34a",
    "OK_L":     "#dcfce7",
    "WARN":     "#d97706",
    "WARN_L":   "#fef3c7",
    "DANGER":   "#dc2626",
    "DANGER_L": "#fee2e2",
    "INFO":     "#0ea5e9",
    "INFO_L":   "#e0f2fe",
    "ROW_EVEN": "#ffffff",
    "ROW_ODD":  "#f8fafc",
    "SEL_BG":   "#bfdbfe",
    "SEL_FG":   "#0f172a",
}

FONTS = {
    "h1":     ("Segoe UI", 16, "bold"),
    "h2":     ("Segoe UI", 12, "bold"),
    "h3":     ("Segoe UI", 11, "bold"),
    "body":   ("Segoe UI", 10),
    "body_b": ("Segoe UI", 10, "bold"),
    "small":  ("Segoe UI", 9),
    "big":    ("Segoe UI", 22, "bold"),
    "btn":    ("Segoe UI", 10, "bold"),
    "mono":   ("Consolas", 9),
}


# Status display helpers (Arabic labels + color tags)
STATUS_AR = {
    "draft":     "مسودة",
    "released":  "صادر",
    "running":   "قيد التنفيذ",
    "paused":    "متوقف",
    "blocked":   "محجوب",
    "done":      "منتهي",
    "cancelled": "ملغى",
    "planned":   "مخطط",
    "skipped":   "متخطى",
    "ready":     "جاهز",
}

STATUS_COLOR = {
    "draft":     (UI["SURFACE2"], UI["TEXT_SEC"]),
    "released":  (UI["INFO_L"],   UI["INFO"]),
    "running":   (UI["OK_L"],     UI["OK"]),
    "paused":    (UI["WARN_L"],   UI["WARN"]),
    "blocked":   (UI["DANGER_L"], UI["DANGER"]),
    "done":      (UI["OK_L"],     UI["OK"]),
    "cancelled": (UI["SURFACE2"], UI["TEXT_SEC"]),
    "planned":   (UI["SURFACE2"], UI["TEXT_SEC"]),
    "skipped":   (UI["SURFACE2"], UI["TEXT_SEC"]),
    "ready":     (UI["INFO_L"],   UI["INFO"]),
}


def status_label(code: str) -> str:
    if code == "draft":
        return "مسودة قديمة"
    if code == "released":
        return "جاهز للتشغيل"
    return STATUS_AR.get(code or "", code or "")


def fmt_ts(ts: Optional[str]) -> str:
    if not ts:
        return "—"
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(ts)


def fmt_minutes(mins: Optional[float]) -> str:
    if mins is None:
        return "—"
    mins = int(round(float(mins)))
    if mins < 60:
        return f"{mins} د"
    h, m = divmod(mins, 60)
    if h < 24:
        return f"{h}س {m:02d}د"
    d, h = divmod(h, 24)
    return f"{d} يوم {h}س"


# =============================================================================
# Database layer
# =============================================================================

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS app_settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS stage_templates (
    id                        INTEGER PRIMARY KEY AUTOINCREMENT,
    code                      TEXT NOT NULL UNIQUE,
    name                      TEXT NOT NULL,
    default_setup_minutes     REAL NOT NULL DEFAULT 0,
    default_per_unit_minutes  REAL NOT NULL DEFAULT 0,
    notes                     TEXT,
    is_active                 INTEGER NOT NULL DEFAULT 1,
    created_at                TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS products (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    code         TEXT NOT NULL UNIQUE,
    name         TEXT NOT NULL,
    product_type TEXT,
    school_name  TEXT,
    size         TEXT,
    color        TEXT,
    notes        TEXT,
    is_active    INTEGER NOT NULL DEFAULT 1,
    created_at   TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS product_stages (
    id                         INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id                 INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
    stage_template_id          INTEGER NOT NULL REFERENCES stage_templates(id) ON DELETE RESTRICT,
    sequence                   INTEGER NOT NULL,
    is_default                 INTEGER NOT NULL DEFAULT 1,
    is_optional                INTEGER NOT NULL DEFAULT 0,
    override_setup_minutes     REAL,
    override_per_unit_minutes  REAL,
    UNIQUE(product_id, sequence)
);

CREATE TABLE IF NOT EXISTS workers (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    code       TEXT NOT NULL UNIQUE,
    name       TEXT NOT NULL,
    notes      TEXT,
    is_active  INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS worker_stages (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    worker_id  INTEGER NOT NULL REFERENCES workers(id) ON DELETE CASCADE,
    stage_template_id INTEGER NOT NULL REFERENCES stage_templates(id) ON DELETE CASCADE,
    UNIQUE(worker_id, stage_template_id)
);

CREATE TABLE IF NOT EXISTS stations (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT NOT NULL UNIQUE,
    station_type TEXT,
    is_active  INTEGER NOT NULL DEFAULT 1,
    notes      TEXT
);

CREATE TABLE IF NOT EXISTS production_orders (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    order_number  TEXT NOT NULL UNIQUE,
    product_id    INTEGER NOT NULL REFERENCES products(id) ON DELETE RESTRICT,
    quantity      INTEGER NOT NULL CHECK (quantity > 0),
    priority      INTEGER NOT NULL DEFAULT 3, -- 1 highest .. 5 lowest
    due_at        TEXT,
    status        TEXT NOT NULL DEFAULT 'draft'
                      CHECK(status IN ('draft','released','running','paused','blocked','done','cancelled')),
    -- Per-order descriptors (filled at order time, with autocomplete from history)
    school_name   TEXT,
    size          TEXT,
    color         TEXT,
    notes         TEXT,
    root_order_id INTEGER REFERENCES production_orders(id) ON DELETE SET NULL,
    parent_order_id INTEGER REFERENCES production_orders(id) ON DELETE SET NULL,
    original_quantity INTEGER,
    split_from_stage_id INTEGER REFERENCES order_stages(id) ON DELETE SET NULL,
    split_reason TEXT,
    created_at    TEXT NOT NULL DEFAULT (datetime('now')),
    released_at   TEXT,
    completed_at  TEXT
);

CREATE TABLE IF NOT EXISTS order_stages (
    id                     INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id               INTEGER NOT NULL REFERENCES production_orders(id) ON DELETE CASCADE,
    stage_template_id      INTEGER NOT NULL REFERENCES stage_templates(id)  ON DELETE RESTRICT,
    sequence               INTEGER NOT NULL,
    name_snapshot          TEXT NOT NULL,
    setup_minutes          REAL NOT NULL DEFAULT 0,
    per_unit_minutes       REAL NOT NULL DEFAULT 0,
    is_optional_selected   INTEGER NOT NULL DEFAULT 1,
    status                 TEXT NOT NULL DEFAULT 'planned'
                              CHECK(status IN ('planned','running','paused','blocked','done','skipped')),
    assigned_worker_id     INTEGER REFERENCES workers(id)  ON DELETE SET NULL,
    station_id             INTEGER REFERENCES stations(id) ON DELETE SET NULL,
    planned_start          TEXT,
    planned_end            TEXT,
    actual_start           TEXT,
    actual_end             TEXT,
    actual_minutes         REAL,
    early_release          INTEGER NOT NULL DEFAULT 0,
    notes                  TEXT,
    UNIQUE(order_id, sequence)
);

CREATE TABLE IF NOT EXISTS order_events (
    id             INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id       INTEGER NOT NULL REFERENCES production_orders(id) ON DELETE CASCADE,
    order_stage_id INTEGER REFERENCES order_stages(id) ON DELETE SET NULL,
    event_type     TEXT NOT NULL,
    actor          TEXT,
    reason         TEXT,
    payload_json   TEXT,
    created_at     TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS order_wizard_suggestions (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    field      TEXT NOT NULL
               CHECK (field IN ('school_name', 'size', 'color', 'quantity')),
    value      TEXT NOT NULL,
    use_count  INTEGER NOT NULL DEFAULT 1,
    last_used  TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(field, value)
);

CREATE INDEX IF NOT EXISTS idx_ows_field ON order_wizard_suggestions(field, use_count DESC);

CREATE INDEX IF NOT EXISTS idx_order_stages_order  ON order_stages(order_id);
CREATE INDEX IF NOT EXISTS idx_order_events_order  ON order_events(order_id);
CREATE INDEX IF NOT EXISTS idx_order_events_stage  ON order_events(order_stage_id);
CREATE INDEX IF NOT EXISTS idx_orders_status       ON production_orders(status);
CREATE INDEX IF NOT EXISTS idx_order_stages_status ON order_stages(status);
"""


SEED_STAGES = [
    # code, name, setup_minutes, per_unit_minutes
    ("CUT",   "قص",          30, 1.0),
    ("SEW",   "حياكة",       20, 4.0),
    ("IRON",  "كوي",         10, 1.5),
    ("FIN",   "تشطيب",        5, 1.0),
    ("QC",    "مراقبة جودة", 10, 0.5),
    ("PACK",  "تعبئة",       10, 0.8),
]

# Default product and its routing (all defaults, none optional for baseline)
SEED_DEFAULT_PRODUCT = {
    "code": "UNI-STD",
    "name": "قميص",
    "product_type": "قميص",
}

# Predefined product types that users pick from when creating a new order.
# The ProductsAdmin tab is what lets them add/remove entries in this catalog
# (via the `products` table). This list is only a starter suggestion for first-run.
COMMON_PRODUCT_TYPES = ["قميص", "بنطلون", "جاكيت", "بلوزة", "تنورة", "فستان", "بليزر", "صدرية"]
# Smart-fill starter values for the order wizard (merged with anything the user
# has actually used before, pulled live from the DB).
COMMON_SIZES = ["4", "6", "8", "10", "12", "14", "16", "S", "M", "L", "XL", "XXL"]
COMMON_COLORS = ["أبيض", "أزرق", "كحلي", "أسود", "رمادي", "أخضر", "وردي", "بيج", "أحمر", "أصفر"]
COMMON_QUANTITIES = ["10", "20", "50", "100", "200", "500", "1000"]
SEED_DEFAULT_ROUTING = [
    # (stage_code, sequence, is_default, is_optional)
    ("CUT",  1, 1, 0),
    ("SEW",  2, 1, 0),
    ("IRON", 3, 1, 0),
    ("FIN",  4, 1, 1),  # optional finishing
    ("QC",   5, 1, 0),
    ("PACK", 6, 1, 0),
]


DEFAULT_SETTINGS = {
    "shop_start":     "08:00",
    "shop_end":       "18:00",
    "weekend_days":   "fri",            # csv of mon,tue,wed,thu,fri,sat,sun
    "order_counter":  "0",
    "app_initialized": "0",
}


class DB:
    """Thin wrapper around sqlite3 with schema/seed and common queries."""

    def __init__(self, path: str = DB_PATH) -> None:
        self.path = path
        self._conn = sqlite3.connect(path, check_same_thread=False, isolation_level=None)
        self._conn.row_factory = sqlite3.Row
        self._conn.execute("PRAGMA foreign_keys = ON")
        self._conn.execute("PRAGMA journal_mode = WAL")
        self._lock = threading.RLock()
        self._init_schema()
        self._seed_if_empty()

    # ----- low-level ----- #

    def execute(self, sql: str, params: Iterable[Any] = ()) -> sqlite3.Cursor:
        with self._lock:
            return self._conn.execute(sql, tuple(params))

    def executescript(self, sql: str) -> None:
        with self._lock:
            self._conn.executescript(sql)

    def query(self, sql: str, params: Iterable[Any] = ()) -> List[sqlite3.Row]:
        with self._lock:
            cur = self._conn.execute(sql, tuple(params))
            return list(cur.fetchall())

    def query_one(self, sql: str, params: Iterable[Any] = ()) -> Optional[sqlite3.Row]:
        rows = self.query(sql, params)
        return rows[0] if rows else None

    def _init_schema(self) -> None:
        self.executescript(SCHEMA_SQL)
        self._migrate()
        for k, v in DEFAULT_SETTINGS.items():
            self.execute(
                "INSERT OR IGNORE INTO app_settings(key,value) VALUES(?,?)", (k, v)
            )

    def _migrate(self) -> None:
        """Forward-compatible column additions for older DBs."""
        pcols = {r["name"] for r in self.query("PRAGMA table_info(products)")}
        for col in ("product_type", "school_name", "size", "color"):
            if col not in pcols:
                self.execute(f"ALTER TABLE products ADD COLUMN {col} TEXT")

        ocols = {r["name"] for r in self.query("PRAGMA table_info(production_orders)")}
        order_column_defs = {
            "school_name": "TEXT",
            "size": "TEXT",
            "color": "TEXT",
            "root_order_id": "INTEGER REFERENCES production_orders(id) ON DELETE SET NULL",
            "parent_order_id": "INTEGER REFERENCES production_orders(id) ON DELETE SET NULL",
            "original_quantity": "INTEGER",
            "split_from_stage_id": "INTEGER REFERENCES order_stages(id) ON DELETE SET NULL",
            "split_reason": "TEXT",
        }
        for col, ddl in order_column_defs.items():
            if col not in ocols:
                self.execute(f"ALTER TABLE production_orders ADD COLUMN {col} {ddl}")
        self.execute(
            "UPDATE production_orders SET original_quantity=quantity WHERE original_quantity IS NULL"
        )
        self.execute(
            "UPDATE production_orders SET root_order_id=id WHERE root_order_id IS NULL"
        )
        self.execute("CREATE INDEX IF NOT EXISTS idx_orders_root ON production_orders(root_order_id)")
        self.execute("CREATE INDEX IF NOT EXISTS idx_orders_parent ON production_orders(parent_order_id)")

        # Stage=Skill model migration: ensure worker_stages exists.
        self.execute(
            """CREATE TABLE IF NOT EXISTS worker_stages (
                   id INTEGER PRIMARY KEY AUTOINCREMENT,
                   worker_id INTEGER NOT NULL REFERENCES workers(id) ON DELETE CASCADE,
                   stage_template_id INTEGER NOT NULL REFERENCES stage_templates(id) ON DELETE CASCADE,
                   UNIQUE(worker_id, stage_template_id)
               )"""
        )
        # Backfill worker_stages from legacy worker_skills + stage.required_skill_id mapping.
        has_worker_skills = bool(
            self.query_one(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='worker_skills'"
            )
        )
        has_required_skill = "required_skill_id" in {
            r["name"] for r in self.query("PRAGMA table_info(stage_templates)")
        }
        if has_worker_skills and has_required_skill:
            self.execute(
                """INSERT OR IGNORE INTO worker_stages(worker_id, stage_template_id)
                   SELECT ws.worker_id, st.id
                   FROM worker_skills ws
                   JOIN stage_templates st ON st.required_skill_id = ws.skill_id
                   WHERE st.required_skill_id IS NOT NULL"""
            )

        # Backfill the seed default product's type for existing DBs
        row = self.query_one(
            "SELECT id, product_type FROM products WHERE code='UNI-STD'"
        )
        if row and not row["product_type"]:
            sp = SEED_DEFAULT_PRODUCT
            self.execute(
                "UPDATE products SET product_type=?, name=? WHERE id=?",
                (sp["product_type"], sp["name"], row["id"]),
            )

        # Autocomplete history survives deleted/completed orders (see order_wizard_suggestions).
        self.execute(
            """CREATE TABLE IF NOT EXISTS order_wizard_suggestions (
                   id         INTEGER PRIMARY KEY AUTOINCREMENT,
                   field      TEXT NOT NULL
                              CHECK (field IN ('school_name', 'size', 'color', 'quantity')),
                   value      TEXT NOT NULL,
                   use_count  INTEGER NOT NULL DEFAULT 1,
                   last_used  TEXT NOT NULL DEFAULT (datetime('now')),
                   UNIQUE(field, value)
               )"""
        )
        self.execute(
            "CREATE INDEX IF NOT EXISTS idx_ows_field "
            "ON order_wizard_suggestions(field, use_count DESC)"
        )
        if self.get_setting("order_wizard_suggest_backfill") != "1":
            self._backfill_order_wizard_suggestions()
            self.set_setting("order_wizard_suggest_backfill", "1")

    def _backfill_order_wizard_suggestions(self) -> None:
        """Seed suggestion table from existing orders (one-time; keeps values after deletions)."""
        for col, fld in (("school_name", "school_name"), ("size", "size"), ("color", "color")):
            for r in self.query(
                f"SELECT TRIM({col}) AS v, COUNT(*) AS n FROM production_orders "
                f"WHERE {col} IS NOT NULL AND TRIM({col})<>'' "
                f"GROUP BY TRIM({col})"
            ):
                v = (r["v"] or "").strip()
                if not v:
                    continue
                self.execute(
                    """INSERT INTO order_wizard_suggestions (field, value, use_count, last_used)
                       VALUES (?, ?, ?, datetime('now'))
                       ON CONFLICT(field, value) DO UPDATE SET
                         use_count = use_count + excluded.use_count,
                         last_used = datetime('now')""",
                    (fld, v, int(r["n"])),
                )
        for r in self.query(
            "SELECT quantity, COUNT(*) AS n FROM production_orders GROUP BY quantity"
        ):
            qv = str(int(r["quantity"])) if r["quantity"] is not None else ""
            if not qv:
                continue
            self.execute(
                """INSERT INTO order_wizard_suggestions (field, value, use_count, last_used)
                   VALUES ('quantity', ?, ?, datetime('now'))
                   ON CONFLICT(field, value) DO UPDATE SET
                     use_count = use_count + excluded.use_count,
                     last_used = datetime('now')""",
                (qv, int(r["n"])),
            )

    def _seed_if_empty(self) -> None:
        if self.get_setting("app_initialized") == "1":
            return
        # Stages
        for code, name, setup, per_unit in SEED_STAGES:
            self.execute(
                """INSERT OR IGNORE INTO stage_templates
                   (code,name,default_setup_minutes,default_per_unit_minutes)
                   VALUES(?,?,?,?)""",
                (code, name, setup, per_unit),
            )
        # Default product type + routing
        sp = SEED_DEFAULT_PRODUCT
        self.execute(
            """INSERT OR IGNORE INTO products(code, name, product_type)
               VALUES(?,?,?)""",
            (sp["code"], sp["name"], sp["product_type"]),
        )
        # Also seed a few more common product types so the new-order dropdown is useful
        # from day one. They all start without a routing; user fills routing per type.
        for t in COMMON_PRODUCT_TYPES:
            if t == sp["product_type"]:
                continue
            code = self.next_product_code()
            self.execute(
                """INSERT OR IGNORE INTO products(code, name, product_type)
                   VALUES(?,?,?)""",
                (code, t, t),
            )
        prod = self.query_one("SELECT id FROM products WHERE code=?", (sp["code"],))
        if prod:
            for stage_code, seq, is_def, is_opt in SEED_DEFAULT_ROUTING:
                st = self.query_one(
                    "SELECT id FROM stage_templates WHERE code=?", (stage_code,)
                )
                if st:
                    self.execute(
                        """INSERT OR IGNORE INTO product_stages
                           (product_id,stage_template_id,sequence,is_default,is_optional)
                           VALUES(?,?,?,?,?)""",
                        (prod["id"], st["id"], seq, is_def, is_opt),
                    )
        self.set_setting("app_initialized", "1")

    # ----- settings ----- #

    def get_setting(self, key: str, default: Optional[str] = None) -> Optional[str]:
        row = self.query_one("SELECT value FROM app_settings WHERE key=?", (key,))
        return row["value"] if row else default

    def set_setting(self, key: str, value: str) -> None:
        self.execute(
            """INSERT INTO app_settings(key,value) VALUES(?,?)
               ON CONFLICT(key) DO UPDATE SET value=excluded.value""",
            (key, value),
        )

    # ----- order number ----- #

    def next_order_number(self) -> str:
        n = int(self.get_setting("order_counter") or "0") + 1
        self.set_setting("order_counter", str(n))
        prefix = datetime.now().strftime("%y%m")
        return f"WO-{prefix}-{n:04d}"

    def distinct_product_values(self, column: str) -> List[str]:
        """Return existing non-empty values in a products column (for autocomplete)."""
        if column not in ("product_type",):
            return []
        rows = self.query(
            f"SELECT DISTINCT {column} AS v FROM products "
            f"WHERE {column} IS NOT NULL AND TRIM({column}) <> '' "
            f"ORDER BY {column}"
        )
        return [r["v"] for r in rows]

    def record_wizard_suggestion(self, field: str, value: str) -> None:
        """Remember a value for new-order comboboxes (kept when orders are deleted or done)."""
        v = (value or "").strip()
        if not v or field not in ("school_name", "size", "color", "quantity"):
            return
        self.execute(
            """INSERT INTO order_wizard_suggestions (field, value, use_count, last_used)
               VALUES (?, ?, 1, datetime('now'))
               ON CONFLICT(field, value) DO UPDATE SET
                 use_count = use_count + 1,
                 last_used = datetime('now')""",
            (field, v),
        )

    def distinct_order_values(self, column: str) -> List[str]:
        """Autocomplete: live orders + persistent history (survives deletions / finished orders)."""
        if column not in ("school_name", "size", "color"):
            return []
        rows = self.query(
            f"""
            SELECT v, SUM(n) AS total FROM (
                SELECT TRIM({column}) AS v, COUNT(*) AS n
                FROM production_orders
                WHERE {column} IS NOT NULL AND TRIM({column}) <> ''
                GROUP BY TRIM({column})
            UNION ALL
                SELECT TRIM(value) AS v, use_count AS n
                FROM order_wizard_suggestions
                WHERE field = ? AND TRIM(value) <> ''
            ) AS combined
            GROUP BY v
            HAVING v IS NOT NULL AND TRIM(v) <> ''
            ORDER BY total DESC, v ASC
            """,
            (column,),
        )
        return [r["v"] for r in rows if r["v"]]

    def distinct_order_quantities(self, limit: int = 15) -> List[str]:
        """Most used quantities: current orders + persistent history."""
        rows = self.query(
            """
            SELECT q, SUM(n) AS total FROM (
                SELECT TRIM(CAST(quantity AS TEXT)) AS q, COUNT(*) AS n
                FROM production_orders
                GROUP BY quantity
            UNION ALL
                SELECT TRIM(value) AS q, use_count AS n
                FROM order_wizard_suggestions
                WHERE field = 'quantity' AND TRIM(value) <> ''
            ) AS combined
            WHERE q IS NOT NULL AND q <> ''
            GROUP BY q
            ORDER BY total DESC, q ASC
            LIMIT ?
            """,
            (limit,),
        )
        return [str(r["q"]) for r in rows]

    def next_product_code(self) -> str:
        """Auto-generate the next product code like P-0001."""
        row = self.query_one(
            "SELECT COUNT(*) AS c FROM products"
        )
        n = int(row["c"]) + 1 if row else 1
        # avoid collision if user already has P-0001 etc.
        while True:
            code = f"P-{n:04d}"
            if not self.query_one("SELECT 1 FROM products WHERE code=?", (code,)):
                return code
            n += 1

    # ----- event log ----- #

    def log_event(
        self,
        order_id: int,
        event_type: str,
        *,
        order_stage_id: Optional[int] = None,
        actor: str = "user",
        reason: Optional[str] = None,
        payload: Optional[Dict[str, Any]] = None,
    ) -> None:
        self.execute(
            """INSERT INTO order_events(order_id,order_stage_id,event_type,actor,reason,payload_json)
               VALUES(?,?,?,?,?,?)""",
            (
                order_id,
                order_stage_id,
                event_type,
                actor,
                reason,
                json.dumps(payload, ensure_ascii=False) if payload else None,
            ),
        )


def compose_product_name(
    product_type: str, school: str, size: str, color: str
) -> str:
    """Build a descriptive display name from the 4 dimensions."""
    parts = [p.strip() for p in (product_type, school, size, color) if p and p.strip()]
    return " · ".join(parts) if parts else "منتج بدون وصف"


# =============================================================================
# Scheduling / ETA
# =============================================================================

WEEKDAY_CODES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]


def parse_hhmm(s: str, fallback: Tuple[int, int]) -> Tuple[int, int]:
    try:
        h, m = s.split(":")
        return int(h), int(m)
    except Exception:
        return fallback


def parse_weekend(csv: str) -> List[int]:
    out: List[int] = []
    for tok in (csv or "").lower().replace(" ", "").split(","):
        if tok in WEEKDAY_CODES:
            out.append(WEEKDAY_CODES.index(tok))
    return out


@dataclass
class ShopHours:
    start_h: int
    start_m: int
    end_h: int
    end_m: int
    weekend: List[int]  # weekday indices (Mon=0 .. Sun=6)

    def daily_minutes(self) -> int:
        return max(
            0, (self.end_h * 60 + self.end_m) - (self.start_h * 60 + self.start_m)
        )

    def is_working_day(self, d: date) -> bool:
        return d.weekday() not in self.weekend

    @classmethod
    def from_db(cls, db: DB) -> "ShopHours":
        sh, sm = parse_hhmm(db.get_setting("shop_start") or "08:00", (8, 0))
        eh, em = parse_hhmm(db.get_setting("shop_end") or "18:00", (18, 0))
        wknd = parse_weekend(db.get_setting("weekend_days") or "fri")
        if eh * 60 + em <= sh * 60 + sm:
            eh, em = sh + 8, sm  # fallback 8h day
        return cls(sh, sm, eh, em, wknd)


def add_working_minutes(start: datetime, minutes: float, hours: ShopHours) -> datetime:
    """Advance `start` by `minutes` working minutes, snapping to shop-hours.

    If `start` falls outside the shop window or on a weekend, it is moved to
    the next shop-start. Work only progresses during shop-hours on working
    days. Weekend days are skipped entirely.
    """
    minutes = max(0.0, float(minutes))
    cur = _snap_to_shop(start, hours)
    remaining = minutes
    daily = hours.daily_minutes()
    if daily <= 0:
        return cur
    while remaining > 0:
        day_end = cur.replace(hour=hours.end_h, minute=hours.end_m, second=0, microsecond=0)
        avail = (day_end - cur).total_seconds() / 60.0
        if avail >= remaining:
            return cur + timedelta(minutes=remaining)
        remaining -= avail
        # jump to next working day
        nxt_day = (cur + timedelta(days=1)).date()
        while not hours.is_working_day(nxt_day):
            nxt_day = nxt_day + timedelta(days=1)
        cur = datetime.combine(nxt_day, dtime(hours.start_h, hours.start_m))
    return cur


def _snap_to_shop(dt: datetime, hours: ShopHours) -> datetime:
    """Move a moment forward to the next valid working instant."""
    cur = dt.replace(second=0, microsecond=0)
    # roll forward to a working day
    while not hours.is_working_day(cur.date()):
        cur = datetime.combine(cur.date() + timedelta(days=1), dtime(hours.start_h, hours.start_m))
    start_of_day = cur.replace(hour=hours.start_h, minute=hours.start_m, second=0, microsecond=0)
    end_of_day = cur.replace(hour=hours.end_h, minute=hours.end_m, second=0, microsecond=0)
    if cur < start_of_day:
        return start_of_day
    if cur >= end_of_day:
        # go to next working day start
        nxt_day = cur.date() + timedelta(days=1)
        while not hours.is_working_day(nxt_day):
            nxt_day = nxt_day + timedelta(days=1)
        return datetime.combine(nxt_day, dtime(hours.start_h, hours.start_m))
    return cur


def stage_duration_minutes(row: sqlite3.Row, quantity: int) -> float:
    setup = float(row["setup_minutes"] or 0)
    per_unit = float(row["per_unit_minutes"] or 0)
    return setup + per_unit * max(1, int(quantity))


def recompute_schedule(db: DB, order_id: int) -> None:
    """Recompute planned_start/planned_end for remaining stages of an order.

    Simple rule-based policy:
      - For each non-done, non-skipped stage in sequence order, planned_start
        is max(now, previous stage's planned_end/actual_end).
      - planned_end = planned_start + (setup + per_unit * qty) clamped to
        shop hours.
      - Workers/stations are not globally resource-loaded here; this is an
        order-local ETA. Global capacity-aware recompute is done by
        recompute_all().
    """
    order = db.query_one("SELECT * FROM production_orders WHERE id=?", (order_id,))
    if not order:
        return
    qty = int(order["quantity"])
    hours = ShopHours.from_db(db)
    stages = db.query("SELECT * FROM order_stages WHERE order_id=? ORDER BY sequence", (order_id,))
    cursor = datetime.now()
    for s in stages:
        if s["status"] == "skipped":
            continue
        if s["status"] == "done":
            if s["actual_end"]:
                try:
                    cursor = max(cursor, datetime.fromisoformat(s["actual_end"]))
                except Exception:
                    pass
            continue
        if not s["is_optional_selected"]:
            continue
        dur = stage_duration_minutes(s, qty)
        if s["status"] == "running" and s["actual_start"]:
            try:
                start = datetime.fromisoformat(s["actual_start"])
            except Exception:
                start = cursor
        else:
            start = _snap_to_shop(cursor, hours)
        end = add_working_minutes(start, dur, hours)
        db.execute(
            "UPDATE order_stages SET planned_start=?, planned_end=? WHERE id=?",
            (start.isoformat(timespec="minutes"), end.isoformat(timespec="minutes"), s["id"]),
        )
        cursor = end


def _remaining_stage_minutes(stages: List[sqlite3.Row], quantity: int) -> List[float]:
    remaining = [0.0] * len(stages)
    tail = 0.0
    for idx in range(len(stages) - 1, -1, -1):
        s = stages[idx]
        if s["is_optional_selected"] and s["status"] not in ("done", "skipped"):
            tail += stage_duration_minutes(s, quantity)
        remaining[idx] = tail
    return remaining


def _candidate_stage_plan(
    db: DB,
    state: Dict[str, Any],
    stage: sqlite3.Row,
    *,
    desired_start: datetime,
    hours: ShopHours,
) -> Dict[str, Any]:
    duration = stage_duration_minutes(stage, int(state["quantity"]))
    existing_assigned = int(stage["assigned_worker_id"]) if stage["assigned_worker_id"] else None

    candidate_rows: List[sqlite3.Row]
    if existing_assigned is not None:
        assigned = db.query_one(
            "SELECT id, name FROM workers WHERE id=? AND is_active=1",
            (existing_assigned,),
        )
        candidate_rows = [assigned] if assigned else []
    else:
        candidate_rows = candidate_workers_for_order_stage(
            db, int(state["product_id"]), int(stage["stage_template_id"])
        )

    best: Optional[Tuple[datetime, datetime, Optional[int], str]] = None
    for candidate in candidate_rows:
        wid = int(candidate["id"])
        slot_start, slot_end = _find_worker_slot(
            db,
            wid,
            desired_start,
            duration,
            hours,
            exclude_stage_id=int(stage["id"]),
        )
        candidate_name = str(candidate["name"])
        choice = (slot_start, slot_end, wid, candidate_name)
        if best is None or (slot_end, slot_start, candidate_name) < (best[1], best[0], best[3]):
            best = choice

    if best is None:
        slot_start = _snap_to_shop(desired_start, hours)
        slot_end = add_working_minutes(slot_start, duration, hours)
        return {
            "stage": stage,
            "slot_start": slot_start,
            "slot_end": slot_end,
            "worker_id": existing_assigned,
            "worker_name": "",
            "desired_start": desired_start,
            "remaining_total": float(state["remaining"][state["next_idx"]]),
            "has_worker": False,
        }

    slot_start, slot_end, wid, wname = best
    return {
        "stage": stage,
        "slot_start": slot_start,
        "slot_end": slot_end,
        "worker_id": wid,
        "worker_name": wname,
        "desired_start": desired_start,
        "remaining_total": float(state["remaining"][state["next_idx"]]),
        "has_worker": True,
    }


def plan_all_active_orders(db: DB) -> None:
    """Globally plan all active orders stage-by-stage across shared workers.

    The planner keeps done/running/paused/blocked work stable, then schedules
    remaining planned stages across all active orders using a global dispatch
    rule:
      - higher priority first
      - started orders, and rework lots from started orders, before untouched orders
      - rework lots before normal planned work at the same priority
      - earlier due dates first
      - shorter remaining work first
      - earlier achievable finish as tie-breaker

    This produces a much better factory-wide plan than scheduling one order at
    a time, while still respecting worker-stage permissions and no-overlap
    capacity rules.
    """
    hours = ShopHours.from_db(db)
    now = _snap_to_shop(datetime.now(), hours)

    db.execute(
        """UPDATE order_stages
           SET planned_start=NULL, planned_end=NULL
           WHERE status='planned'
             AND is_optional_selected=1
             AND order_id IN (
                 SELECT id FROM production_orders
                 WHERE status NOT IN ('done','cancelled')
             )"""
    )

    orders = db.query(
        """SELECT po.id, po.product_id, po.quantity, po.priority, po.due_at,
                  po.created_at, po.status, po.parent_order_id, po.root_order_id,
                  parent.status AS parent_status,
                  root.status AS root_status,
                  COALESCE(root.created_at, po.created_at) AS root_created_at
           FROM production_orders po
           LEFT JOIN production_orders parent ON parent.id = po.parent_order_id
           LEFT JOIN production_orders root ON root.id = po.root_order_id
           WHERE po.status NOT IN ('done','cancelled')
           ORDER BY po.priority, COALESCE(po.due_at,'9999'), COALESCE(root.created_at, po.created_at), po.created_at"""
    )
    if not orders:
        return

    stage_rows = db.query(
        """SELECT os.*
           FROM order_stages os
           JOIN production_orders po ON po.id = os.order_id
           WHERE po.status NOT IN ('done','cancelled')
             AND os.is_optional_selected=1
             AND os.status <> 'skipped'
           ORDER BY os.order_id, os.sequence"""
    )
    stages_by_order: Dict[int, List[sqlite3.Row]] = {}
    for row in stage_rows:
        stages_by_order.setdefault(int(row["order_id"]), []).append(row)

    states: List[Dict[str, Any]] = []
    for order in orders:
        order_id = int(order["id"])
        stages = stages_by_order.get(order_id, [])
        if not stages:
            continue
        active_family = (
            order["status"] in ("running", "paused", "blocked")
            or order["parent_status"] in ("running", "paused", "blocked")
            or order["root_status"] in ("running", "paused", "blocked")
        )
        is_rework_lot = order["parent_order_id"] is not None
        states.append(
            {
                "order_id": order_id,
                "product_id": int(order["product_id"]),
                "quantity": int(order["quantity"]),
                "priority": int(order["priority"]),
                "due_dt": _parse_iso_safe(order["due_at"]),
                "created_dt": _parse_iso_safe(order["root_created_at"]) or _parse_iso_safe(order["created_at"]) or now,
                "started_rank": 0 if active_family else 1,
                "rework_rank": 0 if is_rework_lot else 1,
                "stages": stages,
                "remaining": _remaining_stage_minutes(stages, int(order["quantity"])),
                "next_idx": 0,
                "ready_at": now,
            }
        )

    def advance_state(state: Dict[str, Any]) -> Optional[sqlite3.Row]:
        stages = state["stages"]
        while state["next_idx"] < len(stages):
            s = stages[state["next_idx"]]
            status = str(s["status"])
            if status == "done":
                done_end = _parse_iso_safe(s["actual_end"]) or _parse_iso_safe(s["planned_end"])
                if done_end:
                    state["ready_at"] = max(state["ready_at"], done_end)
                state["next_idx"] += 1
                continue
            if status in ("running", "paused", "blocked"):
                start = _parse_iso_safe(s["actual_start"]) or _parse_iso_safe(s["planned_start"]) or _snap_to_shop(state["ready_at"], hours)
                end = _parse_iso_safe(s["planned_end"]) or _parse_iso_safe(s["actual_end"])
                if not end or end <= start:
                    end = add_working_minutes(start, stage_duration_minutes(s, int(state["quantity"])), hours)
                db.execute(
                    "UPDATE order_stages SET planned_start=?, planned_end=? WHERE id=?",
                    (start.isoformat(timespec="minutes"), end.isoformat(timespec="minutes"), int(s["id"])),
                )
                state["ready_at"] = max(state["ready_at"], start if s["early_release"] else end)
                state["next_idx"] += 1
                continue
            if status == "planned":
                return s
            state["next_idx"] += 1
        return None

    while True:
        candidate_plans: List[Tuple[Tuple[Any, ...], Dict[str, Any], Dict[str, Any]]] = []
        for state in states:
            stage = advance_state(state)
            if not stage:
                continue
            desired_start = _snap_to_shop(state["ready_at"], hours)
            plan = _candidate_stage_plan(db, state, stage, desired_start=desired_start, hours=hours)
            due_key = plan["desired_start"] if state["due_dt"] is None else state["due_dt"]
            score = (
                state["priority"],
                state["started_rank"],
                state["rework_rank"],
                due_key,
                plan["remaining_total"],
                plan["slot_end"],
                state["created_dt"],
                int(stage["sequence"]),
            )
            candidate_plans.append((score, state, plan))

        if not candidate_plans:
            break

        _, state, plan = min(candidate_plans, key=lambda x: x[0])
        stage = plan["stage"]
        worker_id = plan["worker_id"]
        existing_assigned = int(stage["assigned_worker_id"]) if stage["assigned_worker_id"] else None

        if worker_id is not None:
            db.execute(
                "UPDATE order_stages SET planned_start=?, planned_end=?, assigned_worker_id=? WHERE id=?",
                (
                    plan["slot_start"].isoformat(timespec="minutes"),
                    plan["slot_end"].isoformat(timespec="minutes"),
                    worker_id,
                    int(stage["id"]),
                ),
            )
            if existing_assigned != worker_id:
                db.log_event(
                    int(state["order_id"]),
                    "WORKER_ASSIGNED",
                    order_stage_id=int(stage["id"]),
                    actor="system",
                    payload={
                        "worker_id": worker_id,
                        "worker_name": plan["worker_name"],
                        "auto_plan": True,
                        "global_plan": True,
                    },
                )
        else:
            db.execute(
                "UPDATE order_stages SET planned_start=?, planned_end=? WHERE id=?",
                (
                    plan["slot_start"].isoformat(timespec="minutes"),
                    plan["slot_end"].isoformat(timespec="minutes"),
                    int(stage["id"]),
                ),
            )

        state["ready_at"] = max(
            state["ready_at"],
            plan["slot_start"] if stage["early_release"] else plan["slot_end"],
        )
        state["next_idx"] += 1

    for state in states:
        refresh_production_order_status(db, int(state["order_id"]))


def recompute_all(db: DB) -> None:
    """Recompute factory-wide schedule for every active order."""
    plan_all_active_orders(db)


def _parse_iso_safe(ts: Optional[str]) -> Optional[datetime]:
    if not ts:
        return None
    try:
        return datetime.fromisoformat(ts)
    except Exception:
        return None


def _interval_overlap(a_start: datetime, a_end: datetime, b_start: datetime, b_end: datetime) -> bool:
    return a_start < b_end and b_start < a_end


def _stage_interval_bounds(row: sqlite3.Row) -> Tuple[Optional[datetime], Optional[datetime]]:
    start = _parse_iso_safe(row["actual_start"]) or _parse_iso_safe(row["planned_start"])
    end = _parse_iso_safe(row["actual_end"]) or _parse_iso_safe(row["planned_end"])
    if start and end and end > start:
        return start, end
    return start, None


def _worker_stage_conflicts(
    db: DB, worker_id: int, *, exclude_stage_id: Optional[int] = None
) -> List[sqlite3.Row]:
    params: List[Any] = [worker_id]
    sql = (
        """SELECT os.id, os.planned_start, os.planned_end, os.actual_start, os.actual_end, os.status
           FROM order_stages os
           JOIN production_orders po ON po.id = os.order_id
           WHERE os.assigned_worker_id = ?
             AND po.status NOT IN ('done','cancelled')
             AND os.status NOT IN ('done','skipped')"""
    )
    if exclude_stage_id is not None:
        sql += " AND os.id <> ?"
        params.append(exclude_stage_id)
    sql += " ORDER BY COALESCE(os.actual_start, os.planned_start), COALESCE(os.actual_end, os.planned_end)"
    return db.query(sql, tuple(params))


def _worker_is_available_now(
    db: DB, worker_id: int, *, exclude_stage_id: Optional[int] = None
) -> bool:
    row = db.query_one(
        """SELECT 1
           FROM order_stages os
           JOIN production_orders po ON po.id = os.order_id
           WHERE os.assigned_worker_id = ?
             AND po.status NOT IN ('done','cancelled')
             AND os.status = 'running'"""
        + (" AND os.id <> ?" if exclude_stage_id is not None else "")
        + " LIMIT 1",
        (worker_id, exclude_stage_id) if exclude_stage_id is not None else (worker_id,),
    )
    return row is None


def _find_worker_slot(
    db: DB,
    worker_id: int,
    desired_start: datetime,
    duration_minutes: float,
    hours: ShopHours,
    *,
    exclude_stage_id: Optional[int] = None,
) -> Tuple[datetime, datetime]:
    start = _snap_to_shop(desired_start, hours)
    conflicts = _worker_stage_conflicts(db, worker_id, exclude_stage_id=exclude_stage_id)
    while True:
        end = add_working_minutes(start, duration_minutes, hours)
        moved = False
        for conflict in conflicts:
            c_start, c_end = _stage_interval_bounds(conflict)
            if not c_start or not c_end or c_end <= c_start:
                continue
            if _interval_overlap(start, end, c_start, c_end):
                start = _snap_to_shop(max(start, c_end), hours)
                moved = True
                break
        if not moved:
            return start, end


def refresh_production_order_status(db: DB, order_id: int) -> Optional[str]:
    order = db.query_one("SELECT status FROM production_orders WHERE id=?", (order_id,))
    if not order or order["status"] in ("done", "cancelled"):
        return order["status"] if order else None
    if db.query_one(
        "SELECT 1 FROM order_stages WHERE order_id=? AND status='running' LIMIT 1",
        (order_id,),
    ):
        status = "running"
    elif db.query_one(
        "SELECT 1 FROM order_stages WHERE order_id=? AND status='blocked' LIMIT 1",
        (order_id,),
    ):
        status = "blocked"
    elif db.query_one(
        "SELECT 1 FROM order_stages WHERE order_id=? AND status='paused' LIMIT 1",
        (order_id,),
    ):
        status = "paused"
    else:
        status = "released"
    if status != order["status"]:
        db.execute(
            "UPDATE production_orders SET status=? WHERE id=? AND status NOT IN ('done','cancelled')",
            (status, order_id),
        )
    return status


def _available_workers_for_stage_db(
    db: DB, stage_template_id: int, *, exclude_stage_id: Optional[int] = None
) -> List[sqlite3.Row]:
    """DB-level availability query used by auto-dispatch and manual assignment."""
    params: List[Any] = [stage_template_id]
    sql = (
        """SELECT DISTINCT w.id, w.name
           FROM workers w
           JOIN worker_stages ws ON ws.worker_id = w.id
           WHERE ws.stage_template_id=? AND w.is_active=1
             AND NOT EXISTS (
                SELECT 1
                FROM order_stages os
                WHERE os.assigned_worker_id = w.id
                  AND os.status='running'"""
    )
    if exclude_stage_id is not None:
        sql += " AND os.id <> ?"
        params.append(exclude_stage_id)
    sql += ") ORDER BY w.name"
    return db.query(sql, tuple(params))


def candidate_workers_for_order_stage(
    db: DB, product_id: int, stage_template_id: int
) -> List[sqlite3.Row]:
    """Workers explicitly allowed to perform this exact stage."""
    return db.query(
        """SELECT w.id, w.name
           FROM workers w
           JOIN worker_stages ws ON ws.worker_id = w.id
           WHERE ws.stage_template_id = ? AND w.is_active = 1
           ORDER BY w.name""",
        (stage_template_id,),
    )


def auto_allocate_workers_by_timeline(db: DB, order_id: int) -> int:
    """Allocate workers and capacity-aware time windows for an order.

    This is *planning-time* allocation (before pressing Start):
      - For each remaining stage, respect sequence order and shop hours.
      - Pick a worker allowed for that exact stage whose first free slot yields
        the earliest completion.
      - Never force an overlapping fallback assignment.
      - Existing manual assignments are kept, but their planned slot is pushed
        forward if that worker is already busy elsewhere.
    Returns number of new assignments made.
    """
    order_row = db.query_one("SELECT product_id, quantity FROM production_orders WHERE id=?", (order_id,))
    if not order_row:
        return 0
    product_id = int(order_row["product_id"])
    quantity = int(order_row["quantity"])
    hours = ShopHours.from_db(db)

    stages = db.query(
        """SELECT os.*
           FROM order_stages os
           JOIN production_orders po ON po.id=os.order_id
           WHERE os.order_id=?
             AND po.status NOT IN ('done','cancelled')
             AND os.is_optional_selected=1
             AND os.status NOT IN ('done','skipped')
           ORDER BY os.sequence""",
        (order_id,),
    )
    assigned_count = 0
    cursor = datetime.now()
    for s in stages:
        if s["status"] == "done":
            done_end = _parse_iso_safe(s["actual_end"]) or _parse_iso_safe(s["planned_end"])
            if done_end:
                cursor = max(cursor, done_end)
            continue
        if s["status"] in ("running", "paused", "blocked"):
            active_start = _parse_iso_safe(s["actual_start"]) or _parse_iso_safe(s["planned_start"]) or _snap_to_shop(cursor, hours)
            active_end = _parse_iso_safe(s["planned_end"])
            if not active_end or active_end <= active_start:
                active_end = add_working_minutes(active_start, stage_duration_minutes(s, quantity), hours)
            db.execute(
                "UPDATE order_stages SET planned_start=?, planned_end=? WHERE id=?",
                (active_start.isoformat(timespec="minutes"), active_end.isoformat(timespec="minutes"), int(s["id"])),
            )
            cursor = max(cursor, active_end)
            continue

        desired_start = max(
            cursor,
            _parse_iso_safe(s["planned_start"]) or _snap_to_shop(cursor, hours),
        )
        duration = stage_duration_minutes(s, quantity)

        existing_assigned = int(s["assigned_worker_id"]) if s["assigned_worker_id"] else None
        candidate_rows: List[sqlite3.Row]
        if existing_assigned is not None:
            assigned = db.query_one(
                "SELECT id, name FROM workers WHERE id=? AND is_active=1",
                (existing_assigned,),
            )
            candidate_rows = [assigned] if assigned else []
        else:
            candidate_rows = candidate_workers_for_order_stage(
                db, product_id, int(s["stage_template_id"])
            )

        best: Optional[Tuple[datetime, datetime, sqlite3.Row]] = None
        for candidate in candidate_rows:
            slot_start, slot_end = _find_worker_slot(
                db,
                int(candidate["id"]),
                desired_start,
                duration,
                hours,
                exclude_stage_id=int(s["id"]),
            )
            choice = (slot_start, slot_end, candidate)
            if best is None or (slot_end, slot_start, str(candidate["name"])) < (
                best[1], best[0], str(best[2]["name"])
            ):
                best = choice

        if best:
            slot_start, slot_end, picked = best
            wid = int(picked["id"])
            db.execute(
                "UPDATE order_stages SET planned_start=?, planned_end=?, assigned_worker_id=? WHERE id=?",
                (slot_start.isoformat(timespec="minutes"), slot_end.isoformat(timespec="minutes"), wid, int(s["id"])),
            )
            if existing_assigned != wid:
                db.log_event(
                    order_id,
                    "WORKER_ASSIGNED",
                    order_stage_id=int(s["id"]),
                    actor="system",
                    payload={"worker_id": wid, "worker_name": picked["name"], "auto_plan": True},
                )
                assigned_count += 1
            cursor = slot_end
        else:
            slot_start = _snap_to_shop(desired_start, hours)
            slot_end = add_working_minutes(slot_start, duration, hours)
            db.execute(
                "UPDATE order_stages SET planned_start=?, planned_end=? WHERE id=?",
                (slot_start.isoformat(timespec="minutes"), slot_end.isoformat(timespec="minutes"), int(s["id"])),
            )
            cursor = slot_end
    return assigned_count


def sync_order_state(
    db: DB,
    order_id: int,
    *,
    dispatch: bool = False,
    ensure_started: bool = False,
) -> bool:
    """Recompute one order once, then optionally advance execution.

    Centralizing this keeps user actions from triggering duplicate full
    recompute/allocation cycles and makes UI refreshes feel much snappier.
    """
    plan_all_active_orders(db)
    progressed = False
    if dispatch:
        progressed = auto_dispatch_order(db, order_id) or progressed
    if ensure_started:
        progressed = ensure_production_order_started(db, order_id) or progressed
    refresh_production_order_status(db, order_id)
    if progressed:
        plan_all_active_orders(db)
        refresh_production_order_status(db, order_id)
    return progressed


def normalize_order_statuses(db: DB) -> None:
    """Align legacy order statuses with actual stage execution state."""
    active = db.query(
        """SELECT id
           FROM production_orders
           WHERE status NOT IN ('done','cancelled')"""
    )
    for row in active:
        refresh_production_order_status(db, int(row["id"]))


def ensure_production_order_started(db: DB, order_id: int) -> bool:
    """If the order has no running stage, start the first eligible planned stage
    and set order status to running. Used as a hard guarantee after new orders.
    """
    if db.query_one(
        "SELECT 1 AS x FROM order_stages WHERE order_id=? AND status='running' LIMIT 1",
        (order_id,),
    ):
        return True
    order = db.query_one("SELECT * FROM production_orders WHERE id=?", (order_id,))
    if not order or order["status"] in ("done", "cancelled"):
        return False
    nxt = db.query_one(
        """SELECT * FROM order_stages
           WHERE order_id=? AND status='planned' AND is_optional_selected=1
           ORDER BY sequence LIMIT 1""",
        (order_id,),
    )
    if not nxt:
        return False
    pred = db.query_one(
        """SELECT * FROM order_stages
           WHERE order_id=? AND sequence < ? AND is_optional_selected=1
           ORDER BY sequence DESC LIMIT 1""",
        (order_id, int(nxt["sequence"])),
    )
    if pred and pred["status"] != "done" and not pred["early_release"]:
        return False

    product_id = int(order["product_id"])
    now = datetime.now().isoformat(timespec="minutes")
    wid: Optional[int] = None
    wname: str = ""
    if nxt["assigned_worker_id"]:
        wid = int(nxt["assigned_worker_id"])
        wr = db.query_one("SELECT name FROM workers WHERE id=?", (wid,))
        wname = str(wr["name"]) if wr else ""
    if wid is not None and not _worker_is_available_now(db, wid, exclude_stage_id=int(nxt["id"])):
        wid = None
        wname = ""
    if wid is None:
        cands = candidate_workers_for_order_stage(db, product_id, int(nxt["stage_template_id"]))
        av = {int(x["id"]): str(x["name"]) for x in _available_workers_for_stage_db(
            db, int(nxt["stage_template_id"]), exclude_stage_id=int(nxt["id"])
        )}
        for c in cands:
            cid = int(c["id"])
            if cid in av:
                wid, wname = cid, av[cid]
                break
    if wid is None:
        db.log_event(
            order_id,
            "AUTO_START_DEFERRED",
            order_stage_id=int(nxt["id"]),
            actor="system",
            reason="No qualified worker available now; kept queued in pipeline",
        )
        refresh_production_order_status(db, order_id)
        return False

    db.execute(
        "UPDATE order_stages SET assigned_worker_id=COALESCE(assigned_worker_id,?),"
        " status='running', actual_start=COALESCE(actual_start, ?) WHERE id=?",
        (wid, now, int(nxt["id"])),
    )
    db.log_event(
        order_id, "WORKER_ASSIGNED", order_stage_id=int(nxt["id"]), actor="system",
        payload={"worker_id": wid, "worker_name": wname, "ensure_start": True},
    )
    db.execute(
        "UPDATE production_orders SET status='running' WHERE id=? AND status NOT IN ('done','cancelled')",
        (order_id,),
    )
    db.log_event(order_id, "STAGE_STARTED", order_stage_id=int(nxt["id"]), actor="system")
    return True


def auto_dispatch_order(db: DB, order_id: int) -> bool:
    """Auto-assign and auto-start next eligible stage for an order.

    Returns True if it started a stage; False otherwise.
    """
    order = db.query_one("SELECT * FROM production_orders WHERE id=?", (order_id,))
    if not order or order["status"] in ("done", "cancelled"):
        return False

    planned_stages = db.query(
        """SELECT * FROM order_stages
           WHERE order_id=? AND status='planned' AND is_optional_selected=1
           ORDER BY sequence""",
        (order_id,),
    )
    next_stage: Optional[sqlite3.Row] = None
    for candidate in planned_stages:
        pred = db.query_one(
            """SELECT * FROM order_stages
               WHERE order_id=? AND sequence < ? AND is_optional_selected=1
               ORDER BY sequence DESC LIMIT 1""",
            (order_id, int(candidate["sequence"])),
        )
        if pred and pred["status"] != "done" and not pred["early_release"]:
            continue
        next_stage = candidate
        break
    if not next_stage:
        return False

    workers = _available_workers_for_stage_db(
        db, int(next_stage["stage_template_id"]), exclude_stage_id=int(next_stage["id"])
    )
    wid: Optional[int] = None
    wname: str = ""
    if next_stage["assigned_worker_id"]:
        assigned_id = int(next_stage["assigned_worker_id"])
        if _worker_is_available_now(db, assigned_id, exclude_stage_id=int(next_stage["id"])):
            wid = assigned_id
            wr = db.query_one("SELECT name FROM workers WHERE id=?", (wid,))
            wname = str(wr["name"]) if wr else ""
    if wid is None and workers:
        wid = int(workers[0]["id"])
        wname = str(workers[0]["name"])
    if wid is None:
        db.log_event(
            order_id,
            "AUTO_START_DEFERRED",
            order_stage_id=int(next_stage["id"]),
            actor="system",
            reason="No qualified worker available now; kept queued in pipeline",
        )
        refresh_production_order_status(db, order_id)
        return False

    now = datetime.now().isoformat(timespec="minutes")
    db.execute(
        "UPDATE order_stages SET assigned_worker_id=?, status='running', actual_start=COALESCE(actual_start,?) WHERE id=?",
        (wid, now, int(next_stage["id"])),
    )
    db.execute(
        "UPDATE production_orders SET status='running' WHERE id=? AND status NOT IN ('done','cancelled')",
        (order_id,),
    )
    db.log_event(
        order_id,
        "WORKER_ASSIGNED",
        order_stage_id=int(next_stage["id"]),
        actor="system",
        payload={"worker_id": wid, "worker_name": wname, "auto": True},
    )
    db.log_event(order_id, "STAGE_STARTED", order_stage_id=int(next_stage["id"]), actor="system")
    return True


def start_next_stage_in_parallel(
    db: DB,
    order_id: int,
    current_stage_id: int,
    *,
    actor: str = "manager",
) -> bool:
    current_stage = db.query_one(
        "SELECT * FROM order_stages WHERE id=? AND order_id=?",
        (current_stage_id, order_id),
    )
    if not current_stage:
        return False
    if current_stage["status"] not in ("running", "paused"):
        return False
    db.execute(
        "UPDATE order_stages SET early_release=1 WHERE id=?",
        (current_stage_id,),
    )
    db.log_event(
        order_id,
        "STAGE_EARLY_RELEASE",
        order_stage_id=current_stage_id,
        actor=actor,
        payload={"parallel_start": True},
    )
    progressed = auto_dispatch_order(db, order_id)
    if progressed:
        plan_all_active_orders(db)
    refresh_production_order_status(db, order_id)
    return progressed


def complete_stage_and_advance(
    db: DB,
    stage_id: int,
    *,
    actor: str = "user",
    reason: Optional[str] = None,
) -> bool:
    stage = db.query_one("SELECT * FROM order_stages WHERE id=?", (stage_id,))
    if not stage or stage["status"] != "running":
        return False
    order_id = int(stage["order_id"])
    now_dt = datetime.now()
    now = now_dt.isoformat(timespec="minutes")
    actual_minutes: Optional[float] = None
    if stage["actual_start"]:
        try:
            actual_minutes = max(
                0.0,
                (now_dt - datetime.fromisoformat(str(stage["actual_start"]))).total_seconds() / 60.0,
            )
        except Exception:
            actual_minutes = None
    db.execute(
        "UPDATE order_stages SET status='done', actual_end=?, actual_minutes=? WHERE id=?",
        (now, actual_minutes, stage_id),
    )
    db.log_event(
        order_id,
        "STAGE_COMPLETED",
        order_stage_id=stage_id,
        actor=actor,
        reason=reason,
    )
    if db.query_one(
        """SELECT 1
           FROM order_stages
           WHERE order_id=?
             AND is_optional_selected=1
             AND status NOT IN ('done','skipped')
           LIMIT 1""",
        (order_id,),
    ):
        sync_order_state(db, order_id)
    else:
        db.execute(
            "UPDATE production_orders SET status='done', completed_at=? WHERE id=?",
            (now, order_id),
        )
        db.log_event(order_id, "ORDER_COMPLETED", actor="system")
    refresh_production_order_status(db, order_id)
    return True


def _next_rework_lot_number(db: DB, root_order_number: str) -> str:
    for n in range(1, 1000):
        candidate = f"{root_order_number}-R{n}"
        if not db.query_one("SELECT 1 FROM production_orders WHERE order_number=?", (candidate,)):
            return candidate
    raise ValueError("تعذر إنشاء رقم دفعة إعادة عمل جديد")


def split_order_lot(
    db: DB,
    order_id: int,
    split_quantity: int,
    split_from_stage_id: int,
    return_stage_id: int,
    *,
    actor: str = "user",
    reason: Optional[str] = None,
) -> int:
    """Split part of an order into a child rework lot.

    The current order keeps the remaining quantity and continues forward. The
    new child lot gets its own order row and copied routing, with stages before
    `return_stage_id` already marked done so it can move independently from the
    chosen earlier stage.
    """
    order = db.query_one("SELECT * FROM production_orders WHERE id=?", (order_id,))
    if not order:
        raise ValueError("الأمر غير موجود")
    if order["status"] in ("done", "cancelled"):
        raise ValueError("لا يمكن تقسيم أمر منتهي أو ملغي")

    current_qty = int(order["quantity"])
    split_quantity = int(split_quantity)
    if split_quantity <= 0 or split_quantity >= current_qty:
        raise ValueError("كمية التقسيم يجب أن تكون أقل من كمية الدفعة الحالية")

    split_from_stage = db.query_one(
        "SELECT * FROM order_stages WHERE id=? AND order_id=?",
        (split_from_stage_id, order_id),
    )
    if not split_from_stage:
        raise ValueError("مرحلة المشكلة غير موجودة في هذا الأمر")

    return_stage = db.query_one(
        """SELECT * FROM order_stages
           WHERE id=? AND order_id=? AND is_optional_selected=1 AND status<>'skipped'""",
        (return_stage_id, order_id),
    )
    if not return_stage:
        raise ValueError("مرحلة الرجوع غير متاحة في هذا الأمر")
    if int(return_stage["sequence"]) > int(split_from_stage["sequence"]):
        raise ValueError("مرحلة الرجوع يجب أن تكون نفس مرحلة المشكلة أو مرحلة قبلها")

    root_id = int(order["root_order_id"] or order_id)
    root_order = db.query_one("SELECT order_number FROM production_orders WHERE id=?", (root_id,))
    root_order_number = root_order["order_number"] if root_order else order["order_number"]
    child_order_number = _next_rework_lot_number(db, root_order_number)
    remaining_qty = current_qty - split_quantity
    now_iso = datetime.now().isoformat(timespec="minutes")
    return_seq = int(return_stage["sequence"])

    source_stages = db.query(
        "SELECT * FROM order_stages WHERE order_id=? ORDER BY sequence",
        (order_id,),
    )

    with db._lock:
        db.execute("BEGIN IMMEDIATE")
        try:
            db.execute(
                """UPDATE production_orders
                   SET quantity=?,
                       root_order_id=COALESCE(root_order_id, id),
                       original_quantity=COALESCE(original_quantity, ?)
                   WHERE id=?""",
                (remaining_qty, current_qty, order_id),
            )
            db.execute(
                """UPDATE order_stages
                   SET planned_start=NULL, planned_end=NULL
                   WHERE order_id=?
                     AND status IN ('planned','running','paused','blocked')""",
                (order_id,),
            )
            db.execute(
                """INSERT INTO production_orders
                   (order_number, product_id, quantity, priority, due_at, status,
                    school_name, size, color, notes,
                    root_order_id, parent_order_id, original_quantity,
                    split_from_stage_id, split_reason,
                    created_at, released_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    child_order_number,
                    order["product_id"],
                    split_quantity,
                    order["priority"],
                    order["due_at"],
                    "released",
                    order["school_name"],
                    order["size"],
                    order["color"],
                    order["notes"],
                    root_id,
                    order_id,
                    split_quantity,
                    split_from_stage_id,
                    reason,
                    now_iso,
                    now_iso,
                ),
            )
            child_id = int(
                db.query_one(
                    "SELECT id FROM production_orders WHERE order_number=?",
                    (child_order_number,),
                )["id"]
            )

            child_return_stage_id: Optional[int] = None
            for src in source_stages:
                seq = int(src["sequence"])
                src_selected = int(src["is_optional_selected"] or 0) == 1
                if not src_selected or src["status"] == "skipped":
                    child_status = "skipped"
                    child_selected = 0
                    planned_start = planned_end = actual_start = actual_end = None
                    actual_minutes = None
                elif seq < return_seq:
                    child_status = "done"
                    child_selected = 1
                    planned_start = src["planned_start"]
                    planned_end = src["planned_end"]
                    actual_start = src["actual_start"]
                    actual_end = src["actual_end"]
                    actual_minutes = src["actual_minutes"]
                else:
                    child_status = "planned"
                    child_selected = 1
                    planned_start = planned_end = actual_start = actual_end = None
                    actual_minutes = None

                db.execute(
                    """INSERT INTO order_stages
                       (order_id, stage_template_id, sequence, name_snapshot,
                        setup_minutes, per_unit_minutes, is_optional_selected,
                        status, planned_start, planned_end, actual_start,
                        actual_end, actual_minutes, notes)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        child_id,
                        src["stage_template_id"],
                        seq,
                        src["name_snapshot"],
                        src["setup_minutes"],
                        src["per_unit_minutes"],
                        child_selected,
                        child_status,
                        planned_start,
                        planned_end,
                        actual_start,
                        actual_end,
                        actual_minutes,
                        src["notes"],
                    ),
                )
                if seq == return_seq:
                    child_return_stage_id = int(
                        db.query_one(
                            "SELECT id FROM order_stages WHERE order_id=? AND sequence=?",
                            (child_id, seq),
                        )["id"]
                    )

            payload = {
                "child_order_id": child_id,
                "child_order_number": child_order_number,
                "split_quantity": split_quantity,
                "remaining_quantity": remaining_qty,
                "return_stage_id": return_stage_id,
                "return_stage_name": return_stage["name_snapshot"],
                "return_sequence": return_seq,
            }
            db.log_event(
                order_id,
                "LOT_SPLIT",
                order_stage_id=split_from_stage_id,
                actor=actor,
                reason=reason,
                payload=payload,
            )
            db.log_event(
                child_id,
                "LOT_CREATED_FROM_SPLIT",
                order_stage_id=child_return_stage_id,
                actor=actor,
                reason=reason,
                payload={
                    "parent_order_id": order_id,
                    "parent_order_number": order["order_number"],
                    "root_order_id": root_id,
                    "root_order_number": root_order_number,
                    "split_from_stage_id": split_from_stage_id,
                    "split_from_stage_name": split_from_stage["name_snapshot"],
                    "return_stage_name": return_stage["name_snapshot"],
                },
            )
            db.execute("COMMIT")
        except Exception:
            db.execute("ROLLBACK")
            raise

    recompute_all(db)
    return child_id


def find_order_eta(db: DB, order_id: int) -> Optional[datetime]:
    row = db.query_one(
        """SELECT MAX(planned_end) AS last_end FROM order_stages
           WHERE order_id=? AND is_optional_selected=1 AND status<>'skipped'""",
        (order_id,),
    )
    if not row or not row["last_end"]:
        return None
    try:
        return datetime.fromisoformat(row["last_end"])
    except Exception:
        return None


def _month_key(ts: Optional[str]) -> str:
    if ts:
        try:
            return datetime.fromisoformat(str(ts)).strftime("%Y-%m")
        except Exception:
            pass
    return "unknown"


def _write_csv(path: str, rows: List[Dict[str, Any]], fieldnames: List[str]) -> None:
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow({k: row.get(k, "") for k in fieldnames})


def _row_to_dict(row: sqlite3.Row) -> Dict[str, Any]:
    return {k: row[k] for k in row.keys()}


def _write_json(path: str, payload: Dict[str, Any]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _write_jsonl(path: str, records: List[Dict[str, Any]]) -> None:
    with open(path, "w", encoding="utf-8") as f:
        for record in records:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")


def export_monthly_analytics(db: DB) -> None:
    """Write month-by-month AI-ready archives for later production analysis."""
    orders = db.query(
        """SELECT po.id, po.order_number, po.created_at, po.completed_at, po.status,
                  po.quantity, po.priority, po.school_name, po.size, po.color,
                  p.code AS product_code, p.name AS product_name, p.product_type
           FROM production_orders po
           JOIN products p ON p.id = po.product_id
           ORDER BY po.created_at, po.id"""
    )
    stages = db.query(
        """SELECT os.id, os.order_id, os.sequence, os.name_snapshot, os.status,
                  os.setup_minutes, os.per_unit_minutes, os.planned_start, os.planned_end,
                  os.actual_start, os.actual_end, os.actual_minutes,
                  w.name AS worker_name,
                  po.order_number, po.created_at AS order_created_at, po.school_name, po.size, po.color,
                  p.name AS product_name, p.product_type
           FROM order_stages os
           JOIN production_orders po ON po.id = os.order_id
           JOIN products p ON p.id = po.product_id
           LEFT JOIN workers w ON w.id = os.assigned_worker_id
           ORDER BY po.created_at, po.id, os.sequence"""
    )
    events = db.query(
        """SELECT oe.id, oe.order_id, oe.order_stage_id, oe.event_type, oe.actor, oe.reason,
                  oe.payload_json, oe.created_at,
                  po.order_number, po.school_name, po.size, po.color,
                  p.name AS product_name, p.product_type
           FROM order_events oe
           LEFT JOIN production_orders po ON po.id = oe.order_id
           LEFT JOIN products p ON p.id = po.product_id
           ORDER BY oe.created_at, oe.id"""
    )

    months: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}

    for row in orders:
        month = _month_key(row["created_at"])
        months.setdefault(month, {"orders": [], "stages": [], "events": []})
        months[month]["orders"].append({
            "order_id": row["id"],
            "order_number": row["order_number"],
            "created_at": row["created_at"],
            "completed_at": row["completed_at"],
            "status": row["status"],
            "quantity": row["quantity"],
            "priority": row["priority"],
            "product_code": row["product_code"],
            "product_type": row["product_type"] or row["product_name"],
            "product_name": row["product_name"],
            "school_name": row["school_name"],
            "size": row["size"],
            "color": row["color"],
        })

    for row in stages:
        month = _month_key(row["order_created_at"])
        months.setdefault(month, {"orders": [], "stages": [], "events": []})
        months[month]["stages"].append({
            "stage_id": row["id"],
            "order_id": row["order_id"],
            "order_number": row["order_number"],
            "sequence": row["sequence"],
            "stage_name": row["name_snapshot"],
            "status": row["status"],
            "worker_name": row["worker_name"],
            "product_type": row["product_type"] or row["product_name"],
            "product_name": row["product_name"],
            "school_name": row["school_name"],
            "size": row["size"],
            "color": row["color"],
            "planned_start": row["planned_start"],
            "planned_end": row["planned_end"],
            "actual_start": row["actual_start"],
            "actual_end": row["actual_end"],
            "actual_minutes": row["actual_minutes"],
            "setup_minutes": row["setup_minutes"],
            "per_unit_minutes": row["per_unit_minutes"],
        })

    for row in events:
        month = _month_key(row["created_at"])
        months.setdefault(month, {"orders": [], "stages": [], "events": []})
        months[month]["events"].append({
            "event_id": row["id"],
            "created_at": row["created_at"],
            "order_id": row["order_id"],
            "order_number": row["order_number"],
            "order_stage_id": row["order_stage_id"],
            "event_type": row["event_type"],
            "actor": row["actor"],
            "reason": row["reason"],
            "product_type": row["product_type"] or row["product_name"],
            "product_name": row["product_name"],
            "school_name": row["school_name"],
            "size": row["size"],
            "color": row["color"],
            "payload_json": row["payload_json"],
        })

    for month, data in months.items():
        month_dir = os.path.join(ANALYTICS_DIR, month)
        os.makedirs(month_dir, exist_ok=True)
        _write_csv(
            os.path.join(month_dir, "orders.csv"),
            data["orders"],
            [
                "order_id", "order_number", "created_at", "completed_at", "status",
                "quantity", "priority", "product_code", "product_type", "product_name",
                "school_name", "size", "color",
            ],
        )
        _write_csv(
            os.path.join(month_dir, "stages.csv"),
            data["stages"],
            [
                "stage_id", "order_id", "order_number", "sequence", "stage_name", "status",
                "worker_name", "product_type", "product_name", "school_name", "size", "color",
                "planned_start", "planned_end", "actual_start", "actual_end", "actual_minutes",
                "setup_minutes", "per_unit_minutes",
            ],
        )
        _write_csv(
            os.path.join(month_dir, "events.csv"),
            data["events"],
            [
                "event_id", "created_at", "order_id", "order_number", "order_stage_id",
                "event_type", "actor", "reason", "product_type", "product_name",
                "school_name", "size", "color", "payload_json",
            ],
        )

        summary = {
            "month": month,
            "exported_at": datetime.now().isoformat(timespec="seconds"),
            "orders_count": len(data["orders"]),
            "events_count": len(data["events"]),
            "stages_count": len(data["stages"]),
            "total_quantity": sum(int(r.get("quantity") or 0) for r in data["orders"]),
            "completed_orders": sum(1 for r in data["orders"] if r.get("status") == "done"),
            "running_orders": sum(1 for r in data["orders"] if r.get("status") == "running"),
            "blocked_orders": sum(1 for r in data["orders"] if r.get("status") == "blocked"),
            "paused_orders": sum(1 for r in data["orders"] if r.get("status") == "paused"),
            "avg_actual_stage_minutes": round(
                sum(float(r.get("actual_minutes") or 0.0) for r in data["stages"])
                / max(1, sum(1 for r in data["stages"] if r.get("actual_minutes") not in (None, ""))),
                2,
            ),
        }
        _write_json(os.path.join(month_dir, "summary.json"), summary)

        order_numbers = [str(r["order_number"]) for r in data["orders"]]
        stage_order_ids = [int(r["order_id"]) for r in data["stages"]]
        month_order_ids = sorted({int(r["order_id"]) for r in data["orders"]} | set(stage_order_ids))
        if month_order_ids:
            placeholders = ",".join("?" for _ in month_order_ids)
            prod_routes = db.query(
                f"""SELECT ps.*, p.code AS product_code, p.name AS product_name, p.product_type,
                           st.code AS stage_code, st.name AS stage_name
                    FROM product_stages ps
                    JOIN products p ON p.id = ps.product_id
                    JOIN stage_templates st ON st.id = ps.stage_template_id
                    WHERE p.id IN (
                        SELECT DISTINCT product_id FROM production_orders WHERE id IN ({placeholders})
                    )
                    ORDER BY p.id, ps.sequence""",
                tuple(month_order_ids),
            )
        else:
            prod_routes = []

        worker_rows = db.query(
            """SELECT w.*, GROUP_CONCAT(st.code, ',') AS allowed_stage_codes,
                      GROUP_CONCAT(st.name, ' | ') AS allowed_stage_names
               FROM workers w
               LEFT JOIN worker_stages ws ON ws.worker_id = w.id
               LEFT JOIN stage_templates st ON st.id = ws.stage_template_id
               GROUP BY w.id
               ORDER BY w.name"""
        )
        stage_templates = db.query(
            "SELECT * FROM stage_templates ORDER BY name"
        )
        app_settings = db.query(
            "SELECT key, value FROM app_settings ORDER BY key"
        )
        suggestions = db.query(
            "SELECT * FROM order_wizard_suggestions ORDER BY field, use_count DESC, value"
        )

        order_stage_map: Dict[int, List[Dict[str, Any]]] = {}
        for row in data["stages"]:
            order_stage_map.setdefault(int(row["order_id"]), []).append(row)
        order_event_map: Dict[int, List[Dict[str, Any]]] = {}
        for row in data["events"]:
            if row.get("order_id") is None:
                continue
            order_event_map.setdefault(int(row["order_id"]), []).append(row)

        ai_records: List[Dict[str, Any]] = []
        for order_row in data["orders"]:
            oid = int(order_row["order_id"])
            ai_records.append(
                {
                    "record_type": "order_bundle",
                    "month": month,
                    "exported_at": summary["exported_at"],
                    "order": order_row,
                    "stages": order_stage_map.get(oid, []),
                    "events": order_event_map.get(oid, []),
                    "derived": {
                        "event_count": len(order_event_map.get(oid, [])),
                        "stage_count": len(order_stage_map.get(oid, [])),
                        "completed_stage_count": sum(
                            1 for s in order_stage_map.get(oid, []) if s.get("status") == "done"
                        ),
                        "total_actual_minutes": round(
                            sum(float(s.get("actual_minutes") or 0.0) for s in order_stage_map.get(oid, [])),
                            2,
                        ),
                    },
                }
            )

        ai_snapshot = {
            "format": "hosny_factory_ai_archive_v2",
            "month": month,
            "exported_at": summary["exported_at"],
            "summary": summary,
            "settings": [_row_to_dict(r) for r in app_settings],
            "stage_templates": [_row_to_dict(r) for r in stage_templates],
            "workers": [_row_to_dict(r) for r in worker_rows],
            "product_routes": [_row_to_dict(r) for r in prod_routes],
            "wizard_suggestions": [_row_to_dict(r) for r in suggestions],
            "orders": data["orders"],
            "stages": data["stages"],
            "events": data["events"],
        }
        _write_json(os.path.join(month_dir, "ai_snapshot.json"), ai_snapshot)
        _write_jsonl(os.path.join(month_dir, "ai_order_bundles.jsonl"), ai_records)
        _write_json(
            os.path.join(month_dir, "ai_manifest.json"),
            {
                "format": "hosny_factory_ai_archive_v2",
                "month": month,
                "exported_at": summary["exported_at"],
                "files": [
                    "orders.csv",
                    "stages.csv",
                    "events.csv",
                    "summary.json",
                    "ai_snapshot.json",
                    "ai_order_bundles.jsonl",
                ],
                "purpose": "AI-only monthly production telemetry for bottleneck and improvement analysis.",
            },
        )


# =============================================================================
# UI helpers
# =============================================================================
