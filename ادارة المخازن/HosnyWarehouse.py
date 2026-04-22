# filepath: app/warehouse_manager_nosqlite_excel_billing.py
# Python 3.10+

try:
    import logging_setup
    logging_setup.install_crash_logging("HosnyWarehouse")
except Exception:
    pass

import json
import os
import hashlib
import sys, subprocess
import sqlite3
import time
import tempfile
import unicodedata
import webbrowser
from dataclasses import dataclass
from datetime import datetime, date   # +date for calendar
import calendar                       # ADD
from typing import Any, Dict, List, Optional, Sequence, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

DB_PATH = "warehouse_data.sqlite3"
LEGACY_JSON_PATH = "warehouse_data.json"
ADMIN_PASSWORD_PLAIN = "1234"
ADMIN_PASSWORD_HASH_PREFIX = "sha256$"

DEFAULT_BRANCH_POS_NAMES = [
    "POS-ZAY",
    "POS-OCT",
    "POS-OBO",
    "POS-GESR",
    "POS-BAH",
    "POS-CEN",
]
BRANCH_UI_NAME_BY_DEVICE = {
    "POS-ZAY": "فرع زايد",
    "POS-OCT": "فرع اكتوبر",
    "POS-BAH": "فرع بهتيم",
    "POS-CEN": "فرع السنتر",
    "POS-OBO": "فرع العبور",
    "POS-GESR": "فرع جسر السويس",
}
BRANCH_DEVICE_BY_UI_NAME = {v: k for k, v in BRANCH_UI_NAME_BY_DEVICE.items()}
WAREHOUSE_NUMBER_VALUES = ["1", "2", "3", "4", "5", "6", "7"]
WAREHOUSE_NUMBER_LABELS = {
    "1": "مخزن 1",
    "2": "مخزن 2",
    "3": "مخزن 3",
    "4": "مخزن 4",
    "5": "مخزن زايد",
    "6": "مخزن اكتوبر",
    "7": "مخزن العبور",
}
WAREHOUSE_NUMBER_DISPLAY_VALUES = [WAREHOUSE_NUMBER_LABELS[v] for v in WAREHOUSE_NUMBER_VALUES]


def normalize_branch_customer_name(value: Any) -> Optional[str]:
    raw = str(value or "").strip()
    if not raw:
        return None
    if raw.startswith("فرع:"):
        raw = raw.split(":", 1)[1].strip()
    if raw in DEFAULT_BRANCH_POS_NAMES:
        return raw
    return BRANCH_DEVICE_BY_UI_NAME.get(raw)


def canonical_branch_device_name(value: Any, known_devices: Optional[Sequence[str]] = None) -> Optional[str]:
    """Resolve any branch-facing label to canonical POS device name.

    Accepts internal names (POS-*), Arabic UI names, and mixed labels like:
    "فرع: فرع زايد (POS-ZAY)".
    """
    raw = str(value or "").strip()
    if not raw:
        return None
    direct = normalize_branch_customer_name(raw)
    if direct:
        return direct
    if raw.startswith("فرع:"):
        raw = raw.split(":", 1)[1].strip()
    if "(" in raw and ")" in raw:
        inside = raw.rsplit("(", 1)[1].split(")", 1)[0].strip()
        if inside in DEFAULT_BRANCH_POS_NAMES:
            return inside
        if known_devices:
            for d in known_devices:
                if str(d or "").strip().upper() == inside.upper():
                    return str(d or "").strip()
    if known_devices:
        for d in known_devices:
            clean = str(d or "").strip()
            if not clean:
                continue
            if clean.upper() == raw.upper():
                return clean
    return None


def branch_display_name(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    return BRANCH_UI_NAME_BY_DEVICE.get(raw, raw)


def branch_customer_label(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    if raw.startswith("فرع:"):
        raw = raw.split(":", 1)[1].strip()
    return f"فرع: {branch_display_name(raw)}"


def warehouse_display_label(value: Any) -> str:
    raw = str(value or "").strip()
    return WAREHOUSE_NUMBER_LABELS.get(raw, raw)


def warehouse_numeric_value(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    for num, label in WAREHOUSE_NUMBER_LABELS.items():
        if raw == num or raw == label:
            return num
    return raw

ALLOWED_NUMERIC_RANGES = {
    (0, 16): [str(i) for i in range(0, 18, 2)],
    (6, 22): [str(i) for i in range(6, 24, 2)],
    (14, 28): [str(i) for i in range(14, 30, 2)],
    (18, 30): [str(i) for i in range(18, 32, 2)],
    (0, 9):  [str(i) for i in range(0, 10, 1)],
    (32, 62): [str(i) for i in range(32, 63,2)],

}
NUMERIC_RANGE_LABELS = [
    f"{a} → {b}" for (a, b) in ALLOWED_NUMERIC_RANGES.keys()
]

ALPHA_SIZES = ["S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL"]


def merged_numeric_size_labels_from_profile(
    r1s: Optional[int],
    r1e: Optional[int],
    r2s: Optional[int],
    r2e: Optional[int],
) -> List[str]:
    """Union of predefined numeric columns for range 1 and range 2 (no duplicate sizes).

    Overlapping presets such as (6→22) and (14→28) become one sorted list 6…28.
    """
    chunks: List[List[str]] = []
    if r1s is not None and r1e is not None:
        labs = ALLOWED_NUMERIC_RANGES.get((r1s, r1e))
        if labs:
            chunks.append(list(labs))
    if r2s is not None and r2e is not None:
        labs = ALLOWED_NUMERIC_RANGES.get((r2s, r2e))
        if labs:
            chunks.append(list(labs))
    if not chunks:
        return []
    numeric_vals: List[int] = []
    extras: List[str] = []
    for chunk in chunks:
        for s in chunk:
            t = str(s).strip()
            if not t:
                continue
            if t.isdigit():
                numeric_vals.append(int(t))
            elif t not in extras:
                extras.append(t)
    uniq_sorted = sorted(set(numeric_vals))
    out = [str(i) for i in uniq_sorted]
    out.extend(extras)
    return out


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


@dataclass
class StockRow:
    id: int
    item_type: str
    school: str
    color: str
    size: str
    warehouse_no: int
    package_no: int
    unit_price: float
    count: int
    has_badge: int  # 0/1


# ------------------- Excel exporting -------------------

def export_to_excel(path: str, headers: List[str], rows: Sequence[Sequence[Any]]) -> None:
    ext = os.path.splitext(path)[1].lower()
    try:
        import openpyxl  # type: ignore
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        ws.append(headers)
        for row in rows:
            ws.append(list(row))
        for idx, _ in enumerate(headers, start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = 18
        wb.save(path)
        return
    except Exception:
        xml = []
        xml.append('<?xml version="1.0" encoding="UTF-8"?>')
        xml.append(
            '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:x="urn:schemas-microsoft-com:office:excel" '
            'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">'
        )
        xml.append('<Worksheet ss:Name="Sheet1"><Table>')
        xml.append("<Row>")
        for h in headers:
            xml.append(f'<Cell><Data ss:Type="String">{_xml_escape(str(h))}</Data></Cell>')
        xml.append("</Row>")
        for r in rows:
            xml.append("<Row>")
            for v in r:
                if isinstance(v, (int, float)):
                    xml.append(f'<Cell><Data ss:Type="Number">{v}</Data></Cell>')
                else:
                    xml.append(f'<Cell><Data ss:Type="String">{_xml_escape("" if v is None else str(v))}</Data></Cell>')
            xml.append("</Row>")
        xml.append("</Table></Worksheet></Workbook>")
        content = "\n".join(xml)
        if ext not in (".xls", ".xlsx", ".xml"):
            path += ".xls"
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)


def _xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )

def _file_url(path: str) -> str:
    return f"file:///{path.replace(os.sep, '/')}"

def _print_html_auto(path: str, copies: int = 1, parent: Optional[tk.Widget] = None) -> None:
    """
    Try silent HTML print (Windows COM) or shell 'print'. Else open auto-printing copy.
    """
    copies = max(1, int(copies))
    import time

    if sys.platform.startswith("win"):
        try:
            import win32com.client  # type: ignore
            ie = win32com.client.Dispatch("InternetExplorer.Application")
            ie.Visible = False
            url = _file_url(path)
            for _ in range(copies):
                ie.Navigate(url)
                # Wait for full readiness (ReadyState + Busy + Document.ReadyState)
                while True:
                    time.sleep(0.1)
                    try:
                        if not ie.Busy and int(ie.ReadyState) == 4:
                            doc = getattr(ie, "Document", None)
                            if doc is not None and str(getattr(doc, "readyState", "")).lower() == "complete":
                                break
                    except Exception:
                        # If IE throws while navigating, break to try printing anyway
                        break
                time.sleep(0.2)  # tiny settle time
                ie.ExecWB(6, 2)  # 6: PRINT, 2: DONTPROMPTUSER
                time.sleep(0.2)

            ie.Quit()
            return
        except Exception:
            pass
        try:
            for _ in range(copies):
                os.startfile(path, "print")  # type: ignore[attr-defined]
            return
        except Exception:
            pass

    try:
        auto_path = path[:-5] + "_autoprint.html" if path.lower().endswith(".html") else path + "_autoprint.html"
        try:
            with open(path, "r", encoding="utf-8") as f:
                html = f.read()
            edge_meta = '<meta http-equiv="X-UA-Compatible" content="IE=edge" />'
            if "</head>" in html:
                # inject edge meta (once) + autoprint script
                if "X-UA-Compatible" not in html:
                    html = html.replace("</head>", edge_meta + "</head>")
                html = html.replace(
                    "</head>",
                    "<script>window.onload=function(){try{window.print();}catch(e){} setTimeout(()=>window.close(),600);};</script></head>"
                )
            else:
                if "X-UA-Compatible" not in html:
                    html = edge_meta + html
                html = html + "<script>window.onload=function(){try{window.print();}catch(e){} setTimeout(()=>window.close(),600);};</script>"

            with open(auto_path, "w", encoding="utf-8") as f:
                f.write(html)
        except Exception:
            auto_path = path

        for _ in range(copies):
            webbrowser.open_new_tab(_file_url(auto_path))
            time.sleep(0.05)

        if parent is not None:
            show_toast(parent, f"تم فتح الفاتورة للطباعة — عدد النسخ: {copies}")
    except Exception as ex:
        if parent is not None:
            messagebox.showerror("فشل الطباعة", f"{ex}", parent=parent)

def save_bill_as_html(path: str, bill: Dict[str, Any], items: List[Dict[str, Any]]) -> None:
    """Generate a single HTML with 2 pages: customer copy + warehouse copy."""

    def _fmtf(x: Any) -> str:
        try:
            return f"{float(x):.2f}"
        except Exception:
            return "0.00"

    def _wh(v):
        if v in (None, "", 0, "0"):
            return ""
        return str(v)

    # Page 1: Customer copy (no warehouse/package columns)
    rows1 = "\n".join(
        f"""<tr>
            <td>{_html(ln['item_type'])}</td>
            <td>{_html(ln['school'])}</td>
            <td>{_html(ln['color'])}</td>
            <td>{_html(ln['size'])}</td>
            <td style="text-align:right">{_fmtf(ln['unit_price'])}</td>
            <td style="text-align:center">{ln['qty']}</td>
            <td style="text-align:right">{_fmtf(ln['line_total'])}</td>
        </tr>"""
        for ln in items
    )

    # Page 2: Warehouse copy (with warehouse_no + package_no)
    rows2 = "\n".join(
        f"""<tr>
            <td>{_html(ln['item_type'])}</td>
            <td>{_html(ln['school'])}</td>
            <td>{_html(ln['color'])}</td>
            <td>{_html(ln['size'])}</td>
            <td style="text-align:center">{_wh(ln.get('warehouse_no'))}</td>
            <td style="text-align:center">{_wh(ln.get('package_no'))}</td>
            <td style="text-align:right">{_fmtf(ln['unit_price'])}</td>
            <td style="text-align:center">{ln['qty']}</td>
            <td style="text-align:right">{_fmtf(ln['line_total'])}</td>
        </tr>"""
        for ln in items
    )

    customer_name = _html(bill.get('customer') or '')
    bill_date = bill['created_at']
    bill_total = f"{float(bill['total']):.2f}"

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>فاتورة #{bill['id']}</title>
<style>
  @page {{ size: auto; margin: 14mm; }}
  body {{ font-family: "Segoe UI", Tahoma, Arial, "Noto Sans Arabic", sans-serif; margin: 0; }}
  .page {{ padding: 24px; }}
  .page-break {{ page-break-before: always; }}
  .copy-label {{ background: #2563eb; color: white; display: inline-block; padding: 4px 16px;
                 border-radius: 4px; font-size: 14px; margin-bottom: 12px; }}
  table {{ border-collapse: collapse; width: 100%; }}
  thead th {{ background: #eee; }}
  th, td {{ border: 1px solid #777; padding: 6px 8px; }}
  thead {{ display: table-header-group; }}
</style>
</head>
<body>

<!-- ===== PAGE 1: Customer Copy ===== -->
<div class="page">
  <span class="copy-label">نسخة العميل</span>
  <h1>فاتورة رقم: {bill['id']}</h1>
  <div>العميل: {customer_name}</div>
  <div>التاريخ: {bill_date}</div>
  <div>الإجمالي: {bill_total}</div>
  <table>
    <thead><tr>
      <th>النوع</th><th>المدرسة</th><th>اللون</th><th>المقاس</th>
      <th>السعر</th><th>الكمية</th><th>الإجمالي</th>
    </tr></thead>
    <tbody>{rows1}</tbody>
  </table>
</div>

<!-- ===== PAGE 2: Warehouse Copy ===== -->
<div class="page page-break">
  <span class="copy-label" style="background:#059669">نسخة المخزن</span>
  <h1>فاتورة رقم: {bill['id']}</h1>
  <div>العميل: {customer_name}</div>
  <div>التاريخ: {bill_date}</div>
  <div>الإجمالي: {bill_total}</div>
  <table>
    <thead><tr>
      <th>النوع</th><th>المدرسة</th><th>اللون</th><th>المقاس</th>
      <th>رقم المخزن</th><th>رقم الطرد</th>
      <th>السعر</th><th>الكمية</th><th>الإجمالي</th>
    </tr></thead>
    <tbody>{rows2}</tbody>
  </table>
</div>

</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)

def _html(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# Normalize size text for matching (handles Arabic digits & case)
_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
def _normalize_size_label(s: str) -> str:
    return (s or "").strip().translate(_AR_DIGITS).upper()

def _normalize_spec_label(s: Any) -> str:
    """Normalize free-text spec labels for stable grouping/display.

    Handles hidden Unicode differences that make labels look identical
    in UI but different in DB keys (ZW chars, tatweel, diacritics, etc.).
    """
    txt = unicodedata.normalize("NFKC", str(s or ""))
    # Drop invisible direction/zero-width artifacts that often come from copy/paste.
    txt = txt.replace("\u200c", "").replace("\u200d", "").replace("\u200e", "").replace("\u200f", "").replace("\ufeff", "")
    # Drop Arabic tatweel and combining marks (tashkeel/harakat).
    txt = txt.replace("\u0640", "")
    txt = "".join(ch for ch in txt if unicodedata.category(ch) != "Mn")
    return " ".join(txt.strip().split())



# ------------------- SQLite DB -------------------

def apply_focus_highlight(widget, color="#19E72A"):
    default_bg = widget.cget("background")

    def on_focus_in(e):
        try:
            widget.configure(background=color)
        except:
            pass

    def on_focus_out(e):
        try:
            widget.configure(background=default_bg)
        except:
            pass

    widget.bind("<FocusIn>", on_focus_in)
    widget.bind("<FocusOut>", on_focus_out)

class SqliteDatabase:
    """SQLite persistence with package-centric controls."""

    def __init__(self, path: str = DB_PATH, legacy_json: str = LEGACY_JSON_PATH) -> None:
        self.path = path
        self.conn = sqlite3.connect(self.path, isolation_level=None, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row

        self._apply_pragmas()
        self._init_schema()
        self._migrate_from_json_if_empty(legacy_json)


    def _apply_pragmas(self) -> None:
        cur = self.conn.cursor()
        cur.execute("PRAGMA journal_mode=WAL;")
        cur.execute("PRAGMA synchronous=NORMAL;")
        cur.execute("PRAGMA foreign_keys=ON;")
        cur.close()

    def _record_sync_event(
        self,
        event_type: str,
        payload: Dict[str, Any],
        *,
        target_scope: Optional[str] = None,
    ) -> Optional[str]:
        """Append a business event to sync_outbox. Fail-safe.

        Phase 1: events are written but never pushed. If the sync layer
        is not present or the outbox insert fails, the caller continues
        without error — we must never break a business operation.

        Returns the new event_uuid when known, otherwise None.
        """
        try:
            if target_scope:
                import json as _json
                import sqlite3 as _sqlite3

                try:
                    from sync_core import new_uuid as _new_uuid
                except Exception:
                    _new_uuid = None

                def _fallback_uuid() -> str:
                    import uuid as _u
                    return str(_u.uuid4())

                event_uuid = (_new_uuid() if _new_uuid else _fallback_uuid())
                ts = now_iso()
                try:
                    self.conn.execute(
                        """
                        INSERT INTO sync_outbox
                            (event_uuid, event_type, payload_json,
                             created_at, status, attempts, target_scope)
                        VALUES (?, ?, ?, ?, 'pending', 0, ?)
                        """,
                        (
                            event_uuid,
                            event_type,
                            _json.dumps(payload, ensure_ascii=False, default=str),
                            ts,
                            str(target_scope).strip(),
                        ),
                    )
                except _sqlite3.OperationalError:
                    # Legacy outbox schema (no target_scope column) — fall back
                    # to storing the scope inside the payload. The client-side
                    # push loop handles both shapes.
                    scoped_payload = dict(payload)
                    scoped_payload["__target_scope__"] = str(target_scope).strip()
                    self.conn.execute(
                        """
                        INSERT INTO sync_outbox
                            (event_uuid, event_type, payload_json,
                             created_at, status, attempts)
                        VALUES (?, ?, ?, ?, 'pending', 0)
                        """,
                        (
                            event_uuid,
                            event_type,
                            _json.dumps(scoped_payload, ensure_ascii=False, default=str),
                            ts,
                        ),
                    )
                return str(event_uuid)

            from sync_core import record_event
            return record_event(self.conn, event_type, payload)
        except Exception:
            import traceback
            traceback.print_exc()
            return None

    def _init_schema(self) -> None:
        cur = self.conn.cursor()
        cur.executescript(
            """
            -- Default price per item (used when income price is empty)
            CREATE TABLE IF NOT EXISTS item_defaults (
                item_type TEXT PRIMARY KEY,
                default_price REAL NOT NULL
            );
            CREATE TABLE IF NOT EXISTS size_profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,

                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,

                -- numeric ranges (nullable)
                num_start_1 INTEGER,
                num_end_1   INTEGER,

                num_start_2 INTEGER,
                num_end_2   INTEGER,

                -- flags
                has_alpha INTEGER NOT NULL DEFAULT 0,

                updated_at TEXT NOT NULL,

                UNIQUE(item_type, school, color)
            );

            CREATE TABLE IF NOT EXISTS stocks(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                warehouse_no INTEGER NOT NULL,
                package_no INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                count INTEGER NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_stocks_specs
            ON stocks(item_type, school, color, size);
            CREATE INDEX IF NOT EXISTS idx_stocks_loc
            ON stocks(warehouse_no, package_no);
            CREATE INDEX IF NOT EXISTS idx_stocks_count
            ON stocks(count);

            CREATE TABLE IF NOT EXISTS movements(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT NOT NULL,
                direction TEXT NOT NULL, -- IN | OUT | ADJUST_OUT | OUT_FACTORY | PRICE_UPDATE
                stock_id INTEGER,
                qty INTEGER NOT NULL,
                note TEXT,
                bill_id INTEGER,
                item_type TEXT,
                school TEXT,
                color TEXT,
                size TEXT,
                warehouse_no INTEGER,
                package_no INTEGER,
                unit_price REAL
            );
            CREATE INDEX IF NOT EXISTS idx_movements_ts ON movements(ts DESC);
            CREATE INDEX IF NOT EXISTS idx_movements_dir ON movements(direction);
            CREATE INDEX IF NOT EXISTS idx_movements_specs ON movements(item_type,school,color,size);

            -- Audit trail for warehouse-issued PRICE_UPDATE fan-out decisions
            CREATE TABLE IF NOT EXISTS price_sync_audit(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                new_price REAL NOT NULL,
                filters_json TEXT NOT NULL,
                sync_mode TEXT NOT NULL, -- none | all-pos | selected-pos | bill-auto
                targets_json TEXT,        -- JSON list of device names and/or scopes
                event_uuids_json TEXT,  -- JSON list of emitted sync_outbox UUIDs
                note TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_price_sync_audit_created
            ON price_sync_audit(created_at DESC);

            CREATE TABLE IF NOT EXISTS bills(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                customer TEXT,
                total REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_bills_created ON bills(created_at DESC);

            CREATE TABLE IF NOT EXISTS bill_items(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bill_id INTEGER NOT NULL,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                warehouse_no INTEGER NOT NULL,
                package_no INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                qty INTEGER NOT NULL,
                line_total REAL NOT NULL,
                origin TEXT NOT NULL DEFAULT 'STOCK'
            );
            CREATE INDEX IF NOT EXISTS idx_bill_items_bill ON bill_items(bill_id);

            CREATE TABLE IF NOT EXISTS spec_history(
                field TEXT NOT NULL,
                value TEXT NOT NULL,
                PRIMARY KEY(field, value)
            );

            CREATE TABLE IF NOT EXISTS packages(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                warehouse_no INTEGER NOT NULL,
                package_no  INTEGER NOT NULL,
                status TEXT NOT NULL DEFAULT 'OPEN', -- OPEN | CLOSED
                created_at TEXT NOT NULL,
                closed_at  TEXT,
                note TEXT,
                UNIQUE(warehouse_no, package_no)
            );
            CREATE INDEX IF NOT EXISTS idx_packages_warehouse ON packages(warehouse_no, package_no);

            CREATE TABLE IF NOT EXISTS app_settings(
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )

        # ensure 'origin' exists
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bill_items)")}
            if "origin" not in cols:
                self.conn.execute(
                    "ALTER TABLE bill_items ADD COLUMN origin TEXT NOT NULL DEFAULT 'STOCK'"
                )
        except Exception:
            pass

        # ensure has_badge exists
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(stocks)")}
            if "has_badge" not in cols:
                self.conn.execute(
                    "ALTER TABLE stocks ADD COLUMN has_badge INTEGER NOT NULL DEFAULT 0"
                )
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bill_items)")}
            if "has_badge" not in cols:
                self.conn.execute(
                    "ALTER TABLE bill_items ADD COLUMN has_badge INTEGER NOT NULL DEFAULT 0"
                )
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if "has_badge" not in cols:
                self.conn.execute(
                    "ALTER TABLE movements ADD COLUMN has_badge INTEGER"
                )
        except Exception:
            pass

        # Ensure 'status' column on bills
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bills)")}
            if "status" not in cols:
                self.conn.execute(
                    "ALTER TABLE bills ADD COLUMN status TEXT NOT NULL DEFAULT 'CONFIRMED'"
                )
        except Exception:
            pass

        # Price sync audit (additive for older DB files)
        try:
            self.conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS price_sync_audit(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    new_price REAL NOT NULL,
                    filters_json TEXT NOT NULL,
                    sync_mode TEXT NOT NULL,
                    targets_json TEXT,
                    event_uuids_json TEXT,
                    note TEXT
                );
                CREATE INDEX IF NOT EXISTS idx_price_sync_audit_created
                ON price_sync_audit(created_at DESC);
                """
            )
        except Exception:
            pass

        # Returns tables
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS returns (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bill_id INTEGER NOT NULL,
                created_at TEXT NOT NULL,
                note TEXT,
                FOREIGN KEY (bill_id) REFERENCES bills(id)
            );
            CREATE TABLE IF NOT EXISTS return_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                return_id INTEGER NOT NULL,
                bill_item_id INTEGER,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                warehouse_no INTEGER NOT NULL,
                package_no INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                qty INTEGER NOT NULL,
                has_badge INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (return_id) REFERENCES returns(id)
            );
            CREATE INDEX IF NOT EXISTS idx_returns_bill ON returns(bill_id);
            CREATE INDEX IF NOT EXISTS idx_return_items_return ON return_items(return_id);
        """)

        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS branch_inventory_queue (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sync_event_uuid TEXT NOT NULL,
                queue_kind TEXT NOT NULL, -- RETURN | TRANSFER
                source_device TEXT NOT NULL,
                requested_target_device TEXT,
                external_ref TEXT,
                line_index INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                unit_price REAL NOT NULL DEFAULT 0,
                qty INTEGER NOT NULL,
                has_badge INTEGER NOT NULL DEFAULT 0,
                note TEXT,
                status TEXT NOT NULL DEFAULT 'PENDING', -- PENDING | ASSIGNED | DISCARDED | REROUTED
                processed_at TEXT,
                processed_note TEXT,
                processed_warehouse_no INTEGER,
                processed_package_no INTEGER,
                rerouted_target_device TEXT,
                UNIQUE(sync_event_uuid, line_index)
            );
            CREATE INDEX IF NOT EXISTS idx_branch_inventory_queue_status
                ON branch_inventory_queue(status, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_branch_inventory_queue_source
                ON branch_inventory_queue(source_device, created_at DESC);
        """)

        # ------------------- Sync layer (Phase 1) -------------------
        # Additive: creates sync_outbox/inbox/state/device_identity tables
        # and backfills a `uuid` column on every syncable domain table.
        # Failures here must never break the main app.
        try:
            from sync_core import apply_sync_migration, ensure_device_identity
            apply_sync_migration(self.conn)
            ensure_device_identity(
                self.conn,
                default_name="WAREHOUSE-MAIN",
                default_role="warehouse",
            )
            # Performance indexes for common filtered reads (safe/additive).
            self.conn.executescript("""
                CREATE INDEX IF NOT EXISTS idx_stocks_specs_norm
                    ON stocks(LOWER(TRIM(item_type)), LOWER(TRIM(school)), LOWER(TRIM(color)), LOWER(TRIM(size)));
                CREATE INDEX IF NOT EXISTS idx_bills_customer_norm
                    ON bills(LOWER(TRIM(customer)));
                CREATE INDEX IF NOT EXISTS idx_movements_bill_id
                    ON movements(bill_id);
                CREATE INDEX IF NOT EXISTS idx_bill_items_specs
                    ON bill_items(item_type, school, color, size);
                CREATE INDEX IF NOT EXISTS idx_sync_outbox_status_seq
                    ON sync_outbox(status, local_seq);
                CREATE INDEX IF NOT EXISTS idx_sync_inbox_status_seq
                    ON sync_inbox(apply_status, server_seq);
            """)
        except Exception:
            import traceback
            traceback.print_exc()

        cur.close()
        # ------------------- Size Profiles -------------------

    def get_size_profile(self, item_type: str, school: str, color: str):
        cur = self.conn.execute(
            """
            SELECT
                num_start_1,
                num_end_1,
                num_start_2,
                num_end_2,
                has_alpha
            FROM size_profiles
            WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
              AND LOWER(TRIM(school)) = LOWER(TRIM(?))
              AND LOWER(TRIM(color)) = LOWER(TRIM(?))
            """,
            (item_type, school, color),
        )
        row = cur.fetchone()
        if row is None:
            return None
        return (
            row["num_start_1"],
            row["num_end_1"],
            row["num_start_2"],
            row["num_end_2"],
            row["has_alpha"],
        )
    def _get_size_ranges(self, item_type, school, color):
        profile = self.get_size_profile(item_type, school, color)
        if not profile:
            return []

        r1s, r1e, r2s, r2e, has_alpha = profile
        ranges = []

        if r1s is not None and r1e is not None:
            ranges.append({
                "range_type": "NUMERIC",
                "start": r1s,
                "end": r1e,
            })

        if r2s is not None and r2e is not None:
            ranges.append({
                "range_type": "NUMERIC",
                "start": r2s,
                "end": r2e,
            })

        if has_alpha:
            ranges.append({
                "range_type": "ALPHA",
                "alpha_set": ALPHA_SIZES[:],
            })

        return ranges

    def upsert_size_profile(
        self,
        item_type: str,
        school: str,
        color: str,
        *,
        r1_start: Optional[int],
        r1_end: Optional[int],
        r2_start: Optional[int],
        r2_end: Optional[int],
        has_alpha: bool,
    ) -> None:

        """
        Insert or update a size profile for (item_type, school, color).
        """
        self.conn.execute(
            """
            INSERT INTO size_profiles (
                item_type, school, color,
                num_start_1, num_end_1,
                num_start_2, num_end_2,
                has_alpha, updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
            ON CONFLICT(item_type, school, color)
            DO UPDATE SET
                num_start_1 = excluded.num_start_1,
                num_end_1   = excluded.num_end_1,
                num_start_2 = excluded.num_start_2,
                num_end_2   = excluded.num_end_2,
                has_alpha   = excluded.has_alpha,
                updated_at  = datetime('now')
            """,
            (
                item_type, school, color,
                r1_start, r1_end,
                r2_start, r2_end,
                int(has_alpha),
            )
        )

        self.conn.commit()

    # -------- Packages helpers --------
    def auto_reopen_package_if_empty(self, warehouse_no: int, package_no: int) -> None:
        cur = self.conn.execute(
            "SELECT COUNT(*) AS c FROM stocks WHERE warehouse_no=? AND package_no=?",
            (int(warehouse_no), int(package_no)),
        )
        c = int(cur.fetchone()["c"] or 0)

        if c == 0:
            self.conn.execute(
                """
                UPDATE packages
                SET status='OPEN', closed_at=NULL
                WHERE warehouse_no=? AND package_no=? AND status='CLOSED'
                """,
                (int(warehouse_no), int(package_no)),
            )

    def get_effective_price(self, item_type, school, color, size):
        """
        Priority:
        1) exact last price (history)
        2) default price for item
        3) None if unknown
        """
        # 1) history
        try:
            p = self.last_price_for_specs(item_type, school, color, size)
            if p is not None:
                return float(p)
        except Exception:
            pass

        # 2) default price (stored once)
        try:
            cur = self.conn.execute(
                "SELECT default_price FROM item_defaults WHERE item_type=?",
                (item_type,)
            )
            row = cur.fetchone()
            if row and row["default_price"] is not None:
                return float(row["default_price"])
        except Exception:
            pass

        return None
    def ensure_default_price(self, item_type, price):
        """
        Save default price once (only if not exists).
        """
        self.conn.execute(
            """
            INSERT INTO item_defaults (item_type, default_price)
            VALUES (?, ?)
            ON CONFLICT(item_type) DO NOTHING
            """,
            (item_type, float(price))
        )
        self.conn.commit()

    def purge_definition(self, field: str, value: str) -> int:
        """
        Permanently delete a definition from the system.
        Removes it from:
        - stocks
        - spec_history
        So it never appears again in any dropdown.
        """

        if field not in ("item_type", "school", "color", "size"):
            raise ValueError("حقل غير مدعوم")

        value = (value or "").strip()
        if not value:
            return 0

        with self.conn:
            # delete from stocks
            cur = self.conn.execute(
                f"SELECT COUNT(*) FROM stocks WHERE {field} = ?",
                (value,),
            )
            count = int(cur.fetchone()[0] or 0)

            self.conn.execute(
                f"DELETE FROM stocks WHERE {field} = ?",
                (value,),
            )

            # delete from autocomplete history
            self.conn.execute(
                "DELETE FROM spec_history WHERE field = ? AND value = ?",
                (field, value),
            )

        return count


    # AFTER (NEW) — list all schools (from stocks or history/bill_items)
    def list_schools_all(self) -> List[str]:
        cur = self.conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT TRIM(school) AS s FROM stocks
                UNION
                SELECT DISTINCT TRIM(school) AS s FROM bill_items
                UNION
                SELECT DISTINCT TRIM(value) AS s FROM spec_history WHERE field='school'
            """)
            vals = [r["s"] for r in cur.fetchall() if r["s"]]
            vals.sort(key=lambda v: v.lower())
            return vals
        finally:
            cur.close()

    # AFTER (NEW) — items for a school: distinct (item_type, color) seen anywhere
    def list_items_for_school(self, school: str) -> List[Tuple[str, str]]:
        sc = (school or "").strip()
        if not sc:
            return []
        cur = self.conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT TRIM(item_type) AS it, TRIM(color) AS cl
                FROM stocks WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
                UNION
                SELECT DISTINCT TRIM(item_type) AS it, TRIM(color) AS cl
                FROM bill_items WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
            """, (sc, sc))
            rows = [(r["it"], r["cl"]) for r in cur.fetchall() if r["it"] and r["cl"]]
            rows.sort(key=lambda t: (t[0].lower(), t[1].lower()))
            return rows
        finally:
            cur.close()

        # AFTER (NEW) — sizes for a (school, item_type, color) with current count + last price
    def _size_row(self, school: str, item_type: str, color: str, size: str) -> Dict[str, Any]:
        """Return a dict with size, current stock count, and last known price."""
        cur = self.conn.execute(
            """SELECT COALESCE(SUM(count), 0) AS total_count
               FROM stocks
               WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                 AND LOWER(TRIM(school))    = LOWER(TRIM(?))
                 AND LOWER(TRIM(color))     = LOWER(TRIM(?))
                 AND LOWER(TRIM(size))      = LOWER(TRIM(?))""",
            (item_type.strip(), school.strip(), color.strip(), size.strip()),
        )
        row = cur.fetchone()
        total_count = int(row["total_count"]) if row else 0
        last_price = self.last_price_for_specs(item_type, school, color, size)
        return {
            "size": size,
            "count": total_count,
            "last_price": last_price,
        }

    def list_sizes_for_item(self, school, item_type, color):
        ranges = self._get_size_ranges(item_type, school, color)
        out = []

        for r in ranges:
            if r["range_type"] == "NUMERIC":
                for sz in range(r["start"], r["end"] + 1):
                    out.append(self._size_row(school, item_type, color, str(sz)))

            elif r["range_type"] == "ALPHA":
                for sz in r["alpha_set"]:
                    out.append(self._size_row(school, item_type, color, sz))

        return out


    def last_price_for_specs(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
    ) -> Optional[float]:
        """
        Return the most recent unit_price for the exact specs.
        Priority:
        1) bill_items (actual sold price)
        2) stocks (last income price)
        """

        it = (item_type or "").strip()
        sc = (school or "").strip()
        cl = (color or "").strip()
        sz_raw = (size or "").strip()

        if not (it and sc and cl and sz_raw):
            return None

        sz_norm = _normalize_size_label(sz_raw)
        cur = self.conn.cursor()

        try:
            # ---- 1) LAST SOLD PRICE (bill_items) ----
            if sz_norm and sz_norm != sz_raw:
                cur.execute(
                    """
                    SELECT unit_price
                    FROM bill_items
                    WHERE LOWER(TRIM(item_type)) = LOWER(?)
                    AND LOWER(TRIM(school))    = LOWER(?)
                    AND LOWER(TRIM(color))     = LOWER(?)
                    AND (LOWER(TRIM(size))     = LOWER(?) OR LOWER(TRIM(size)) = LOWER(?))
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (it, sc, cl, sz_raw, sz_norm),
                )
            else:
                cur.execute(
                    """
                    SELECT unit_price
                    FROM bill_items
                    WHERE LOWER(TRIM(item_type)) = LOWER(?)
                    AND LOWER(TRIM(school))    = LOWER(?)
                    AND LOWER(TRIM(color))     = LOWER(?)
                    AND LOWER(TRIM(size))      = LOWER(?)
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (it, sc, cl, sz_raw),
                )

            r = cur.fetchone()
            if r and r[0] is not None:
                return float(r[0])

            # ---- 2) FALLBACK: LAST STOCK PRICE ----
            if sz_norm and sz_norm != sz_raw:
                cur.execute(
                    """
                    SELECT unit_price
                    FROM stocks
                    WHERE LOWER(TRIM(item_type)) = LOWER(?)
                    AND LOWER(TRIM(school))    = LOWER(?)
                    AND LOWER(TRIM(color))     = LOWER(?)
                    AND (LOWER(TRIM(size))     = LOWER(?) OR LOWER(TRIM(size)) = LOWER(?))
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (it, sc, cl, sz_raw, sz_norm),
                )
            else:
                cur.execute(
                    """
                    SELECT unit_price
                    FROM stocks
                    WHERE LOWER(TRIM(item_type)) = LOWER(?)
                    AND LOWER(TRIM(school))    = LOWER(?)
                    AND LOWER(TRIM(color))     = LOWER(?)
                    AND LOWER(TRIM(size))      = LOWER(?)
                    ORDER BY id DESC
                    LIMIT 1
                    """,
                    (it, sc, cl, sz_raw),
                )

            r = cur.fetchone()
            if r and r[0] is not None:
                return float(r[0])

            return None

        finally:
            cur.close()

    def update_specs_by_ids(
        self,
        ids: Sequence[int],
        *,
        item_type: Optional[str] = None,
        school: Optional[str] = None,
        color: Optional[str] = None,
        size: Optional[str] = None,
        note: str = "Edit specs (by ids)",
    ) -> int:
        ids = [int(x) for x in ids if str(x).strip()]
        if not ids:
            return 0

        sets = []
        args: List[Any] = []
        changes = {}

        def _norm(v: Optional[str]) -> Optional[str]:
            v = (v or "").strip()
            return v if v else None

        it = _norm(item_type)
        sc = _norm(school)
        cl = _norm(color)
        sz = _norm(size)

        if it is not None:
            sets.append("item_type = ?"); args.append(it); changes["item_type"] = it
        if sc is not None:
            sets.append("school = ?"); args.append(sc); changes["school"] = sc
        if cl is not None:
            sets.append("color = ?"); args.append(cl); changes["color"] = cl
        if sz is not None:
            sets.append("size = ?"); args.append(sz); changes["size"] = sz

        if not sets:
            return 0

        ph = ",".join("?" for _ in ids)
        with self.conn:
            cur = self.conn.execute(f"SELECT COUNT(*) AS c FROM stocks WHERE id IN ({ph})", ids)
            count = int(cur.fetchone()["c"] or 0)
            if count == 0:
                return 0

            old_specs = self.conn.execute(
                f"SELECT DISTINCT item_type, school, color, size FROM stocks WHERE id IN ({ph})",
                ids,
            ).fetchall()

            self.conn.execute(
                f"UPDATE stocks SET {', '.join(sets)} WHERE id IN ({ph})",
                (*args, *ids),
            )

            self._cascade_spec_rename(old_specs, changes)

            self._upsert_history(changes)
        self.cleanup_unused_specs()
        return count


    def _cascade_spec_rename(
        self,
        old_specs: Sequence[Any],
        changes: Dict[str, Any],
    ) -> None:
        """Propagate a completed spec rename from `stocks` to historical tables.

        For each old (item_type, school, color, size) tuple we just rewrote,
        if no `stocks` row still references that tuple, it was a full rename —
        cascade the same field changes to `movements`, `bill_items`, and
        (best-effort) the POS mirror tables so audit views stay consistent.
        """
        if not changes or not old_specs:
            return
        sets_parts: List[str] = []
        new_vals: List[Any] = []
        for fld in ("item_type", "school", "color", "size"):
            if fld in changes:
                sets_parts.append(f"{fld} = ?")
                new_vals.append(changes[fld])
        if not sets_parts:
            return
        set_sql = ", ".join(sets_parts)

        for row in old_specs:
            old_it = row["item_type"] if isinstance(row, sqlite3.Row) else row[0]
            old_sc = row["school"]    if isinstance(row, sqlite3.Row) else row[1]
            old_cl = row["color"]     if isinstance(row, sqlite3.Row) else row[2]
            old_sz = row["size"]      if isinstance(row, sqlite3.Row) else row[3]

            still = self.conn.execute(
                "SELECT 1 FROM stocks WHERE item_type=? AND school=? AND color=? AND size=? LIMIT 1",
                (old_it, old_sc, old_cl, old_sz),
            ).fetchone()
            if still:
                continue

            where_sql = "item_type=? AND school=? AND color=? AND size=?"
            where_args = (old_it, old_sc, old_cl, old_sz)

            for tbl in ("movements", "bill_items", "pos_stocks_mirror", "pos_reservations_mirror"):
                try:
                    self.conn.execute(
                        f"UPDATE {tbl} SET {set_sql} WHERE {where_sql}",
                        (*new_vals, *where_args),
                    )
                except sqlite3.OperationalError:
                    pass

    def update_specs_in_package(
        self,
        warehouse_no: int,
        package_no: int,
        *,
        item_type: Optional[str] = None,
        school: Optional[str] = None,
        color: Optional[str] = None,
        size: Optional[str] = None,
        note: str = "Edit specs (package)",
    ) -> int:
        w = int(warehouse_no); p = int(package_no)
        if w < 1 or p < 1:
            raise ValueError("رقم المخزن والعبوة يجب أن يكونا >= 1.")

        sets = []
        args: List[Any] = []
        changes = {}

        def _norm(v: Optional[str]) -> Optional[str]:
            v = (v or "").strip()
            return v if v else None

        it = _norm(item_type)
        sc = _norm(school)
        cl = _norm(color)
        sz = _norm(size)

        if it is not None:
            sets.append("item_type = ?"); args.append(it); changes["item_type"] = it
        if sc is not None:
            sets.append("school = ?"); args.append(sc); changes["school"] = sc
        if cl is not None:
            sets.append("color = ?"); args.append(cl); changes["color"] = cl
        if sz is not None:
            sets.append("size = ?"); args.append(sz); changes["size"] = sz

        if not sets:
            return 0

        with self.conn:
            cur = self.conn.execute(
                "SELECT COUNT(*) AS c FROM stocks WHERE warehouse_no=? AND package_no=?",
                (w, p),
            )
            count = int(cur.fetchone()["c"] or 0)
            if count == 0:
                return 0

            old_specs = self.conn.execute(
                "SELECT DISTINCT item_type, school, color, size FROM stocks "
                "WHERE warehouse_no=? AND package_no=?",
                (w, p),
            ).fetchall()

            self.conn.execute(
                f"UPDATE stocks SET {', '.join(sets)} WHERE warehouse_no=? AND package_no=?",
                (*args, w, p),
            )

            self._cascade_spec_rename(old_specs, changes)

            self._upsert_history(changes)
        self.cleanup_unused_specs()
        return count

    # --- NEW: context-aware distincts ---------------------------------
    def get_distinct_filtered(self, target: str, constraints: Dict[str, Any]) -> List[str]:
        """
        Return DISTINCT values for `target` constrained by the other filters.
        Includes rows with count 0 so products stay visible in pickers after stock runs out.
        constraints keys may include: item_type, school, color, size, warehouse_no, package_no.
        """
        valid = {"item_type", "school", "color", "size"}
        if target not in valid:
            return []

        where: List[str] = ["1=1"]
        args: List[Any] = []

        def _is_list(x): return isinstance(x, (list, tuple))

        for fld in ("item_type", "school", "color", "size"):
            if fld == target:
                continue
            v = constraints.get(fld)
            if v in (None, ""):
                continue
            if _is_list(v) and v:
                ph = ",".join("?" for _ in v)
                where.append(f"LOWER({fld}) IN ({ph})")
                args += [str(s).strip().lower() for s in v]
            elif isinstance(v, str) and v.strip():
                where.append(f"LOWER({fld}) = LOWER(?)")
                args.append(v.strip())

        for nf in ("warehouse_no", "package_no"):
            v = constraints.get(nf)
            if v in (None, ""):
                continue
            if _is_list(v) and v:
                ph = ",".join("?" for _ in v)
                where.append(f"{nf} IN ({ph})")
                args += [int(x) for x in v]
            elif str(v).strip():
                where.append(f"{nf} = ?")
                args.append(int(v))

        cur = self.conn.cursor()
        try:
            cur.execute(
                f"SELECT DISTINCT TRIM({target}) AS v FROM stocks WHERE {' AND '.join(where)}"
                f" ORDER BY LOWER(TRIM({target})) ASC",
                args,
            )

            return [r["v"] for r in cur.fetchall() if r["v"] not in (None, "")]
        finally:
            cur.close()

    def last_color_for_school(self, school: str) -> Optional[str]:
        school = (school or "").strip()
        if not school:
            return None
        cur = self.conn.cursor()
        try:
            cur.execute(
                "SELECT color FROM stocks WHERE LOWER(school)=LOWER(?) AND color<>'' ORDER BY id DESC LIMIT 1",
                (school,)
            )
            r = cur.fetchone()
            if r and r["color"]:
                return str(r["color"])
            cur.execute(
                "SELECT color FROM bill_items WHERE LOWER(school)=LOWER(?) AND color<>'' ORDER BY id DESC LIMIT 1",
                (school,)
            )
            r = cur.fetchone()
            if r and r["color"]:
                return str(r["color"])
            return None
        finally:
            cur.close()

    def _price_update_filters_for_sync(self, filters: Dict[str, Any]) -> Dict[str, Any]:
        """Normalize filters for PRICE_UPDATE payloads (POS applier shape)."""
        out: Dict[str, Any] = {}
        for k, v in (filters or {}).items():
            if v in (None, "", []):
                continue
            out[k] = v
        return out

    def _record_price_sync_audit(
        self,
        *,
        new_price: float,
        filters: Dict[str, Any],
        sync_mode: str,
        targets: List[str],
        event_uuids: List[str],
        note: str,
    ) -> None:
        try:
            self.conn.execute(
                """
                INSERT INTO price_sync_audit
                    (created_at, new_price, filters_json, sync_mode, targets_json, event_uuids_json, note)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    now_iso(),
                    float(new_price),
                    json.dumps(filters, ensure_ascii=False, default=str),
                    str(sync_mode or ""),
                    json.dumps(targets, ensure_ascii=False, default=str),
                    json.dumps(event_uuids, ensure_ascii=False, default=str),
                    note or "",
                ),
            )
        except Exception:
            # Never break pricing operations if audit insert fails.
            import traceback
            traceback.print_exc()

    def emit_price_update_sync_events(
        self,
        *,
        filters: Dict[str, Any],
        new_price: float,
        note: str,
        sync_mode: str = "all-pos",
        sync_pos_devices: Optional[Sequence[str]] = None,
        audit_mode: str = "manual",
    ) -> List[str]:
        """Emit one or more PRICE_UPDATE outbox rows with explicit target_scope.

        sync_mode:
          - none: do not emit sync events
          - all-pos: fan-out to every POS (server-side scope `all-pos`)
          - selected-pos: one event per selected POS device name (`pos:<name>`)
          - bill-auto: internal helper mode (same as all-pos/selected depending on callers)
        """
        cleaned_filters = self._price_update_filters_for_sync(filters)

        # POS applier refuses completely unfiltered updates; mirror that guard here.
        spec_keys = ("item_type", "school", "color", "size")
        if not any(str(cleaned_filters.get(k) or "").strip() for k in spec_keys):
            return []

        mode = (sync_mode or "all-pos").strip().lower()
        if mode == "none":
            self._record_price_sync_audit(
                new_price=float(new_price),
                filters=cleaned_filters,
                sync_mode=str(audit_mode or mode),
                targets=[],
                event_uuids=[],
                note=str(note or ""),
            )
            return []

        targets: List[str] = []
        uuids: List[str] = []

        if mode == "all-pos":
            targets = ["all-pos"]
            eu = self._record_sync_event(
                "PRICE_UPDATE",
                {
                    "new_price":     float(new_price),
                    "filters":       cleaned_filters,
                    "updated_count": None,
                    "note":          note,
                },
                target_scope="all-pos",
            )
            if eu:
                uuids.append(str(eu))
        elif mode == "selected-pos":
            devs = [str(x).strip() for x in (sync_pos_devices or []) if str(x).strip()]
            # de-dupe while preserving order
            seen = set()
            ordered: List[str] = []
            for d in devs:
                if d in seen:
                    continue
                seen.add(d)
                ordered.append(d)
            for d in ordered:
                scope = f"pos:{d}"
                targets.append(scope)
                eu = self._record_sync_event(
                    "PRICE_UPDATE",
                    {
                        "new_price":     float(new_price),
                        "filters":       cleaned_filters,
                        "updated_count": None,
                        "note":          note,
                    },
                    target_scope=scope,
                )
                if eu:
                    uuids.append(str(eu))
        else:
            # Unknown mode -> safest explicit fan-out (matches server default intent)
            targets = ["all-pos"]
            eu = self._record_sync_event(
                "PRICE_UPDATE",
                {
                    "new_price":     float(new_price),
                    "filters":       cleaned_filters,
                    "updated_count": None,
                    "note":          note,
                },
                target_scope="all-pos",
            )
            if eu:
                uuids.append(str(eu))

        self._record_price_sync_audit(
            new_price=float(new_price),
            filters=cleaned_filters,
            sync_mode=str(audit_mode or mode),
            targets=targets,
            event_uuids=uuids,
            note=str(note or ""),
        )
        return uuids

    def list_price_sync_audit(self, limit: int = 200) -> List[Dict[str, Any]]:
        lim = max(1, min(int(limit or 200), 2000))
        try:
            cur = self.conn.execute(
                """
                SELECT id, created_at, new_price, filters_json, sync_mode, targets_json, event_uuids_json, note
                  FROM price_sync_audit
                 ORDER BY id DESC
                 LIMIT ?
                """,
                (lim,),
            )
            return [dict(r) for r in cur.fetchall()]
        except Exception:
            return []

    def update_prices(
        self,
        filters: Dict[str, Any],
        new_price: float,
        note: str = "Price update",
        *,
        price_sync_mode: str = "all-pos",
        price_sync_pos_devices: Optional[Sequence[str]] = None,
        emit_price_sync: bool = True,
    ) -> int:
        if new_price is None or float(new_price) < 0:
            raise ValueError("New price must be a non-negative number.")
        where: List[str] = []
        args: List[Any] = []

        if "id" in filters and filters["id"]:
            where.append("id = ?")
            args.append(int(filters["id"]))
        else:
            # Important: trim DB columns to ignore stray spaces in stored data
            for k in ("item_type", "school", "color", "size"):
                v_raw = (filters.get(k) or "").strip()
                if not v_raw:
                    continue
                if k == "size":
                    # dual match: original text AND normalized (Arabic→Latin digits, uppercased)
                    v_norm = _normalize_size_label(v_raw)
                    if v_norm and v_norm != v_raw:
                        where.append("(LOWER(TRIM(size)) = LOWER(?) OR LOWER(TRIM(size)) = LOWER(?))")
                        args.extend([v_raw, v_norm])
                    else:
                        where.append("LOWER(TRIM(size)) = LOWER(?)")
                        args.append(v_raw)
                else:
                    where.append(f"LOWER(TRIM({k})) = LOWER(?)")
                    args.append(v_raw)

            for k in ("warehouse_no", "package_no"):
                v = filters.get(k)
                if v in (None, ""):
                    continue

                # ---- HANDLE GROUPED PACKAGE NUMBERS ----
                if k == "package_no" and isinstance(v, str) and "," in v:
                    nums = [int(x.strip()) for x in v.split(",") if x.strip().isdigit()]
                    if not nums:
                        continue
                    placeholders = ",".join("?" for _ in nums)
                    where.append(f"{k} IN ({placeholders})")
                    args.extend(nums)
                else:
                    where.append(f"{k} = ?")
                    args.append(int(v))


        if not where:
            raise ValueError("No filter provided for price update.")

        cur = self.conn.cursor()
        cur.execute(f"SELECT * FROM stocks WHERE {' AND '.join(where)}", args)
        rows = cur.fetchall()
        if not rows:
            cur.close()
            return 0

        with self.conn:
            self.conn.execute(
                f"UPDATE stocks SET unit_price = ? WHERE {' AND '.join(where)}",
                (float(new_price), *args),
            )
            ts = now_iso()
            for r in rows:
                self.conn.execute(
                    """INSERT INTO movements
                    (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        ts, "PRICE_UPDATE", int(r["id"]), 0, note, None,
                        r["item_type"], r["school"], r["color"], r["size"],
                        int(r["warehouse_no"]), int(r["package_no"]),
                        float(new_price), int(r["has_badge"] or 0),
                    ),
                )
            if emit_price_sync:
                cleaned = {k: v for k, v in (filters or {}).items() if v not in (None, "", [])}
                spec_keys = ("item_type", "school", "color", "size")
                has_specs = any(str(cleaned.get(k) or "").strip() for k in spec_keys)

                emit_filters: List[Dict[str, Any]]
                if has_specs:
                    emit_filters = [cleaned]
                else:
                    groups: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}
                    for r in rows:
                        it = str(r["item_type"] or "").strip()
                        sc = str(r["school"] or "").strip()
                        cl = str(r["color"] or "").strip()
                        sz = str(r["size"] or "").strip()
                        if not (it and sc and cl and sz):
                            continue
                        key = (it, sc, cl, sz)
                        if key not in groups:
                            flt = {
                                "item_type": it,
                                "school": sc,
                                "color": cl,
                                "size": sz,
                            }
                            # carry over non-spec constraints (except row id)
                            for ck, cv in cleaned.items():
                                if ck in ("id",) or ck in spec_keys:
                                    continue
                                if cv in (None, "", []):
                                    continue
                                flt[ck] = cv
                            groups[key] = flt
                    emit_filters = list(groups.values()) if groups else []

                for flt in emit_filters:
                    self.emit_price_update_sync_events(
                        filters=flt,
                        new_price=float(new_price),
                        note=note,
                        sync_mode=str(price_sync_mode or "all-pos"),
                        sync_pos_devices=price_sync_pos_devices,
                        audit_mode="manual",
                    )
        cur.close()
        return len(rows)

    def _get_package_row(self, warehouse_no: int, package_no: int) -> Optional[sqlite3.Row]:
        cur = self.conn.execute(
            "SELECT * FROM packages WHERE warehouse_no=? AND package_no=?", (int(warehouse_no), int(package_no))
        )
        row = cur.fetchone()
        return row

    def ensure_package_open(self, warehouse_no: int, package_no: int, note: Optional[str] = None) -> None:
        if int(package_no) < 1:
            raise ValueError("رقم العبوة يجب أن يكون 1 أو أكبر.")
        row = self._get_package_row(warehouse_no, package_no)
        if row is None:
            # Create as OPEN if never existed
            self.conn.execute(
                "INSERT INTO packages(warehouse_no,package_no,status,created_at,note) VALUES(?,?,?,?,?)",
                (int(warehouse_no), int(package_no), "OPEN", now_iso(), note),
            )
            return

        status = (row["status"] or "").upper()
        if status == "OPEN":
            return

        # If CLOSED: auto-reopen when the package is empty (no stock rows left)
        cur = self.conn.execute(
            "SELECT COUNT(*) AS c FROM stocks WHERE warehouse_no=? AND package_no=?",
            (int(warehouse_no), int(package_no)),
        )
        c = int(cur.fetchone()["c"] or 0)
        if c == 0:
            self.conn.execute(
                "UPDATE packages SET status='OPEN', closed_at=NULL, note=COALESCE(note,'') || ? WHERE warehouse_no=? AND package_no=?",
                (f" | reopened {now_iso()}", int(warehouse_no), int(package_no)),
            )
            return

        # Still has stock: keep it closed
        raise ValueError("لا يمكن الإضافة: هذه العبوة مغلقة وبداخلها مخزون.")


    def close_package(self, warehouse_no: int, package_no: int, force: bool = False) -> None:
        row = self._get_package_row(warehouse_no, package_no)
        if row is None:
            if force:
                self.conn.execute(
                    "INSERT INTO packages(warehouse_no,package_no,status,created_at,closed_at) VALUES(?,?,?,?,?)",
                    (int(warehouse_no), int(package_no), "CLOSED", now_iso(), now_iso()),
                )
                return
            else:
                raise ValueError("لا توجد حاوية بهذا الرقم لإغلاقها.")
        if (row["status"] or "").upper() == "CLOSED":
            return
        self.conn.execute(
            "UPDATE packages SET status='CLOSED', closed_at=? WHERE warehouse_no=? AND package_no=?",
            (now_iso(), int(warehouse_no), int(package_no)),
        )

    def package_status(self, warehouse_no: int, package_no: int) -> Optional[str]:
        row = self._get_package_row(warehouse_no, package_no)
        return None if row is None else (row["status"] or "").upper()

    def package_numbers_summary(self, warehouse_no: int) -> Dict[str, Any]:
        w = int(warehouse_no)
        cur = self.conn.cursor()
        cur.execute("SELECT package_no, status FROM packages WHERE warehouse_no=?", (w,))
        pkg_rows = cur.fetchall()
        cur.execute("SELECT MAX(package_no) AS m FROM stocks WHERE warehouse_no=?", (w,))
        max_from_stocks = (cur.fetchone()["m"] or 0)
        cur.execute("SELECT MAX(package_no) AS m FROM packages WHERE warehouse_no=?", (w,))
        max_from_packages = (cur.fetchone()["m"] or 0)
        max_no = max(int(max_from_stocks or 0), int(max_from_packages or 0))
        used = {int(r["package_no"]) for r in pkg_rows}
        cur.execute("SELECT DISTINCT package_no FROM stocks WHERE warehouse_no=?", (w,))
        used |= {int(r["package_no"]) for r in cur.fetchall()}
        gaps = [n for n in range(1, max_no + 1) if n not in used]
        open_pkgs = sorted([int(r["package_no"]) for r in pkg_rows if (r["status"] or "").upper() == "OPEN"])
        next_num = max_no + 1 if max_no >= 1 else 1
        cur.close()
        return {"free": gaps, "open": open_pkgs, "next": next_num, "max": max_no}

    # -------- Migration (best-effort) --------
    def reopen_package(self, warehouse_no: int, package_no: int, *, require_empty: bool = False, note: str = "Manual reopen") -> None:
        """
        Manually reopen a package (admin action).
        - If require_empty=True, refuse unless the package has no stock rows.
        - Creates the package row if it never existed (and marks it OPEN).
        """
        w = int(warehouse_no); p = int(package_no)
        if p < 1:
            raise ValueError("رقم العبوة يجب أن يكون 1 أو أكبر.")

        row = self._get_package_row(w, p)
        if row is None:
            # create it as OPEN
            self.conn.execute(
                "INSERT INTO packages(warehouse_no,package_no,status,created_at,closed_at,note) VALUES(?,?,?,?,?,?)",
                (w, p, "OPEN", now_iso(), None, f"{note} {now_iso()}"),
            )
            return

        if require_empty:
            cur = self.conn.execute("SELECT COUNT(*) AS c FROM stocks WHERE warehouse_no=? AND package_no=?", (w, p))
            c = int(cur.fetchone()["c"] or 0)
            if c > 0:
                raise ValueError("لا يمكن إعادة الفتح مع وجود مخزون داخل العبوة (خيار 'فارغة فقط' مفعل).")

        self.conn.execute(
            "UPDATE packages SET status='OPEN', closed_at=NULL, note=COALESCE(note,'') || ? WHERE warehouse_no=? AND package_no=?",
            (f" | {note} {now_iso()}", w, p),
        )

    def _migrate_from_json_if_empty(self, legacy_path: str) -> None:
        cur = self.conn.cursor()
        cur.execute("SELECT COUNT(*) FROM stocks")
        has_data = cur.fetchone()[0] > 0
        cur.execute("SELECT COUNT(*) FROM bills")
        has_data = has_data or cur.fetchone()[0] > 0
        cur.close()
        if has_data or not os.path.exists(legacy_path):
            return
        try:
            with open(legacy_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            with self.conn:
                for s in data.get("stocks", []):
                    self.conn.execute(
                        """INSERT INTO stocks(id,item_type,school,color,size,warehouse_no,package_no,unit_price,count)
                           VALUES(?,?,?,?,?,?,?,?,?)""",
                        (
                            int(s["id"]), s["item_type"], s["school"], s["color"], s["size"],
                            int(s["warehouse_no"]), int(s["package_no"]),
                            float(s["unit_price"]), int(s["count"]),
                        ),
                    )
                for m in data.get("movements", []):
                    self.conn.execute(
                        """INSERT INTO movements
                           (id,ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            int(m["id"]), m["ts"], m["direction"], m.get("stock_id"), int(m["qty"]),
                            m.get("note"), m.get("bill_id"), m.get("item_type"), m.get("school"),
                            m.get("color"), m.get("size"), m.get("warehouse_no"), m.get("package_no"),
                            float(m.get("unit_price")) if m.get("unit_price") is not None else None,
                        ),
                    )
                for b in data.get("bills", []):
                    self.conn.execute(
                        "INSERT INTO bills(id,created_at,customer,total) VALUES(?,?,?,?)",
                        (int(b["id"]), b["created_at"], b.get("customer"), float(b["total"])),
                    )
                for bi in data.get("bill_items", []):
                    self.conn.execute(
                        """INSERT INTO bill_items
                           (id,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,qty,line_total)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            int(bi["id"]), int(bi["bill_id"]), bi["item_type"], bi["school"], bi["color"], bi["size"],
                            int(bi["warehouse_no"]), int(bi["package_no"]), float(bi["unit_price"]),
                            int(bi["qty"]), float(bi["line_total"]),
                        ),
                    )
                sh = data.get("spec_history", {})
                for field in ("item_type", "school", "color", "size"):
                    for val in sh.get(field, []):
                        self.conn.execute(
                            "INSERT OR IGNORE INTO spec_history(field,value) VALUES(?,?)", (field, str(val).strip())
                        )
            try:
                os.rename(legacy_path, legacy_path + ".migrated_backup")
            except Exception:
                pass
        except Exception:
            pass

    # -------- Helpers --------
    def list_customers(self) -> List[str]:
        cur = self.conn.cursor()
        cur.execute("""
            SELECT DISTINCT TRIM(customer) AS c
            FROM bills
            WHERE customer IS NOT NULL AND TRIM(customer) <> ''
            ORDER BY LOWER(c)
        """)
        vals = [r["c"] for r in cur.fetchall() if r["c"]]
        cur.close()
        return vals

    def _upsert_history(self, specs: Dict[str, Any]) -> None:
        with self.conn:
            for field in ("item_type", "school", "color", "size"):
                val = str(specs.get(field, "")).strip()
                if val:
                    self.conn.execute(
                        "INSERT OR IGNORE INTO spec_history(field,value) VALUES(?,?)", (field, val)
                    )

    def get_distinct(self, field: str) -> List[str]:
        if field not in ("item_type", "school", "color", "size"):
            return []

        cur = self.conn.cursor()
        try:
            cur.execute(f"""
                SELECT DISTINCT TRIM({field}) AS v
                FROM stocks
                WHERE count > 0
                ORDER BY LOWER(v)
            """)
            return [r["v"] for r in cur.fetchall() if r["v"]]
        finally:
            cur.close()


    # -------- Stocks --------
    def add_stock(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
        warehouse_no: int,
        package_no: int,
        unit_price: Optional[float],
        count: int,
        has_badge: int = 0,
    ) -> int:

        if count <= 0:
            raise ValueError("Count must be > 0")

        # ---- Resolve effective price ----
        price: Optional[float] = None
        user_provided_price = unit_price is not None and str(unit_price) != ""

        if user_provided_price:
            price = float(unit_price)
            if price < 0:
                raise ValueError("Price must be ≥ 0")
        else:
            price = self.get_effective_price(item_type, school, color, size)

        if price is None:
            raise ValueError("يجب إدخال السعر مرة واحدة على الأقل لهذا الصنف.")

        # Save default price ONLY if user explicitly typed it
        if user_provided_price:
            self.ensure_default_price(item_type, price)

        self.ensure_package_open(warehouse_no, package_no)

        with self.conn:
            cur = self.conn.execute(
                """INSERT INTO stocks
                (item_type,school,color,size,warehouse_no,package_no,unit_price,count,has_badge)
                VALUES(?,?,?,?,?,?,?,?,?)""",
                (
                    item_type.strip(),
                    school.strip(),
                    color.strip(),
                    size.strip(),
                    int(warehouse_no),
                    int(package_no),
                    float(price),
                    int(count),
                    int(has_badge),
                ),
            )

            stock_id = cur.lastrowid

            self.conn.execute(
                """INSERT INTO movements
                (ts,direction,stock_id,qty,note,bill_id,
                    item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    now_iso(),
                    "IN",
                    stock_id,
                    int(count),
                    "Income add",
                    None,
                    item_type.strip(),
                    school.strip(),
                    color.strip(),
                    size.strip(),
                    int(warehouse_no),
                    int(package_no),
                    float(price),
                    int(has_badge),
                ),
            )

            self._upsert_history(
                {
                    "item_type": item_type,
                    "school": school,
                    "color": color,
                    "size": size,
                }
            )

            stock_uuid_row = self.conn.execute(
                "SELECT uuid FROM stocks WHERE id=?", (int(stock_id),)
            ).fetchone()
            self._record_sync_event("STOCK_INCOME", {
                "stock_uuid":   stock_uuid_row[0] if stock_uuid_row else None,
                "stock_id":     int(stock_id),
                "item_type":    item_type.strip(),
                "school":       school.strip(),
                "color":        color.strip(),
                "size":         size.strip(),
                "warehouse_no": int(warehouse_no),
                "package_no":   int(package_no),
                "unit_price":   float(price),
                "count":        int(count),
                "has_badge":    int(has_badge),
            })

        return int(stock_id)
    def cleanup_unused_specs(self):
        """
        Remove unused item_type / school / color only.
        Size is intentionally excluded.
        """
        cur = self.conn.cursor()
        try:
            for field in ("item_type", "school", "color"):
                cur.execute(f"""
                    DELETE FROM spec_history
                    WHERE field = ?
                    AND value NOT IN (
                        SELECT DISTINCT TRIM({field}) FROM stocks
                        UNION
                        SELECT DISTINCT TRIM({field}) FROM bill_items
                    )
                """, (field,))
            self.conn.commit()
        finally:
            cur.close()

    def _filters_where(self, filters: Dict[str, Any], prefix: str = "") -> Tuple[str, List[Any]]:
        """
        Supports:
        - Single-value filters (strings/ints) for item_type/school/color/size/warehouse_no/package_no
        - Exactly one multi-value filter (list) among the above keys.
        Includes count = 0 rows so catalog grids and inventory lists keep zero-stock products.
        """
        where: List[str] = ["1=1"]
        args: List[Any] = []

        # Detect multi-field usage and enforce "only one field can be used" rule
        candidate_keys_txt = ("item_type", "school", "color", "size")
        candidate_keys_num = ("warehouse_no", "package_no")
        all_keys = candidate_keys_txt + candidate_keys_num

        multi_keys = [k for k in all_keys if isinstance(filters.get(k), list) and len(filters.get(k) or []) > 0]
        if len(multi_keys) > 1:
            # Hard guard (UI already prevents this, but keep DB side correct)
            raise ValueError("يمكن اختيار أكثر من قيمة لحقل واحد فقط، وليس لأكثر من حقل.")

        # Apply single-value filters
        for k in candidate_keys_txt:
            v = filters.get(k)
            if v and not isinstance(v, list):
                v = str(v).strip()
                if v:
                    where.append(f"LOWER(TRIM({prefix}{k})) = LOWER(?)")
                    args.append(v)

        for k in candidate_keys_num:
            v = filters.get(k)
            if v not in (None, "", []) and not isinstance(v, list):
                where.append(f"{prefix}{k} = ?")
                args.append(int(v))

        # Apply the one allowed multi-value filter (if any)
        if multi_keys:
            mk = multi_keys[0]
            vals = [x for x in (filters.get(mk) or []) if x not in (None, "")]
            if vals:
                placeholders = ",".join(["?"] * len(vals))
                if mk in candidate_keys_txt:
                    where.append(f"LOWER({prefix}{mk}) IN ({placeholders})")
                    args.extend([str(x).strip().lower() for x in vals])
                else:
                    where.append(f"{prefix}{mk} IN ({placeholders})")
                    args.extend([int(x) for x in vals])

        return (" AND ".join(where)) if where else "1=1", args


    def search_stocks(self, filters: Dict[str, Any]) -> List[StockRow]:
        where, args = self._filters_where(filters)
        cur = self.conn.cursor()
        cur.execute(
            f"""SELECT id,item_type,school,color,size,warehouse_no,package_no,unit_price,count,has_badge
                FROM stocks
                WHERE {where}
                ORDER BY id ASC""",
            args,
        )
        rows = [
            StockRow(
                id=r["id"], item_type=r["item_type"], school=r["school"], color=r["color"],
                size=r["size"], warehouse_no=r["warehouse_no"], package_no=r["package_no"],
                unit_price=r["unit_price"], count=r["count"], has_badge=int(r["has_badge"] or 0),
            )

            for r in cur.fetchall()
        ]
        cur.close()
        return rows

    def current_inventory(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        where, args = self._filters_where(filters, prefix="s.")

        cur = self.conn.cursor()
        cur.execute(
            f"""
            SELECT
                MIN(s.id)              AS id,
                s.item_type,
                s.school,
                s.color,
                s.size,
                s.warehouse_no,
                GROUP_CONCAT(s.package_no, ', ') AS package_no,
                s.unit_price,
                SUM(s.count)           AS count,
                SUM(s.count * s.unit_price) AS value,
                MAX(s.has_badge)       AS has_badge
            FROM stocks s
            WHERE {where}
            GROUP BY
                s.item_type,
                s.school,
                s.color,
                s.size,
                s.warehouse_no,
                s.unit_price
            ORDER BY
                s.item_type,
                s.school,
                s.color,

                -- 1) numeric sizes FIRST (digits-only)
                CASE
                    WHEN TRIM(s.size) NOT GLOB '*[^0-9]*' THEN 0
                    ELSE 1
                END,

                -- 2) numeric sizes sorted numerically
                CASE
                    WHEN TRIM(s.size) NOT GLOB '*[^0-9]*'
                    THEN CAST(s.size AS INTEGER)
                    ELSE NULL
                END,

                -- 3) alpha sizes AFTER numerics
                CASE UPPER(TRIM(s.size))
                    WHEN 'XXS' THEN 1
                    WHEN 'XS'  THEN 2
                    WHEN 'S'   THEN 3
                    WHEN 'M'   THEN 4
                    WHEN 'L'   THEN 5
                    WHEN 'XL'  THEN 6
                    WHEN '2XL' THEN 7
                    WHEN '3XL' THEN 8
                    WHEN '4XL' THEN 9
                    WHEN '5XL' THEN 10
                    ELSE 99
                END



            """,
            args,
        )

        rows = []
        for r in cur.fetchall():
            rows.append(
                dict(
                    id=r["id"],
                    item_type=r["item_type"],
                    school=r["school"],
                    color=r["color"],
                    size=r["size"],
                    warehouse_no=r["warehouse_no"],
                    package_no=r["package_no"],  # ← now "93, 121, 137"
                    unit_price=r["unit_price"],
                    count=int(r["count"]),
                    value=float(r["value"]),
                    has_badge=int(r["has_badge"] or 0),
                )
            )

        cur.close()
        return rows

    # -------- Billing --------
    def create_bill(
        self,
        customer: str,
        bill_lines: List[Dict[str, Any]],
        target_pos: Optional[str] = None,
    ) -> int:
        """
        Creates a bill from planned lines.

        Each planned line can:
        - Pull from a specific stock row via stock_id, OR
        - Match by specs (item_type/school/color/size and optional warehouse/package)

        If `allow_factory_fill` is True and stock is insufficient (or absent),
        the remainder is booked as FACTORY (no stock row). We also upsert the specs
        into spec_history so they appear later in autocompletes.

        Carries `has_badge` from stock rows when taking from stock; falls back to the
        line-provided flag for factory portions.
        """
        if not bill_lines:
            raise ValueError("Bill has no items")

        def _candidates_for_line(line: Dict[str, Any]) -> List[sqlite3.Row]:
            cur = self.conn.cursor()
            try:
                if line.get("stock_id"):
                    cur.execute(
                        "SELECT * FROM stocks WHERE id=? AND count>0 ORDER BY id ASC",
                        (int(line["stock_id"]),),
                    )
                else:
                    # build a flexible matcher by specs
                    where_parts = ["count > 0"]
                    args: List[Any] = []
                    for k in ("item_type", "school", "color", "size"):
                        v = (line.get(k) or "").strip()
                        if v:
                            where_parts.append(f"LOWER(TRIM({k})) = LOWER(?)")
                            args.append(v)
                    for k in ("warehouse_no", "package_no"):
                        v = line.get(k)
                        if v not in (None, "", 0, "0"):
                            where_parts.append(f"{k} = ?")
                            args.append(int(v))
                    cur.execute(f"SELECT * FROM stocks WHERE {' AND '.join(where_parts)} ORDER BY id ASC", args)
                return cur.fetchall()
            finally:
                cur.close()

        with self.conn:
            # create bill shell
            bill_cur = self.conn.execute(
                "INSERT INTO bills(created_at,customer,total,status) VALUES(?,?,?,?)",
                (now_iso(), (customer or "").strip() or None, 0.0, "CONFIRMED"),
            )
            bill_id = int(bill_cur.lastrowid)
            total = 0.0
            affected_packages: set = set()
            user_priced_specs: Dict[Tuple[str, str, str, str], float] = {}

            for line in bill_lines:
                # --- Parse & validate quantity ---
                try:
                    qty_needed = int(line["qty"])
                except Exception:
                    raise ValueError("Qty must be > 0")
                if qty_needed <= 0:
                    raise ValueError("Qty must be > 0")

                # Decide mode early
                allow_factory = bool(line.get("allow_factory_fill"))

                # --- Locate stock candidates ---
                # Enforce FACTORY-ONLY: if it's a factory line with NO stock_id, do not pull from stock.
                if allow_factory and not line.get("stock_id"):
                    cands = []
                else:
                    cands = _candidates_for_line(line)

                # --- Greedy take from available stock ---
                remaining = qty_needed
                chunks: List[Tuple[sqlite3.Row, int]] = []
                for s in cands:
                    if remaining <= 0:
                        break
                    take = min(int(s["count"]), remaining)
                    if take > 0:
                        chunks.append((s, take))
                        remaining -= take

                if remaining > 0 and not allow_factory:
                    # Friendly error in Arabic with context
                    raise ValueError(
                        f"لا توجد كمية كافية للصنف "
                        f"{line.get('item_type','(النوع?)')} / {line.get('school','(المدرسة?)')} / {line.get('size','(المقاس?)')} "
                        f"(المطلوب {qty_needed}، المتاح {qty_needed-remaining})."
                    )


                taken = qty_needed - remaining

                # --- Decide unit price base ---
                price_base: Optional[float] = None
                if "unit_price" in line and line["unit_price"] is not None and str(line["unit_price"]) != "":
                    price_base = float(line["unit_price"])
                elif taken > 0:
                    # If pulling from stock and no explicit price provided, use weighted average
                    value_sum = sum(float(s["unit_price"]) * take for s, take in chunks)
                    qty_sum = sum(take for _, take in chunks)
                    if qty_sum > 0:
                        price_base = round(value_sum / qty_sum, 4)
                # If still None and it's factory-only, default to 0.0 (user could have typed price in the dialog)
                if price_base is None:
                    price_base = 0.0

                user_set_price = bool(line.get("user_set_price"))
                if user_set_price:
                    it = str(line.get("item_type") or "").strip()
                    sc = str(line.get("school") or "").strip()
                    cl = str(line.get("color") or "").strip()
                    sz = str(line.get("size") or "").strip()
                    if it and sc and cl and sz:
                        user_priced_specs[(it, sc, cl, sz)] = float(price_base)

                # Derive has_badge for the STOCK part from the first chunk; for factory we fall back to line
                stock_badge = int(chunks[0][0]["has_badge"]) if chunks else int(line.get("has_badge") or 0)

                # --- Insert bill_items for STOCK portion (one row per chunk with actual WH/PKG) ---
                if taken > 0:
                    for s, take in chunks:
                        stock_price = float(price_base)  # billed price
                        line_total_chunk = stock_price * int(take)
                        self.conn.execute(
                            """INSERT INTO bill_items
                            (bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,qty,line_total,origin,has_badge)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (
                                bill_id,
                                line["item_type"], line["school"], line["color"], line["size"],
                                int(s["warehouse_no"]), int(s["package_no"]),
                                stock_price, int(take), line_total_chunk, "STOCK",
                                int(s["has_badge"] or 0),
                            ),
                        )
                        total += line_total_chunk


                # --- Insert bill_items for FACTORY remainder (if any) ---
                if remaining > 0 and allow_factory:
                    # Remember these specs in history so they show in dropdowns later
                    self._upsert_history({
                        "item_type": line.get("item_type"),
                        "school":    line.get("school"),
                        "color":     line.get("color"),
                        "size":      line.get("size"),
                    })

                    factory_price = float(price_base)
                    factory_line_total = factory_price * remaining
                    self.conn.execute(
                        """INSERT INTO bill_items
                        (bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,qty,line_total,origin,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            bill_id,
                            line["item_type"], line["school"], line["color"], line["size"],
                            int(line.get("warehouse_no") or 0), int(line.get("package_no") or 0),
                            factory_price, int(remaining), factory_line_total, "FACTORY",
                            int(line.get("has_badge") or 0),
                        ),
                    )
                    total += factory_line_total

                # --- Apply stock deductions + movements for STOCK part ---
                for s, take in chunks:
                    self.conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?", (int(take), int(s["id"])))
                    if user_set_price and int(take) > 0:
                        self.conn.execute(
                            "UPDATE stocks SET unit_price = ? WHERE id = ?",
                            (float(price_base), int(s["id"])),
                        )
                    affected_packages.add((int(s["warehouse_no"]), int(s["package_no"])))
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            now_iso(), "OUT", int(s["id"]), int(take), "Bill deduction", bill_id,
                            s["item_type"], s["school"], s["color"], s["size"],
                            int(s["warehouse_no"]), int(s["package_no"]), float(s["unit_price"]),
                            int(s["has_badge"] or 0),
                        ),
                    )

                # --- Movements for FACTORY part ---
                if remaining > 0 and allow_factory:
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            now_iso(), "OUT_FACTORY", None, int(remaining), "Factory direct", bill_id,
                            line["item_type"], line["school"], line["color"], line["size"],
                            int(line.get("warehouse_no") or 0), int(line.get("package_no") or 0),
                            float(price_base),
                            int(line.get("has_badge") or 0),
                        ),
                    )

            # Keep stock rows at count 0 so products remain in pickers; trim orphan specs only.
            self.cleanup_unused_specs()
            self.conn.execute("UPDATE bills SET total=? WHERE id=?", (float(total), bill_id))
            for wh, pkg in affected_packages:
                self.auto_reopen_package_if_empty(wh, pkg)

            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id=?", (bill_id,)
            ).fetchone()
            items_payload = [
                {
                    "item_type": r[0], "school": r[1], "color": r[2], "size": r[3],
                    "warehouse_no": int(r[4] or 0), "package_no": int(r[5] or 0),
                    "unit_price": float(r[6]), "qty": int(r[7]),
                    "line_total": float(r[8]), "origin": r[9],
                    "has_badge": int(r[10] or 0),
                }
                for r in self.conn.execute(
                    "SELECT item_type,school,color,size,warehouse_no,package_no,"
                    "unit_price,qty,line_total,origin,has_badge "
                    "FROM bill_items WHERE bill_id=?", (bill_id,)
                ).fetchall()
            ]

            tp = (target_pos or "").strip()
            if tp:
                # Phase 3: a bill whose "customer" is a POS device is
                # actually a branch shipment. Emit STOCK_TRANSFER_OUT
                # scoped to that POS instead of the usual SALE_CREATED.
                # The local bill/bill_items/movements rows remain as
                # the warehouse-side audit trail of what was shipped.
                shipment_items = [
                    {
                        "item_type":  it["item_type"],
                        "school":     it["school"],
                        "color":      it["color"],
                        "size":       it["size"],
                        "unit_price": float(it["unit_price"]),
                        "qty":        int(it["qty"]),
                    }
                    for it in items_payload if int(it["qty"]) > 0
                ]
                self._record_branch_shipment_event(
                    shipment_uuid=(bill_uuid_row[0] if bill_uuid_row else None) or "",
                    target_name=tp,
                    note=f"bill #{bill_id}",
                    lines=shipment_items,
                )
            else:
                self._record_sync_event("SALE_CREATED", {
                    "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                    "bill_id":   bill_id,
                    "customer":  (customer or "").strip() or None,
                    "total":     float(total),
                    "items":     items_payload,
                })

            # If the cashier typed a manual price in the bill flow, push that
            # canonical spec price to POS devices (scoped for branch shipments).
            if user_priced_specs:
                tp2 = (target_pos or "").strip()
                if tp2:
                    sync_mode = "selected-pos"
                    sync_devs = [tp2]
                else:
                    sync_mode = "all-pos"
                    sync_devs = None

                for (it, sc, cl, sz), pr in user_priced_specs.items():
                    self.emit_price_update_sync_events(
                        filters={"item_type": it, "school": sc, "color": cl, "size": sz},
                        new_price=float(pr),
                        note=f"Bill-priced default (bill #{bill_id})",
                        sync_mode=sync_mode,
                        sync_pos_devices=sync_devs,
                        audit_mode="bill-auto",
                    )

            return bill_id

    # -------- Bill history APIs --------
    def list_bills(self) -> List[Dict[str, Any]]:
        cur = self.conn.cursor()
        cur.execute("SELECT id,created_at,customer,total,COALESCE(status,'CONFIRMED') AS status FROM bills ORDER BY id DESC")
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        for row in rows:
            customer = (row.get("customer") or "").strip()
            row["bill_kind"] = "BRANCH_SHIPMENT" if normalize_branch_customer_name(customer) else "NORMAL"
        return rows

    def list_bill_items(self, bill_id: int) -> List[Dict[str, Any]]:
        cur = self.conn.cursor()
        cur.execute(
            """SELECT item_type,school,color,size,warehouse_no,package_no,unit_price,qty,line_total,origin,has_badge
            FROM bill_items WHERE bill_id=?""",
            (int(bill_id),),
        )
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows

    # -------- Bill Lifecycle (Draft / Confirm / Void) --------

    def create_draft_bill(self, customer: str, bill_lines: List[Dict[str, Any]]) -> int:
        """Create a DRAFT bill — saves bill data but does NOT deduct stock."""
        if not bill_lines:
            raise ValueError("Bill has no items")
        with self.conn:
            bill_cur = self.conn.execute(
                "INSERT INTO bills(created_at,customer,total,status) VALUES(?,?,?,?)",
                (now_iso(), (customer or "").strip() or None, 0.0, "DRAFT"),
            )
            bill_id = int(bill_cur.lastrowid)
            total = 0.0
            for line in bill_lines:
                qty = int(line["qty"])
                if qty <= 0:
                    raise ValueError("Qty must be > 0")
                price = float(line.get("unit_price") or 0.0)
                line_total = price * qty
                allow_factory = bool(line.get("allow_factory_fill"))
                origin = "FACTORY" if allow_factory else "STOCK"
                has_badge = int(line.get("has_badge") or 0)
                self.conn.execute(
                    """INSERT INTO bill_items
                    (bill_id,item_type,school,color,size,warehouse_no,package_no,
                     unit_price,qty,line_total,origin,has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (bill_id, line["item_type"], line["school"], line["color"],
                     line["size"], int(line.get("warehouse_no") or 0),
                     int(line.get("package_no") or 0), price, qty, line_total,
                     origin, has_badge),
                )
                total += line_total
            self.conn.execute("UPDATE bills SET total=? WHERE id=?", (float(total), bill_id))
            return bill_id

    def confirm_draft_bill(self, bill_id: int) -> None:
        """Confirm a DRAFT bill — deduct stock and record movements."""
        cur = self.conn.execute(
            "SELECT COALESCE(status,'CONFIRMED') AS status FROM bills WHERE id=?", (int(bill_id),))
        row = cur.fetchone()
        if not row:
            raise ValueError("الفاتورة غير موجودة")
        if row["status"] != "DRAFT":
            raise ValueError("هذه الفاتورة ليست مسودة")
        items = self.list_bill_items(bill_id)
        if not items:
            raise ValueError("الفاتورة لا تحتوي على بنود")
        with self.conn:
            affected_packages: set = set()
            for item in items:
                if item.get("origin") == "FACTORY":
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,
                         item_type,school,color,size,warehouse_no,package_no,
                         unit_price,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (now_iso(), "OUT_FACTORY", None, int(item["qty"]),
                         "Draft confirmed (factory)", bill_id,
                         item["item_type"], item["school"], item["color"],
                         item["size"], int(item.get("warehouse_no") or 0),
                         int(item.get("package_no") or 0),
                         float(item["unit_price"]),
                         int(item.get("has_badge") or 0)),
                    )
                    continue
                qty_needed = int(item["qty"])
                where_parts = ["count > 0"]
                args: List[Any] = []
                for k in ("item_type", "school", "color", "size"):
                    v = (item.get(k) or "").strip()
                    if v:
                        where_parts.append(f"LOWER(TRIM({k})) = LOWER(?)")
                        args.append(v)
                cands = self.conn.execute(
                    f"SELECT * FROM stocks WHERE {' AND '.join(where_parts)} ORDER BY id ASC",
                    args
                ).fetchall()
                remaining = qty_needed
                for s in cands:
                    if remaining <= 0:
                        break
                    take = min(int(s["count"]), remaining)
                    if take > 0:
                        self.conn.execute(
                            "UPDATE stocks SET count = count - ? WHERE id = ?",
                            (take, int(s["id"]))
                        )
                        self.conn.execute(
                            """INSERT INTO movements
                            (ts,direction,stock_id,qty,note,bill_id,
                             item_type,school,color,size,warehouse_no,package_no,
                             unit_price,has_badge)
                            VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (now_iso(), "OUT", int(s["id"]), take,
                             "Draft confirmed", bill_id,
                             s["item_type"], s["school"], s["color"], s["size"],
                             int(s["warehouse_no"]), int(s["package_no"]),
                             float(s["unit_price"]),
                             int(s["has_badge"] or 0)),
                        )
                        affected_packages.add((int(s["warehouse_no"]), int(s["package_no"])))
                        remaining -= take
                if remaining > 0:
                    raise ValueError(
                        f"كمية غير كافية للصنف {item['item_type']} / {item['school']} / {item['size']} "
                        f"(المطلوب {qty_needed}، المتاح {qty_needed - remaining})"
                    )
            self.cleanup_unused_specs()
            self.conn.execute(
                "UPDATE bills SET status='CONFIRMED' WHERE id=?", (int(bill_id),)
            )
            for wh, pkg in affected_packages:
                self.auto_reopen_package_if_empty(wh, pkg)

    def void_bill(self, bill_id: int) -> None:
        """Void a confirmed bill — return items to stock."""
        cur = self.conn.execute(
            "SELECT COALESCE(status,'CONFIRMED') AS status FROM bills WHERE id=?", (int(bill_id),))
        row = cur.fetchone()
        if not row:
            raise ValueError("الفاتورة غير موجودة")
        if row["status"] != "CONFIRMED":
            raise ValueError("يمكن إلغاء الفواتير المؤكدة فقط")
        items = self.list_bill_items(bill_id)
        with self.conn:
            for item in items:
                if item.get("origin") == "FACTORY":
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,
                         item_type,school,color,size,warehouse_no,package_no,
                         unit_price,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (now_iso(), "VOID", None, int(item["qty"]),
                         "Bill voided (factory)", bill_id,
                         item["item_type"], item["school"], item["color"],
                         item["size"], int(item.get("warehouse_no") or 0),
                         int(item.get("package_no") or 0),
                         float(item["unit_price"]),
                         int(item.get("has_badge") or 0)),
                    )
                    continue
                qty = int(item["qty"])
                wh = int(item.get("warehouse_no") or 0)
                pkg = int(item.get("package_no") or 0)
                if wh > 0 and pkg > 0:
                    self.ensure_package_open(wh, pkg)
                cur2 = self.conn.execute(
                    """INSERT INTO stocks
                    (item_type,school,color,size,warehouse_no,package_no,
                     unit_price,count,has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?)""",
                    (item["item_type"], item["school"], item["color"],
                     item["size"], wh, pkg, float(item["unit_price"]),
                     qty, int(item.get("has_badge") or 0)),
                )
                stock_id = cur2.lastrowid
                self.conn.execute(
                    """INSERT INTO movements
                    (ts,direction,stock_id,qty,note,bill_id,
                     item_type,school,color,size,warehouse_no,package_no,
                     unit_price,has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), "VOID", stock_id, qty,
                     "Bill voided (stock returned)", bill_id,
                     item["item_type"], item["school"], item["color"],
                     item["size"], wh, pkg, float(item["unit_price"]),
                     int(item.get("has_badge") or 0)),
                )
            self.conn.execute(
                "UPDATE bills SET status='VOID' WHERE id=?", (int(bill_id),)
            )
            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id=?", (int(bill_id),)
            ).fetchone()
            self._record_sync_event("SALE_VOIDED", {
                "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                "bill_id":   int(bill_id),
                "items": [
                    {
                        "item_type":    item["item_type"],
                        "school":       item["school"],
                        "color":        item["color"],
                        "size":         item["size"],
                        "warehouse_no": int(item.get("warehouse_no") or 0),
                        "package_no":   int(item.get("package_no") or 0),
                        "unit_price":   float(item["unit_price"]),
                        "qty":          int(item["qty"]),
                        "origin":       item.get("origin"),
                        "has_badge":    int(item.get("has_badge") or 0),
                    }
                    for item in items
                ],
            })

    def delete_draft_bill(self, bill_id: int) -> None:
        """Delete a draft bill entirely (no stock was deducted)."""
        cur = self.conn.execute(
            "SELECT COALESCE(status,'CONFIRMED') AS status FROM bills WHERE id=?", (int(bill_id),))
        row = cur.fetchone()
        if not row:
            raise ValueError("الفاتورة غير موجودة")
        if row["status"] != "DRAFT":
            raise ValueError("يمكن حذف المسودات فقط")
        with self.conn:
            self.conn.execute("DELETE FROM bill_items WHERE bill_id=?", (int(bill_id),))
            self.conn.execute("DELETE FROM bills WHERE id=?", (int(bill_id),))

    # -------- Returns --------

    def process_return(self, bill_id: int, return_lines: List[Dict[str, Any]], note: str = "") -> int:
        """Process a return against a confirmed bill. Returns the return_id."""
        cur = self.conn.execute(
            "SELECT customer, COALESCE(status,'CONFIRMED') AS status FROM bills WHERE id=?", (int(bill_id),))
        row = cur.fetchone()
        if not row:
            raise ValueError("الفاتورة غير موجودة")
        if row["status"] != "CONFIRMED":
            raise ValueError("المرتجعات متاحة فقط للفواتير المؤكدة")
        if not return_lines:
            raise ValueError("لا توجد أصناف للإرجاع")
        branch_target = normalize_branch_customer_name(row["customer"])
        with self.conn:
            ret_cur = self.conn.execute(
                "INSERT INTO returns(bill_id, created_at, note) VALUES(?,?,?)",
                (int(bill_id), now_iso(), note or None),
            )
            return_id = int(ret_cur.lastrowid)
            for idx, line in enumerate(return_lines):
                qty = int(line["qty"])
                if qty <= 0:
                    continue
                wh = int(line.get("warehouse_no") or 0)
                pkg = int(line.get("package_no") or 0)
                price = float(line.get("unit_price") or 0.0)
                has_badge = int(line.get("has_badge") or 0)
                self.conn.execute(
                    """INSERT INTO return_items
                    (return_id, bill_item_id, item_type, school, color, size,
                     warehouse_no, package_no, unit_price, qty, has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (return_id, line.get("bill_item_id"), line["item_type"],
                     line["school"], line["color"], line["size"],
                     wh, pkg, price, qty, has_badge),
                )
                if branch_target:
                    self.conn.execute(
                        """
                        INSERT INTO branch_inventory_queue
                            (sync_event_uuid, queue_kind, source_device, requested_target_device,
                             external_ref, line_index, created_at, item_type, school, color,
                             size, unit_price, qty, has_badge, note, status)
                        VALUES (?, 'RETURN', ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'PENDING')
                        """,
                        (
                            f"manual-return:{return_id}:{idx}",
                            branch_target,
                            f"warehouse-return-{return_id}",
                            idx,
                            now_iso(),
                            line["item_type"],
                            line["school"],
                            line["color"],
                            line["size"],
                            price,
                            qty,
                            has_badge,
                            (note or "").strip() or f"مرتجع فاتورة شحن إلى {branch_target}",
                        ),
                    )
                else:
                    if wh > 0 and pkg > 0:
                        self.ensure_package_open(wh, pkg)
                    stock_cur = self.conn.execute(
                        """INSERT INTO stocks
                        (item_type,school,color,size,warehouse_no,package_no,
                         unit_price,count,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?)""",
                        (line["item_type"], line["school"], line["color"],
                         line["size"], wh, pkg, price, qty, has_badge),
                    )
                    stock_id = stock_cur.lastrowid
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,
                         item_type,school,color,size,warehouse_no,package_no,
                         unit_price,has_badge)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (now_iso(), "RETURN_IN", stock_id, qty,
                         f"Return #{return_id}", bill_id,
                         line["item_type"], line["school"], line["color"],
                         line["size"], wh, pkg, price, has_badge),
                    )
            return_uuid_row = self.conn.execute(
                "SELECT uuid FROM returns WHERE id=?", (return_id,)
            ).fetchone()
            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id=?", (int(bill_id),)
            ).fetchone()
            self._record_sync_event("SALE_RETURNED", {
                "return_uuid": return_uuid_row[0] if return_uuid_row else None,
                "return_id":   return_id,
                "bill_uuid":   bill_uuid_row[0] if bill_uuid_row else None,
                "bill_id":     int(bill_id),
                "note":        note or None,
                "lines": [
                    {
                        "item_type":    ln["item_type"],
                        "school":       ln["school"],
                        "color":        ln["color"],
                        "size":         ln["size"],
                        "warehouse_no": int(ln.get("warehouse_no") or 0),
                        "package_no":   int(ln.get("package_no") or 0),
                        "unit_price":   float(ln.get("unit_price") or 0.0),
                        "qty":          int(ln["qty"]),
                        "has_badge":    int(ln.get("has_badge") or 0),
                    }
                    for ln in return_lines if int(ln.get("qty") or 0) > 0
                ],
            })
        return return_id

    def list_returns_for_bill(self, bill_id: int) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            "SELECT id, bill_id, created_at, note FROM returns WHERE bill_id=? ORDER BY id DESC",
            (int(bill_id),))
        return [dict(r) for r in cur.fetchall()]

    def list_return_items(self, return_id: int) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            """SELECT item_type, school, color, size, warehouse_no, package_no,
                      unit_price, qty, has_badge
               FROM return_items WHERE return_id=?""",
            (int(return_id),))
        return [dict(r) for r in cur.fetchall()]

    # -------- Warehouse Transfer --------

    def transfer_stock(self, stock_id: int, qty: int,
                       dest_warehouse_no: int, dest_package_no: int,
                       note: str = "Warehouse transfer") -> Tuple[int, int]:
        """Move qty items from stock_id to dest warehouse/package."""
        if qty <= 0:
            raise ValueError("الكمية يجب أن تكون أكبر من صفر")
        with self.conn:
            cur = self.conn.execute("SELECT * FROM stocks WHERE id=?", (int(stock_id),))
            s = cur.fetchone()
            if not s:
                raise ValueError("صف المخزون غير موجود")
            if int(s["count"]) < qty:
                raise ValueError(f"الكمية المتاحة ({s['count']}) أقل من المطلوب ({qty})")
            src_wh = int(s["warehouse_no"])
            src_pkg = int(s["package_no"])
            if src_wh == dest_warehouse_no and src_pkg == dest_package_no:
                raise ValueError("المصدر والوجهة متطابقان")
            self.ensure_package_open(dest_warehouse_no, dest_package_no)
            self.conn.execute(
                "UPDATE stocks SET count = count - ? WHERE id = ?",
                (qty, int(stock_id)))
            dest_cur = self.conn.execute(
                """INSERT INTO stocks
                (item_type,school,color,size,warehouse_no,package_no,
                 unit_price,count,has_badge)
                VALUES(?,?,?,?,?,?,?,?,?)""",
                (s["item_type"], s["school"], s["color"], s["size"],
                 int(dest_warehouse_no), int(dest_package_no),
                 float(s["unit_price"]), qty, int(s["has_badge"] or 0)))
            dest_stock_id = dest_cur.lastrowid
            out_cur = self.conn.execute(
                """INSERT INTO movements
                (ts,direction,stock_id,qty,note,bill_id,
                 item_type,school,color,size,warehouse_no,package_no,
                 unit_price,has_badge)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (now_iso(), "TRANSFER_OUT", int(stock_id), qty,
                 f"{note} -> WH{dest_warehouse_no}/PKG{dest_package_no}",
                 None, s["item_type"], s["school"], s["color"], s["size"],
                 src_wh, src_pkg, float(s["unit_price"]),
                 int(s["has_badge"] or 0)))
            out_id = out_cur.lastrowid
            in_cur = self.conn.execute(
                """INSERT INTO movements
                (ts,direction,stock_id,qty,note,bill_id,
                 item_type,school,color,size,warehouse_no,package_no,
                 unit_price,has_badge)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (now_iso(), "TRANSFER_IN", dest_stock_id, qty,
                 f"{note} <- WH{src_wh}/PKG{src_pkg}",
                 None, s["item_type"], s["school"], s["color"], s["size"],
                 int(dest_warehouse_no), int(dest_package_no),
                 float(s["unit_price"]), int(s["has_badge"] or 0)))
            in_id = in_cur.lastrowid
            self.auto_reopen_package_if_empty(src_wh, src_pkg)
            dest_uuid_row = self.conn.execute(
                "SELECT uuid FROM stocks WHERE id=?", (int(dest_stock_id),)
            ).fetchone()
            src_uuid_row = self.conn.execute(
                "SELECT uuid FROM stocks WHERE id=?", (int(stock_id),)
            ).fetchone()
            self._record_sync_event("STOCK_TRANSFER", {
                "src_stock_uuid":  src_uuid_row[0] if src_uuid_row else None,
                "src_stock_id":    int(stock_id),
                "src_warehouse_no": src_wh,
                "src_package_no":   src_pkg,
                "dest_stock_uuid": dest_uuid_row[0] if dest_uuid_row else None,
                "dest_stock_id":   int(dest_stock_id),
                "dest_warehouse_no": int(dest_warehouse_no),
                "dest_package_no":   int(dest_package_no),
                "qty":             int(qty),
                "item_type":       s["item_type"],
                "school":          s["school"],
                "color":           s["color"],
                "size":            s["size"],
                "unit_price":      float(s["unit_price"]),
                "has_badge":       int(s["has_badge"] or 0),
                "note":            note,
            })
        return (out_id, in_id)

    # -------- Branch Shipment (Phase 3) --------

    def list_known_pos_device_names(self) -> List[str]:
        """Return the cached list of POS device names for the bill
        dialog's Customer dropdown. The cache is populated by
        sync_client.refresh_device_list() on each sync cycle.
        """
        names = set(DEFAULT_BRANCH_POS_NAMES)
        try:
            rows = self.conn.execute(
                "SELECT device_name FROM known_devices "
                "WHERE role = 'pos' ORDER BY device_name"
            ).fetchall()
        except sqlite3.OperationalError:
            return sorted(names)
        names.update(r[0] for r in rows if r and r[0])
        return sorted(names)

    def list_branch_inventory_queue(self, status: Optional[str] = "PENDING") -> List[Dict[str, Any]]:
        sql = """
            SELECT id, sync_event_uuid, queue_kind, source_device,
                   requested_target_device, external_ref, line_index, created_at,
                   item_type, school, color, size, unit_price, qty, has_badge,
                   note, status, processed_at, processed_note,
                   processed_warehouse_no, processed_package_no, rerouted_target_device
              FROM branch_inventory_queue
        """
        args: List[Any] = []
        if status:
            sql += " WHERE status = ?"
            args.append(status)
        sql += " ORDER BY created_at DESC, id DESC"
        return [dict(r) for r in self.conn.execute(sql, args).fetchall()]

    def resolve_pos_mirror_device_sql_filter(
        self, source_device: Optional[str],
    ) -> Tuple[str, List[str]]:
        """Map a POS display name (or UUID) to stored `source_device` values for mirror/ledger SQL.

        Sync payloads often store `device_uuid` while the UI lists `device_name`
        from `known_devices`. Expand so filters match either form.
        """
        label = (source_device or "").strip()
        if not label:
            return "", []
        keys: List[str] = []
        seen_lower = set()

        def _add(k: str) -> None:
            k = (k or "").strip()
            if not k:
                return
            lo = k.lower()
            if lo in seen_lower:
                return
            seen_lower.add(lo)
            keys.append(k)

        _add(label)
        try:
            cur = self.conn.execute(
                """
                SELECT device_name, device_uuid
                  FROM known_devices
                 WHERE LOWER(TRIM(device_name)) = LOWER(?)
                    OR TRIM(device_uuid) = TRIM(?)
                    OR LOWER(TRIM(device_uuid)) = LOWER(?)
                """,
                (label, label, label),
            )
            for row in cur.fetchall():
                _add(str(row["device_name"] or ""))
                _add(str(row["device_uuid"] or ""))
        except sqlite3.OperationalError:
            pass
        if not keys:
            return "", []
        ph = ",".join("?" * len(keys))
        return f" AND source_device IN ({ph})", keys

    def display_name_for_sync_source(self, raw: Optional[Any]) -> str:
        """Human-readable POS name for a mirror/ledger `source_device` value."""
        s = str(raw or "").strip()
        if not s:
            return ""
        try:
            row = self.conn.execute(
                """
                SELECT device_name FROM known_devices
                 WHERE TRIM(device_uuid) = TRIM(?)
                    OR LOWER(TRIM(device_name)) = LOWER(?)
                 LIMIT 1
                """,
                (s, s),
            ).fetchone()
            if row and row[0]:
                return str(row[0]).strip()
        except sqlite3.OperationalError:
            pass
        return s[:22] + "…" if len(s) > 26 else s

    def list_pos_reservations_mirror_device_picklist(self) -> List[str]:
        """Combobox values: known POS names plus raw identifiers from mirror/ledger tables."""
        names: set = set(self.list_known_pos_device_names())
        for q in (
            """
            SELECT DISTINCT source_device FROM pos_reservations_mirror
             WHERE source_device IS NOT NULL AND TRIM(source_device) != ''
            """,
            """
            SELECT DISTINCT source_device FROM pos_financial_ledger
             WHERE source_device IS NOT NULL AND TRIM(source_device) != ''
            """,
        ):
            try:
                for r in self.conn.execute(q).fetchall():
                    v = str(r[0] or "").strip()
                    if v:
                        names.add(v)
            except sqlite3.OperationalError:
                pass
        return [""] + sorted(names, key=lambda x: x.lower())

    def list_pos_reservations_mirror(
        self,
        source_device: Optional[str] = None,
        active_only: bool = False,
    ) -> List[Dict[str, Any]]:
        """Rows mirrored from POS sync events (after warehouse apply)."""
        sql = "SELECT * FROM pos_reservations_mirror WHERE 1=1"
        args: List[Any] = []
        frag, frag_args = self.resolve_pos_mirror_device_sql_filter(source_device)
        if frag:
            sql += frag
            args.extend(frag_args)
        if active_only:
            sql += " AND status = 'معلق'"
        sql += " ORDER BY updated_at DESC, id DESC"
        try:
            return [dict(r) for r in self.conn.execute(sql, args).fetchall()]
        except sqlite3.OperationalError:
            return []

    def list_pos_reservations_mirror_aggregated(
        self,
        source_device: Optional[str] = None,
        active_only: bool = False,
    ) -> List[Dict[str, Any]]:
        """Sum reservation mirror lines by product (type/school/color/size) across POS devices.

        Multiple reservation lines (same or different معرّف) that share the same normalized
        product keys are combined; ``agg_qty`` is the total quantity and
        ``reservation_line_count`` is how many mirror rows were merged.
        """
        sql = """
            SELECT TRIM(COALESCE(item_type, '')) AS item_type,
                   TRIM(COALESCE(school, '')) AS school,
                   TRIM(COALESCE(color, '')) AS color,
                   TRIM(COALESCE(CAST(size AS TEXT), '')) AS size,
                   SUM(COALESCE(qty, 0)) AS agg_qty,
                   CASE WHEN SUM(COALESCE(qty, 0)) > 0
                        THEN ROUND(SUM(COALESCE(total_amount, 0)) / SUM(COALESCE(qty, 0)), 2)
                        ELSE MAX(unit_price) END AS avg_unit_price,
                   SUM(COALESCE(total_amount, 0)) AS sum_total_amount,
                   SUM(COALESCE(paid_amount, 0)) AS sum_paid_amount,
                   COUNT(*) AS reservation_line_count,
                   COUNT(DISTINCT source_device) AS pos_device_count,
                   MAX(updated_at) AS last_updated
              FROM pos_reservations_mirror
             WHERE 1=1
        """
        args: List[Any] = []
        frag, frag_args = self.resolve_pos_mirror_device_sql_filter(source_device)
        if frag:
            sql += frag
            args.extend(frag_args)
        if active_only:
            sql += " AND status = 'معلق'"
        sql += """
          GROUP BY 1, 2, 3, 4
          ORDER BY MAX(updated_at) DESC, 1, 2, 3, 4
        """
        try:
            return [dict(r) for r in self.conn.execute(sql, args).fetchall()]
        except sqlite3.OperationalError:
            return []

    def list_pos_financial_summary_by_day(
        self,
        source_device: Optional[str] = None,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Aggregate pos_financial_ledger by calendar day (YYYY-MM-DD)."""
        sql = """
            SELECT day,
                   SUM(CASE WHEN category = 'sale' THEN amount ELSE 0 END) AS sales_amt,
                   SUM(CASE WHEN category = 'return_bill' THEN amount ELSE 0 END) AS returns_amt,
                   SUM(CASE WHEN category = 'void_bill' THEN amount ELSE 0 END) AS voids_amt,
                   SUM(CASE WHEN category = 'exchange_net' THEN amount ELSE 0 END) AS exchange_amt,
                   SUM(CASE WHEN category = 'reservation_downpayment' THEN amount ELSE 0 END) AS res_dep_amt,
                   SUM(CASE WHEN category = 'reservation_payment' THEN amount ELSE 0 END) AS res_pay_amt,
                   SUM(CASE WHEN category = 'reservation_collect' THEN amount ELSE 0 END) AS res_coll_amt,
                   SUM(amount) AS net_amt
              FROM pos_financial_ledger
             WHERE 1=1
        """
        args: List[Any] = []
        frag, frag_args = self.resolve_pos_mirror_device_sql_filter(source_device)
        if frag:
            sql += frag
            args.extend(frag_args)
        if date_from:
            sql += " AND day >= ?"
            args.append(date_from[:10])
        if date_to:
            sql += " AND day <= ?"
            args.append(date_to[:10])
        sql += " GROUP BY day ORDER BY day DESC"
        try:
            return [dict(r) for r in self.conn.execute(sql, args).fetchall()]
        except sqlite3.OperationalError:
            return []

    def list_pos_financial_ledger_detail(
        self,
        day: str,
        source_device: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        sql = """
            SELECT id, source_device, event_uuid, event_type, category, amount,
                   day, related_id, meta_json, created_at
              FROM pos_financial_ledger
             WHERE day = ?
        """
        args: List[Any] = [day[:10]]
        frag, frag_args = self.resolve_pos_mirror_device_sql_filter(source_device)
        if frag:
            sql += frag
            args.extend(frag_args)
        sql += " ORDER BY id ASC"
        try:
            return [dict(r) for r in self.conn.execute(sql, args).fetchall()]
        except sqlite3.OperationalError:
            return []

    def list_branch_cycle_reconciliation(
        self,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """Per-branch cycle summary:
        - shipped bill value from warehouse
        - received cash net from POS financial ledger
        - current branch stock value from latest snapshot
        - cycle gap = shipped - cash - stock_value
        """
        def _in_range(ts: str) -> bool:
            d = (str(ts or "").strip()[:10] if ts else "")
            if not d:
                return False
            if date_from and d < date_from[:10]:
                return False
            if date_to and d > date_to[:10]:
                return False
            return True

        out: Dict[str, Dict[str, Any]] = {}
        approved = list(DEFAULT_BRANCH_POS_NAMES)
        for dev in approved:
            out[dev] = {
                "branch_device": dev,
                "branch_name": branch_display_name(dev),
                "shipment_bills_count": 0,
                "shipment_qty": 0,
                "shipment_value": 0.0,
                "cash_net": 0.0,
                "stock_qty": 0,
                "stock_value": 0.0,
                "cycle_gap": 0.0,
            }

        # 1) Warehouse branch shipments
        try:
            bills = self.list_bills()
        except Exception:
            bills = []
        for b in bills:
            if (b.get("status") or "CONFIRMED") != "CONFIRMED":
                continue
            if not _in_range(str(b.get("created_at") or "")):
                continue
            branch_dev = normalize_branch_customer_name(b.get("customer"))
            if not branch_dev or branch_dev not in out:
                continue
            bid = int(b.get("id") or 0)
            lines = self.list_bill_items(bid)
            bucket = out[branch_dev]
            bucket["shipment_bills_count"] = int(bucket["shipment_bills_count"]) + 1
            for ln in lines:
                q = int(ln.get("qty") or 0)
                v = float(ln.get("line_total") or 0.0)
                bucket["shipment_qty"] = int(bucket["shipment_qty"]) + q
                bucket["shipment_value"] = float(bucket["shipment_value"]) + v

        # 2) Net cash received from branch POS ledgers
        for dev in approved:
            frag, frag_args = self.resolve_pos_mirror_device_sql_filter(dev)
            sql = "SELECT COALESCE(SUM(amount),0) FROM pos_financial_ledger WHERE 1=1"
            args: List[Any] = []
            if frag:
                sql += frag
                args.extend(frag_args)
            if date_from:
                sql += " AND day >= ?"
                args.append(date_from[:10])
            if date_to:
                sql += " AND day <= ?"
                args.append(date_to[:10])
            try:
                row = self.conn.execute(sql, args).fetchone()
                out[dev]["cash_net"] = float((row[0] if row else 0.0) or 0.0)
            except sqlite3.OperationalError:
                out[dev]["cash_net"] = 0.0

        # 3) Current stock value from latest snapshot source per branch
        latest_src: Dict[str, str] = {}
        try:
            snaps = self.conn.execute(
                """
                SELECT pm.source_device, pm.snapshot_at,
                       COALESCE(kd.device_name, pm.source_device) AS branch_name
                FROM pos_stocks_snapshot_meta pm
                LEFT JOIN known_devices kd
                    ON kd.device_name = pm.source_device
                    OR kd.device_uuid = pm.source_device
                """
            ).fetchall()
            latest_at: Dict[str, str] = {}
            for r in snaps:
                src = str(r[0] or "").strip()
                snap_at = str(r[1] or "").strip()
                branch = str(r[2] or "").strip()
                if branch not in out:
                    continue
                if branch not in latest_at or snap_at > latest_at[branch]:
                    latest_at[branch] = snap_at
                    latest_src[branch] = src
        except sqlite3.OperationalError:
            latest_src = {}

        for dev in approved:
            src = latest_src.get(dev)
            if not src:
                continue
            try:
                row = self.conn.execute(
                    """
                    SELECT COALESCE(SUM(count),0), COALESCE(SUM(count*unit_price),0)
                    FROM pos_stocks_mirror
                    WHERE source_device = ?
                    """,
                    (src,),
                ).fetchone()
                out[dev]["stock_qty"] = int((row[0] if row else 0) or 0)
                out[dev]["stock_value"] = float((row[1] if row else 0.0) or 0.0)
            except sqlite3.OperationalError:
                pass

        for dev in approved:
            b = out[dev]
            b["cycle_gap"] = float(b["shipment_value"]) - float(b["cash_net"]) - float(b["stock_value"])

        rows = list(out.values())
        rows.sort(key=lambda r: str(r.get("branch_name") or ""))
        return rows

    def get_branch_inventory_queue_item(self, queue_id: int) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            """
            SELECT *
              FROM branch_inventory_queue
             WHERE id = ?
            """,
            (int(queue_id),),
        ).fetchone()
        return None if row is None else dict(row)

    def assign_branch_inventory_queue_item(
        self,
        queue_id: int,
        warehouse_no: int,
        package_no: int,
        note: str = "",
    ) -> None:
        item = self.get_branch_inventory_queue_item(queue_id)
        if not item:
            raise ValueError("العنصر المطلوب غير موجود.")
        if (item.get("status") or "").upper() != "PENDING":
            raise ValueError("تمت معالجة هذا العنصر بالفعل.")

        w = int(warehouse_no)
        p = int(package_no)
        if w < 1 or p < 1:
            raise ValueError("رقم المخزن والعبوة يجب أن يكونا 1 أو أكبر.")

        with self.conn:
            self.add_stock(
                item["item_type"],
                item["school"],
                item["color"],
                item["size"],
                w,
                p,
                float(item.get("unit_price") or 0),
                int(item["qty"]),
                has_badge=int(item.get("has_badge") or 0),
            )
            self.conn.execute(
                """
                UPDATE branch_inventory_queue
                   SET status = 'ASSIGNED',
                       processed_at = ?,
                       processed_note = ?,
                       processed_warehouse_no = ?,
                       processed_package_no = ?
                 WHERE id = ?
                """,
                (
                    now_iso(),
                    (note or "").strip() or f"Assigned to warehouse {w} / package {p}",
                    w,
                    p,
                    int(queue_id),
                ),
            )

    def discard_branch_inventory_queue_item(self, queue_id: int, note: str = "") -> None:
        item = self.get_branch_inventory_queue_item(queue_id)
        if not item:
            raise ValueError("العنصر المطلوب غير موجود.")
        if (item.get("status") or "").upper() != "PENDING":
            raise ValueError("تمت معالجة هذا العنصر بالفعل.")

        with self.conn:
            self.conn.execute(
                """
                UPDATE branch_inventory_queue
                   SET status = 'DISCARDED',
                       processed_at = ?,
                       processed_note = ?
                 WHERE id = ?
                """,
                (
                    now_iso(),
                    (note or "").strip() or "Discarded as defective",
                    int(queue_id),
                ),
            )

    def reroute_branch_inventory_queue_item(
        self,
        queue_id: int,
        target_name: str,
        note: str = "",
    ) -> None:
        item = self.get_branch_inventory_queue_item(queue_id)
        if not item:
            raise ValueError("العنصر المطلوب غير موجود.")
        if (item.get("status") or "").upper() != "PENDING":
            raise ValueError("تمت معالجة هذا العنصر بالفعل.")

        target = (target_name or "").strip()
        if not target:
            raise ValueError("اسم الفرع المطلوب غير صالح.")
        if target == item.get("source_device"):
            raise ValueError("لا يمكن إعادة التوجيه إلى نفس الفرع المرسل.")

        with self.conn:
            self._record_branch_shipment_event(
                shipment_uuid=item.get("external_ref") or item.get("sync_event_uuid") or "",
                target_name=target,
                note=(note or "").strip() or f"rerouted from {item.get('source_device')}",
                lines=[{
                    "item_type": item["item_type"],
                    "school": item["school"],
                    "color": item["color"],
                    "size": item["size"],
                    "unit_price": float(item.get("unit_price") or 0),
                    "qty": int(item["qty"]),
                }],
            )
            self.conn.execute(
                """
                UPDATE branch_inventory_queue
                   SET status = 'REROUTED',
                       processed_at = ?,
                       processed_note = ?,
                       rerouted_target_device = ?
                 WHERE id = ?
                """,
                (
                    now_iso(),
                    (note or "").strip() or f"Rerouted to {target}",
                    target,
                    int(queue_id),
                ),
            )

    def _branch_shipment_size_profiles(self, lines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        keys = {
            (
                str(line.get("item_type") or "").strip(),
                str(line.get("school") or "").strip(),
                str(line.get("color") or "").strip(),
            )
            for line in (lines or [])
            if str(line.get("item_type") or "").strip()
            and str(line.get("school") or "").strip()
            and str(line.get("color") or "").strip()
        }
        profiles: List[Dict[str, Any]] = []
        for item_type, school, color in sorted(keys):
            profile = self.get_size_profile(item_type, school, color)
            if not profile:
                continue
            profiles.append({
                "item_type": item_type,
                "school": school,
                "color": color,
                "num_start_1": profile[0],
                "num_end_1": profile[1],
                "num_start_2": profile[2],
                "num_end_2": profile[3],
                "has_alpha": int(profile[4] or 0),
            })
        return profiles

    def _record_branch_shipment_event(
        self,
        shipment_uuid: str,
        target_name: str,
        note: str,
        lines: List[Dict[str, Any]],
    ) -> None:
        """Emit STOCK_TRANSFER_OUT scoped to the target POS.

        Separate helper so the payload shape lives in one place and is
        trivial to extend (e.g. add catalog hints) without touching
        create_branch_shipment.
        """
        import json as _json
        import sqlite3 as _sqlite3

        try:
            from sync_core import new_uuid as _new_uuid
        except Exception:
            _new_uuid = None

        def _fallback_uuid() -> str:
            import uuid as _u
            return str(_u.uuid4())

        device_name = None
        try:
            ident = self.conn.execute(
                "SELECT device_name FROM device_identity WHERE id = 1"
            ).fetchone()
            if ident is not None:
                device_name = ident[0]
        except Exception:
            pass

        known = []
        try:
            known = self.list_known_pos_device_names() or []
        except Exception:
            known = []
        canonical_target = canonical_branch_device_name(target_name, known) or (target_name or "").strip()
        if not canonical_target:
            raise ValueError("اسم الفرع المستهدف غير صالح.")

        payload = {
            "shipment_uuid":   shipment_uuid,
            "from_device":     device_name or "WAREHOUSE-MAIN",
            "note":             note or "",
            "items":           lines,
            "size_profiles":   self._branch_shipment_size_profiles(lines),
        }
        target_scope = f"pos:{canonical_target}"

        event_uuid = (_new_uuid() if _new_uuid else _fallback_uuid())
        try:
            self.conn.execute(
                """
                INSERT INTO sync_outbox
                    (event_uuid, event_type, payload_json,
                     created_at, status, attempts, target_scope)
                VALUES (?, ?, ?, ?, 'pending', 0, ?)
                """,
                (
                    event_uuid,
                    "STOCK_TRANSFER_OUT",
                    _json.dumps(payload, ensure_ascii=False, default=str),
                    now_iso(),
                    target_scope,
                ),
            )
        except _sqlite3.OperationalError:
            # Legacy outbox schema (no target_scope column) — fall back
            # to storing the scope inside the payload. The client-side
            # push loop handles both shapes.
            payload["__target_scope__"] = target_scope
            self.conn.execute(
                """
                INSERT INTO sync_outbox
                    (event_uuid, event_type, payload_json,
                     created_at, status, attempts)
                VALUES (?, ?, ?, ?, 'pending', 0)
                """,
                (
                    event_uuid,
                    "STOCK_TRANSFER_OUT",
                    _json.dumps(payload, ensure_ascii=False, default=str),
                    now_iso(),
                ),
            )

    # -------- Stock Audit --------

    def apply_stock_adjustments(self, adjustments: List[Dict[str, Any]],
                                note: str = "Physical count adjustment") -> int:
        """Apply stock count adjustments from a physical audit."""
        count = 0
        applied_events: List[Dict[str, Any]] = []
        with self.conn:
            for adj in adjustments:
                stock_id = int(adj["stock_id"])
                expected = int(adj["expected"])
                actual = int(adj["actual"])
                diff = actual - expected
                if diff == 0:
                    continue
                cur = self.conn.execute("SELECT * FROM stocks WHERE id=?", (stock_id,))
                s = cur.fetchone()
                if not s:
                    continue
                direction = "ADJUST_IN" if diff > 0 else "ADJUST_OUT"
                abs_diff = abs(diff)
                if diff > 0:
                    self.conn.execute(
                        "UPDATE stocks SET count = count + ? WHERE id = ?",
                        (abs_diff, stock_id))
                else:
                    self.conn.execute(
                        "UPDATE stocks SET count = count - ? WHERE id = ?",
                        (abs_diff, stock_id))
                self.conn.execute(
                    """INSERT INTO movements
                    (ts,direction,stock_id,qty,note,bill_id,
                     item_type,school,color,size,warehouse_no,package_no,
                     unit_price,has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), direction, stock_id, abs_diff,
                     f"{note} (expected {expected}, actual {actual})",
                     None, s["item_type"], s["school"], s["color"],
                     s["size"], int(s["warehouse_no"]),
                     int(s["package_no"]), float(s["unit_price"]),
                     int(s["has_badge"] or 0)))
                applied_events.append({
                    "stock_uuid":   s["uuid"] if "uuid" in s.keys() else None,
                    "stock_id":     stock_id,
                    "direction":    direction,
                    "qty":          abs_diff,
                    "expected":     expected,
                    "actual":       actual,
                    "item_type":    s["item_type"],
                    "school":       s["school"],
                    "color":        s["color"],
                    "size":         s["size"],
                    "warehouse_no": int(s["warehouse_no"]),
                    "package_no":   int(s["package_no"]),
                    "unit_price":   float(s["unit_price"]),
                    "has_badge":    int(s["has_badge"] or 0),
                })
                count += 1
            if applied_events:
                self._record_sync_event("STOCK_AUDIT_APPLIED", {
                    "note":         note,
                    "applied_count": count,
                    "adjustments":  applied_events,
                })
        return count

    # -------- Excel export for inventory --------
    def export_inventory_excel(self, path: str, rows: Sequence[Dict[str, Any]]) -> None:
        headers = [
            "id","item_type","school","color","size","warehouse_no","package_no","has_badge","unit_price","count","value",
        ]
        table = []
        for r in rows:
            value = r.get("value", float(r["unit_price"]) * int(r["count"]))
            table.append(
                [
                    r["id"], r["item_type"], r["school"], r["color"], r["size"],
                    r["warehouse_no"], r["package_no"],
                    ("✓" if int(r.get("has_badge") or 0) else ""),
                    float(r["unit_price"]),
                    int(r["count"]), float(value),
                ]
            )
        export_to_excel(path, headers, table)

    def get_app_setting(self, key: str, default: Optional[str] = None) -> Optional[str]:
        row = self.conn.execute(
            "SELECT value FROM app_settings WHERE key = ?",
            (str(key),),
        ).fetchone()
        if row is None:
            return default
        return str(row["value"])

    def set_app_setting(self, key: str, value: Any) -> None:
        with self.conn:
            self.conn.execute(
                "INSERT INTO app_settings(key, value) VALUES(?, ?) "
                "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
                (str(key), str(value)),
            )

    def verify_admin_password(self, plain: str) -> bool:
        stored = self.get_app_setting("admin_password", ADMIN_PASSWORD_PLAIN) or ADMIN_PASSWORD_PLAIN
        if str(stored).startswith(ADMIN_PASSWORD_HASH_PREFIX):
            digest = hashlib.sha256(str(plain).encode("utf-8")).hexdigest()
            return str(stored) == f"{ADMIN_PASSWORD_HASH_PREFIX}{digest}"
        ok = str(plain) == str(stored)
        if ok:
            self.set_admin_password(str(plain))
        return ok

    def set_admin_password(self, plain: str) -> None:
        digest = hashlib.sha256(str(plain).encode("utf-8")).hexdigest()
        self.set_app_setting("admin_password", f"{ADMIN_PASSWORD_HASH_PREFIX}{digest}")

    def remove_from_stock(self, stock_id: int, qty: Optional[int], note: str = "Admin remove") -> int:
        with self.conn:
            cur = self.conn.execute("SELECT * FROM stocks WHERE id=?", (int(stock_id),))
            s = cur.fetchone()
            if not s:
                raise ValueError("Stock row not found.")
            current = int(s["count"])
            if current <= 0:
                raise ValueError("No quantity available to remove.")
            take = current if qty is None else min(int(qty), current)
            if take <= 0:
                raise ValueError("Quantity must be > 0")
            self.conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?", (int(take), int(stock_id)))
            self.conn.execute(
                """INSERT INTO movements
                   (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    now_iso(), "ADJUST_OUT", int(stock_id), int(take), note, None,
                    s["item_type"], s["school"], s["color"], s["size"],
                    int(s["warehouse_no"]), int(s["package_no"]), float(s["unit_price"]),
                    int(s["has_badge"] or 0),
                ),
            )
            self._record_sync_event("STOCK_ADJUST", {
                "stock_uuid":   s["uuid"] if "uuid" in s.keys() else None,
                "stock_id":     int(stock_id),
                "direction":    "OUT",
                "qty":          int(take),
                "note":         note,
                "item_type":    s["item_type"],
                "school":       s["school"],
                "color":        s["color"],
                "size":         s["size"],
                "warehouse_no": int(s["warehouse_no"]),
                "package_no":   int(s["package_no"]),
                "unit_price":   float(s["unit_price"]),
                "has_badge":    int(s["has_badge"] or 0),
            })
            self.auto_reopen_package_if_empty(
                int(s["warehouse_no"]),
                int(s["package_no"])
            )
            return int(take)

    def list_movements(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        where = ["1=1"]
        args: List[Any] = []
        if filters.get("direction"):
            where.append("direction = ?")
            args.append(filters["direction"])
        for k in ("item_type", "school", "color", "size"):
            v = (filters.get(k) or "").strip()
            if v:
                where.append(f"LOWER(TRIM(m.{k})) = LOWER(TRIM(?))")
                args.append(v)

        # NEW: filter by customer via bills table
        cust = (filters.get("customer") or "").strip()
        if cust:
            where.append("bill_id IN (SELECT id FROM bills WHERE LOWER(TRIM(customer)) = LOWER(TRIM(?)))")
            args.append(cust)

        # free-text search across common fields
        txt = (filters.get("text") or "").strip()
        if txt:
            like = f"%{txt}%"
            where.append("""(
                LOWER(COALESCE(note,''))      LIKE LOWER(?)
            OR LOWER(COALESCE(item_type,'')) LIKE LOWER(?)
            OR LOWER(COALESCE(school,''))    LIKE LOWER(?)
            OR LOWER(COALESCE(color,''))     LIKE LOWER(?)
            OR LOWER(COALESCE(size,''))      LIKE LOWER(?)
            OR LOWER(COALESCE(b.customer,'')) LIKE LOWER(?)
            )""")
            args += [like, like, like, like, like, like]

        # date range (YYYY-MM-DD)
        df = (filters.get("date_from") or "").strip()
        dt = (filters.get("date_to")   or "").strip()
        if df:
            where.append("date(ts) >= date(?)")
            args.append(df)
        if dt:
            where.append("date(ts) <= date(?)")
            args.append(dt)

        cur = self.conn.cursor()
        cur.execute(f"""
            SELECT m.id, m.ts, m.direction, m.stock_id, m.qty, m.note, m.bill_id,
                m.item_type, m.school, m.color, m.size, m.warehouse_no, m.package_no,
                m.unit_price, m.has_badge, COALESCE(b.customer,'') AS customer
            FROM movements m
            LEFT JOIN bills b ON b.id = m.bill_id
            WHERE {' AND '.join(where)}
            ORDER BY m.ts DESC, m.id DESC
        """, args)
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows

    def list_movement_item_totals(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Aggregate movement/stock metrics per item specs across warehouse + branches."""
        def _date_clause(col: str) -> Tuple[str, List[Any]]:
            parts: List[str] = []
            args: List[Any] = []
            df = (filters.get("date_from") or "").strip()
            dt = (filters.get("date_to") or "").strip()
            if df:
                parts.append(f"date({col}) >= date(?)")
                args.append(df)
            if dt:
                parts.append(f"date({col}) <= date(?)")
                args.append(dt)
            return (" AND " + " AND ".join(parts), args) if parts else ("", [])

        def _to_key(row: Any) -> Tuple[str, str, str, str]:
            return (
                _normalize_spec_label(row[0]),
                _normalize_spec_label(row[1]),
                _normalize_spec_label(row[2]),
                _normalize_size_label(str(row[3] or "")),
            )

        def _spec_match(key: Tuple[str, str, str, str]) -> bool:
            checks = (
                ("item_type", key[0]),
                ("school", key[1]),
                ("color", key[2]),
                ("size", key[3]),
            )
            for fld, actual in checks:
                want = str(filters.get(fld) or "").strip()
                if fld == "size":
                    want = _normalize_size_label(want)
                else:
                    want = _normalize_spec_label(want)
                if want and actual.lower() != want.lower():
                    return False
            txt = str(filters.get("text") or "").strip().lower()
            if txt:
                hay = " ".join(key).lower()
                if txt not in hay:
                    return False
            return True

        cur = self.conn.cursor()
        metrics: Dict[Tuple[str, str, str, str], Dict[str, int]] = {}

        def _bucket(key: Tuple[str, str, str, str]) -> Dict[str, int]:
            if key not in metrics:
                metrics[key] = {
                    "incoming_qty": 0,
                    "sold_warehouse_qty": 0,
                    "branch_shipped_qty": 0,
                    "branch_sold_qty_est": 0,
                    "reserved_qty": 0,
                    "warehouse_qty": 0,
                    "branch_qty": 0,
                    "remaining_total_qty": 0,
                    "sold_total_qty": 0,
                }
            return metrics[key]

        # Current warehouse stock
        for r in cur.execute(
            """
            SELECT item_type, school, color, size, COALESCE(SUM(count),0)
            FROM stocks
            WHERE count > 0
            GROUP BY item_type, school, color, size
            """
        ).fetchall():
            key = _to_key(r)
            _bucket(key)["warehouse_qty"] = int(r[4] or 0)

        # Current branch stock mirror (deduplicated by logical branch):
        # We choose ONE latest snapshot source per branch (canonical device),
        # then aggregate mirror rows only from those selected sources.
        # This avoids doubling when the same branch exists under both
        # device_name and device_uuid in mirror history.
        branch_spec_totals: Dict[Tuple[str, str, str, str], int] = {}
        try:
            snaps = cur.execute(
                """
                SELECT
                    pm.source_device,
                    pm.snapshot_at,
                    COALESCE(kd.device_name, pm.source_device) AS canonical_branch
                FROM pos_stocks_snapshot_meta pm
                LEFT JOIN known_devices kd
                    ON kd.device_name = pm.source_device
                    OR kd.device_uuid = pm.source_device
                """
            ).fetchall()
            approved_branches = set(DEFAULT_BRANCH_POS_NAMES)
            latest_source_by_branch: Dict[str, Tuple[str, str]] = {}
            for rr in snaps:
                src = str(rr[0] or "").strip()
                snap_at = str(rr[1] or "").strip()
                branch = str(rr[2] or "").strip()
                if not src or not branch:
                    continue
                # Ignore unknown/test devices (e.g. POS-01) so movement monitor
                # reflects only configured real branches.
                if branch not in approved_branches:
                    continue
                prev = latest_source_by_branch.get(branch)
                if prev is None or snap_at > prev[1]:
                    latest_source_by_branch[branch] = (src, snap_at)

            selected_sources = [v[0] for v in latest_source_by_branch.values()]
            if selected_sources:
                ph = ",".join("?" for _ in selected_sources)
                raw_rows = cur.execute(
                    f"""
                    SELECT item_type, school, color, size, COALESCE(SUM(count),0) AS qty
                    FROM pos_stocks_mirror
                    WHERE source_device IN ({ph})
                    GROUP BY item_type, school, color, size
                    """,
                    tuple(selected_sources),
                ).fetchall()
                for rr in raw_rows:
                    key4 = _to_key(rr)
                    branch_spec_totals[key4] = int(rr[4] or 0)
        except sqlite3.OperationalError:
            branch_spec_totals = {}
        for key, qty in branch_spec_totals.items():
            _bucket(key)["branch_qty"] = int(qty or 0)

        # Reserved qty mirrored from branches (active only)
        try:
            rows = cur.execute(
                """
                SELECT item_type, school, color, size, COALESCE(SUM(qty),0)
                FROM pos_reservations_mirror
                WHERE status = 'معلق'
                GROUP BY item_type, school, color, size
                """
            ).fetchall()
        except sqlite3.OperationalError:
            rows = []
        for r in rows:
            key = _to_key(r)
            _bucket(key)["reserved_qty"] = int(r[4] or 0)

        # Warehouse incoming flow (period-filtered):
        # - normal stock incoming movements (IN)
        # - factory-direct sold lines (OUT_FACTORY, non-branch customer)
        #   are counted as virtual incoming so per-item math stays coherent
        #   in the movement monitor (incoming/sold/remaining).
        d_sql, d_args = _date_clause("m.ts")
        rows = cur.execute(
            f"""
            SELECT m.item_type, m.school, m.color, m.size, COALESCE(SUM(m.qty),0)
            FROM movements m
            LEFT JOIN bills b ON b.id = m.bill_id
            WHERE (
                    m.direction = 'IN'
                 OR (
                        m.direction = 'OUT_FACTORY'
                    AND (
                            b.id IS NULL
                         OR LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('فرع:%'))
                        )
                    )
                  )
            {d_sql}
            GROUP BY m.item_type, m.school, m.color, m.size
            """,
            d_args,
        ).fetchall()
        for r in rows:
            key = _to_key(r)
            _bucket(key)["incoming_qty"] = int(r[4] or 0)

        # Warehouse sold (exclude shipment bills whose customer is a branch)
        d_sql, d_args = _date_clause("m.ts")
        rows = cur.execute(
            f"""
            SELECT m.item_type, m.school, m.color, m.size, COALESCE(SUM(m.qty),0)
            FROM movements m
            LEFT JOIN bills b ON b.id = m.bill_id
            WHERE m.direction IN ('OUT', 'OUT_FACTORY')
              AND (
                    b.id IS NULL
                 OR LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('فرع:%'))
              )
            {d_sql}
            GROUP BY m.item_type, m.school, m.color, m.size
            """,
            d_args,
        ).fetchall()
        for r in rows:
            key = _to_key(r)
            _bucket(key)["sold_warehouse_qty"] = int(r[4] or 0)

        # Shipped to branches from warehouse bills (period-filtered on bill date)
        d_sql, d_args = _date_clause("b.created_at")
        rows = cur.execute(
            f"""
            SELECT bi.item_type, bi.school, bi.color, bi.size, COALESCE(SUM(bi.qty),0)
            FROM bill_items bi
            JOIN bills b ON b.id = bi.bill_id
            WHERE COALESCE(b.status,'CONFIRMED') = 'CONFIRMED'
              AND LOWER(TRIM(COALESCE(b.customer,''))) LIKE LOWER(TRIM('فرع:%'))
            {d_sql}
            GROUP BY bi.item_type, bi.school, bi.color, bi.size
            """,
            d_args,
        ).fetchall()
        for r in rows:
            key = _to_key(r)
            _bucket(key)["branch_shipped_qty"] = int(r[4] or 0)

        cur.close()

        out: List[Dict[str, Any]] = []
        for key, m in metrics.items():
            if not _spec_match(key):
                continue
            branch_sold = max(0, int(m["branch_shipped_qty"]) - int(m["branch_qty"]))
            sold_total = int(m["sold_warehouse_qty"]) + branch_sold
            remaining_total = int(m["warehouse_qty"]) + int(m["branch_qty"])
            out.append(
                {
                    "item_type": key[0],
                    "school": key[1],
                    "color": key[2],
                    "size": key[3],
                    "incoming_qty": int(m["incoming_qty"]),
                    "sold_branch_qty": int(branch_sold),
                    "sold_warehouse_qty": int(m["sold_warehouse_qty"]),
                    "sold_total_qty": int(sold_total),
                    "reserved_qty": int(m["reserved_qty"]),
                    "remaining_branch_qty": int(m["branch_qty"]),
                    "remaining_warehouse_qty": int(m["warehouse_qty"]),
                    "remaining_total_qty": int(remaining_total),
                }
            )

        out.sort(
            key=lambda r: (
                (r.get("item_type") or "").lower(),
                (r.get("school") or "").lower(),
                (r.get("color") or "").lower(),
                (r.get("size") or "").lower(),
            )
        )
        return out



# ------------------- UI helpers -------------------
class LabeledStaticCombo(ttk.Frame):
    def __init__(self, master, text: str, values: List[str], value_map: Optional[Dict[str, str]] = None, **kwargs):
        super().__init__(master)
        ttk.Label(self, text=text).pack(side=tk.TOP, anchor="w")
        self.var = tk.StringVar()
        self._value_map = dict(value_map or {})
        self._reverse_value_map = {v: k for k, v in self._value_map.items()}
        self.cb = ttk.Combobox(self, textvariable=self.var, values=values, state="readonly", **kwargs)
        self.cb.pack(fill=tk.X)

    def get(self) -> str:
        raw = (self.var.get() or "").strip()
        return self._reverse_value_map.get(raw, raw)

    def set(self, v: str) -> None:
        raw = (v or "").strip()
        self.var.set(self._value_map.get(raw, raw))

class LabeledCombobox(ttk.Frame):
    """
    Entry + caret button + Listbox popup (no focus theft).
    - Button “▼” toggles popup; Entry keeps focus.
    - Opens on first typed letter; live filters while typing.
    - Emits <<ComboboxSelected>> when a value is picked (keeps existing app binds working).
    - Public API: get(), set(), .cb (Entry), .btn (Button).
    """
    FILTER_IDLE_MS = 60
    MATCH_MODE = "startswith"     # or "contains"
    MIN_CHARS_TO_OPEN = 1
    OPEN_ALL_ON_BUTTON = True     # caret opens full list even if field empty

    def __init__(self, master, text: str, db: 'SqliteDatabase', field: str, **kwargs):
        super().__init__(master)
        self.db = db
        self.field = field

        ttk.Label(self, text=text).pack(side=tk.TOP, anchor="w")

        # --- Input row: Entry + tiny caret button ---
        row = ttk.Frame(self)
        row.pack(fill=tk.X)

        self.var = tk.StringVar()
        self.cb = ttk.Entry(row, textvariable=self.var, **kwargs)
        self.cb.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Small caret button; no focus grab
        self.btn = ttk.Button(row, text="▼", width=2, takefocus=0, command=self._on_caret_click)
        # make it flat-ish if theme supports it
        try:
            self.btn.configure(style="Toolbutton")
        except Exception:
            pass
        self.btn.pack(side=tk.RIGHT, padx=(4, 0))

        # State
        self._all_values: List[str] = []
        self._supplier = None  # NEW: callable that returns List[str]
        self._debounce_job: Optional[str] = None
        self._suspend_set = False
        self._opened = False
        self._mouse_inside_popup = False

        # Popup (lazy)
        self._popup: Optional[tk.Toplevel] = None
        self._list: Optional[tk.Listbox] = None
        self._ysb: Optional[ttk.Scrollbar] = None

        # Events (Entry)
        self.cb.bind("<FocusIn>", self._on_focus, add="+")
        self.cb.bind("<KeyRelease>", self._on_key_release, add="+")
        self.cb.bind("<Down>", self._on_down, add="+")
        self.cb.bind("<Escape>", self._on_escape, add="+")
        self.cb.bind("<FocusOut>", self._on_focus_out, add="+")

        # Keep popup aligned on window move/resize
        self.winfo_toplevel().bind("<Configure>", self._reposition_popup_safely, add="+")

 # ---------- Public API ----------
    def get(self) -> str:
        return (self.var.get() or "").strip()

    def set(self, v: str) -> None:
        self._suspend_set = True
        try:
            self.var.set(v or "")
        finally:
            self._suspend_set = False
        self._close_popup()

    def set_supplier(self, fn):  # NEW
        """fn: () -> List[str]"""
        self._supplier = fn
        self.refresh_values()

    def refresh_values(self):  # NEW
        """Force re-pull options and refresh open popup if needed."""
        self._ensure_data(force=True)
        if self._opened:
            self._fill_list(self._filtered(self.get()))

    # ---------- Data ----------
    def _ensure_data(self, force: bool = False) -> None:
        if self._supplier:
            try:
                vals = self._supplier() or []
            except Exception:
                vals = []
            # refresh if first time or changed or explicitly forced
            if force or (vals != self._all_values):
                self._all_values = vals[:]
            return

        if not self._all_values or force:
            try:
                # fallback ONLY if no supplier is attached
                self._all_values = self.db.get_distinct_filtered(
                    self.field,
                    {}
                ) or []
            except Exception:
                self._all_values = []


    # ---------- Popup lifecycle ----------
    def _create_popup(self) -> None:
        if self._popup and self._popup.winfo_exists():
            return
        self._popup = tk.Toplevel(self)
        self._popup.overrideredirect(True)
        self._popup.attributes("-topmost", True)
        self._popup.withdraw()

        frm = ttk.Frame(self._popup, borderwidth=1, relief="solid")
        frm.pack(fill=tk.BOTH, expand=True)

        self._list = tk.Listbox(frm, activestyle="none")
        self._list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._ysb = ttk.Scrollbar(frm, orient="vertical", command=self._list.yview)
        self._ysb.pack(side=tk.RIGHT, fill=tk.Y)
        self._list.configure(yscrollcommand=self._ysb.set)

        # Mouse interactions
        self._list.bind("<ButtonRelease-1>", self._on_list_click, add="+")
        self._list.bind("<Motion>", lambda e: self._set_active_under_mouse(e), add="+")
        self._popup.bind("<Enter>", lambda *_: self._mark_mouse(True), add="+")
        self._popup.bind("<Leave>", lambda *_: self._mark_mouse(False), add="+")

    def _open_popup(self, *, items: Optional[List[str]] = None) -> None:
        if self._opened and items is None:
            return
        self._create_popup()
        if items is None:
            items = self._filtered(self.get())
        self._fill_list(items or [])
        if not self._list or self._list.size() == 0:
            self._close_popup()
            return

        self._position_popup()
        self._popup.deiconify()
        self._opened = True
        self.after_idle(self._restore_entry_focus)

    def _close_popup(self) -> None:
        if self._popup and self._popup.winfo_exists():
            try:
                self._popup.withdraw()
            except Exception:
                pass
        self._opened = False

    def _position_popup(self) -> None:
        if not (self._popup and self._popup.winfo_exists()):
            return
        try:
            x = self.cb.winfo_rootx()
            y = self.cb.winfo_rooty() + self.cb.winfo_height()
            w = self.cb.winfo_width() + self.btn.winfo_width() + 4  # span both widgets
            rows = min(8, max(1, self._list.size() if self._list else 0))
            row_h = max(18, int(self.cb.winfo_fpixels("1.2m")))
            h = rows * row_h + 2
            self._popup.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _reposition_popup_safely(self, *_):
        if self._opened:
            self.after_idle(self._position_popup)

    # ---------- Filtering ----------
    def _filtered(self, text: str) -> List[str]:
        self._ensure_data()
        tl = (text or "").lower()
        if not tl:
            return self._all_values[:]
        if self.MATCH_MODE == "contains":
            return [v for v in self._all_values if tl in str(v).lower()]
        return [v for v in self._all_values if str(v).lower().startswith(tl)]

    def _fill_list(self, items: List[str]) -> None:
        if not self._list:
            return
        self._list.delete(0, tk.END)
        for v in items:
            self._list.insert(tk.END, v)
        if items:
            self._list.activate(0)
            self._list.selection_clear(0, tk.END)
        self._position_popup()

    # ---------- Entry events ----------
    def _on_focus(self, *_):
        self._ensure_data()

    def _on_key_release(self, event):
        if self._suspend_set:
            return
        if event.keysym in ("Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R",
                            "Left", "Right", "Home", "End", "Prior", "Next"):
            return
        if event.keysym == "Up":
            if self._opened and self._list and self._list.size() > 0:
                idx = self._list.index("active")
                idx = max(0, idx - 1)
                self._list.activate(idx); self._ensure_visible(idx)
            return "break"
        if event.keysym == "Down":
            return self._on_down(event)
        if event.keysym in ("Return", "KP_Enter"):
            if self._opened and self._list and self._list.size() > 0:
                self._commit_active_selection()
                return "break"
            return

        # Debounced filter
        if self._debounce_job:
            try: self.after_cancel(self._debounce_job)
            except Exception: pass
        self._debounce_job = self.after(self.FILTER_IDLE_MS, self._apply_filter)

    def _apply_filter(self):
        text = self.get()
        items = self._filtered(text)
        if items:
            if self._opened:
                self._fill_list(items)
            else:
                if len(text) >= self.MIN_CHARS_TO_OPEN:
                    self._open_popup(items=items)
        else:
            self._close_popup()

    def _on_down(self, _event):
        if not self._opened:
            # open even if empty filter (shows all)
            items = self._filtered(self.get())
            if not items and self.OPEN_ALL_ON_BUTTON:
                self._ensure_data()
                items = self._all_values[:]
            self._open_popup(items=items)
            return "break"
        if self._list and self._list.size() > 0:
            idx = self._list.index("active")
            idx = min(self._list.size()-1, max(0, idx + 1))
            self._list.activate(idx); self._ensure_visible(idx)
        return "break"

    def _on_escape(self, *_):
        self._close_popup()

    def _on_focus_out(self, *_):
        # Close only if focus truly left (not going into popup)
        self.after(20, self._maybe_close_after_blur)

    def _maybe_close_after_blur(self):
        if self._opened and not self._mouse_inside_popup and self.focus_get() is not self.cb:
            self._close_popup()

    def _restore_entry_focus(self):
        try:
            self.cb.focus_set()
            insert_pos = self.cb.index("insert")
            self.cb.icursor(insert_pos)
            try:
                self.cb.selection_clear()
            except TypeError:
                self.cb.selection_clear(0, tk.END)
        except Exception:
            pass

    # ---------- Caret button ----------
    def _on_caret_click(self):
        # Toggle popup. If opening and empty input, show full list.
        if self._opened:
            self._close_popup()
        else:
            if self.OPEN_ALL_ON_BUTTON and not self.get():
                self._ensure_data()
                self._open_popup(items=self._all_values[:])
            else:
                self._open_popup()
        # ensure Entry keeps focus for continuous typing
        self.after_idle(self._restore_entry_focus)

    # ---------- Listbox interactions ----------
    def _on_list_click(self, *_):
        self._commit_active_selection()

    def _commit_active_selection(self):
        if not (self._list and self._list.size() > 0):
            return
        idx = self._list.index("active")
        if 0 <= idx < self._list.size():
            val = self._list.get(idx)
            self._suspend_set = True
            try:
                self.var.set(val)
            finally:
                self._suspend_set = False
            self._close_popup()
            try:
                self.cb.event_generate("<<ComboboxSelected>>")
            except Exception:
                pass
            self.after_idle(self._restore_entry_focus)

    def _set_active_under_mouse(self, event):
        if not self._list:
            return
        try:
            idx = self._list.nearest(event.y)
            if 0 <= idx < self._list.size():
                self._list.activate(idx)
        except Exception:
            pass

    def _ensure_visible(self, idx: int):
        if not self._list:
            return
        top = self._list.nearest(0)
        bot = self._list.nearest(self._list.winfo_height())
        if idx <= top:
            self._list.see(max(0, idx-1))
        elif idx >= bot:
            self._list.see(min(self._list.size()-1, idx+1))

    def _mark_mouse(self, inside: bool):
        self._mouse_inside_popup = inside

class LabeledEntry(ttk.Frame):
    def __init__(self, master, text: str, **kwargs):
        super().__init__(master)
        ttk.Label(self, text=text).pack(side=tk.TOP, anchor="w")
        self.var = tk.StringVar()
        self.entry = ttk.Entry(self, textvariable=self.var, **kwargs)
        self.entry.pack(fill=tk.X)

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, v: str) -> None:
        self.var.set(v)

class DateField(ttk.Frame):
    """Entry + '📅' button; popup calendar; returns YYYY-MM-DD; empty allowed."""
    def __init__(self, master, label: str):
        super().__init__(master)
        ttk.Label(self, text=label).pack(side=tk.TOP, anchor="w")
        row = ttk.Frame(self); row.pack(fill=tk.X)
        self.var = tk.StringVar()
        self.entry = ttk.Entry(row, textvariable=self.var, width=14)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn = ttk.Button(row, text="📅", width=3, command=self._open)
        self.btn.pack(side=tk.LEFT, padx=(4,0))
        self._popup: Optional[tk.Toplevel] = None
        self.entry.bind("<Return>", lambda e: self._validate())
        self.entry.bind("<Escape>", lambda e: self._close())
        self.entry.bind("<FocusOut>", lambda e: self._close())

    def _open(self):
        self._close()
        tp = tk.Toplevel(self)
        self._popup = tp
        tp.wm_overrideredirect(True)
        tp.attributes("-topmost", True)
        x = self.entry.winfo_rootx(); y = self.entry.winfo_rooty() + self.entry.winfo_height()
        tp.geometry(f"+{x}+{y}")
        tp.bind("<FocusOut>", lambda e: self._close())
        tp.bind("<Escape>", lambda e: self._close())
        try:
            tp.focus_force()   # ensure focus so FocusOut reliably fires
        except Exception:
            pass


        frm = ttk.Frame(tp, padding=6, borderwidth=1, relief="solid"); frm.pack()
        today = date.today()
        try:
            base = datetime.strptime(self.var.get(), "%Y-%m-%d").date()
        except Exception:
            base = today
        self._year = tk.IntVar(value=base.year)
        self._month = tk.IntVar(value=base.month)

        hdr = ttk.Frame(frm); hdr.pack(fill=tk.X)
        ttk.Button(hdr, text="◀", width=2, command=self._prev_month).pack(side=tk.LEFT)
        self._title = ttk.Label(hdr, font=("", 10, "bold")); self._title.pack(side=tk.LEFT, padx=6)
        ttk.Button(hdr, text="▶", width=2, command=self._next_month).pack(side=tk.RIGHT)

        self._grid = ttk.Frame(frm); self._grid.pack()
        self._render()

    def _render(self):
        for w in self._grid.winfo_children(): w.destroy()
        y, m = self._year.get(), self._month.get()
        self._title.config(text=f"{calendar.month_name[m]} {y}")
        hdr = ttk.Frame(self._grid); hdr.grid(row=0, column=0, columnspan=7, pady=(0,4))
        for i, wk in enumerate(["Mo","Tu","We","Th","Fr","Sa","Su"]):
            ttk.Label(hdr, text=wk, width=3, anchor="center").grid(row=0, column=i)
        for r, week in enumerate(calendar.Calendar(firstweekday=0).monthdayscalendar(y, m), start=1):
            for c, d_ in enumerate(week):
                if d_ == 0:
                    ttk.Label(self._grid, text="   ", width=3).grid(row=r, column=c)
                else:
                    ttk.Button(self._grid, text=f"{d_:02d}", width=3,
                               command=lambda dd=d_: self._pick(y, m, dd)).grid(row=r, column=c, padx=1, pady=1)

    def _prev_month(self):
        m, y = self._month.get() - 1, self._year.get()
        if m < 1: m, y = 12, y - 1
        self._month.set(m); self._year.set(y); self._render()

    def _next_month(self):
        m, y = self._month.get() + 1, self._year.get()
        if m > 12: m, y = 1, y + 1
        self._month.set(m); self._year.set(y); self._render()

    def _pick(self, y: int, m: int, d_: int):
        self.var.set(f"{y:04d}-{m:02d}-{d_:02d}")
        self._close()

    def _validate(self):
        txt = self.var.get().strip()
        if not txt: return
        try:
            datetime.strptime(txt, "%Y-%m-%d")
        except Exception:
            messagebox.showerror("صيغة التاريخ", "الرجاء إدخال التاريخ بصيغة YYYY-MM-DD أو استخدم التقويم.", parent=self)
            self.entry.focus_set()

    def _close(self):
        if self._popup and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None

    def get(self) -> str:
        return self.var.get().strip()

    def set(self, v: str) -> None:
        self.var.set(v or "")
class SizesGrid(ttk.Frame):
    """
    - Uses the fixed SIZE_LABELS
    - Converts ONLY 4XL & 5XL into editable slots
    - Adds TWO EXTRA empty editable slots
    """

    COLS = 6

    def __init__(self, master, sizes: List[str], **kwargs):
        super().__init__(master, **kwargs)

        # Add two empty extra slots
        self.fixed_sizes = sizes[:]                     # original 22 sizes
        self.extra_sizes = ["__custom1__", "__custom2__"]
        self.sizes = self.fixed_sizes + self.extra_sizes

        # Determine which indices are editable
        self.idx_4xl = self.fixed_sizes.index("4XL") if "4XL" in self.fixed_sizes else None

        self.idx_5xl = self.fixed_sizes.index("5XL") if "5XL" in self.fixed_sizes else None


        self.editable_indices = set(filter(
            lambda x: x is not None,
            {
                self.idx_4xl,
                self.idx_5xl,
                len(self.fixed_sizes),
                len(self.fixed_sizes) + 1,
            }
        ))

        self._vars = {}

        for i, sz in enumerate(self.sizes):
            r = i // self.COLS
            c = i % self.COLS

            cell = ttk.Frame(self, padding=(4, 4))
            cell.grid(row=r, column=c, sticky="nsew", padx=2, pady=2)

            is_custom = i in self.editable_indices

            # --- size label / entry ---
            if is_custom:
                v_label = tk.StringVar(value="" if sz.startswith("__") else sz)
                ttk.Entry(cell, textvariable=v_label, width=7, justify="center").pack()
            else:
                v_label = tk.StringVar(value=sz)
                ttk.Label(cell, text=sz, anchor="center").pack()

            # qty
            ttk.Label(cell, text="الكمية", font=("", 8)).pack()
            v_qty = tk.StringVar()
            ttk.Entry(cell, textvariable=v_qty, width=6).pack()

            # price
            ttk.Label(cell, text="السعر", font=("", 8)).pack()
            v_price = tk.StringVar()
            ttk.Entry(cell, textvariable=v_price, width=8).pack()

            self._vars[sz] = (v_qty, v_price, v_label, is_custom)

        # resize columns equally
        for c in range(self.COLS):
            self.columnconfigure(c, weight=1)

    def get_rows(self):
        out = []
        for sz in self.sizes:
            v_qty, v_price, v_label, is_custom = self._vars[sz]

            size_label = v_label.get().strip() if is_custom else sz
            if not size_label:
                continue

            qty_txt = v_qty.get().strip()
            if not qty_txt:
                continue  # ← EMPTY instead of zero

            try:
                qty = int(qty_txt)
            except Exception:
                continue

            if qty <= 0:
                continue  # ← ZERO NOT PRINTED

            price_txt = v_price.get().strip()
            try:
                price = float(price_txt) if price_txt else None
            except Exception:
                price = None

            out.append({
                "size": size_label,
                "qty": qty,
                "price": price
            })

        return out

    def set_price_for_size(self, size_label: str, price):
        """Set the price entry for a specific size (only if currently empty)."""
        for sz, (v_qty, v_price, v_label, is_custom) in self._vars.items():
            actual_label = v_label.get().strip() if is_custom else sz
            if actual_label == size_label and not v_price.get().strip():
                try:
                    v_price.set(f"{float(price):.2f}")
                except Exception:
                    pass

    def set_price_for_all(self, price):
        """Set the price entry for ALL sizes (only where currently empty)."""
        for _, (v_qty, v_price, v_label, is_custom) in self._vars.items():
            if not v_price.get().strip():
                try:
                    v_price.set(f"{float(price):.2f}")
                except Exception:
                    pass

# ------------------- Scroll Wheel Utility -------------------

def _bind_mousewheel(widget, scrollable=None):
    """
    Enable hover-to-scroll on any scrollable tkinter widget.
    widget    – the visual element the user hovers over
    scrollable – the target that receives yview_scroll (defaults to widget)
    """
    target = scrollable or widget

    def _on_mousewheel(event):
        try:
            if hasattr(event, "delta") and event.delta:
                target.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif hasattr(event, "num"):
                if event.num == 4:
                    target.yview_scroll(-3, "units")
                elif event.num == 5:
                    target.yview_scroll(3, "units")
        except Exception:
            pass
        return "break"

    def _on_enter(_e=None):
        try:
            widget.bind_all("<MouseWheel>", _on_mousewheel)
            widget.bind_all("<Button-4>", _on_mousewheel)
            widget.bind_all("<Button-5>", _on_mousewheel)
        except Exception:
            pass

    def _on_leave(_e=None):
        try:
            widget.unbind_all("<MouseWheel>")
            widget.unbind_all("<Button-4>")
            widget.unbind_all("<Button-5>")
        except Exception:
            pass

    widget.bind("<Enter>", _on_enter, add="+")
    widget.bind("<Leave>", _on_leave, add="+")

# ------------------- UI Enhancement Helpers -------------------

# ---- Premium UI Design System ----
_UI = {
    "BG":       "#F8FAFC",
    "SURFACE":  "#FFFFFF",
    "SURFACE2": "#F1F5F9",
    "BORDER":   "#E2E8F0",
    "ACCENT":   "#0F172A",
    "ACCENT_H": "#1E293B",
    "BRAND":    "#3B82F6",
    "BRAND_H":  "#2563EB",
    "BRAND_L":  "#EFF6FF",
    "TEXT":     "#0F172A",
    "TEXT_SEC": "#475569",
    "TEXT_DIM": "#94A3B8",
    "OK":       "#059669",
    "OK_H":     "#047857",
    "OK_L":     "#ECFDF5",
    "WARN":     "#D97706",
    "WARN_H":   "#B45309",
    "WARN_L":   "#FFFBEB",
    "DANGER":   "#DC2626",
    "DANGER_H": "#B91C1C",
    "DANGER_L": "#FEF2F2",
    "ROW_EVEN": "#FFFFFF",
    "ROW_ODD":  "#F8FAFC",
    "SEL_BG":   "#DBEAFE",
    "SEL_FG":   "#1E40AF",
}

_FONTS = {
    "h1":      ("Segoe UI", 14, "bold"),
    "h2":      ("Segoe UI", 12, "bold"),
    "h3":      ("Segoe UI", 11, "bold"),
    "body":    ("Segoe UI", 9),
    "body_b":  ("Segoe UI", 9, "bold"),
    "small":   ("Segoe UI", 8),
    "caption": ("Segoe UI", 7),
    "big_num": ("Segoe UI", 20, "bold"),
    "price":   ("Segoe UI", 16, "bold"),
    "btn_lg":  ("Segoe UI", 11, "bold"),
    "btn_md":  ("Segoe UI", 9, "bold"),
    "btn_sm":  ("Segoe UI", 8),
}

def _add_hover(btn, enter_bg, leave_bg, enter_fg=None, leave_fg=None):
    """Add hover color transition to a tk.Button."""
    def on_enter(e):
        btn.config(bg=enter_bg)
        if enter_fg: btn.config(fg=enter_fg)
    def on_leave(e):
        btn.config(bg=leave_bg)
        if leave_fg: btn.config(fg=leave_fg)
    btn.bind("<Enter>", on_enter, add="+")
    btn.bind("<Leave>", on_leave, add="+")

def _make_card(parent, pad=1):
    """Create a card frame with subtle border effect."""
    outer = tk.Frame(parent, bg=_UI["BORDER"], padx=pad, pady=pad)
    inner = tk.Frame(outer, bg=_UI["SURFACE"])
    inner.pack(fill=tk.BOTH, expand=True)
    return outer, inner


class ToastNotification(tk.Toplevel):
    """Auto-dismissing notification popup that fades out."""
    def __init__(self, parent, message, duration=2500, bg="#0F172A", fg="white"):
        super().__init__(parent)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=bg)
        container = tk.Frame(self, bg=bg, padx=16, pady=10)
        container.pack()
        tk.Label(container, text=message, bg=bg, fg=fg,
                 font=("Segoe UI", 10, "bold"), padx=8).pack(side=tk.LEFT)
        self.update_idletasks()
        px = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        py = parent.winfo_rooty() + 10
        self.geometry(f"+{px}+{py}")
        try:
            self.attributes("-alpha", 0.95)
        except Exception:
            pass
        self.after(duration, self._fade_out)

    def _fade_out(self):
        try:
            alpha = float(self.attributes("-alpha"))
            if alpha > 0.1:
                self.attributes("-alpha", alpha - 0.1)
                self.after(50, self._fade_out)
            else:
                self.destroy()
        except Exception:
            try: self.destroy()
            except Exception: pass


def show_toast(parent, message, bg="#0F172A", fg="white", duration=2500):
    """Convenience function to show a toast notification."""
    try:
        ToastNotification(parent.winfo_toplevel(), message, duration=duration, bg=bg, fg=fg)
    except Exception:
        pass


class ToolTip:
    """Hover tooltip for any widget."""
    def __init__(self, widget, text, delay=400):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._tip = None
        self._job = None
        widget.bind("<Enter>", self._on_enter, add="+")
        widget.bind("<Leave>", self._on_leave, add="+")

    def _on_enter(self, event):
        self._job = self.widget.after(self.delay, self._show)

    def _on_leave(self, event):
        if self._job:
            self.widget.after_cancel(self._job)
            self._job = None
        self._hide()

    def _show(self):
        if self._tip:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tk.Toplevel(self.widget)
        self._tip.overrideredirect(True)
        self._tip.attributes("-topmost", True)
        lbl = tk.Label(self._tip, text=self.text, bg=_UI["ACCENT"], fg="#F1F5F9",
                       font=("Segoe UI", 9), padx=10, pady=6, relief="flat", borderwidth=0)
        lbl.pack()
        self._tip.geometry(f"+{x}+{y}")

    def _hide(self):
        if self._tip:
            try: self._tip.destroy()
            except Exception: pass
            self._tip = None


def apply_zebra_tags(tree):
    """Apply alternating row colors to a treeview."""
    tree.tag_configure("oddrow", background=_UI["ROW_ODD"])
    tree.tag_configure("evenrow", background=_UI["ROW_EVEN"])
    for i, item in enumerate(tree.get_children("")):
        tag = "evenrow" if i % 2 == 0 else "oddrow"
        tree.item(item, tags=(tag,))


def add_context_menu(tree, extra_commands=None):
    """Add right-click context menu with copy options to a treeview."""
    menu = tk.Menu(tree, tearoff=0)

    def _copy_cell():
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        col = tree.identify_column(tree.winfo_pointerx() - tree.winfo_rootx())
        try:
            idx = int(col.replace('#', '')) - 1
            if 0 <= idx < len(vals):
                tree.clipboard_clear()
                tree.clipboard_append(str(vals[idx]))
        except Exception:
            pass

    def _copy_row():
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        tree.clipboard_clear()
        tree.clipboard_append("\t".join(str(v) for v in vals))

    menu.add_command(label="نسخ الخلية", command=_copy_cell)
    menu.add_command(label="نسخ الصف", command=_copy_row)

    if extra_commands:
        menu.add_separator()
        for label, cmd in extra_commands:
            menu.add_command(label=label, command=cmd)

    def _show_menu(event):
        row = tree.identify_row(event.y)
        if row:
            tree.selection_set(row)
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    tree.bind("<Button-3>", _show_menu)
    return menu


# ------------------- Income Frame -------------------
class IncomeFrame(ttk.Frame):
    
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        # income-only size range state (NOT saved)
        self._income_r1 = tk.StringVar()
        self._income_r2 = tk.StringVar()
        self._income_has_alpha = tk.BooleanVar(value=False)

        self._build()

    def _build(self):
        ttk.Label(self, text="وارد (إضافة أصناف جديدة)", font=_FONTS["h2"]).pack(anchor="w", pady=(0, 4))

        pkg_frame = ttk.LabelFrame(self, text="حاوية (حسب المخزن/العبوة)")
        pkg_frame.pack(fill=tk.X, pady=(0, 4))

        self.wh  = LabeledStaticCombo(
            pkg_frame, "رقم المخزن",
            values=["", *WAREHOUSE_NUMBER_DISPLAY_VALUES],
            value_map=WAREHOUSE_NUMBER_LABELS,
        )
        self.pkg = LabeledEntry(pkg_frame, "رقم العبوة")
        self.pkg_hint_var  = tk.StringVar(value="")
        self.pkg_status_var= tk.StringVar(value="")
        self.pkg_count_var = tk.StringVar(value="0")

        self.wh.grid(row=0, column=0, padx=6, pady=3, sticky="ew")
        self.pkg.grid(row=0, column=1, padx=6, pady=3, sticky="ew")
        self._income_r1.trace_add("write", lambda *_: self._rebuild_sizes_grid())
        self._income_r2.trace_add("write", lambda *_: self._rebuild_sizes_grid())
        self._income_has_alpha.trace_add("write", lambda *_: self._rebuild_sizes_grid())

        hint_box = ttk.Frame(pkg_frame)
        hint_box.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 2))
        ttk.Label(hint_box, textvariable=self.pkg_hint_var).pack(anchor="w")

        stat_box = ttk.Frame(pkg_frame)
        stat_box.grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 3))
        ttk.Label(stat_box, text="الحالة:").pack(side=tk.LEFT)
        ttk.Label(stat_box, textvariable=self.pkg_status_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(stat_box, text="عدد صفوف المخزون داخل العبوة:").pack(side=tk.LEFT)
        ttk.Label(stat_box, textvariable=self.pkg_count_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(4, 0))

        pkg_frame.columnconfigure(0, weight=1)
        pkg_frame.columnconfigure(1, weight=1)

        grid = ttk.LabelFrame(self, text="بيانات الصنف")
        grid.pack(fill=tk.BOTH, expand=True)

        # Specs comboboxes (no single 'size' field now)
        self.item_type = LabeledCombobox(grid, "النوع", self.db, "item_type")
        self.school    = LabeledCombobox(grid, "المدرسة", self.db, "school")
        self.color     = LabeledCombobox(grid, "اللون", self.db, "color")
        # ---------- Income size ranges (single compact row) ----------
        ranges_box = ttk.Frame(grid)
        ranges_box.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(2, 2))

        ttk.Label(ranges_box, text="النطاق الأول").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            ranges_box,
            textvariable=self._income_r1,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=14
        ).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(ranges_box, text="النطاق الثاني").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            ranges_box,
            textvariable=self._income_r2,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=14
        ).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Checkbutton(
            ranges_box,
            text="مقاسات حرفية (S → 5XL)",
            variable=self._income_has_alpha
        ).pack(side=tk.LEFT)

        # --- Cascading filter suppliers ---
        self.item_type.set_supplier(lambda: self.db.get_distinct_filtered(
            "item_type", self._income_constraints("item_type")) or [])
        self.school.set_supplier(lambda: self.db.get_distinct_filtered(
            "school", self._income_constraints("school")) or [])
        self.color.set_supplier(lambda: self.db.get_distinct_filtered(
            "color", self._income_constraints("color")) or [])

        # Bind to cascade + auto-load size profile + fill prices when specs change
        def _on_spec_change(e=None):
            self._on_income_filter_changed()
            self._auto_load_size_profile()
            self._rebuild_sizes_grid()
            self._auto_fill_price_for_grid()

        for w in (self.item_type.cb, self.school.cb, self.color.cb):
            w.bind("<<ComboboxSelected>>", _on_spec_change, add="+")


        # ---------- Scrollable sizes grid container ----------
        sizes_container = ttk.Frame(grid)
        sizes_container.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=6, pady=6)

        # Canvas + scrollbar
        sizes_canvas = tk.Canvas(sizes_container, highlightthickness=0)
        sizes_scroll = ttk.Scrollbar(sizes_container, orient="vertical", command=sizes_canvas.yview)
        sizes_canvas.configure(yscrollcommand=sizes_scroll.set)

        sizes_scroll.pack(side="right", fill="y")
        sizes_canvas.pack(side="left", fill="both", expand=True)

        # Inner frame that holds the SizesGrid
        self._sizes_inner = ttk.Frame(sizes_canvas)
        sizes_window = sizes_canvas.create_window((0, 0), window=self._sizes_inner, anchor="nw")

        # start with empty grid (no specs selected yet)
        self.sizes_grid = None

        # Make canvas resize and scroll correctly
        def _on_sizes_config(event=None):
            try:
                sizes_canvas.configure(scrollregion=sizes_canvas.bbox("all"))
                sizes_canvas.itemconfigure(sizes_window, width=sizes_canvas.winfo_width())
            except Exception:
                pass

        self._sizes_inner.bind("<Configure>", _on_sizes_config)

        sizes_canvas.bind("<Configure>", _on_sizes_config)

        # ---------- Mouse wheel support ----------
        self._sizes_canvas = sizes_canvas
        _bind_mousewheel(sizes_container, sizes_canvas)

        # NEW: badge checkbox — inside the sizes_container so it doesn't overlap
        self.has_badge = tk.BooleanVar(value=False)

        # layout of spec fields + sizes grid
        # Row 0: item_type + school
        # Row 1: color
        # Row 2: ranges_box (already placed above)
        # Row 3: sizes_container (canvas — expandable)
        # Row 4: badge
        # Row 5: buttons
        self.item_type.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        self.school.grid(   row=0, column=1, padx=6, pady=4, sticky="ew")
        self.color.grid(    row=1, column=0, padx=6, pady=4, sticky="ew")

        badge_row = ttk.Frame(grid)
        badge_row.grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 2))
        ttk.Checkbutton(badge_row, text="بادج", variable=self.has_badge).pack(side=tk.LEFT)

        # ensure columns expand equally
        for c in range(2):
            grid.columnconfigure(c, weight=1)

        # Give the sizes row weight so it expands to fill available space
        grid.rowconfigure(3, weight=1)   # sizes_container canvas

        # Put package buttons INSIDE the 'grid' labelframe
        pkg_btns = ttk.Frame(grid)
        pkg_btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 4))

        # keep the same buttons inside it
        _btn_close_pkg = ttk.Button(pkg_btns, text="إغلاق العبوة", command=self._close_current_package); _btn_close_pkg.pack(side=tk.RIGHT)
        ToolTip(_btn_close_pkg, "إغلاق العبوة الحالية ومنع الإضافة إليها")
        _btn_add = ttk.Button(pkg_btns, text="إضافة", command=self._on_add); _btn_add.pack(side=tk.LEFT)
        ToolTip(_btn_add, "إضافة الصنف بالكمية المحددة إلى المخزون")

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, pady=(12, 0))
        _btn_reset = ttk.Button(btns, text="تفريغ (يبقي المخزن/العبوة)", command=self._on_reset_keep_pkg); _btn_reset.pack(side=tk.LEFT, padx=8)
        ToolTip(_btn_reset, "مسح الحقول مع الإبقاء على رقم المخزن والعبوة")
        # removed "إضافة مقاسات متعددة…" button (inline grid used instead)

        # bindings
        self.wh.cb.bind("<<ComboboxSelected>>", lambda *_: (self._refresh_pkg_hints(), self._refresh_pkg_status()))
        self.pkg.var.trace_add("write", lambda *_: self._refresh_pkg_status())

        # initial
        self._refresh_pkg_hints()
        self._refresh_pkg_status()

    def _income_constraints(self, exclude_field: str) -> dict:
        """Build a constraints dict from the other two spec fields (for cascading)."""
        c = {}
        if exclude_field != "item_type":
            v = (self.item_type.get() or "").strip()
            if v:
                c["item_type"] = v
        if exclude_field != "school":
            v = (self.school.get() or "").strip()
            if v:
                c["school"] = v
        if exclude_field != "color":
            v = (self.color.get() or "").strip()
            if v:
                c["color"] = v
        return c

    def _on_income_filter_changed(self):
        """Validate current values and refresh cascading combo options."""
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        # Validate each value against the other two
        if it:
            valid = self.db.get_distinct_filtered("item_type", self._income_constraints("item_type"))
            if it not in valid:
                self.item_type.set("")
        if sc:
            valid = self.db.get_distinct_filtered("school", self._income_constraints("school"))
            if sc not in valid:
                self.school.set("")
        if cl:
            valid = self.db.get_distinct_filtered("color", self._income_constraints("color"))
            if cl not in valid:
                self.color.set("")
        # Refresh all combo dropdown lists
        self.item_type.refresh_values()
        self.school.refresh_values()
        self.color.refresh_values()

    def _auto_load_size_profile(self):
        """Auto-set range dropdowns from saved size_profile when item+school+color are all set."""
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        if not (it and sc and cl):
            return
        try:
            profile = self.db.get_size_profile(it, sc, cl)
            if not profile:
                return
            r1s, r1e, r2s, r2e, has_alpha = profile
            # Set range 1
            if r1s is not None and r1e is not None:
                label1 = f"{r1s} → {r1e}"
                if label1 in NUMERIC_RANGE_LABELS:
                    self._income_r1.set(label1)
            else:
                self._income_r1.set("")
            # Set range 2
            if r2s is not None and r2e is not None:
                label2 = f"{r2s} → {r2e}"
                if label2 in NUMERIC_RANGE_LABELS:
                    self._income_r2.set(label2)
            else:
                self._income_r2.set("")
            # Set alpha
            self._income_has_alpha.set(bool(has_alpha))
        except Exception:
            pass

    def _rebuild_sizes_grid(self):
        try:
            self.sizes_grid.destroy()
        except Exception:
            pass

        numeric_ranges = []

        def _parse_range(label: str):
            if not label:
                return None
            a, b = label.split("→")
            return int(a.strip()), int(b.strip())

        # collect selected ranges
        for lbl in (self._income_r1.get(), self._income_r2.get()):
            pair = _parse_range(lbl)
            if pair:
                numeric_ranges.append(pair)

        # --- MERGE OVERLAPPING RANGES ---
        numeric_ranges.sort()
        merged = []

        for start, end in numeric_ranges:
            if not merged:
                merged.append([start, end])
            else:
                last_start, last_end = merged[-1]
                if start <= last_end:  # overlap or touch
                    merged[-1][1] = max(last_end, end)
                else:
                    merged.append([start, end])

        # --- EXPAND MERGED RANGES ---
        sizes = []
        for start, end in merged:
            sizes.extend(str(x) for x in range(start, end + 1, 2))

        # alpha sizes
        if self._income_has_alpha.get():
            sizes.extend(ALPHA_SIZES)

        self.sizes_grid = SizesGrid(self._sizes_inner, sizes)
        self.sizes_grid.pack(fill="both", expand=False)

        self._sizes_inner.update_idletasks()



    def _auto_fill_price_for_grid(self):
        """Fill empty price cells in the sizes grid from last known prices."""
        if not self.sizes_grid:
            return
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        if not (it and sc and cl):
            return
        try:
            # Query per-size prices for each size in the grid
            found_any = False
            for sz in self.sizes_grid.fixed_sizes:
                p = self.db.last_price_for_specs(it, sc, cl, sz)
                if p is not None:
                    self.sizes_grid.set_price_for_size(sz, p)
                    found_any = True
            # If no per-size prices found, try the first size as a fallback for all
            if not found_any and self.sizes_grid.fixed_sizes:
                first = self.sizes_grid.fixed_sizes[0]
                p = self.db.last_price_for_specs(it, sc, cl, first)
                if p is not None:
                    self.sizes_grid.set_price_for_all(p)
        except Exception:
            pass

    def _parse_int_or_none(self, s: str) -> Optional[int]:
        s = warehouse_numeric_value(s)
        s = (s or "").strip()
        if not s:
            return None
        try:
            v = int(s)
            return v
        except Exception:
            return None

    def _refresh_pkg_hints(self):
        w = self._parse_int_or_none(self.wh.get())
        if not w or w < 1:
            self.pkg_hint_var.set("أدخل رقم المخزن أولاً.")
            return
        try:
            info = self.db.package_numbers_summary(w)
        except Exception as ex:
            self.pkg_hint_var.set(str(ex))
            return
        if info["free"]:
            free_txt = "، ".join(map(str, info["free"][:20]))
            more = "" if len(info["free"]) <= 20 else " …"
            hint = f"الأرقام المتاحة (فارغة): {free_txt}{more}"
        else:
            hint = f"لا توجد فراغات. الرقم التالي: {info['next']}"
        if info.get("open"):
            hint += f" | حاويات مفتوحة: {', '.join(map(str, info['open'][:20]))}"
        self.pkg_hint_var.set(hint)
        self._refresh_pkg_status()

    def _refresh_pkg_status(self):
        w = self._parse_int_or_none(self.wh.get())
        p = self._parse_int_or_none(self.pkg.get())
        if not (w and p):
            self.pkg_status_var.set("")
            self.pkg_count_var.set("0")
            return
        status = self.db.package_status(w, p)
        if status is None:
            self.pkg_status_var.set("غير موجودة (سيتم فتحها تلقائياً عند أول إضافة)")
        elif status == "OPEN":
            self.pkg_status_var.set("OPEN")
        else:
            self.pkg_status_var.set("CLOSED")
        try:
            cur = self.db.conn.execute(
                "SELECT COUNT(*) AS c FROM stocks WHERE warehouse_no=? AND package_no=?",
                (int(w), int(p)),
            )
            c = cur.fetchone()["c"] or 0
        except Exception:
            c = 0
        self.pkg_count_var.set(str(int(c)))

    def _close_current_package(self):
        w = self._parse_int_or_none(self.wh.get())
        p = self._parse_int_or_none(self.pkg.get())
        if not (w and p):
            messagebox.showwarning("بيانات ناقصة", "أدخل رقم المخزن والعبوة أولاً.")
            return
        if not messagebox.askyesno("تأكيد", f"إغلاق العبوة {p} في المخزن {w} ؟\nلن يُسمح بأي إضافات لاحقاً."):
            return
        try:
            self.db.close_package(w, p)
            show_toast(self, "تم إغلاق العبوة بنجاح")
            self._on_reset_keep_pkg()
            self._refresh_pkg_hints()
            try:
                self.item_type.cb.focus_set()
            except Exception:
                pass
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def _on_reset_keep_pkg(self):
        for w in (self.item_type, self.school, self.color):
            w.set("")
        self._income_r1.set("")
        self._income_r2.set("")
        self._income_has_alpha.set(False)

        # clear sizes grid entries and badge
        if self.sizes_grid:
            self.sizes_grid.destroy()
            self.sizes_grid = None

        self._sizes_inner.update_idletasks()

        self.has_badge.set(False)

        self._refresh_pkg_status()

    def _on_add(self):
        w = self._parse_int_or_none(self.wh.get())
        p = self._parse_int_or_none(self.pkg.get())
        if not (w and p and w >= 1 and p >= 1):
            messagebox.showerror("بيانات ناقصة", "أدخل رقم المخزن والعبوة بشكل صحيح (>= 1).")
            return
        try:
            item_type = self.item_type.get() or self._err("النوع مطلوب")
            school = self.school.get() or self._err("المدرسة مطلوبة")
            color = self.color.get() or self._err("اللون مطلوب")
        except RuntimeError as ex:
            messagebox.showerror("بيانات ناقصة", str(ex))
            return

        rows = self.sizes_grid.get_rows()
        # keep only rows with qty > 0
        rows = [r for r in rows if int(r.get("qty") or 0) > 0]
        if not rows:
            messagebox.showwarning("فارغ", "لم تُدخل أي كميات في شبكة المقاسات.", parent=self)
            return

        # ensure package open
        try:
            self.db.ensure_package_open(w, p)
        except Exception as ex:
            messagebox.showerror("فشل الإضافة", str(ex))
            return

        added = 0
        errors: List[str] = []
        for r in rows:
            size = str(r["size"]).strip()
            qty = int(r["qty"])
            price = r["price"]

            if price is None:
                price = self.db.get_effective_price(item_type, school, color, size)

            if price is None:
                errors.append(f"{size}: السعر غير معروف (أدخله مرة واحدة على الأقل).")
                continue

            try:
                stock_id = self.db.add_stock(
                    item_type=item_type,
                    school=school,
                    color=color,
                    size=size,
                    warehouse_no=w,
                    package_no=p,
                    unit_price=float(price),
                    count=qty,
                    has_badge=1 if self.has_badge.get() else 0,
                )

                # save default price ONCE
                self.db.ensure_default_price(item_type, price)

                added += 1
            except Exception as ex:
                errors.append(f"{size}: {ex}")

        msg_parts = []
        if added:
            msg_parts.append(f"تمت إضافة {added} صفًا للمخزون.")
        if errors:
            msg_parts.append("أخطاء:\n" + "\n".join(errors))
        if not msg_parts:
            show_toast(self, "لم تُجر أي تغييرات", bg="#f59e0b")
        else:
            show_toast(self, "\n".join(msg_parts), duration=3500)

        if added:
            # clear specs
            self.item_type.set("")
            self.school.set("")
            self.color.set("")

            # clear income-only ranges
            self._income_r1.set("")
            self._income_r2.set("")
            self._income_has_alpha.set(False)

            # remove sizes grid completely
            if self.sizes_grid:
                try:
                    self.sizes_grid.destroy()
                except Exception:
                    pass
                self.sizes_grid = None

            # clear badge
            self.has_badge.set(False)

            # refresh package info only
            self._refresh_pkg_hints()
            self._refresh_pkg_status()

            # optional UX: focus back to item type
            try:
                self.item_type.cb.focus_set()
            except Exception:
                pass


    @staticmethod
    def _err(msg: str) -> None:
        raise RuntimeError(msg)

# ------------------- Outcome (Bill) Frame -------------------

class OutcomeFrame(ttk.Frame):
    """
    POS-style bill creator:
      1) choose school  ->  2) choose item type  ->  3) choose color
      4) choose size (shows all defined sizes with current counts) + qty/price -> Add to bill
    Left panel: search + filters + button grid.  Right panel: bill table.
    """
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=0)
        self.db = db
        self.bill_lines: List[Dict[str, Any]] = []

        # selection state
        self._sel_school: Optional[str] = None
        self._sel_item: Optional[str] = None
        self._sel_color: Optional[str] = None
        self._sel_size: Optional[str]  = None

        self._sizes_cache: List[Dict[str, Any]] = []
        self._build()

    # ---- Premium button palettes ----
    _BTN_PRODUCT = {
        "bg": "#FFFFFF", "fg": _UI["TEXT"], "font": ("Segoe UI", 9, "bold"),
        "bd": 0, "relief": "flat", "padx": 12, "pady": 6, "cursor": "hand2",
        "activebackground": _UI["BRAND_L"], "activeforeground": _UI["SEL_FG"],
        "highlightbackground": _UI["BORDER"], "highlightthickness": 1,
    }
    _BTN_BACK = {
        "bg": _UI["SURFACE"], "fg": _UI["TEXT_SEC"], "font": ("Segoe UI", 8),
        "bd": 0, "cursor": "hand2", "padx": 8, "pady": 4,
        "highlightbackground": _UI["BORDER"], "highlightthickness": 1,
        "activebackground": _UI["SURFACE2"], "activeforeground": _UI["TEXT"],
    }
    _BTN_BLUE = _BTN_PRODUCT   # alias for compatibility
    _BTN_GRAY = _BTN_BACK      # alias for compatibility

    # ---------------- UI build ----------------
    def _build(self):
        # ============ MAIN HORIZONTAL SPLIT ============
        hsplit = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        hsplit.pack(fill=tk.BOTH, expand=True)

        # ============ LEFT PANEL (Product Selection) ============
        left = tk.Frame(hsplit, bg=_UI["BG"])
        hsplit.add(left, weight=5)

        # -- Search bar (card-like) --
        sf = tk.Frame(left, bg=_UI["SURFACE"],
                      highlightbackground=_UI["BORDER"], highlightthickness=1,
                      padx=12, pady=8)
        sf.pack(fill=tk.X, padx=12, pady=(12, 8))

        self._search_var = tk.StringVar()
        _se = tk.Entry(sf, textvariable=self._search_var, font=_FONTS["body"],
                       bg=_UI["SURFACE"], fg=_UI["TEXT"], bd=0,
                       highlightthickness=0, insertbackground=_UI["TEXT"])
        _se.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(0, 8))
        _se.bind("<Return>", lambda e: self._do_search())

        _clear_btn = tk.Button(sf, text="X", command=lambda: (self._search_var.set(""), self._clear_quick_filters()),
                               bg=_UI["SURFACE"], fg=_UI["TEXT_DIM"], bd=0,
                               font=("Segoe UI", 10), cursor="hand2",
                               activebackground=_UI["SURFACE"])
        _clear_btn.pack(side=tk.LEFT)
        _add_hover(_clear_btn, _UI["SURFACE"], _UI["SURFACE"], _UI["DANGER"], _UI["TEXT_DIM"])

        # -- Cascading quick filters (inline on secondary surface) --
        qf = tk.Frame(left, bg=_UI["SURFACE2"])
        qf.pack(fill=tk.X, padx=12, pady=(0, 8))
        qf_inner = tk.Frame(qf, bg=_UI["SURFACE2"])
        qf_inner.pack(fill=tk.X, padx=8, pady=6)

        for _, (label_text, attr, width) in enumerate([
            ("المدرسة", "_filter_school", 18),
            ("النوع", "_filter_type", 16),
            ("اللون", "_filter_color", 12),
        ]):
            pill = tk.Frame(qf_inner, bg=_UI["SURFACE2"])
            pill.pack(side=tk.RIGHT, padx=6)
            tk.Label(pill, text=label_text, bg=_UI["SURFACE2"], fg=_UI["TEXT_SEC"],
                     font=_FONTS["small"]).pack(side=tk.RIGHT, padx=(0, 4))
            cb = ttk.Combobox(pill, width=width, state="readonly")
            cb.pack(side=tk.RIGHT)
            setattr(self, attr, cb)

        clear_f = tk.Button(qf_inner, text="مسح", command=self._clear_quick_filters,
                            bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"], font=_FONTS["small"],
                            bd=0, padx=12, pady=3, cursor="hand2",
                            highlightbackground=_UI["BORDER"], highlightthickness=1,
                            activebackground=_UI["SURFACE2"])
        clear_f.pack(side=tk.LEFT, padx=4)
        _add_hover(clear_f, _UI["SURFACE2"], _UI["SURFACE"])

        self._filter_school.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed("school"))
        self._filter_type.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed("item_type"))
        self._filter_color.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed("color"))
        self._refresh_filter_combos()

        # -- Breadcrumb --
        crumb = tk.Frame(left, bg=_UI["BG"])
        crumb.pack(fill=tk.X, padx=14, pady=(4, 2))
        self._crumb_var = tk.StringVar(value="اختر المدرسة")
        tk.Label(crumb, textvariable=self._crumb_var, bg=_UI["BG"], fg=_UI["TEXT"],
                 font=_FONTS["h3"]).pack(anchor="w")

        # -- Scrollable button grid --
        grid_container = tk.Frame(left, bg=_UI["BG"])
        grid_container.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))

        self._canvas = tk.Canvas(grid_container, highlightthickness=0, bd=0,
                                 bg=_UI["BG"])
        self._scroll_y = ttk.Scrollbar(grid_container, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._scroll_y.set)

        # --- Size button styles ---
        try:
            s = ttk.Style(self)
            s.configure("Size.TButton", padding=5, font=("Segoe UI", 9))
            s.configure("SizeZero.TButton",
                        background=_UI["DANGER_L"], foreground=_UI["TEXT"],
                        bordercolor=_UI["DANGER"], focusthickness=0, padding=8)
            s.map("SizeZero.TButton",
                background=[("active", "#FECACA"), ("pressed", "#FECACA")],
                bordercolor=[("focus", _UI["DANGER"])])
            s.configure("SizeSelected.TButton",
                        background=_UI["OK_L"], foreground=_UI["TEXT"],
                        bordercolor=_UI["OK"], focusthickness=0, padding=8)
            s.map("SizeSelected.TButton",
                background=[("active", "#A7F3D0"), ("pressed", "#A7F3D0")],
                bordercolor=[("focus", _UI["OK"])])
            s.configure("SizeSelectedZero.TButton",
                        background=_UI["WARN_L"], foreground=_UI["TEXT"],
                        bordercolor=_UI["WARN"], focusthickness=0, padding=8)
            s.map("SizeSelectedZero.TButton",
                background=[("active", "#FDE68A"), ("pressed", "#FDE68A")],
                bordercolor=[("focus", _UI["WARN"])])
        except Exception:
            pass

        self._size_btns = {}
        self._selected_size_btn = None

        self._grid_host = tk.Frame(self._canvas, bg=_UI["BG"])
        self._grid_window = self._canvas.create_window((0, 0), window=self._grid_host, anchor="nw")
        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        def _on_configure(_e=None):
            self._canvas.configure(scrollregion=self._canvas.bbox("all"))
            try:
                self._canvas.itemconfigure(self._grid_window, width=self._canvas.winfo_width())
            except Exception:
                pass
        self._grid_host.bind("<Configure>", _on_configure)
        self._canvas.bind("<Configure>", _on_configure)

        # Mousewheel: direct binding on canvas + grid_host; buttons bound in _mk_grid_buttons
        def _mw(e):
            try:
                self._canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
            except Exception:
                pass
            return "break"
        self._grid_mw = _mw
        self._canvas.bind("<MouseWheel>", _mw)
        self._grid_host.bind("<MouseWheel>", _mw)

        # -- Favorites bar --
        self._fav_frame = ttk.LabelFrame(left, text="الأكثر مبيعاً")
        self._fav_frame.pack(fill=tk.X, padx=12, pady=(0, 4))
        self._fav_inner = ttk.Frame(self._fav_frame)
        self._fav_inner.pack(fill=tk.X, padx=4, pady=4)
        self._refresh_favorites()

        # -- Action row (qty / price / add) --
        act = tk.Frame(left, bg=_UI["SURFACE"],
                       highlightbackground=_UI["BORDER"], highlightthickness=1)
        act.pack(fill=tk.X, padx=12, pady=(0, 8))
        act_inner = tk.Frame(act, bg=_UI["SURFACE"])
        act_inner.pack(fill=tk.X, padx=10, pady=8)

        # Quantity
        tk.Label(act_inner, text="الكمية:", bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
                 font=_FONTS["small"]).pack(side=tk.LEFT, padx=(0, 4))
        self.qty_var = tk.StringVar(value="1")
        _qm = tk.Button(act_inner, text="-",
                         command=lambda: self.qty_var.set(str(max(1, int(self.qty_var.get() or 1) - 1))),
                         bg=_UI["SURFACE"], fg=_UI["TEXT"], font=("Segoe UI", 11, "bold"),
                         bd=0, width=3, cursor="hand2",
                         highlightbackground=_UI["BORDER"], highlightthickness=1,
                         activebackground=_UI["SURFACE2"])
        _qm.pack(side=tk.LEFT)
        _add_hover(_qm, _UI["SURFACE2"], _UI["SURFACE"])
        ttk.Entry(act_inner, textvariable=self.qty_var, width=5, justify="center",
                  font=_FONTS["body"]).pack(side=tk.LEFT, padx=2)
        _qp = tk.Button(act_inner, text="+",
                         command=lambda: self.qty_var.set(str(int(self.qty_var.get() or 1) + 1)),
                         bg=_UI["SURFACE"], fg=_UI["TEXT"], font=("Segoe UI", 11, "bold"),
                         bd=0, width=3, cursor="hand2",
                         highlightbackground=_UI["BORDER"], highlightthickness=1,
                         activebackground=_UI["SURFACE2"])
        _qp.pack(side=tk.LEFT, padx=(0, 16))
        _add_hover(_qp, _UI["SURFACE2"], _UI["SURFACE"])

        # Price
        tk.Label(act_inner, text="السعر:", bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
                 font=_FONTS["small"]).pack(side=tk.LEFT, padx=(0, 4))
        self.price_var = tk.StringVar(value="")
        self._price_entry = ttk.Entry(act_inner, textvariable=self.price_var, width=10,
                                       font=_FONTS["body"])
        self._price_entry.pack(side=tk.LEFT, padx=(0, 16))
        self._price_user_edited = False
        self._price_entry.bind("<KeyRelease>", lambda e: setattr(self, "_price_user_edited", True), add="+")
        self._price_entry.bind("<FocusIn>", lambda e: setattr(self, "_price_user_edited", True), add="+")

        # Factory checkbox
        self.instant_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(act_inner, text="من المصنع", variable=self.instant_mode).pack(side=tk.LEFT, padx=(0, 8))

        # Add to bill button (brand blue)
        _add_btn = tk.Button(act_inner, text="أضف إلى الفاتورة", command=self._add_current_selection,
                             bg=_UI["BRAND"], fg="#FFFFFF", font=_FONTS["btn_md"],
                             bd=0, padx=20, pady=8, cursor="hand2",
                             activebackground=_UI["BRAND_H"], activeforeground="#FFFFFF")
        _add_btn.pack(side=tk.RIGHT)
        _add_hover(_add_btn, _UI["BRAND_H"], _UI["BRAND"])

        # ============ RIGHT PANEL (Cart / Bill) ============
        right = tk.Frame(hsplit, bg=_UI["SURFACE"])
        hsplit.add(right, weight=5)

        # Set initial sash so bill panel gets enough room
        def _set_sash(_e=None):
            try:
                w = hsplit.winfo_width()
                if w > 100:
                    hsplit.sashpos(0, int(w * 0.48))
                    hsplit.unbind("<Map>")
            except Exception:
                pass
        hsplit.bind("<Map>", _set_sash)

        # -- Customer --
        cust_frame = tk.Frame(right, bg=_UI["SURFACE"])
        cust_frame.pack(fill=tk.X, padx=10, pady=(10, 6))
        tk.Label(cust_frame, text="العميل:", bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
                 font=_FONTS["body"]).pack(side=tk.RIGHT, padx=(4, 0))
        self.customer = LabeledCombobox(cust_frame, "", self.db, "item_type")

        def _customer_supplier():
            # Branch POS devices appear at the top with a "فرع: " prefix
            # so the user can tell a shipment bill apart from a sale
            # bill at a glance. Picking one routes the bill through
            # Phase-3 STOCK_TRANSFER_OUT on save.
            pos_names = []
            try:
                pos_names = self.db.list_known_pos_device_names() or []
            except Exception:
                pos_names = []
            customers = []
            try:
                customers = self.db.list_customers() or []
            except Exception:
                customers = []

            branch_names = []
            seen_branch = set()
            for name in pos_names:
                clean = (name or "").strip()
                if not clean or clean in seen_branch:
                    continue
                seen_branch.add(clean)
                branch_names.append(clean)

            branch_set = set(branch_names)
            clean_customers = []
            seen_customers = set()
            for customer in customers:
                raw = (customer or "").strip()
                if not raw:
                    continue
                if normalize_branch_customer_name(raw):
                    continue
                if raw in seen_customers:
                    continue
                seen_customers.add(raw)
                clean_customers.append(raw)

            return [branch_customer_label(n) for n in branch_names] + clean_customers

        self.customer.set_supplier(_customer_supplier)
        for ev in ("<FocusIn>", "<Button-1>", "<KeyRelease>"):
            self.customer.cb.bind(ev, lambda e: self.customer.refresh_values(), add="+")
        self.customer.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        # -- Bill table --
        self.bill_table = ttk.Treeview(
            right,
            columns=("type", "school", "color", "size", "price", "qty", "total"),
            show="headings", height=12,
        )
        for col, txt, w in [
            ("type", "النوع", 100), ("school", "المدرسة", 100), ("color", "اللون", 70),
            ("size", "المقاس", 55), ("price", "السعر", 65), ("qty", "الكمية", 50),
            ("total", "الإجمالي", 75),
        ]:
            self.bill_table.heading(col, text=txt)
            self.bill_table.column(col, width=w, anchor="center")
        self.bill_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))
        add_context_menu(self.bill_table)
        _bind_mousewheel(self.bill_table)

        # -- Bill action buttons (ghost style) --
        bar = tk.Frame(right, bg=_UI["SURFACE"])
        bar.pack(fill=tk.X, padx=10, pady=(0, 6))

        _bm = tk.Button(bar, text="-", command=self._decrement_bill_line,
                         bg=_UI["SURFACE"], fg=_UI["TEXT"], font=("Segoe UI", 11, "bold"),
                         bd=0, width=3, cursor="hand2",
                         highlightbackground=_UI["BORDER"], highlightthickness=1,
                         activebackground=_UI["SURFACE2"])
        _bm.pack(side=tk.LEFT)
        _add_hover(_bm, _UI["SURFACE2"], _UI["SURFACE"])

        _bp = tk.Button(bar, text="+", command=self._increment_bill_line,
                         bg=_UI["SURFACE"], fg=_UI["TEXT"], font=("Segoe UI", 11, "bold"),
                         bd=0, width=3, cursor="hand2",
                         highlightbackground=_UI["BORDER"], highlightthickness=1,
                         activebackground=_UI["SURFACE2"])
        _bp.pack(side=tk.LEFT, padx=(2, 12))
        _add_hover(_bp, _UI["SURFACE2"], _UI["SURFACE"])

        _del = tk.Button(bar, text="حذف السطر", command=self._remove_bill_line,
                         bg=_UI["SURFACE"], fg=_UI["DANGER"], font=("Segoe UI", 9, "bold"),
                         bd=0, padx=12, pady=4, cursor="hand2",
                         highlightbackground=_UI["DANGER_L"], highlightthickness=1,
                         activebackground=_UI["DANGER_L"])
        _del.pack(side=tk.LEFT, padx=(0, 4))
        _add_hover(_del, _UI["DANGER_L"], _UI["SURFACE"])

        _clr = tk.Button(bar, text="تفريغ", command=self._clear_bill,
                         bg=_UI["SURFACE"], fg=_UI["WARN"], font=("Segoe UI", 9, "bold"),
                         bd=0, padx=12, pady=4, cursor="hand2",
                         highlightbackground=_UI["WARN_L"], highlightthickness=1,
                         activebackground=_UI["WARN_L"])
        _clr.pack(side=tk.LEFT)
        _add_hover(_clr, _UI["WARN_L"], _UI["SURFACE"])

        # -- Total display (green accent card) --
        total_card = tk.Frame(right, bg=_UI["OK_L"],
                              highlightbackground="#BBF7D0", highlightthickness=1)
        total_card.pack(fill=tk.X, padx=10, pady=(0, 8))
        total_inner = tk.Frame(total_card, bg=_UI["OK_L"])
        total_inner.pack(fill=tk.X, padx=16, pady=12)
        tk.Label(total_inner, text="الإجمالي", bg=_UI["OK_L"], fg="#065F46",
                 font=_FONTS["body"]).pack(side=tk.RIGHT, padx=(0, 12))
        self.total_var = tk.StringVar(value="0.00")
        tk.Label(total_inner, textvariable=self.total_var, bg=_UI["OK_L"], fg="#065F46",
                 font=_FONTS["price"]).pack(side=tk.RIGHT)

        # -- Confirm button (full width, premium green) --
        _conf = tk.Button(right, text="تأكيد الفاتورة / الحجز  (F8)", command=self._finalize_bill,
                          bg=_UI["OK"], fg="white", font=_FONTS["btn_lg"],
                          bd=0, pady=14, cursor="hand2",
                          activebackground=_UI["OK_H"], activeforeground="white")
        _conf.pack(fill=tk.X, padx=10, pady=(0, 4))
        _add_hover(_conf, _UI["OK_H"], _UI["OK"])

        # Draft button
        ttk.Button(right, text="حفظ كمسودة", command=self._save_as_draft)\
            .pack(fill=tk.X, padx=10, pady=(0, 8))

        # ---- initial stage: show SCHOOLS first ----
        self._render_schools()

    # ---------------- Stage renders ----------------
    def _clear_grid(self):
        for w in self._grid_host.winfo_children():
            w.destroy()
        try:
            self._canvas.yview_moveto(0)
        except Exception:
            pass

    def _bind_grid_scroll(self):
        """Bind mousewheel to every widget inside the scrollable grid."""
        mw = self._grid_mw
        def _rec(w):
            try:
                w.bind("<MouseWheel>", mw)
            except Exception:
                pass
            for ch in w.winfo_children():
                _rec(ch)
        _rec(self._grid_host)

    def _mk_grid_buttons(self, labels: List[str], on_click, *, cols: int = 5):
        """Create a grid of product-tile buttons inside the scrollable host."""
        if not labels:
            tk.Label(self._grid_host, text="لا توجد بيانات.", bg=_UI["BG"],
                     fg=_UI["TEXT_DIM"], font=_FONTS["body"]).pack(pady=20)
            return
        row = None
        for i, label in enumerate(labels):
            if i % cols == 0:
                row = tk.Frame(self._grid_host, bg=_UI["BG"])
                row.pack(fill=tk.X, padx=4, pady=2)
            btn = tk.Button(row, text=label, command=lambda v=label: on_click(v),
                            **self._BTN_PRODUCT)
            btn.pack(side=tk.LEFT, padx=4, pady=4, ipadx=4, ipady=2)
            _add_hover(btn, _UI["BRAND_L"], _UI["SURFACE"],
                       _UI["SEL_FG"], _UI["TEXT"])

    # ---------------- Cascading Quick Filters ----------------
    def _refresh_filter_combos(self):
        """Repopulate each filter combobox based on the other two selections."""
        try:
            constraints: Dict[str, Any] = {}
            school = self._filter_school.get().strip()
            item_type = self._filter_type.get().strip()
            color = self._filter_color.get().strip()

            if school:
                constraints["school"] = school
            if item_type:
                constraints["item_type"] = item_type
            if color:
                constraints["color"] = color

            # School dropdown: filtered by type + color
            sc = {k: v for k, v in constraints.items() if k != "school"}
            self._filter_school["values"] = self.db.get_distinct_filtered("school", sc)

            # Type dropdown: filtered by school + color
            tc = {k: v for k, v in constraints.items() if k != "item_type"}
            self._filter_type["values"] = self.db.get_distinct_filtered("item_type", tc)

            # Color dropdown: filtered by school + type
            cc = {k: v for k, v in constraints.items() if k != "color"}
            self._filter_color["values"] = self.db.get_distinct_filtered("color", cc)
        except Exception:
            pass

    def _on_filter_changed(self, changed_field: str):
        """Called when any quick-filter combobox selection changes."""
        school = self._filter_school.get().strip() or None
        item_type = self._filter_type.get().strip() or None
        color = self._filter_color.get().strip() or None

        # Invalidate any selection that is no longer valid given the others
        # Check school against current item_type + color
        if school:
            c = {}
            if item_type: c["item_type"] = item_type
            if color: c["color"] = color
            if school not in self.db.get_distinct_filtered("school", c):
                self._filter_school.set("")
                school = None

        # Check item_type against current school + color
        if item_type:
            c = {}
            if school: c["school"] = school
            if color: c["color"] = color
            if item_type not in self.db.get_distinct_filtered("item_type", c):
                self._filter_type.set("")
                item_type = None

        # Check color against current school + item_type
        if color:
            c = {}
            if school: c["school"] = school
            if item_type: c["item_type"] = item_type
            if color not in self.db.get_distinct_filtered("color", c):
                self._filter_color.set("")
                color = None

        self._refresh_filter_combos()
        self._sel_school = school
        self._sel_item = item_type
        self._sel_color = color
        self._sel_size = None
        self._price_user_edited = False

        # Render: School → Item → Color → Size
        if school and item_type and color:
            self._render_sizes()
        elif school and item_type:
            self._render_colors()
        elif school:
            self._render_items()
        else:
            self._render_schools()

    def _clear_quick_filters(self):
        """Reset all quick-filter comboboxes and return to initial view."""
        self._filter_school.set("")
        self._filter_type.set("")
        self._filter_color.set("")
        self._sel_school = None
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._price_user_edited = False
        self._refresh_filter_combos()
        self._render_schools()

    def _sync_filters_to_combos(self):
        """Sync internal selection state to filter comboboxes."""
        self._filter_school.set(self._sel_school or "")
        self._filter_type.set(self._sel_item or "")
        self._filter_color.set(self._sel_color or "")
        self._refresh_filter_combos()

    def _render_schools(self):
        """Initial view: show all schools."""
        self._sel_school = None
        self._sel_size = None
        if not self._sel_item and not self._sel_color:
            self._sel_item = None
            self._sel_color = None
            self._crumb_var.set("اختر المدرسة")
        else:
            self._crumb_var.set("اختر المدرسة (بعد تضييق النتائج)")
        self._clear_grid()
        try:
            constraints: Dict[str, Any] = {}
            if self._sel_item:
                constraints["item_type"] = self._sel_item
            if self._sel_color:
                constraints["color"] = self._sel_color
            schools = sorted({
                r["school"]
                for r in self.db.current_inventory(constraints)
                if r.get("school")
            })
        except Exception:
            schools = []
        self._mk_grid_buttons(schools, self._select_school, cols=4)
        self._bind_grid_scroll()

    def _render_school_search_results(self, school_names: List[str]) -> None:
        """Show only schools in `school_names` (e.g. search matched several)."""
        names = sorted({s for s in (school_names or []) if (s or "").strip()})
        self._sel_school = None
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._crumb_var.set(f"نتائج البحث — اختر المدرسة ({len(names)})")
        self._clear_grid()
        if not names:
            self._crumb_var.set("اختر المدرسة")
            self._bind_grid_scroll()
            return
        self._mk_grid_buttons(names, self._select_school, cols=4)
        tk.Button(
            self._grid_host,
            text="◀ عرض كل المدارس",
            command=self._render_schools,
            **self._BTN_GRAY,
        ).pack(anchor="w", padx=4, pady=4)
        self._bind_grid_scroll()

    # ---------------- Search & Favorites ----------------
    def _do_search(self):
        """Jump to a school or item type matching the search text."""
        q = (self._search_var.get() or "").strip().lower()
        if not q:
            return
        # Prefer item search first so a common text in a school name does not
        # unexpectedly jump to the wrong school.
        try:
            matches: List[Tuple[str, str]] = []
            for school in self.db.list_schools_all():
                for item_type, _color in self.db.list_items_for_school(school):
                    if q in (item_type or "").lower():
                        matches.append((school, item_type))
            if matches:
                exact_items = sorted({it for _sc, it in matches if (it or "").lower() == q})
                start_items = sorted({it for _sc, it in matches if (it or "").lower().startswith(q)})
                chosen_item = (exact_items or start_items or [matches[0][1]])[0]
                schools_for_item = sorted({sc for sc, it in matches if it == chosen_item})
                self._sel_item = chosen_item
                self._sel_color = None
                self._sel_size = None
                if len(schools_for_item) == 1:
                    self._sel_school = schools_for_item[0]
                    self._sync_filters_to_combos()
                    self._render_colors()
                    show_toast(self, f"تم الانتقال إلى: {chosen_item} / {schools_for_item[0]}")
                else:
                    self._sel_school = None
                    self._sync_filters_to_combos()
                    self._render_schools()
                    show_toast(self, f"تم العثور على {chosen_item} في أكثر من مدرسة")
                return
        except Exception:
            pass
        # Then try matching school name(s) by substring (all matches, not only the first).
        try:
            schools = sorted({
                r["school"] for r in self.db.current_inventory({}) if r.get("school")
            })
            matching_schools = [s for s in schools if q in (s or "").lower()]
            if not matching_schools:
                show_toast(self, "لم يتم العثور على نتائج", bg="#dc2626")
                return
            if len(matching_schools) == 1:
                self._select_school(matching_schools[0])
                show_toast(self, f"تم الانتقال إلى: {matching_schools[0]}")
                return
            self._sel_school = None
            self._sel_item = None
            self._sel_color = None
            self._sel_size = None
            self._price_user_edited = False
            self._sync_filters_to_combos()
            self._render_school_search_results(matching_schools)
            show_toast(
                self,
                f"يوجد {len(matching_schools)} مدرسة مطابقة للبحث — اختر المدرسة من القائمة",
            )
            return
        except Exception:
            pass
        show_toast(self, "لم يتم العثور على نتائج", bg="#dc2626")

    def _refresh_favorites(self):
        """Show top 6 most sold item/school/color combos as quick-access buttons."""
        for w in self._fav_inner.winfo_children():
            w.destroy()
        try:
            cur = self.db.conn.cursor()
            cur.execute("""
                SELECT bi.item_type, bi.school, bi.color, SUM(bi.qty) as total_qty
                FROM bill_items bi
                GROUP BY bi.item_type, bi.school, bi.color
                ORDER BY total_qty DESC
                LIMIT 6
            """)
            rows = cur.fetchall()
            cur.close()
            if not rows:
                ttk.Label(self._fav_inner, text="لا توجد بيانات بعد").pack(side=tk.LEFT)
                return
            for r in rows:
                label = f"{r[0]} / {r[1]} / {r[2]}"
                btn = ttk.Button(self._fav_inner, text=label,
                                 command=lambda it=r[0], sc=r[1], cl=r[2]: self._jump_to_favorite(it, sc, cl))
                btn.pack(side=tk.RIGHT, padx=3, pady=2)
                ToolTip(btn, f"الكمية المباعة: {r[3]}")
        except Exception:
            pass

    def _jump_to_favorite(self, item_type, school, color):
        """Jump directly to the size selection for a favorite item."""
        self._sel_item = item_type
        self._sel_school = school
        self._sel_color = color
        self._price_user_edited = False
        self._sync_filters_to_combos()
        self._render_sizes()
        show_toast(self, f"تم التحديد: {item_type} / {school} / {color}")

    def _render_items(self):
        """Show item types for the selected school (or all if none selected)."""
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._clear_grid()

        if self._sel_school:
            self._crumb_var.set(f"المدرسة: {self._sel_school}  ⟶  اختر النوع")
            try:
                items = sorted({
                    it for (it, cl) in self.db.list_items_for_school(self._sel_school) if it
                })
            except Exception:
                items = []
            self._mk_grid_buttons(items, self._select_item, cols=4)
            # Back to schools
            tk.Button(self._grid_host, text="◀ رجوع إلى المدارس", command=self._render_schools,
                      **self._BTN_GRAY).pack(anchor="w", padx=4, pady=4)
            self._bind_grid_scroll()
        else:
            self._crumb_var.set("اختر النوع")
            try:
                items = sorted({
                    r["item_type"]
                    for r in self.db.current_inventory({})
                    if r.get("item_type")
                })
            except Exception:
                items = []
            self._mk_grid_buttons(items, self._select_item, cols=4)
            self._bind_grid_scroll()


    def _select_item(self, item_type: str):
        self._sel_item = item_type
        self._price_user_edited = False
        self._sync_filters_to_combos()
        # After choosing item, pick color next
        self._render_colors()

    def _select_school(self, school: str):
        self._sel_school = school
        self._price_user_edited = False
        self._sync_filters_to_combos()
        # after choosing a school, next pick item type
        self._render_items()

    def _render_colors(self):
        self._sel_color = None
        self._sel_size = None
        self._crumb_var.set(f"{self._sel_school}  ⟶  {self._sel_item}  ⟶  اختر اللون")
        self._clear_grid()
        pairs = self.db.list_items_for_school(self._sel_school or "")
        colors = sorted({cl for (it, cl) in pairs if it == self._sel_item})
        self._mk_grid_buttons(colors, self._select_color, cols=4)
        # Back to items
        tk.Button(self._grid_host, text="◀ رجوع إلى الأنواع", command=self._render_items,
                  **self._BTN_GRAY).pack(anchor="w", padx=4, pady=4)
        self._bind_grid_scroll()

    def _select_color(self, color: str):
        self._sel_color = color
        self._price_user_edited = False
        self._sync_filters_to_combos()
        self._render_sizes()

    @staticmethod
    def _spec_match(a: str, b: str) -> bool:
        return (a or "").strip().casefold() == (b or "").strip().casefold()

    def _pending_stock_out_qty(self, school: str, item: str, color: str, size: str) -> int:
        """Qty already on the draft bill that will deduct warehouse stock (excludes مصنع مباشر lines)."""
        total = 0
        for ln in self.bill_lines:
            if bool(ln.get("allow_factory_fill")):
                continue
            if not self._spec_match(ln.get("school"), school):
                continue
            if not self._spec_match(ln.get("item_type"), item):
                continue
            if not self._spec_match(ln.get("color"), color):
                continue
            if not self._spec_match(ln.get("size"), size):
                continue
            try:
                total += int(ln.get("qty") or 0)
            except Exception:
                pass
        return total

    def _refresh_size_grid_if_current(self, *, preserve_size: Optional[str] = None) -> None:
        if self._sel_school and self._sel_item and self._sel_color:
            try:
                self._render_sizes(preserve_size=preserve_size)
            except Exception:
                pass

    def _render_sizes(self, preserve_size: Optional[str] = None):
        self._sel_size = None
        self._crumb_var.set(
            f"{self._sel_school}  ⟶  {self._sel_item}  ⟶  {self._sel_color}  ⟶  اختر المقاس"
        )
        self._clear_grid()

        # Get sizes with counts/prices
        raw_sizes = self._get_sizes_for_bill(
            self._sel_school or "",
            self._sel_item or "",
            self._sel_color or "",
        )

        # Build cache with counts/prices if stock exists
        stock_rows = {}
        try:
            for r in self.db.current_inventory({
                "school": self._sel_school,
                "item_type": self._sel_item,
                "color": self._sel_color,
            }):
                sz = str(r.get("size") or "").strip()
                if not sz:
                    continue
                stock_rows.setdefault(sz, {
                    "count": 0,
                    "last_price": r.get("unit_price"),
                })
                stock_rows[sz]["count"] += int(r.get("count") or 0)
        except Exception:
            pass

        self._sizes_cache = []
        for sz in raw_sizes:
            r = stock_rows.get(sz, {})
            base = int(r.get("count", 0) or 0)
            pending = self._pending_stock_out_qty(
                self._sel_school or "", self._sel_item or "", self._sel_color or "", str(sz))
            eff = max(0, base - pending)
            self._sizes_cache.append({
                "size": sz,
                "count": eff,
                "last_price": r.get("last_price"),
            })

        # --- style for zero-count buttons (light red) ---
        try:
            s = ttk.Style(self)
            s.configure("Zero.TButton",
                        background="#f8d7da",   # light red
                        foreground="#000000",
                        bordercolor="#b91c1c",
                        focusthickness=1, focuscolor="#b91c1c", padding=6)
            s.map("Zero.TButton",
                background=[("active", "#f1b0b7"), ("pressed", "#2ee956")],
                bordercolor=[("focus", "#b91c1c")])
        except Exception:
            pass

        # Build the grid (wrap at 5 columns so it fills vertical space)
        cols = 4
        row = None
        for i, r in enumerate(self._sizes_cache):
            if i % cols == 0:
                row = ttk.Frame(self._grid_host)
                row.pack(fill=tk.X, padx=4, pady=2)

            sz = str(r.get("size") or "")
            cnt = int(r.get("count") or 0)
            label = f"{sz} ({cnt})"
            style = "Zero.TButton" if cnt == 0 else "TButton"

            btn = ttk.Button(row, text=label, style=style)
            btn.configure(command=lambda v=label, b=btn: self._on_size_click(v, b))
            btn.pack(side=tk.LEFT, padx=3, pady=3)

            # Double-click: quick add +1 to bill
            btn.bind("<Double-1>", lambda e, v=label, b=btn: self._on_size_double_click(v, b))
            ToolTip(btn, f"نقر: اختيار | نقر مزدوج: إضافة سريعة +1")

            # save button reference
            self._size_btns[label] = (btn, (cnt == 0))  # second value tracks zero-count


        # Back link
        tk.Button(self._grid_host, text="◀ رجوع إلى الألوان", command=self._render_colors,
                  **self._BTN_GRAY).pack(anchor="w", padx=4, pady=4)

        if preserve_size:
            psz = str(preserve_size).strip()
            for lbl in list(self._size_btns.keys()):
                if lbl.startswith(psz + " ("):
                    btn, _ = self._size_btns[lbl]
                    self._on_size_click(lbl, btn)
                    break

        self._bind_grid_scroll()

    def _on_size_click(self, label_with_count: str, btn):
        # 1) Extract raw size
        size = label_with_count.split(" (", 1)[0].strip()
        self._sel_size = size
        self._price_user_edited = False

        # 2) Reset previous selected button
        if self._selected_size_btn is not None:
            old_btn, was_zero = self._selected_size_btn
            if old_btn.winfo_exists():
                old_btn.configure(style="Zero.TButton" if was_zero else "TButton")

        # 3) Mark this button as selected (GREEN)
        btn.configure(style="SizeSelected.TButton")

        # 4) Remember selected button
        was_zero = "(0)" in label_with_count
        self._selected_size_btn = (btn, was_zero)

        # 5) Compute price
        try:
            computed = self._compute_price_for_size(size)
            if computed is not None and not self._price_user_edited:
                self.price_var.set(f"{float(computed):.2f}")
                self._price_user_edited = False
            else:
                if not self._price_user_edited and computed is None:
                    self.price_var.set("")
        except Exception:
            pass

    def _on_size_double_click(self, label_with_count: str, btn):
        """Double-click on a size button: select it and add +1 to the bill immediately."""
        self._on_size_click(label_with_count, btn)
        # Auto-set qty to 1 and try to add
        self.qty_var.set("1")
        price_txt = (self.price_var.get() or "").strip()
        if price_txt:
            self._add_current_selection()
            show_toast(self, f"+1 {self._sel_size}", duration=1200)
        else:
            show_toast(self, "لا يوجد سعر محدد - اختر السعر أولاً", bg="#f59e0b", fg="#000")

    def _compute_price_for_size(self, target_size: str) -> Optional[float]:

        import math
        try:
            def _ceil_to_next_10(x: float) -> float:
                # Round up to the next multiple of 10, keep as an integer-like float.
                return float(int(math.ceil(x / 10.0) * 10))

            # Map size -> last_price from cached sizes
            known = {}
            for r in (self._sizes_cache or []):
                s = str(r.get("size") or "")
                lp = r.get("last_price")
                if lp is not None:
                    try:
                        known[s] = float(lp)
                    except Exception:
                        pass

            # 1) exact match (history) — return as-is (preserve original value)
            if target_size in known:
                return float(known[target_size])

            # Build index mapping from sizes cache order
            index_of = {s: i for i, s in enumerate([r.get("size") for r in (self._sizes_cache or [])])}

            # Collect (idx, price) for sizes we know (and that exist in index_of)
            points = []
            for s, p in known.items():
                if s in index_of:
                    points.append((index_of[s], float(p)))

            if len(points) >= 2:
                # Linear least-squares fit (simple formula without numpy)
                xs = [float(x) for x, _ in points]
                ys = [float(y) for _, y in points]
                mean_x = sum(xs) / len(xs)
                mean_y = sum(ys) / len(ys)
                cov = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
                varx = sum((x - mean_x) ** 2 for x in xs)
                slope = 0.0 if varx == 0 else cov / varx
                intercept = mean_y - slope * mean_x

                # predict using target index if available
                if target_size in index_of:
                    tx = float(index_of[target_size])
                    pred = slope * tx + intercept
                    # sanity: if pred is negative or nan, ignore and fallback
                    if pred is not None and not (pred != pred) and pred >= 0:
                        # **round up to next multiple of 10**
                        return _ceil_to_next_10(pred)

            # If we reach here and have at least one known price -> pick nearest known by index
            if points and target_size in index_of:
                tgt_idx = index_of[target_size]
                nearest = min(points, key=lambda ip: abs(ip[0] - tgt_idx))
                # nearest price derived from history — treat as a calculated fallback and ceil to next 10
                return _ceil_to_next_10(nearest[1])

            # If points exist but target not indexable -> return the median known price (ceil to next 10)
            if points:
                prices = sorted([p for _, p in points])
                mid = prices[len(prices) // 2]
                return _ceil_to_next_10(mid)

            # 4) fallback: try DB default / general last price (pass empty size)
            try:
                p = self.db.last_price_for_specs(self._sel_item or "", self._sel_school or "", self._sel_color or "", "")
                if p is not None:
                    return float(round(float(p), 2))
            except Exception:
                pass

            # optional: try db.default_price_for_item if that helper exists
            try:
                if hasattr(self.db, "default_price_for_item"):
                    p = self.db.default_price_for_item(self._sel_item or "")
                    if p is not None:
                        return float(round(float(p), 2))
            except Exception:
                pass

        except Exception:
            pass

        return None

    def _select_size(self, label_with_count: str):
        # label is "SIZE (count)" -> extract raw size
        size = label_with_count.split(" (", 1)[0].strip()
        self._sel_size = size
        self._price_user_edited = False

        # Compute price according to priority:
        #  1) exact history, 2) estimate from other sizes, 3) default price
        try:
            computed = self._compute_price_for_size(size)
            # Respect manual user edits: if user edited the price we don't overwrite it.
            if computed is not None and not getattr(self, "_price_user_edited", False):
                self.price_var.set(f"{float(computed):.2f}")
                # mark programmatic set as NOT user-edited
                self._price_user_edited = False
            else:
                # if no computed price and price not user-edited, clear the box
                if not getattr(self, "_price_user_edited", False) and computed is None:
                    self.price_var.set("")
        except Exception:
            pass


    # ---------------- Bill ops ----------------
    def _add_current_selection(self):
        # validate selection path
        if not all([self._sel_school, self._sel_item, self._sel_color, self._sel_size]):
            messagebox.showwarning("اختر أولًا", "اختر المدرسة ثم النوع ثم اللون ثم المقاس.")
            return
        # qty
        try:
            qty = int((self.qty_var.get() or "0").strip())
            if qty <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("كمية غير صالحة", "الكمية يجب أن تكون عددًا صحيحًا موجبًا.")
            return
        # price
        try:
            price = float((self.price_var.get() or "").strip())
            if price < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("سعر غير صالح", "أدخل سعرًا رقميًا غير سالب.")
            return

        user_set_flag = bool(getattr(self, "_price_user_edited", False))
        allow_factory = bool(self.instant_mode.get())
        line = {
            "item_type":  self._sel_item,
            "school":     self._sel_school,
            "color":      self._sel_color,
            "size":       self._sel_size,
            "unit_price": price,
            "qty":        qty,
            "allow_factory_fill": allow_factory,
            "has_badge":  0,          # will be refined below if unique candidate exists
            "user_set_price": user_set_flag,
            # we'll also set warehouse_no/package_no below when uniquely determined
        }

        # ---- NEW: prefill WH/PKG/BADGE when uniquely determined ----
        if allow_factory:
            # factory: explicitly mark 0/0 so they appear in the table
            line["warehouse_no"] = 0
            line["package_no"]   = 0
            # keep has_badge as typed (defaults to 0) or leave 0
        else:
            # look for stock candidates for this exact spec
            cands = self.db.search_stocks({
                "item_type": self._sel_item,
                "school":    self._sel_school,
                "color":     self._sel_color,
                "size":      self._sel_size,
            })
            # restrict to rows that still have quantity
            cands = [r for r in cands if int(r.count) > 0]
            unique_triplets = {(r.warehouse_no, r.package_no, int(r.has_badge or 0)) for r in cands}
            if len(unique_triplets) == 1:
                w, p, b = next(iter(unique_triplets))
                line["warehouse_no"] = int(w)
                line["package_no"]   = int(p)
                line["has_badge"]    = int(b)
            # else: multiple packages match; leave wh/pkg unset so the table shows blank

        # ---- UPDATED: merge rule also considers wh/pkg/badge (when present) ----
        def _same_line(x: dict) -> bool:
            same_specs = (
                x.get("stock_id") is None
                and x.get("item_type") == line["item_type"]
                and x.get("school")    == line["school"]
                and x.get("color")     == line["color"]
                and x.get("size")      == line["size"]
                and float(x.get("unit_price", 0.0)) == price
                and bool(x.get("user_set_price")) is user_set_flag
                and bool(x.get("allow_factory_fill")) is allow_factory
            )
            # treat missing as None; only require equality if either line set a value
            wh_ok  = (x.get("warehouse_no") or None) == (line.get("warehouse_no") or None)
            pkg_ok = (x.get("package_no")  or None) == (line.get("package_no")  or None)
            badge_ok = int(x.get("has_badge") or 0) == int(line.get("has_badge") or 0)
            return same_specs and wh_ok and pkg_ok and badge_ok


        for existing in self.bill_lines:
            if _same_line(existing):
                existing["qty"] += qty
                break
        else:
            self.bill_lines.append(line)

        # refresh UI: keep entries as-is, just reset qty to 1
        self._sync_bill_table()
        self.qty_var.set("1")
        self._price_user_edited = False
        self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _remove_bill_line(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.bill_lines):
            self.bill_lines.pop(idx)
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _clear_bill(self):
        self.bill_lines.clear()
        self._sync_bill_table()
        self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _increment_bill_line(self):
        """Increase qty of selected bill line by 1."""
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.bill_lines):
            self.bill_lines[idx]["qty"] = int(self.bill_lines[idx]["qty"]) + 1
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _decrement_bill_line(self):
        """Decrease qty of selected bill line by 1 (min 1)."""
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.bill_lines):
            q = int(self.bill_lines[idx]["qty"])
            if q > 1:
                self.bill_lines[idx]["qty"] = q - 1
                self._sync_bill_table()
                self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _sync_bill_table(self):
        self.bill_table.delete(*self.bill_table.get_children())
        total = 0.0
        for idx, ln in enumerate(self.bill_lines):
            line_total = float(ln["unit_price"]) * int(ln["qty"])
            total += line_total
            self.bill_table.insert(
                "", tk.END, iid=str(idx),
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"],
                        f"{float(ln['unit_price']):.2f}", ln["qty"], f"{line_total:.2f}")
            )
        self.total_var.set(f"{total:.2f}")
        apply_zebra_tags(self.bill_table)


    # ---------------- Finalize / print ----------------
    def _finalize_bill(self):
        if not self.bill_lines:
            show_toast(self, "أضف بندًا واحدًا على الأقل", bg="#f59e0b", fg="#000")
            return
        # Customer is required
        customer_raw = (self.customer.get() or "").strip()
        if not customer_raw:
            show_toast(self, "يرجى اختيار اسم عميل أولاً", bg="#dc2626")
            return

        # Phase 3: "فرع: <branch ui name>" or raw device name → shipment.
        known_pos = set(self.db.list_known_pos_device_names() or [])
        norm_target = canonical_branch_device_name(customer_raw, known_pos)
        if norm_target:
            target_pos = norm_target
            customer = branch_customer_label(norm_target)
        elif customer_raw in known_pos:
            # Accept a raw device name too, not only the prefixed dropdown
            # value, so manually typed branch names still sync as shipments.
            target_pos = customer_raw
            customer = branch_customer_label(customer_raw)
        else:
            target_pos = None
            customer = customer_raw

        # ===== CONFIRMATION OVERLAY =====
        total = self.total_var.get()
        n_items = len(self.bill_lines)

        overlay = tk.Toplevel(self)
        overlay.title("تأكيد الفاتورة" if not target_pos else "تأكيد شحنة إلى فرع")
        overlay.transient(self.winfo_toplevel())
        overlay.grab_set()
        overlay.resizable(False, False)

        frm = ttk.Frame(overlay, padding=20)
        frm.pack(fill=tk.BOTH, expand=True)

        header_text = "مراجعة الفاتورة قبل الحفظ" if not target_pos else "مراجعة شحنة الفرع قبل الحفظ"
        ttk.Label(frm, text=header_text, font=("Segoe UI", 13, "bold")).pack(pady=(0, 12))
        if target_pos:
            ttk.Label(frm, text=f"الفرع المُستلِم: {target_pos}",
                      font=("Segoe UI", 11, "bold"),
                      foreground="#0ea5e9").pack(anchor="w", padx=8)
        else:
            ttk.Label(frm, text=f"العميل: {customer}", font=("Segoe UI", 11)).pack(anchor="w", padx=8)
        ttk.Label(frm, text=f"عدد البنود: {n_items}", font=("Segoe UI", 11)).pack(anchor="w", padx=8)
        ttk.Label(frm, text=f"الإجمالي: {total}", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(4, 12))

        # Summary of items
        summary = ttk.Treeview(frm, columns=("item", "size", "qty", "price"), show="headings", height=min(6, n_items))
        for col, txt, w in [("item", "الصنف", 180), ("size", "المقاس", 70), ("qty", "الكمية", 60), ("price", "السعر", 80)]:
            summary.heading(col, text=txt)
            summary.column(col, width=w, anchor="center")
        for ln in self.bill_lines:
            summary.insert("", tk.END, values=(
                f"{ln['item_type']} / {ln['school']}",
                ln["size"], ln["qty"], f"{float(ln['unit_price']):.2f}"
            ))
        summary.pack(fill=tk.X, padx=8, pady=(0, 12))
        apply_zebra_tags(summary)
        _bind_mousewheel(summary)

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X)

        def _do_save(print_after):
            overlay.destroy()
            try:
                bill_id = self.db.create_bill(
                    customer, self.bill_lines, target_pos=target_pos
                )
            except Exception as ex:
                messagebox.showerror("فشل الحفظ", str(ex))
                return
            self._clear_bill()
            self._refresh_favorites()
            if print_after:
                self._direct_print_bill(bill_id)
                show_toast(self, f"الفاتورة #{bill_id} تم الحفظ والطباعة")
            elif target_pos:
                show_toast(self, f"تم إنشاء شحنة الفرع #{bill_id} إلى {target_pos}")
            else:
                show_toast(self, f"تم حفظ الفاتورة #{bill_id}")

        ttk.Button(btns, text="إلغاء", command=overlay.destroy).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="حفظ بدون طباعة", command=lambda: _do_save(False)).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="حفظ وطباعة", command=lambda: _do_save(True)).pack(side=tk.RIGHT, padx=4)

        # Center overlay on parent
        overlay.update_idletasks()
        pw = self.winfo_toplevel().winfo_width()
        ph = self.winfo_toplevel().winfo_height()
        px = self.winfo_toplevel().winfo_rootx()
        py = self.winfo_toplevel().winfo_rooty()
        ow = overlay.winfo_width()
        oh = overlay.winfo_height()
        overlay.geometry(f"+{px + (pw - ow) // 2}+{py + (ph - oh) // 2}")

    def _save_as_draft(self):
        if not self.bill_lines:
            show_toast(self, "أضف بند واحد على الأقل", bg="#f59e0b", fg="#000")
            return
        customer = (self.customer.get() or "").strip()
        if not customer:
            show_toast(self, "يرجى اختيار اسم عميل أولاً", bg="#dc2626")
            return
        try:
            bill_id = self.db.create_draft_bill(customer, self.bill_lines)
        except Exception as ex:
            messagebox.showerror("فشل الحفظ", str(ex))
            return
        self._clear_bill()
        show_toast(self, f"تم حفظ المسودة #{bill_id}")

    def _print_bill(self, bill_id: int):
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("فشل الطباعة", "لم يتم العثور على الفاتورة.")
            return
        tmp_dir = tempfile.gettempdir()
        path = os.path.join(tmp_dir, f"bill_{bill_id}.html")
        save_bill_as_html(path, bill, items)
        webbrowser.open_new_tab(f"file:///{path.replace(os.sep, '/')}")

    def _get_sizes_for_bill(self, school: str, item: str, color: str) -> List[str]:
        """
        Priority:
        1) Size profile (numeric + alpha)
        2) Available sizes from stock
        """
        profile = self.db.get_size_profile(item, school, color)

        sizes: List[str] = []

        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            sizes.extend(
                merged_numeric_size_labels_from_profile(r1s, r1e, r2s, r2e),
            )
            if has_alpha:
                sizes.extend(ALPHA_SIZES)
            return sizes


        try:
            rows = self.db.current_inventory({
                "school": school,
                "item_type": item,
                "color": color,
            })

            seen = set()
            sizes = []
            for r in rows:
                sz = str(r.get("size") or "").strip()
                if sz and sz not in seen:
                    seen.add(sz)
                    sizes.append(sz)

            return sizes
        except Exception:
            return []


    def _direct_print_bill(self, bill_id: int, copies: int = 2):
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("فشل الطباعة", "لم يتم العثور على الفاتورة.")
            return
        tmp_dir = tempfile.gettempdir()
        path = os.path.join(tmp_dir, f"bill_{bill_id}.html")
        save_bill_as_html(path, bill, items)
        _print_html_auto(path, copies=max(1, int(copies)), parent=self)

class FactoryItemDialog(tk.Toplevel):
    """
    Create an ad-hoc bill line that ships directly from the factory
    (no stock row needed). Writes a planned line into OutcomeFrame.bill_lines.
    """
    def __init__(self, master: tk.Widget, db: SqliteDatabase, preset: Optional[Dict[str, str]] = None):
        super().__init__(master)
        self.db = db
        self.title("إضافة بند من المصنع")
        self.transient(master)
        self.grab_set()
        self.geometry("520x420")
        self.resizable(False, False)

        preset = preset or {}
        frm = ttk.Frame(self, padding=10); frm.pack(fill=tk.BOTH, expand=True)

        # Specs (combos for consistency with the app)
        self.t = LabeledCombobox(frm, "النوع",   db, "item_type"); self.t.set(preset.get("item_type",""))
        self.s = LabeledCombobox(frm, "المدرسة", db, "school");    self.s.set(preset.get("school",""))
        self.c = LabeledCombobox(frm, "اللون",   db, "color");     self.c.set(preset.get("color",""))
        self.z = LabeledCombobox(frm, "المقاس",  db, "size");      self.z.set(preset.get("size",""))
        for i, w in enumerate((self.t, self.s, self.c, self.z)):
            w.grid(row=i//2, column=i%2, padx=6, pady=6, sticky="ew")
        # NEW: auto price detection on edit/select/typing
        for w in (self.t.cb, self.s.cb, self.c.cb, self.z.cb):
            w.bind("<<ComboboxSelected>>", lambda e: self._auto_fill_price(), add="+")
            w.bind("<FocusOut>",           lambda e: self._auto_fill_price(), add="+")
            w.bind("<KeyRelease>",         lambda e: self._auto_fill_price(), add="+")
        frm.columnconfigure(0, weight=1); frm.columnconfigure(1, weight=1)

        # WH/PKG (optional, just for labeling/tracking)
        self.whv = tk.StringVar(value=preset.get("warehouse_no",""))
        self.pkv = tk.StringVar(value=preset.get("package_no",""))

        row_wp = ttk.Frame(frm); row_wp.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,6))
        ttk.Label(row_wp, text="المخزن:").pack(side=tk.LEFT)
        ttk.Entry(row_wp, textvariable=self.whv, width=8).pack(side=tk.LEFT, padx=(4,10))
        ttk.Label(row_wp, text="العبوة:").pack(side=tk.LEFT)
        ttk.Entry(row_wp, textvariable=self.pkv, width=10).pack(side=tk.LEFT, padx=(4,0))

        # Price/Qty + badge
        self.qv = tk.StringVar(value=preset.get("qty","1"))
        self.pv = tk.StringVar(value=preset.get("unit_price",""))
        self.badge = tk.BooleanVar(value=bool(int(preset.get("has_badge","0") or 0)))

        grid2 = ttk.Frame(frm); grid2.grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        ttk.Label(grid2, text="الكمية:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(grid2, textvariable=self.qv, width=12).grid(row=0, column=1, sticky="w")
        ttk.Label(grid2, text="سعر الوحدة:").grid(row=0, column=2, sticky="e", padx=12, pady=4)
        ttk.Entry(grid2, textvariable=self.pv, width=12).grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(grid2, text="بادج", variable=self.badge).grid(row=0, column=4, sticky="w", padx=(12,0))

        # Buttons
        btns = ttk.Frame(frm); btns.grid(row=4, column=0, columnspan=2, sticky="e", padx=6, pady=(10,0))
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="إضافة", command=self._ok).pack(side=tk.RIGHT, padx=8)

        # Autofill color from school (like the rest of the app)
        # self.s.cb.bind("<<ComboboxSelected>>", lambda e: self._maybe_autofill_color(), add="+")
        # self.s.cb.bind("<FocusOut>",           lambda e: self._maybe_autofill_color(), add="+")
        self.after(10, lambda: (self.t.cb.focus_set(), self._auto_fill_price()))

        self._result: Optional[Dict[str, str]] = None

    # def _maybe_autofill_color(self):
    #     school = (self.s.get() or "").strip()
    #     if not school:
    #         return
    #     try:
    #         last_color = self.db.last_color_for_school(school)
    #         if last_color:
    #             self.c.set(last_color)
    #     except Exception:
    #         pass
    # AFTER (NEW method inside FactoryItemDialog)
    def _auto_fill_price(self):
        """
        Prefill price for current specs if the price box is empty.
        Why: keep 'factory item' dialog consistent with Income auto-price behavior.
        """
        if (self.pv.get() or "").strip():
            return  # respect user-entered price
        it = (self.t.get() or "").strip()
        sc = (self.s.get() or "").strip()
        cl = (self.c.get() or "").strip()
        sz = (self.z.get() or "").strip()
        if not (it and sc and cl and sz):
            return
        try:
            p = self.db.last_price_for_specs(it, sc, cl, sz)
            if p is not None:
                self.pv.set(f"{p:.2f}")
        except Exception:
            # silent: don't block user flow
            pass

    def _ok(self):
        try:
            item_type = self.t.get() or self._err("النوع مطلوب")
            school    = self.s.get() or self._err("المدرسة مطلوبة")
            color     = self.c.get() or self._err("اللون مطلوب")
            size      = self.z.get() or self._err("المقاس مطلوب")
            qty       = int((self.qv.get() or "0").strip());   assert qty > 0
            price     = float((self.pv.get() or "").strip());  assert price >= 0.0
            wh        = int(self.whv.get()) if (self.whv.get() or "").strip() else 0
            pkg       = int(self.pkv.get()) if (self.pkv.get() or "").strip() else 0
        except AssertionError:
            messagebox.showerror("بيانات غير صالحة", "تحقق من الكمية (>0) والسعر (>=0).", parent=self); return
        except Exception as ex:
            messagebox.showerror("بيانات ناقصة", str(ex), parent=self); return

        self._result = {
            "item_type": item_type, "school": school, "color": color, "size": size,
            "warehouse_no": wh, "package_no": pkg,
            "unit_price": price, "qty": qty,
            "allow_factory_fill": True,          # key point: force factory path
            "has_badge": 1 if self.badge.get() else 0,
        }
        self.destroy()

    @staticmethod
    def _err(msg: str) -> None:
        raise RuntimeError(msg)

    def run(self) -> Optional[Dict[str, str]]:
        self.wait_window()
        return self._result


def open_price_sync_audit_dialog(parent: tk.Misc, db: "SqliteDatabase", *, limit: int = 200) -> None:
    """Read-only UI for warehouse-side PRICE_UPDATE fan-out decisions."""
    dlg = tk.Toplevel(parent)
    dlg.title("سجل تعديلات أسعار الفروع")
    try:
        dlg.transient(parent.winfo_toplevel())
    except Exception:
        pass
    dlg.grab_set()

    top = ttk.Frame(dlg, padding=10)
    top.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        top,
        text="يعرض هذا السجل قرارات «تعديل السعر» التي تمت مزامنتها مع فروع البيع (POS) — بما فيها أحداث الفاتورة عند إدخال سعر يدوي.",
        wraplength=860,
    ).pack(anchor="w", pady=(0, 8))

    # Treeview must live inside `grid` — `top` already uses pack(); mixing grid+pack on one parent raises TclError.
    grid = ttk.Frame(top)
    grid.pack(fill=tk.BOTH, expand=True)

    cols = ("ts", "price", "mode", "targets", "uuids", "filters", "note")
    tv = ttk.Treeview(grid, columns=cols, show="headings", height=16)
    for c, t, w in [
        ("ts", "الوقت", 170),
        ("price", "السعر", 90),
        ("mode", "وضع السجل", 120),
        ("targets", "النطاق/الفروع", 260),
        ("uuids", "معرّفات الحدث", 260),
        ("filters", "عوامل التصفية", 320),
        ("note", "ملاحظة", 220),
    ]:
        tv.heading(c, text=t)
        tv.column(c, width=w, anchor="center")

    ysb = ttk.Scrollbar(grid, orient="vertical", command=tv.yview)
    xsb = ttk.Scrollbar(grid, orient="horizontal", command=tv.xview)
    tv.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
    tv.grid(row=0, column=0, sticky="nsew")
    ysb.grid(row=0, column=1, sticky="ns")
    xsb.grid(row=1, column=0, sticky="ew")
    grid.rowconfigure(0, weight=1)
    grid.columnconfigure(0, weight=1)
    apply_zebra_tags(tv)
    _bind_mousewheel(tv)

    def _short_json(txt: str, max_len: int = 220) -> str:
        s = (txt or "").strip()
        if len(s) <= max_len:
            return s
        return s[: max_len - 1] + "…"

    def _reload():
        tv.delete(*tv.get_children())
        rows = db.list_price_sync_audit(limit=limit)
        for r in rows:
            tv.insert(
                "",
                tk.END,
                values=(
                    str(r.get("created_at") or ""),
                    f"{float(r.get('new_price') or 0):.2f}",
                    str(r.get("sync_mode") or ""),
                    _short_json(str(r.get("targets_json") or ""), 260),
                    _short_json(str(r.get("event_uuids_json") or ""), 260),
                    _short_json(str(r.get("filters_json") or ""), 320),
                    str(r.get("note") or ""),
                ),
            )

    btns = ttk.Frame(top)
    btns.pack(fill=tk.X, pady=(8, 0))
    ttk.Button(btns, text="تحديث", command=_reload).pack(side=tk.RIGHT)
    ttk.Button(btns, text="إغلاق", command=dlg.destroy).pack(side=tk.RIGHT, padx=(0, 8))

    _reload()
    dlg.update_idletasks()
    try:
        dlg.geometry("1100x640")
    except Exception:
        pass


# ------------------- Inventory Window -------------------

class InventoryWindow(tk.Toplevel):
    """
    Inventory window with single-field multi-select filters:
    - You can select multiple values for ONE field at a time (type/school/color/size/warehouse/package).
    - As soon as a field has a multi-selection, the other filters are disabled.
    - Clear button resets everything and re-enables all filters.
    """
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("المخزون")
        self.geometry("1000x560")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  المخزون", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)

        # Track multi selections
        self.multi: Dict[str, List[Any]] = {
            "item_type": [],
            "school": [],
            "color": [],
            "size": [],
            "warehouse_no": [],
            "package_no": [],
        }

        # Keep references to widgets to enable/disable them together
        self._multi_btns: Dict[str, ttk.Button] = {}
        self._field_widgets: Dict[str, tk.Widget] = {}

        self._build()

    # ---------- UI ----------

    def _build(self):
        filters = ttk.LabelFrame(self, text="تصنيف")
        filters.pack(fill=tk.X, padx=8, pady=8)

        # Controls
        self.f_type   = LabeledCombobox(filters, "النوع",    self.db, "item_type")
        self.f_school = LabeledCombobox(filters, "المدرسة", self.db, "school")
        self.f_color  = LabeledCombobox(filters, "اللون",    self.db, "color")
        self.f_size   = LabeledCombobox(filters, "المقاس",   self.db, "size")
        self.f_wh     = LabeledStaticCombo(
            filters, "رقم المخزن",
            values=["", *WAREHOUSE_NUMBER_DISPLAY_VALUES],
            value_map=WAREHOUSE_NUMBER_LABELS,
        )
        self.f_pkg    = LabeledEntry(filters, "رقم العبوة")

        def _constraints(exclude=None):
            d = {
                "item_type": self.f_type.get(),
                "school":    self.f_school.get(),
                "color":     self.f_color.get(),
                "size":      self.f_size.get(),
                "warehouse_no": self.f_wh.get() or None,
                "package_no":  self.f_pkg.get() or None,
            }
            if exclude:
                d.pop(exclude, None)
            return d

        self.f_type.set_supplier(  lambda: self.db.get_distinct_filtered("item_type", _constraints(exclude="item_type")))
        self.f_school.set_supplier(lambda: self.db.get_distinct_filtered("school",    _constraints(exclude="school")))
        self.f_color.set_supplier( lambda: self.db.get_distinct_filtered("color",     _constraints(exclude="color")))
        self.f_size.set_supplier(  lambda: self.db.get_distinct_filtered("size",      _constraints(exclude="size")))

        def _refresh_all_suppliers(*_):
            self.f_type.refresh_values()
            self.f_school.refresh_values()
            self.f_color.refresh_values()
            self.f_size.refresh_values()

        for w in (self.f_type, self.f_school, self.f_color, self.f_size):
            for ev in ("<<ComboboxSelected>>", "<KeyRelease>"):
                w.cb.bind(ev, lambda e: (_refresh_all_suppliers(), self._schedule_refresh()), add="+")
        self.f_wh.cb.bind("<<ComboboxSelected>>", lambda e: (_refresh_all_suppliers(), self._schedule_refresh()), add="+")
        self.f_pkg.var.trace_add("write", lambda *_: (_refresh_all_suppliers(), self._schedule_refresh()))


        widgets = [self.f_type, self.f_school, self.f_color, self.f_size, self.f_wh, self.f_pkg]
        fields  = ["item_type","school","color","size","warehouse_no","package_no"]

        # Place widgets (same grid as before)
        for i, (w, fld) in enumerate(zip(widgets, fields)):
            r, c = divmod(i, 3)
            w.grid(row=r*2, column=c, padx=6, pady=(6, 0), sticky="ew")  # widget row
            filters.columnconfigure(c, weight=1)

            # Add "اختيار متعدد…" button under each filter
            from functools import partial  # put this at the top of the file once

            def _safe_open(field):
                try:
                    self._open_multi_dialog(field)
                except Exception as ex:
                    messagebox.showerror("خطأ", str(ex), parent=self)

            btn = ttk.Button(filters, text="اختيار متعدد…", command=partial(_safe_open, fld))
            btn.grid(row=r*2+1, column=c, sticky="w", padx=6, pady=(2, 8))
            self._multi_btns[fld] = btn

            # Keep a handle to enable/disable
            # Map to inner entry/combobox widget for state toggling
            if hasattr(w, "cb"):   # LabeledCombobox/LabeledStaticCombo
                self._field_widgets[fld] = w.cb
            elif hasattr(w, "entry"):  # LabeledEntry
                self._field_widgets[fld] = w.entry
            else:
                self._field_widgets[fld] = w

        # Live refresh when single-value filters change
        for cb in (self.f_type.cb, self.f_school.cb, self.f_color.cb, self.f_size.cb):
            cb.bind("<<ComboboxSelected>>", lambda e: self._schedule_refresh(), add="+")
            cb.bind("<KeyRelease>",         lambda e: self._schedule_refresh(), add="+")
        self.f_wh.cb.bind("<<ComboboxSelected>>", lambda e: self._schedule_refresh(), add="+")
        self.f_pkg.var.trace_add("write", lambda *_: self._schedule_refresh())

        # Buttons
        btns = ttk.Frame(filters)
        btns.grid(row=4, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))
        _b1 = ttk.Button(btns, text="بحث", command=self._refresh); _b1.pack(side=tk.LEFT)
        ToolTip(_b1, "تطبيق الفلاتر وعرض النتائج")
        _b2 = ttk.Button(btns, text="مسح", command=self._clear_all); _b2.pack(side=tk.LEFT, padx=8)
        ToolTip(_b2, "مسح جميع الفلاتر")
        _b3 = ttk.Button(btns, text="تصدير إلى إكسل", command=self._export_excel); _b3.pack(side=tk.LEFT, padx=8)
        ToolTip(_b3, "تصدير النتائج الحالية إلى ملف إكسل")
        _b4 = ttk.Button(btns, text="طباعة جداول المقاسات", command=self._print_size_sheets); _b4.pack(side=tk.LEFT, padx=8)
        ToolTip(_b4, "طباعة جداول المقاسات لكل نوع/مدرسة/لون")
        _b5 = ttk.Button(btns, text="تعديل نطاقات المقاسات…", command=self._edit_size_ranges_dialog); _b5.pack(side=tk.LEFT, padx=8)
        ToolTip(_b5, "تعديل نطاقات المقاسات المعروضة في الجداول")


        # Table
        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "wh", "pkg", "badge", "price", "count", "value"),
            show="headings",
            selectmode="extended",
        )
        for col, txt, w in [
            ("id","المعرّف",60), ("type","النوع",140), ("school","المدرسة",160),
            ("color","اللون",80), ("size","المقاس",70), ("wh","المخزن",50),
            ("pkg","العبوة",60), ("badge","بادج",60), ("price","السعر",80), ("count","الكمية",70), ("value","القيمة",90),
        ]:

            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        xsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        add_context_menu(self.table)
        _bind_mousewheel(self.table)

        # Totals / actions
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.sum_qty = tk.StringVar(value="0")
        self.sum_val = tk.StringVar(value="0.00")
        ttk.Label(bar, text="إجمالي الكمية:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_qty, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(bar, text="إجمالي القيمة:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_val, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        _ba = ttk.Button(bar, text="سجل أسعار الفروع…", command=self._open_price_branch_audit)
        _ba.pack(side=tk.RIGHT)
        ToolTip(_ba, "عرض سجل مزامنة تعديلات الأسعار إلى فروع البيع (POS)")
        _bp = ttk.Button(bar, text="تعديل السعر…", command=self._edit_price_dialog); _bp.pack(side=tk.RIGHT, padx=(8, 0))
        ToolTip(_bp, "تعديل سعر الأصناف المطابقة للفلاتر")
        _bs = ttk.Button(bar, text="تعديل المواصفات…", command=self._edit_specs_dialog); _bs.pack(side=tk.RIGHT, padx=(8, 0))
        ToolTip(_bs, "تعديل المخزن/العبوة/البادج للأصناف المطابقة")
        _bd = ttk.Button(bar, text="حذف المحدد…", command=self._remove_selected_dialog); _bd.pack(side=tk.RIGHT, padx=(8, 0))
        ToolTip(_bd, "حذف الصفوف المحددة من المخزون")

        self._refresh()

    def _open_price_branch_audit(self):
        open_price_sync_audit_dialog(self, self.db)

    # ---------- Multi-select dialog ----------
    def _get_selected_profile_keys(self):
        """
        Returns (item_type, school, color) if selection is valid,
        otherwise shows message and returns None.
        """
        sel = self.table.selection()
        if not sel:
            return None

        keys = set()
        for iid in sel:
            vals = self.table.item(iid, "values")
            # table columns:
            # 0=id, 1=type, 2=school, 3=color, 4=size, ...
            keys.add((vals[1], vals[2], vals[3]))

        if len(keys) != 1:
            messagebox.showwarning(
                "تحديد غير صالح",
                "يجب أن تكون الصفوف المحددة من نفس (النوع، المدرسة، اللون).",
                parent=self
            )
            return None

        return keys.pop()
    def _edit_size_ranges_dialog(self):

        # 1) Try to get from selected rows
        picked = self._get_selected_profile_keys()

        if picked:
            item_type, school, color = picked
        else:
            # 2) Fallback to filter controls (OLD behavior)
            item_type = (self.f_type.get() or "").strip()
            school    = (self.f_school.get() or "").strip()
            color     = (self.f_color.get() or "").strip()

            if not (item_type and school and color):
                messagebox.showwarning(
                    "حدد الصنف",
                    "حدد صفوفًا من الجدول أو اختر (النوع، المدرسة، اللون) أولاً.",
                    parent=self
                )
                return


        # Fetch existing profile (if any)

        dlg = tk.Toplevel(self)
        dlg.title("تعديل نطاقات المقاسات")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frm,
            text=f"{item_type} / {school} / {color}",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        # --- Variables ---
        r1_start = tk.StringVar()
        r1_end   = tk.StringVar()
        r2_start = tk.StringVar()
        r2_end   = tk.StringVar()
        has_alpha_var = tk.BooleanVar(value=False)




        # --- UI layout ---
        def _row(lbl, var, r, c):
            ttk.Label(frm, text=lbl).grid(row=r, column=c, sticky="e", padx=4, pady=4)
            ttk.Entry(frm, textvariable=var, width=8).grid(row=r, column=c+1, sticky="w", padx=4)

        ttk.Label(frm, text="النطاق الأول").grid(row=0, column=0, sticky="w")

        r1_var = tk.StringVar()
        r1_combo = ttk.Combobox(
            frm,
            textvariable=r1_var,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=12,
        )
        r1_combo.grid(row=0, column=1, columnspan=2, sticky="w")


        ttk.Label(frm, text="النطاق الثاني (اختياري)").grid(row=1, column=0, sticky="w")

        r2_var = tk.StringVar()
        r2_combo = ttk.Combobox(
            frm,
            textvariable=r2_var,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=12,
        )
        r2_combo.grid(row=1, column=1, columnspan=2, sticky="w")
        profile = self.db.get_size_profile(item_type, school, color)

        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile

            if r1s is not None and r1e is not None:
                r1_var.set(f"{r1s} → {r1e}")

            if r2s is not None and r2e is not None:
                r2_var.set(f"{r2s} → {r2e}")

            has_alpha_var.set(bool(has_alpha))

        ttk.Checkbutton(
            frm,
            text="تفعيل المقاسات بالحروف (S / M / L ...)",
            variable=has_alpha_var
        ).grid(row=2, column=0, columnspan=3, sticky="w")


        # --- Buttons ---
        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=4, sticky="e", pady=(10,0))

        ALLOWED_RANGE_KEYS = set(ALLOWED_NUMERIC_RANGES.keys())

        def _parse_range_label(v: str):
            v = (v or "").strip()
            if not v:
                return None, None
            a, b = v.split("→")
            return int(a.strip()), int(b.strip())


        def on_save():
            try:
                r1s, r1e = _parse_range_label(r1_var.get())
                r2s, r2e = _parse_range_label(r2_var.get())

                self.db.upsert_size_profile(
                    item_type,
                    school,
                    color,
                    r1_start=r1s,
                    r1_end=r1e,
                    r2_start=r2s,
                    r2_end=r2e,
                    has_alpha=has_alpha_var.get(),
                )
                show_toast(dlg, "تم حفظ نطاقات المقاسات")
                dlg.destroy()
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)


        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=on_save).pack(side=tk.RIGHT, padx=6)

    def _open_multi_dialog(self, field: str):
        # Source values for each field
        if field in ("item_type", "school", "color", "size"):
            try:
                values = self.db.get_distinct(field)
            except Exception:
                values = []
        elif field == "warehouse_no":
            values = list(WAREHOUSE_NUMBER_DISPLAY_VALUES)
        else:  # package_no
            try:
                cur = self.db.conn.execute("SELECT DISTINCT package_no FROM stocks ORDER BY package_no ASC")
                values = [str(r[0]) for r in cur.fetchall()]
            except Exception:
                values = []

        preselected = [str(x) for x in self.multi[field]]
        if field == "warehouse_no":
            preselected = [warehouse_display_label(x) for x in preselected]
        dlg = MultiSelectDialog(self, title="اختيار متعدد", values=values, preselected=preselected)
        picked = dlg.run()
        if picked is None:
            return  # cancelled

        # Store (as natural type)
        if field in ("warehouse_no", "package_no"):
            if field == "warehouse_no":
                self.multi[field] = [int(warehouse_numeric_value(x)) for x in picked]
            else:
                self.multi[field] = [int(x) for x in picked]
        else:
            self.multi[field] = [str(x) for x in picked]

        # Update button label with count
        btn = self._multi_btns.get(field)
        if btn:
            btn.configure(text=f"اختيار متعدد… ({len(self.multi[field])})" if self.multi[field] else "اختيار متعدد…")

        # Enforce single-active-field rule in the UI
        self._enforce_single_active_field()
        self._refresh()

    def _enforce_single_active_field(self):
        active_fields = [k for k, v in self.multi.items() if v]
        # If one field has multi-values, disable other five single inputs + their multi buttons
        if len(active_fields) >= 1:
            active = active_fields[0]
            for fld, w in self._field_widgets.items():
                is_active = (fld == active)
                # Single-inputs: keep enabled only for active field when it has multi selected,
                # otherwise allow normal editing.
                if self.multi[active]:
                    state = ("normal" if is_active else "disabled")
                else:
                    state = "normal"
                try:
                    w.configure(state=state)
                except Exception:
                    pass

            for fld, b in self._multi_btns.items():
                if self.multi[active]:
                    b.configure(state=("normal" if fld == active else "disabled"))
                else:
                    b.configure(state="normal")

            # Also clear the *text* single-inputs for non-active fields (avoid accidental combos)
            if self.multi[active]:
                for fld in self._field_widgets:
                    if fld != active:
                        self._clear_single_field_text(fld)
        else:
            # No multi selections: everything enabled
            for w in self._field_widgets.values():
                try: w.configure(state="normal")
                except Exception: pass
            for b in self._multi_btns.values():
                b.configure(state="normal")

    def _clear_single_field_text(self, field: str):
        if field == "item_type": self.f_type.set("")
        elif field == "school": self.f_school.set("")
        elif field == "color": self.f_color.set("")
        elif field == "size": self.f_size.set("")
        elif field == "warehouse_no": self.f_wh.set("")
        elif field == "package_no": self.f_pkg.set("")

    # ---------- Filters / Refresh ----------
    def _print_size_sheets(self):
        try:
            rows = self.db.current_inventory(self._filters())
        except Exception as ex:
            messagebox.showerror("فشل", str(ex), parent=self)
            return

        if not rows:
            show_toast(self, "لا توجد صفوف مطابقة للطباعة", bg="#f59e0b")
            return
        from collections import defaultdict, OrderedDict

        # school -> (item_type, color) -> rows
        school_groups = OrderedDict()

        for r in rows:
            school = (r.get("school") or "").strip()
            item   = (r.get("item_type") or "").strip()
            color  = (r.get("color") or "").strip()

            school_groups.setdefault(school, OrderedDict())
            school_groups[school].setdefault((item, color), []).append(r)


        def build_size_ranges_from_profile(profile):
            """
            Returns:
                numeric_tables: List[List[str]]
                alpha_labels:   List[str]
            """
            numeric_tables = []
            alpha_labels = []

            if profile is None:
                return numeric_tables, alpha_labels

            r1_start, r1_end, r2_start, r2_end, has_alpha = profile

            merged = merged_numeric_size_labels_from_profile(
                r1_start, r1_end, r2_start, r2_end,
            )
            if merged:
                numeric_tables.append(merged[:])

            # ---- ALPHA SIZES (SEPARATE TABLE) ----
            if has_alpha:
                alpha_labels = ALPHA_SIZES[:]

            return numeric_tables, alpha_labels




        tables_html = []

        for sch, item_groups in school_groups.items():


            for (t, clr), items in item_groups.items():


                size_counts = defaultdict(int)
                used_sizes = set()

                for r in items:
                    sz = _normalize_size_label(r.get("size") or "")
                    size_counts[sz] += int(r.get("count") or 0)
                    used_sizes.add(sz)

                profile = self.db.get_size_profile(t, sch, clr)
                numeric_tables, alpha_labels = build_size_ranges_from_profile(profile)

                # 🔴 FALLBACK: no profile → use actual available sizes
                if not numeric_tables and not alpha_labels:
                    all_sizes = sorted(
                        { _normalize_size_label(r.get("size") or "") for r in items }
                    )

                    numeric = [s for s in all_sizes if s.isdigit()]
                    alpha   = [s for s in all_sizes if not s.isdigit()]

                    if numeric:
                        numeric_tables = [numeric]
                    if alpha:
                        alpha_labels = alpha

                def row_counts(labels):
                    out = []
                    for lbl in labels:
                        v = int(size_counts.get(lbl, 0))
                        out.append("" if v == 0 else str(v))
                    return out

                head = f"""
                <div class="hdr">
                    <span>النوع: {_html(t)}</span>
                    <span>المدرسة: {_html(sch)}</span>
                    <span>اللون: {_html(clr)}</span>
                </div>
                """

                def build_table(chunk):
                    return f"""
                    <table class="grid">
                    <tbody>
                        <tr>{''.join(f'<th>{_html(x)}</th>' for x in chunk)}</tr>
                        <tr>{''.join(f'<td class="num">{v}</td>' for v in row_counts(chunk))}</tr>
                        <tr>{''.join('<td>&nbsp;</td>' for _ in chunk)}</tr>
                    </tbody>
                    </table>
                    """

                tables = []

                # ---------- numeric tables ----------
                for numeric_labels in numeric_tables:
                    for i in range(0, len(numeric_labels), 15):
                        tables.append(build_table(numeric_labels[i:i + 15]))

                # ---------- alpha table ----------
                if alpha_labels:
                    tables.append("""
                    <div style="margin-top:6px;font-weight:600">المقاسات بالحروف</div>
                    """ + build_table(alpha_labels))

                tables_html.append(f'<section class="sheet">{head}{"".join(tables)}</section>')



            html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<title>جداول المقاسات</title>
<style>
@page {{ size: A4; margin: 12mm; }}
* {{ box-sizing: border-box; }}
body {{
    font-family: "Segoe UI", Tahoma, Arial, "Noto Sans Arabic", sans-serif;
    margin: 0;
    direction: rtl;
}}
.sheet {{ page-break-inside: avoid; margin-bottom: 10mm; }}
.hdr {{ display:flex; justify-content:space-between; font-weight:600; margin: 6px 2px 8px; }}
.grid {{ border-collapse: collapse; width: 100%; table-layout: fixed; margin-bottom: 6px; }}
.grid th, .grid td {{
    border: 1px solid #555;
    padding: 6px 4px;
    text-align: center;
}}
.grid th {{ background: #eee; }}
.num {{ font-variant-numeric: tabular-nums; }}
</style>
</head>
<body>
{''.join(tables_html)}
<script>
window.onload = function() {{
try {{ window.print(); }} catch(e) {{}}
}};
</script>
</body>
</html>
"""

        import tempfile, os
        path = os.path.join(
            tempfile.gettempdir(),
            f"size_sheets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
        )

        with open(path, "w", encoding="utf-8") as f:
            f.write(html)

        _print_html_auto(path, copies=1, parent=self)

    def _schedule_refresh(self, delay_ms: int = 250):
        if hasattr(self, "_inv_job") and self._inv_job:
            self.after_cancel(self._inv_job)
        self._inv_job = self.after(delay_ms, self._refresh)

    def _filters(self) -> Dict[str, Any]:
        """
        If any multi list is non-empty, only that field is used.
        Otherwise, use single inputs.
        """
        # Find active multi field (if any)
        active_multi = [k for k, v in self.multi.items() if v]
        if active_multi:
            fld = active_multi[0]
            # Build filters with only that multi field
            f: Dict[str, Any] = {
                "item_type": None, "school": None, "color": None, "size": None,
                "warehouse_no": None, "package_no": None,
            }
            f[fld] = self.multi[fld][:]
            return f

        # No multi: build from single-value controls
        return {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "color": self.f_color.get() or None,
            "size": self.f_size.get() or None,
            "warehouse_no": (self.f_wh.get() or None),
            "package_no": (self.f_pkg.get() or None),
        }

    def _refresh(self):
        try:
            rows = self.db.current_inventory(self._filters())
        except Exception as ex:
            messagebox.showerror("فشل البحث", str(ex), parent=self)
            return

        self.table.delete(*self.table.get_children())
        total_qty = 0
        total_value = 0.0
        for r in rows:
            self.table.insert(
                "", tk.END,
                values=(r["id"], r["item_type"], r["school"], r["color"], r["size"],
                        r["warehouse_no"], r["package_no"],
                        ("✓" if int(r.get("has_badge") or 0) else ""),
                        f"{float(r['unit_price']):.2f}",
                        r["count"], f"{float(r['value']):.2f}")
            )

            total_qty += int(r["count"])
            total_value += float(r["value"])
        self.sum_qty.set(str(total_qty))
        self.sum_val.set(f"{total_value:.2f}")
        apply_zebra_tags(self.table)

    def _clear_all(self):
        # Clear single inputs
        for w in (self.f_type, self.f_school, self.f_color, self.f_size):
            w.set("")
        self.f_wh.set("")
        self.f_pkg.set("")

        # Clear multi selections & button labels
        for k in self.multi.keys():
            self.multi[k] = []
        for b in self._multi_btns.values():
            b.configure(text="اختيار متعدد…")

        # Re-enable everything
        self._enforce_single_active_field()
        self._refresh()

    # ---------- Existing actions kept as-is ----------

    def _export_excel(self):
        try:
            db_rows = self.db.current_inventory(self._filters())
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)
            return

        path = filedialog.asksaveasfilename(
            title="تصدير المخزون إلى إكسل",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        try:
            self.db.export_inventory_excel(path, db_rows)
            show_toast(self, f"تم حفظ المخزون إلى إكسل بنجاح")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)

    def _remove_selected_dialog(self):
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning(
                "لم يتم التحديد",
                "اختر صفًا واحدًا أو أكثر من المخزون أولًا.",
                parent=self
            )
            return

        # collect selected rows
        rows = [self.table.item(i, "values") for i in sel]
        ids = [int(r[0]) for r in rows]

        # build label
        if len(rows) == 1:
            r = rows[0]
            label = f"{r[1]} / {r[2]} / {r[3]} / {r[4]}"
            available = int(r[9])
            title = "حذف من المخزون (يتطلب كلمة مرور)"
        else:
            label = f"عدد الصفوف المحددة: {len(rows)}"
            available = sum(int(r[9]) for r in rows)
            title = "حذف متعدد من المخزون (يتطلب كلمة مرور)"

        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=label, font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(frm, text=f"إجمالي المتوفر: {available}")\
            .grid(row=1, column=0, columnspan=2, sticky="w")

        ttk.Label(frm, text="الكمية المطلوب حذفها (فارغ = الكل):")\
            .grid(row=2, column=0, sticky="e", padx=4, pady=8)

        qty_var = tk.StringVar()
        ttk.Entry(frm, textvariable=qty_var, width=12)\
            .grid(row=2, column=1, sticky="w", padx=4, pady=8)

        ttk.Label(frm, text="كلمة المرور:")\
            .grid(row=3, column=0, sticky="e", padx=4, pady=4)

        pw_var = tk.StringVar()
        ttk.Entry(frm, textvariable=pw_var, show="*")\
            .grid(row=3, column=1, sticky="w", padx=4, pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))

        def on_ok():
            if not self.db.verify_admin_password(pw_var.get()):
                messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=dlg)
                return

            qtext = qty_var.get().strip()
            qty = None
            if qtext:
                try:
                    qty = int(qtext)
                    if qty <= 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror(
                        "كمية غير صالحة",
                        "أدخل عددًا موجبًا أو اتركه فارغًا للكل.",
                        parent=dlg
                    )
                    return

            try:
                total_removed = 0
                for stock_id in ids:
                    total_removed += self.db.remove_from_stock(
                        stock_id, qty, note="Admin remove (bulk)"
                    )

                dlg.destroy()
                show_toast(self, f"تم حذف {total_removed} وحدة من {len(ids)} صف(وف)")
                self._refresh()

            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حذف", command=on_ok).pack(side=tk.RIGHT, padx=6)

    # file: ui/edit_price_dialog.py

    def _edit_price_dialog(self):
        row = 0
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("لم يتم التحديد", "اختر صفًا من المخزون أولًا.", parent=self)
            return

        rows = [self.table.item(i, "values") for i in sel]
        first = rows[0]

        current_price = float(first[8])
        multi = len(rows) > 1

        dlg = tk.Toplevel(self)
        dlg.title("تعديل السعر")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # ---- Header ----
        if multi:
            ttk.Label(
                frm,
                text=f"عدد الصفوف المحددة: {len(rows)}",
                font=("Segoe UI", 10, "bold")
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 2))
            row += 1
            ttk.Label(
                frm,
                text="يمكنك تحديد صفوف بنفس النوع والمدرسة ومقاسات مختلفة؛ سيتم تحديث كل صف ومزامنة الفروع لكل مقاس على حدة عند تفعيل المزامنة.",
                wraplength=520,
                foreground="#64748B",
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
        else:
            ttk.Label(
                frm,
                text=f"الصنف: {first[1]} / {first[2]} / {first[3]} / {first[4]}"
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1

            ttk.Label(
                frm,
                text=f"المخزن/العبوة: {first[5]} / {first[6]}"
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1

            ttk.Label(
                frm,
                text=f"السعر الحالي: {current_price:.2f}"
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 10))
            row += 1

        # ---- New price ----
        ttk.Label(frm, text="السعر الجديد:").grid(row=row, column=0, sticky="e", padx=4, pady=6)
        price_var = tk.StringVar(value=f"{current_price:.2f}")
        ttk.Entry(frm, textvariable=price_var, width=16).grid(
            row=row, column=1, sticky="w", padx=4, pady=6
        )
        row += 1

        # ---- POS sync scope ----
        pos_names: List[str] = []
        try:
            pos_names = list(self.db.list_known_pos_device_names() or [])
        except Exception:
            pos_names = []

        sync_box = ttk.LabelFrame(frm, text="مزامنة سعر الفروع (POS)")
        sync_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 10))
        pos_sync_var = tk.StringVar(value="all-pos")

        ttk.Radiobutton(
            sync_box,
            text="كل فروع البيع",
            variable=pos_sync_var,
            value="all-pos",
        ).pack(anchor="w", padx=8, pady=(6, 2))

        ttk.Radiobutton(
            sync_box,
            text="فروع محددة…",
            variable=pos_sync_var,
            value="selected-pos",
        ).pack(anchor="w", padx=8, pady=2)

        lb_frame = ttk.Frame(sync_box)
        lb_frame.pack(fill=tk.BOTH, expand=True, padx=(18, 8), pady=(0, 6))
        pos_list = tk.Listbox(lb_frame, height=min(8, max(3, len(pos_names))), selectmode=tk.EXTENDED, exportselection=False)
        sb = ttk.Scrollbar(lb_frame, orient="vertical", command=pos_list.yview)
        pos_list.configure(yscrollcommand=sb.set)
        pos_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        for nm in pos_names:
            pos_list.insert(tk.END, nm)

        ttk.Radiobutton(
            sync_box,
            text="لا (تحديث المخزن فقط)",
            variable=pos_sync_var,
            value="none",
        ).pack(anchor="w", padx=8, pady=(2, 8))

        def _sync_ui_state(*_):
            st = (pos_sync_var.get() or "").strip()
            if st == "selected-pos":
                pos_list.configure(state="normal")
            else:
                pos_list.configure(state="disabled")

        try:
            pos_sync_var.trace_add("write", lambda *_: _sync_ui_state())
        except Exception:
            pass

        _sync_ui_state()
        row += 1

        scope_var = tk.StringVar(value="row")

        # ---- Scope (single-row only) ----
        if not multi:
            scope_box = ttk.LabelFrame(frm, text="النطاق")
            scope_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 10))

            ttk.Radiobutton(
                scope_box,
                text="الصف المحدد فقط",
                variable=scope_var,
                value="row"
            ).pack(anchor="w", padx=8, pady=4)

            ttk.Radiobutton(
                scope_box,
                text="كل الصفوف بنفس (النوع/المدرسة/المقاس) داخل نفس المخزن/العبوة",
                variable=scope_var,
                value="same_pkg"
            ).pack(anchor="w", padx=8, pady=4)

            row += 1

        # ---- OK handler ----
        def on_ok():
            try:
                new_price = float(price_var.get())
                if new_price < 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("قيمة غير صالحة", "أدخل سعرًا رقميًا غير سالب.", parent=dlg)
                return

            try:
                updated_total = 0

                sync_mode = (pos_sync_var.get() or "all-pos").strip()
                if sync_mode not in ("all-pos", "selected-pos", "none"):
                    sync_mode = "all-pos"

                selected_pos: List[str] = []
                if sync_mode == "selected-pos":
                    for i in pos_list.curselection():
                        try:
                            txt = str(pos_list.get(i) or "").strip()
                        except Exception:
                            txt = ""
                        if txt:
                            selected_pos.append(txt)
                    if not selected_pos:
                        messagebox.showwarning(
                            "اختر الفروع",
                            "اختر فرعًا واحدًا على الأقل من قائمة فروع البيع، أو غيّر خيار المزامنة إلى «كل الفروع».",
                            parent=dlg,
                        )
                        return

                emit_sync = sync_mode != "none"
                sync_devs = selected_pos if sync_mode == "selected-pos" else None

                if multi:
                    # Same (type, school, color, size) + single wh/pkg → one SQL + one fan-out batch.
                    keys = {(v[1], v[2], v[3], v[4]) for v in rows}
                    whs = sorted({int(v[5]) for v in rows})
                    pkgs = sorted({int(v[6]) for v in rows})
                    if len(keys) == 1 and len(whs) == 1 and len(pkgs) == 1:
                        it, sc, cl, sz = next(iter(keys))
                        flt: Dict[str, Any] = {
                            "item_type": it,
                            "school": sc,
                            "color": cl,
                            "size": sz,
                            "warehouse_no": whs[0],
                            "package_no": pkgs[0],
                        }
                        updated_total += self.db.update_prices(
                            flt,
                            new_price,
                            note="Price update (multi-selection)",
                            price_sync_mode=sync_mode,
                            price_sync_pos_devices=sync_devs,
                            emit_price_sync=emit_sync,
                        )
                    else:
                        # Mixed sizes/colours/locations: update each stock row so POS gets
                        # one PRICE_UPDATE per (item_type, school, color, size) from that row.
                        for vals in rows:
                            updated_total += self.db.update_prices(
                                {"id": int(vals[0])},
                                new_price,
                                note="Price update (multi-selection)",
                                price_sync_mode=sync_mode,
                                price_sync_pos_devices=sync_devs,
                                emit_price_sync=emit_sync,
                            )
                else:
                    if scope_var.get() == "row":
                        updated_total = self.db.update_prices(
                            {"id": int(first[0])},
                            new_price,
                            note="Price update (single row)",
                            price_sync_mode=sync_mode,
                            price_sync_pos_devices=sync_devs,
                            emit_price_sync=emit_sync,
                        )
                    else:
                        updated_total = self.db.update_prices(
                            {
                                "item_type": first[1],
                                "school": first[2],
                                "size": first[4],
                                "warehouse_no": first[5],
                                "package_no": first[6],
                            },
                            new_price,
                            note="Price update (same type/school/size)",
                            price_sync_mode=sync_mode,
                            price_sync_pos_devices=sync_devs,
                            emit_price_sync=emit_sync,
                        )

                dlg.destroy()
                if updated_total == 0:
                    show_toast(self, "لم يتم العثور على صفوف مطابقة", bg="#f59e0b")
                else:
                    show_toast(self, f"تم تحديث السعر في {updated_total} صف(وف)")
                self._refresh()

            except Exception as ex:
                messagebox.showerror("فشل التحديث", str(ex), parent=dlg)

        # ---- Buttons (ONCE, ALWAYS) ----
        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=on_ok).pack(side=tk.RIGHT, padx=6)

        dlg.update_idletasks()

    def _edit_specs_dialog(self):
        """
        Edit item_type / school / color / size for:
        - Selected rows (if any selection exists), OR
        - All rows in the package indicated by the current filters (warehouse_no + package_no)
        when there is no selection.
        """
        sel = self.table.selection()
        ids: List[int] = []
        if sel:
            # target selected rows
            for iid in sel:
                vals = self.table.item(iid, "values")
                ids.append(int(vals[0]))
            scope_text = f"عدد الصفوف المحددة: {len(ids)}"
            scope_mode = "ids"
        else:
            # fall back to package filters
            wh_txt = (self.f_wh.get() or "").strip()
            pkg_txt = (self.f_pkg.get() or "").strip()
            if not (wh_txt and pkg_txt and wh_txt.isdigit() and pkg_txt.isdigit()):
                messagebox.showwarning(
                    "حدد النطاق",
                    "اختر صفوفًا من الجدول أو أدخل (المخزن/العبوة) في أعلى الشاشة لتطبيق التعديل على العبوة كلها.",
                    parent=self,
                )
                return
            w = int(wh_txt); p = int(pkg_txt)
            scope_text = f"النطاق: المخزن {w} / العبوة {p}"
            scope_mode = "pkg"
            ids = [w, p]  # stash just to reuse locals (won't be used as ids)

        # --- Dialog UI ---
        dlg = tk.Toplevel(self)
        dlg.title("تعديل المواصفات")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=scope_text, font=("", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,8))
        ttk.Label(frm, text="اترك الحقل فارغًا إذا كنت لا تريد تغييره.").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,10))

        it_var = tk.StringVar()
        sc_var = tk.StringVar()
        cl_var = tk.StringVar()
        sz_var = tk.StringVar()

        ttk.Label(frm, text="النوع (جديد):").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=it_var, width=28).grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="المدرسة (جديد):").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=sc_var, width=28).grid(row=3, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="اللون (جديد):").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=cl_var, width=28).grid(row=4, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="المقاس (جديد):").grid(row=5, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=sz_var, width=28).grid(row=5, column=1, sticky="w", padx=6, pady=4)

        btns = ttk.Frame(frm); btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)

        def on_ok():
            # Gather non-empty changes
            changes = {}
            if it_var.get().strip(): changes["item_type"] = it_var.get().strip()
            if sc_var.get().strip(): changes["school"]    = sc_var.get().strip()
            if cl_var.get().strip(): changes["color"]     = cl_var.get().strip()
            if sz_var.get().strip(): changes["size"]      = sz_var.get().strip()

            if not changes:
                messagebox.showwarning("لا تغييرات", "لم تُدخل أي قيم جديدة.", parent=dlg)
                return

            try:
                if scope_mode == "ids":
                    updated = self.db.update_specs_by_ids(ids, **changes)
                else:
                    w, p = int(self.f_wh.get()), int(self.f_pkg.get())
                    updated = self.db.update_specs_in_package(w, p, **changes)

                dlg.destroy()
                if updated == 0:
                    show_toast(self, "لم يتم العثور على صفوف مطابقة", bg="#f59e0b")
                else:
                    show_toast(self, f"تم تعديل المواصفات في {updated} صف(وف)")
                self._refresh()
            except Exception as ex:
                messagebox.showerror("فشل التحديث", str(ex), parent=dlg)

        ttk.Button(btns, text="حفظ", command=on_ok).pack(side=tk.RIGHT, padx=6)

class MultiSelectDialog(tk.Toplevel):
    def __init__(self, master, title: str, values: List[str], preselected: Optional[List[str]] = None):
        super().__init__(master)
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)
        self.title(title)
        self._alive = True
        self._after_id = None
        self._result = None

        self.transient(master)
        self.grab_set()
        try:
            self.focus_force()
        except Exception:
            pass

        self.geometry("380x460")
        self.resizable(False, False)

        self._values = list(dict.fromkeys(values or []))  # unique, keep order
        self._ok = False

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        # Filter box
        frow = ttk.Frame(frm)
        frow.pack(fill=tk.X, pady=(0,6))
        ttk.Label(frow, text="بحث:").pack(side=tk.LEFT)
        self.q = tk.StringVar()
        ent = ttk.Entry(frow, textvariable=self.q)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6,0))
        ent.bind("<KeyRelease>", lambda e: self._refill())

        # Listbox (multi select)
        self.listbox = tk.Listbox(frm, selectmode="extended", activestyle="none")
        self.listbox.pack(fill=tk.BOTH, expand=True)
        ysb = ttk.Scrollbar(frm, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=ysb.set)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons
        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(8,0))
        ttk.Button(btns, text="مسح التحديد", command=self._clear_sel).pack(side=tk.LEFT)
        ttk.Button(btns, text="إلغاء", command=self._on_cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="موافق", command=self._on_ok).pack(side=tk.RIGHT, padx=6)

        self._refill()

        # Preselect
        if preselected:
            lower_map = {v.lower(): i for i, v in enumerate(self._shown)}
            sel_ids = [lower_map.get(v.lower()) for v in preselected if lower_map.get(v.lower()) is not None]
            for i in sel_ids:
                self.listbox.selection_set(i)

        self._after_id = self.after(10, ent.focus_set)


    def _refill(self):
        if not self._alive:
            return
        if not self.winfo_exists():
            return
        if not self.listbox.winfo_exists():
            return

        q = (self.q.get() or "").strip().lower()
        if q:
            self._shown = [v for v in self._values if q in str(v).lower()]
        else:
            self._shown = self._values[:]

        self.listbox.delete(0, tk.END)
        for v in self._shown:
            self.listbox.insert(tk.END, v)



    def _clear_sel(self):
        self.listbox.selection_clear(0, tk.END)

    def _cleanup(self):
        self._alive = False
        try:
            if self._after_id:
                self.after_cancel(self._after_id)
        except Exception:
            pass

    def _on_ok(self):
        try:
            sel_idx = list(self.listbox.curselection())
            self._result = [self._shown[i] for i in sel_idx]
        except Exception:
            self._result = []

        self._ok = True
        self._cleanup()
        self.destroy()



    def _on_cancel(self):
        self._result = None
        self._cleanup()
        self.destroy()



    def run(self) -> Optional[List[str]]:
        self.wait_window()
        return self._result


# ------------------- Bills History Window -------------------

class BillsHistoryWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("سجل الفواتير")
        self.geometry("1000x560")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  سجل الفواتير", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="سجل الفواتير", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)
        _br = ttk.Button(top, text="تحديث", command=self._refresh); _br.pack(side=tk.RIGHT)
        ToolTip(_br, "تحديث قائمة الفواتير")
        _bo = ttk.Button(top, text="فتح", command=self._print_selected); _bo.pack(side=tk.RIGHT, padx=8)
        ToolTip(_bo, "فتح الفاتورة المحددة للطباعة")
        _be = ttk.Button(top, text="تصدير المحدد إلى إكسل", command=self._export_selected); _be.pack(side=tk.RIGHT)
        ToolTip(_be, "تصدير الفاتورة المحددة إلى ملف إكسل")

        # Bill lifecycle buttons
        _bc = ttk.Button(top, text="تأكيد المسودة", command=self._confirm_draft); _bc.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bc, "تأكيد المسودة وخصم الكميات من المخزون")
        _bv = ttk.Button(top, text="إلغاء الفاتورة", command=self._void_bill); _bv.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bv, "إلغاء فاتورة مؤكدة وإرجاع الكميات")
        _bdel = ttk.Button(top, text="حذف المسودة", command=self._delete_draft); _bdel.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bdel, "حذف المسودة نهائياً")
        _bret = ttk.Button(top, text="مرتجع", command=self._process_return); _bret.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bret, "معالجة مرتجع لفاتورة مؤكدة")

        bills_wrap = ttk.Frame(self)
        bills_wrap.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 6))
        self.bills_table = ttk.Treeview(
            bills_wrap, columns=("id", "created_at", "kind", "customer", "total", "status"), show="headings", height=10
        )
        for col, txt, w in [("id","المعرّف",80), ("created_at","التاريخ",180), ("kind","النوع",120), ("customer","العميل",200), ("total","الإجمالي",120), ("status","الحالة",90)]:
            self.bills_table.heading(col, text=txt)
            self.bills_table.column(col, width=w, anchor="center")
        bills_ysb = ttk.Scrollbar(bills_wrap, orient="vertical", command=self.bills_table.yview)
        bills_xsb = ttk.Scrollbar(bills_wrap, orient="horizontal", command=self.bills_table.xview)
        self.bills_table.configure(yscrollcommand=bills_ysb.set, xscrollcommand=bills_xsb.set)
        self.bills_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bills_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        bills_xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.bills_table.bind("<<TreeviewSelect>>", lambda e: self._load_items())
        add_context_menu(self.bills_table)
        _bind_mousewheel(self.bills_table)

        ttk.Label(self, text="بنود الفاتورة").pack(anchor="w", padx=8)

        items_wrap = ttk.Frame(self)
        items_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        # NEW: show badge column
        self.items_table = ttk.Treeview(
            items_wrap,
            columns=("type", "school", "color", "size", "origin", "badge", "wh", "pkg", "price", "qty", "total"),
            show="headings",
            height=12,
        )
        for col, txt, w in [
            ("type","النوع",140), ("school","المدرسة",160), ("color","اللون",80), ("size","المقاس",70),
            ("origin","المصدر",90), ("badge","بادج",60), ("wh","المخزن",50), ("pkg","العبوة",60), ("price","السعر",80),
            ("qty","الكمية",60), ("total","إجمالي ",100),
        ]:
            self.items_table.heading(col, text=txt)
            self.items_table.column(col, width=w, anchor="center")
        items_ysb = ttk.Scrollbar(items_wrap, orient="vertical", command=self.items_table.yview)
        items_xsb = ttk.Scrollbar(items_wrap, orient="horizontal", command=self.items_table.xview)
        self.items_table.configure(yscrollcommand=items_ysb.set, xscrollcommand=items_xsb.set)
        self.items_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        items_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        items_xsb.pack(side=tk.BOTTOM, fill=tk.X)
        add_context_menu(self.items_table)
        _bind_mousewheel(self.items_table)

        self._refresh()

    def _refresh(self):
        self.bills_table.delete(*self.bills_table.get_children())
        status_map = {"DRAFT": "مسودة", "CONFIRMED": "مؤكدة", "VOID": "ملغاة"}
        for b in self.db.list_bills():
            status_txt = status_map.get(b.get("status", "CONFIRMED"), b.get("status", ""))
            kind_txt = "شحن فرع" if b.get("bill_kind") == "BRANCH_SHIPMENT" else "فاتورة"
            self.bills_table.insert(
                "", tk.END, iid=str(b["id"]),
                values=(b["id"], b["created_at"], kind_txt, b.get("customer") or "",
                        f"{float(b['total']):.2f}", status_txt)
            )
        apply_zebra_tags(self.bills_table)
        # Color-code by status
        self.bills_table.tag_configure("draft", background="#fef3c7")
        self.bills_table.tag_configure("void", background="#fee2e2")
        for child in self.bills_table.get_children():
            vals = self.bills_table.item(child, "values")
            if len(vals) >= 5:
                if vals[4] == "مسودة":
                    self.bills_table.item(child, tags=("draft",))
                elif vals[4] == "ملغاة":
                    self.bills_table.item(child, tags=("void",))
        self.items_table.delete(*self.items_table.get_children())

    def _confirm_draft(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.", parent=self)
            return
        if not messagebox.askyesno("تأكيد", "هل تريد تأكيد هذه المسودة وخصم الكميات من المخزون؟", parent=self):
            return
        try:
            self.db.confirm_draft_bill(bill_id)
            show_toast(self, f"تم تأكيد الفاتورة #{bill_id}")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل التأكيد", str(ex), parent=self)

    def _void_bill(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.", parent=self)
            return
        if not messagebox.askyesno("تأكيد الإلغاء", "هل تريد إلغاء هذه الفاتورة وإرجاع الكميات؟", parent=self):
            return
        try:
            self.db.void_bill(bill_id)
            show_toast(self, f"تم إلغاء الفاتورة #{bill_id}")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل الإلغاء", str(ex), parent=self)

    def _delete_draft(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.", parent=self)
            return
        if not messagebox.askyesno("تأكيد الحذف", "هل تريد حذف هذه المسودة نهائياً؟", parent=self):
            return
        try:
            self.db.delete_draft_bill(bill_id)
            show_toast(self, f"تم حذف المسودة #{bill_id}")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل الحذف", str(ex), parent=self)

    def _process_return(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.", parent=self)
            return
        bill = None
        for b in self.db.list_bills():
            if int(b["id"]) == bill_id:
                bill = b
                break
        if not bill or bill.get("status") != "CONFIRMED":
            messagebox.showwarning("غير متاح", "المرتجعات متاحة فقط للفواتير المؤكدة.", parent=self)
            return
        ReturnDialog(self, self.db, bill_id, on_done=self._refresh)

    def _get_selected_bill_id(self) -> Optional[int]:
        sel = self.bills_table.selection()
        if not sel:
            return None
        return int(sel[0])

    def _load_items(self):
        bill_id = self._get_selected_bill_id()
        self.items_table.delete(*self.items_table.get_children())
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولًا.")
            return
        items = self.db.list_bill_items(bill_id)
        for ln in items:
            origin_txt = "من المخزون" if ln.get("origin") == "STOCK" else ("من المصنع" if ln.get("origin") == "FACTORY" else "")
            wh_txt  = "" if (ln.get("warehouse_no") in (None, "", 0, "0")) else ln.get("warehouse_no")
            pkg_txt = "" if (ln.get("package_no")  in (None, "", 0, "0")) else ln.get("package_no")
            self.items_table.insert(
                "", tk.END,
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"], origin_txt,
                        "✓" if int(ln.get("has_badge") or 0) else "",
                        wh_txt, pkg_txt, f"{float(ln['unit_price']):.2f}",
                        ln["qty"], f"{float(ln['line_total']):.2f}")
            )
        apply_zebra_tags(self.items_table)

    def _export_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولًا.")
            return
        items = self.db.list_bill_items(bill_id)
        if not items:
            messagebox.showwarning("فارغ", "لا تحتوي هذه الفاتورة على بنود.")
            return
        path = filedialog.asksaveasfilename(
            title="تصدير الفاتورة إلى إكسل",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"bill_{bill_id}.xlsx",
        )
        if not path:
            return
        headers = ["type", "school", "color", "size", "origin", "has_badge", "warehouse_no", "package_no", "unit_price", "qty", "line_total"]
        def _origin_txt(o: Optional[str]) -> str:
            return "من المخزون" if o == "STOCK" else ("من المصنع" if o == "FACTORY" else "")
        rows = [
            [
                ln["item_type"], ln["school"], ln["color"], ln["size"],
                _origin_txt(ln.get("origin")), ("✓" if int(ln.get("has_badge") or 0) else ""),
                ln["warehouse_no"], ln["package_no"],
                float(ln["unit_price"]), int(ln["qty"]), float(ln["line_total"]),
            ]
            for ln in items
        ]
        try:
            export_to_excel(path, headers, rows)
            show_toast(self, "تم تصدير الفاتورة إلى إكسل بنجاح")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex))

    def _print_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولًا.")
            return
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == bill_id)
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("فشل الطباعة", "لم يتم العثور على الفاتورة.")
            return
        tmp_dir = tempfile.gettempdir()
        path = os.path.join(tmp_dir, f"bill_{bill_id}.html")
        save_bill_as_html(path, bill, items)
        _print_html_auto(path, copies=1, parent=self)


# ------------------- Return Dialog -------------------

class ReturnDialog(tk.Toplevel):
    """Select items and quantities to return from a confirmed bill."""
    def __init__(self, master, db: SqliteDatabase, bill_id: int, on_done=None):
        super().__init__(master)
        self.db = db
        self.bill_id = bill_id
        self.on_done = on_done
        self.title(f"مرتجع - فاتورة #{bill_id}")
        self.geometry("800x500")
        self.transient(master)
        self.grab_set()
        self._items = self.db.list_bill_items(bill_id)
        self._build()

    def _build(self):
        ttk.Label(self, text=f"اختر الأصناف والكميات المرتجعة من الفاتورة #{self.bill_id}",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=8)

        cols = ("idx", "type", "school", "color", "size", "qty_orig", "return_qty")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        for col, txt, w in [
            ("idx", "#", 40), ("type", "النوع", 140), ("school", "المدرسة", 140),
            ("color", "اللون", 80), ("size", "المقاس", 70),
            ("qty_orig", "الكمية الأصلية", 90), ("return_qty", "كمية الإرجاع", 90)
        ]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor="center")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        add_context_menu(self.tree)
        _bind_mousewheel(self.tree)

        for i, item in enumerate(self._items):
            self.tree.insert("", tk.END, iid=str(i), values=(
                i + 1, item["item_type"], item["school"], item["color"],
                item["size"], item["qty"], 0
            ))
        apply_zebra_tags(self.tree)

        edit_frame = ttk.Frame(self)
        edit_frame.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(edit_frame, text="كمية الإرجاع للبند المحدد:").pack(side=tk.LEFT)
        self._ret_var = tk.StringVar(value="0")
        self._ret_entry = ttk.Entry(edit_frame, textvariable=self._ret_var, width=8)
        self._ret_entry.pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="تعيين", command=self._set_return_qty).pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="إرجاع الكل", command=self._return_all).pack(side=tk.LEFT, padx=4)

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.LEFT)
        ttk.Button(btns, text="تأكيد المرتجع", command=self._confirm).pack(side=tk.RIGHT)

    def _set_return_qty(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        try:
            ret_qty = int(self._ret_var.get())
        except ValueError:
            return
        orig_qty = int(self._items[idx]["qty"])
        ret_qty = max(0, min(ret_qty, orig_qty))
        vals = list(self.tree.item(sel[0], "values"))
        vals[6] = ret_qty
        self.tree.item(sel[0], values=vals)

    def _return_all(self):
        for child in self.tree.get_children():
            vals = list(self.tree.item(child, "values"))
            vals[6] = vals[5]  # set return_qty = qty_orig
            self.tree.item(child, values=vals)

    def _confirm(self):
        return_lines = []
        for child in self.tree.get_children():
            vals = self.tree.item(child, "values")
            ret_qty = int(vals[6])
            if ret_qty <= 0:
                continue
            idx = int(vals[0]) - 1
            item = self._items[idx]
            return_lines.append({
                "item_type": item["item_type"],
                "school": item["school"],
                "color": item["color"],
                "size": item["size"],
                "warehouse_no": item.get("warehouse_no", 0),
                "package_no": item.get("package_no", 0),
                "unit_price": item["unit_price"],
                "qty": ret_qty,
                "has_badge": item.get("has_badge", 0),
            })
        if not return_lines:
            messagebox.showwarning("لا شيء", "لم يتم تحديد أي كميات للإرجاع.", parent=self)
            return
        try:
            ret_id = self.db.process_return(self.bill_id, return_lines)
            show_toast(self, f"تم المرتجع #{ret_id} بنجاح")
            self.destroy()
            if self.on_done:
                self.on_done()
        except Exception as ex:
            messagebox.showerror("فشل المرتجع", str(ex), parent=self)

# ------------------- Movements Window -------------------

def _summarize_movement_rows(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Aggregate counts/amounts for the movements log footer (filtered rows only)."""
    n = 0
    qty_in = 0
    qty_out = 0
    qty_adj_out = 0
    qty_return_in = 0
    qty_reserve = 0
    val_out = 0.0
    deliver_cash = 0.0
    for m in rows:
        n += 1
        d = str(m.get("direction") or "").strip().upper()
        q = int(m.get("qty") or 0)
        try:
            up = float(m.get("unit_price") or 0)
        except (TypeError, ValueError):
            up = 0.0
        if d == "IN":
            qty_in += q
        elif d in ("OUT", "OUT_FACTORY"):
            qty_out += q
            val_out += float(q) * up
        elif d == "ADJUST_OUT":
            qty_adj_out += q
            val_out += float(q) * up
        elif d == "RETURN_IN":
            qty_return_in += q
        elif d == "RESERVE":
            qty_reserve += q
        elif d == "DELIVER_PAY":
            deliver_cash += up
    income_moves = val_out + deliver_cash
    return {
        "n": n,
        "qty_in": qty_in,
        "qty_out": qty_out,
        "qty_adj_out": qty_adj_out,
        "qty_return_in": qty_return_in,
        "qty_reserve": qty_reserve,
        "val_out": val_out,
        "deliver_cash": deliver_cash,
        "income_moves": income_moves,
    }


class MovementsWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("سجل الحركات")
        self.geometry("1180x620")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  سجل الحركات", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        top = ttk.LabelFrame(self, text="تصنيف")
        top.pack(fill=tk.X, padx=8, pady=8)
        self.ftype  = LabeledCombobox(top, "النوع",   self.db, "item_type");  self.ftype.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        self.fsch   = LabeledCombobox(top, "المدرسة", self.db, "school");     self.fsch.grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        self.fclr   = LabeledCombobox(top, "اللون",   self.db, "color");      self.fclr.grid(row=0, column=2, padx=6, pady=4, sticky="ew")
        self.fsiz   = LabeledCombobox(top, "المقاس",  self.db, "size");       self.fsiz.grid(row=0, column=3, padx=6, pady=4, sticky="ew")
        def _constraints(exclude: Optional[str] = None) -> Dict[str, Any]:
            data = {
                "item_type": self.ftype.get() or None,
                "school": self.fsch.get() or None,
                "color": self.fclr.get() or None,
                "size": self.fsiz.get() or None,
            }
            if exclude:
                data.pop(exclude, None)
            return {k: v for k, v in data.items() if v not in (None, "")}
        self.ftype.set_supplier(lambda: self.db.get_distinct_filtered("item_type", _constraints("item_type")))
        self.fsch.set_supplier(lambda: self.db.get_distinct_filtered("school", _constraints("school")))
        self.fclr.set_supplier(lambda: self.db.get_distinct_filtered("color", _constraints("color")))
        self.fsiz.set_supplier(lambda: self.db.get_distinct_filtered("size", _constraints("size")))
        def _refresh_filter_values(*_):
            self.ftype.refresh_values()
            self.fsch.refresh_values()
            self.fclr.refresh_values()
            self.fsiz.refresh_values()
        for widget in (self.ftype, self.fsch, self.fclr, self.fsiz):
            for ev in ("<<ComboboxSelected>>", "<KeyRelease>"):
                widget.cb.bind(ev, lambda _e: (_refresh_filter_values(), self._refresh()), add="+")

        self.df = DateField(top, "من (YYYY-MM-DD)"); self.df.grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.dt = DateField(top, "إلى");             self.dt.grid(row=1, column=1, padx=6, pady=4, sticky="w")


        ttk.Label(top, text="بحث").grid(row=1, column=2, sticky="e", padx=4, pady=4)
        self.txt = ttk.Entry(top); self.txt.grid(row=1, column=3, sticky="ew", padx=6, pady=4)
        top.columnconfigure(3, weight=1)

        btns = ttk.Frame(top)
        btns.grid(row=0, column=4, rowspan=2, sticky="e", padx=6, pady=4)
        _mr = ttk.Button(btns, text="تحديث", command=self._refresh); _mr.pack(side=tk.LEFT)
        ToolTip(_mr, "تحديث قائمة الحركات")
        _mc = ttk.Button(btns, text="مسح", command=self._clear); _mc.pack(side=tk.LEFT, padx=6)
        ToolTip(_mc, "مسح جميع الفلاتر")
        _mp = ttk.Button(btns, text="طباعة", command=self._print_report); _mp.pack(side=tk.LEFT, padx=6)
        ToolTip(_mp, "طباعة التقرير التجميعي المعروض")
        _me = ttk.Button(btns, text="تصدير إلى إكسل", command=self._export); _me.pack(side=tk.LEFT)
        ToolTip(_me, "تصدير الحركات إلى ملف إكسل")

        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        # NEW: add badge column
        self.table = ttk.Treeview(
            table_wrap,
            columns=("type","school","color","size","incoming","sold","remaining","reserved"),
            show="headings",
            height=14,
        )
        for col, txt, w in [
            ("type","النوع",160), ("school","المدرسة",170), ("color","اللون",120), ("size","المقاس",90),
            ("incoming","وارد",95), ("sold","مباع",95), ("remaining","متبقي",95), ("reserved","حجوزات",90),
        ]:
            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        xsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        add_context_menu(self.table)
        _bind_mousewheel(self.table)

        sum_fr = ttk.LabelFrame(self, text="ملخص النتائج المعروضة (تجميعي حسب المنتج)")
        sum_fr.pack(fill=tk.X, padx=8, pady=(0, 8))
        self._summary_var = tk.StringVar(value="")
        ttk.Label(
            sum_fr, textvariable=self._summary_var, justify="right", wraplength=1120,
            font=("Segoe UI", 9),
        ).pack(anchor="e", padx=10, pady=8)

        self._refresh()

    def _filters(self) -> Dict[str, Any]:
        return {
            "item_type": self.ftype.get(),
            "school":    self.fsch.get(),
            "color":     self.fclr.get(),
            "size":      self.fsiz.get(),
            "date_from": self.df.get() or None,           # DateField
            "date_to":   self.dt.get() or None,           # DateField
            "text":      (self.txt.get().strip() or None),
        }


    def _refresh(self):
        try:
            rows = self.db.list_movement_item_totals(self._filters())
        except Exception as ex:
            messagebox.showerror("فشل البحث", str(ex), parent=self)
            self._summary_var.set("")
            return

        self.table.delete(*self.table.get_children())
        for r in rows:
            self.table.insert(
                "", tk.END,
                values=(
                    r.get("item_type",""), r.get("school",""), r.get("color",""), r.get("size",""),
                    int(r.get("incoming_qty") or 0),
                    int(r.get("sold_total_qty") or 0),
                    int(r.get("remaining_total_qty") or 0),
                    int(r.get("reserved_qty") or 0),
                )
            )
        apply_zebra_tags(self.table)
        s = {
            "n": len(rows),
            "incoming_qty": sum(int(r.get("incoming_qty") or 0) for r in rows),
            "sold_total_qty": sum(int(r.get("sold_total_qty") or 0) for r in rows),
            "reserved_qty": sum(int(r.get("reserved_qty") or 0) for r in rows),
            "remaining_total_qty": sum(int(r.get("remaining_total_qty") or 0) for r in rows),
        }
        self._summary_var.set(
            f"عدد المنتجات المعروضة: {s['n']}  |  "
            f"إجمالي الوارد: {s['incoming_qty']}  |  "
            f"إجمالي المباع: {s['sold_total_qty']}  |  "
            f"إجمالي المتبقي (المخزن + الفروع): {s['remaining_total_qty']}  |  "
            f"إجمالي الحجوزات: {s['reserved_qty']}"
        )

    def _clear(self):
        for w in (self.ftype, self.fsch, self.fclr, self.fsiz):
            w.set("")
        self.df.set("")               # DateField
        self.dt.set("")               # DateField
        self.txt.delete(0, tk.END)
        self._refresh()


    def _export(self):
        rows = self.db.list_movement_item_totals(self._filters())
        if not rows:
            messagebox.showwarning("فارغ", "لا توجد صفوف للتصدير.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="تصدير الحركات إلى إكسل",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"movement_totals_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        headers = [
            "item_type", "school", "color", "size",
            "incoming_qty", "sold_total_qty", "remaining_total_qty", "reserved_qty",
        ]
        table = [[
            m.get("item_type",""), m.get("school",""), m.get("color",""), m.get("size",""),
            int(m.get("incoming_qty") or 0),
            int(m.get("sold_total_qty") or 0),
            int(m.get("remaining_total_qty") or 0),
            int(m.get("reserved_qty") or 0),
        ] for m in rows]
        try:
            export_to_excel(path, headers, table)
            show_toast(self, "تم حفظ الحركات إلى إكسل بنجاح")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)

    def _print_report(self):
        rows = self.db.list_movement_item_totals(self._filters())
        if not rows:
            messagebox.showwarning("فارغ", "لا توجد صفوف للطباعة.", parent=self)
            return

        def _h(v: Any) -> str:
            s = str(v if v is not None else "")
            return (
                s.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;")
                 .replace('"', "&quot;")
            )

        f = self._filters()
        filter_bits = []
        for label, key in (
            ("النوع", "item_type"),
            ("المدرسة", "school"),
            ("اللون", "color"),
            ("المقاس", "size"),
            ("من", "date_from"),
            ("إلى", "date_to"),
            ("بحث", "text"),
        ):
            val = str(f.get(key) or "").strip()
            if val:
                filter_bits.append(f"{label}: {_h(val)}")
        filters_html = " | ".join(filter_bits) if filter_bits else "بدون فلاتر"

        totals = {
            "incoming": sum(int(r.get("incoming_qty") or 0) for r in rows),
            "sold": sum(int(r.get("sold_total_qty") or 0) for r in rows),
            "remaining": sum(int(r.get("remaining_total_qty") or 0) for r in rows),
            "reserved": sum(int(r.get("reserved_qty") or 0) for r in rows),
        }

        tr_html = []
        for r in rows:
            tr_html.append(
                "<tr>"
                f"<td>{_h(r.get('item_type',''))}</td>"
                f"<td>{_h(r.get('school',''))}</td>"
                f"<td>{_h(r.get('color',''))}</td>"
                f"<td>{_h(r.get('size',''))}</td>"
                f"<td class='num'>{int(r.get('incoming_qty') or 0)}</td>"
                f"<td class='num'>{int(r.get('sold_total_qty') or 0)}</td>"
                f"<td class='num'>{int(r.get('remaining_total_qty') or 0)}</td>"
                f"<td class='num'>{int(r.get('reserved_qty') or 0)}</td>"
                "</tr>"
            )

        html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<title>تقرير حركة الأصناف - تجميعي</title>
<style>
@page {{ size: A4 landscape; margin: 10mm; }}
* {{ box-sizing: border-box; }}
body {{
  margin: 0;
  font-family: "Segoe UI", Tahoma, Arial, "Noto Sans Arabic", sans-serif;
  color: #0f172a;
  direction: rtl;
}}
.wrap {{ padding: 8px; }}
.title {{ font-size: 20px; font-weight: 700; margin-bottom: 4px; }}
.meta {{ font-size: 12px; color: #334155; margin-bottom: 3px; }}
.summary {{
  background: #f8fafc;
  border: 1px solid #cbd5e1;
  padding: 8px;
  margin: 8px 0 10px;
  font-weight: 600;
  font-size: 13px;
}}
table {{
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  font-size: 12px;
}}
th, td {{
  border: 1px solid #64748b;
  padding: 6px 4px;
  text-align: center;
  vertical-align: middle;
  word-wrap: break-word;
}}
th {{
  background: #e2e8f0;
  font-weight: 700;
}}
tbody tr:nth-child(even) {{ background: #f8fafc; }}
.num {{ font-variant-numeric: tabular-nums; }}
</style>
</head>
<body>
  <div class="wrap">
    <div class="title">تقرير حركة الأصناف (تجميعي)</div>
    <div class="meta">وقت التقرير: {_h(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>
    <div class="meta">الفلاتر: {filters_html}</div>
    <div class="summary">
      عدد المنتجات: {len(rows)} |
      إجمالي الوارد: {totals['incoming']} |
      إجمالي المباع: {totals['sold']} |
      إجمالي المتبقي: {totals['remaining']} |
      إجمالي الحجوزات: {totals['reserved']}
    </div>
    <table>
      <thead>
        <tr>
          <th>النوع</th>
          <th>المدرسة</th>
          <th>اللون</th>
          <th>المقاس</th>
          <th>وارد</th>
          <th>مباع</th>
          <th>متبقي</th>
          <th>حجوزات</th>
        </tr>
      </thead>
      <tbody>
        {''.join(tr_html)}
      </tbody>
    </table>
  </div>
</body>
</html>
"""

        path = os.path.join(
            tempfile.gettempdir(),
            f"movement_totals_print_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
        )
        with open(path, "w", encoding="utf-8") as fobj:
            fobj.write(html)
        _print_html_auto(path, copies=1, parent=self)

class AdminWindow(tk.Toplevel):
    """
    Password-protected admin tools.
    First tool: Reopen a package (with optional 'only if empty').
    """
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("إعدادات المدير")
        self.geometry("520x380")
        self.resizable(False, False)
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  إعدادات المدير", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        # Password gate at top
        gate = ttk.LabelFrame(self, text="التحقق من كلمة المرور")
        gate.pack(fill=tk.X, padx=10, pady=(10, 8))
        ttk.Label(gate, text="كلمة المرور:").grid(row=0, column=0, sticky="e", padx=6, pady=8)
        self.pw_var = tk.StringVar()
        pw_entry = ttk.Entry(gate, textvariable=self.pw_var, show="*")
        pw_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=8)
        gate.columnconfigure(1, weight=1)

        # Section: Reopen package
        rp = ttk.LabelFrame(self, text="إعادة فتح حاوية")
        rp.pack(fill=tk.X, padx=10, pady=(0, 8))

        self.wh_var  = tk.StringVar()
        self.pkg_var = tk.StringVar()
        self.empty_only = tk.BooleanVar(value=False)

        ttk.Label(rp, text="رقم المخزن:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Entry(rp, textvariable=self.wh_var, width=10).grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(rp, text="رقم العبوة:").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        ttk.Entry(rp, textvariable=self.pkg_var, width=12).grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Checkbutton(rp, text="إعادة الفتح إذا كانت فارغة فقط", variable=self.empty_only).grid(
            row=1, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 6)
        )

        act = ttk.Frame(rp); act.grid(row=2, column=0, columnspan=4, sticky="e", padx=6, pady=(6, 6))
        _arp = ttk.Button(act, text="إعادة فتح", command=self._do_reopen); _arp.pack(side=tk.RIGHT)
        ToolTip(_arp, "إعادة فتح العبوة المغلقة للإضافة عليها")

        # Future admin tools can go in more frames below...
            # -------- Purge definition (HARD DELETE) --------
        pd = ttk.LabelFrame(self, text="حذف تعريف نهائي (إزالة من النظام)")
        pd.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.purge_field = tk.StringVar(value="color")
        self.purge_value = tk.StringVar()

        ttk.Label(pd, text="الحقل:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Combobox(
            pd,
            textvariable=self.purge_field,
            values=["item_type", "school", "color", "size"],
            state="readonly",
            width=14
        ).grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(pd, text="القيمة:").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        ttk.Entry(pd, textvariable=self.purge_value, width=26)\
            .grid(row=0, column=3, sticky="w", padx=6, pady=6)

        _apd = ttk.Button(pd, text="حذف نهائي", command=self._purge_definition); _apd.grid(row=0, column=4, padx=8)
        ToolTip(_apd, "حذف القيمة نهائياً من المخزون (لا يمكن التراجع!)")


    def _do_reopen(self):
        # Verify password
        if not self.db.verify_admin_password(self.pw_var.get()):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return

        # Parse inputs
        try:
            w = int((self.wh_var.get() or "").strip())
            p = int((self.pkg_var.get() or "").strip())
            if w < 1 or p < 1:
                raise ValueError
        except Exception:
            messagebox.showerror("بيانات غير صالحة", "أدخل أرقامًا صحيحة للمخزن والعبوة (>= 1).", parent=self)
            return

        # Perform action
        try:
            self.db.reopen_package(w, p, require_empty=bool(self.empty_only.get()), note="Admin manual reopen")
            show_toast(self, f"تمت إعادة فتح العبوة {p} في المخزن {w}")
        except Exception as ex:
            messagebox.showerror("فشل", str(ex), parent=self)
    def _purge_definition(self):
        # Verify admin password
        if not self.db.verify_admin_password(self.pw_var.get()):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return

        field = self.purge_field.get()
        value = (self.purge_value.get() or "").strip()

        if not value:
            messagebox.showerror("بيانات ناقصة", "أدخل قيمة للحذف.", parent=self)
            return

        if not messagebox.askyesno(
            "تأكيد الحذف النهائي",
            f"سيتم حذف '{value}' من '{field}' نهائيًا.\n\n"
            "⚠ هذا سيحذف كل الصفوف المرتبطة به من المخزون.\n\n"
            "هل أنت متأكد؟",
            parent=self
        ):
            return

        try:
            removed = self.db.purge_definition(field, value)
            if removed == 0:
                show_toast(self, "القيمة غير موجودة في المخزون", bg="#f59e0b")
            else:
                show_toast(self, f"تم حذف {removed} صف(وف) من المخزون نهائياً")
            self.purge_value.set("")
        except Exception as ex:
            messagebox.showerror("فشل الحذف", str(ex), parent=self)



# ------------------- Transfer Window -------------------

class TransferWindow(tk.Toplevel):
    """Select source stock items, choose destination warehouse/package, confirm transfer."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("تحويل بين المخازن")
        self.geometry("950x560")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  تحويل بين المخازن", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        # Source filters
        src_frame = ttk.LabelFrame(self, text="اختر المصدر")
        src_frame.pack(fill=tk.X, padx=8, pady=8)

        row1 = ttk.Frame(src_frame)
        row1.pack(fill=tk.X, padx=4, pady=4)
        self.f_type = LabeledCombobox(row1, "النوع", self.db, "item_type")
        self.f_type.pack(side=tk.LEFT, padx=4)
        self.f_school = LabeledCombobox(row1, "المدرسة", self.db, "school")
        self.f_school.pack(side=tk.LEFT, padx=4)
        self.f_color = LabeledCombobox(row1, "اللون", self.db, "color")
        self.f_color.pack(side=tk.LEFT, padx=4)
        self.f_size = LabeledCombobox(row1, "المقاس", self.db, "size")
        self.f_size.pack(side=tk.LEFT, padx=4)

        row2 = ttk.Frame(src_frame)
        row2.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(row2, text="المخزن:").pack(side=tk.LEFT)
        self.f_wh = ttk.Combobox(row2, values=["", "1", "2", "3", "4"], width=5, state="readonly")
        self.f_wh.pack(side=tk.LEFT, padx=4)
        ttk.Label(row2, text="العبوة:").pack(side=tk.LEFT)
        self.f_pkg = ttk.Entry(row2, width=8)
        self.f_pkg.pack(side=tk.LEFT, padx=4)
        _bsrc = ttk.Button(row2, text="بحث", command=self._load_source); _bsrc.pack(side=tk.LEFT, padx=8)
        ToolTip(_bsrc, "عرض الأصناف المتاحة للتحويل")

        # Source table
        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.src_table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "wh", "pkg", "badge", "price", "count"),
            show="headings", height=10)
        for col, txt, w in [
            ("id", "ID", 50), ("type", "النوع", 120), ("school", "المدرسة", 120),
            ("color", "اللون", 70), ("size", "المقاس", 60), ("wh", "المخزن", 50),
            ("pkg", "العبوة", 50), ("badge", "بادج", 40), ("price", "السعر", 70),
            ("count", "الكمية", 60),
        ]:
            self.src_table.heading(col, text=txt)
            self.src_table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.src_table.yview)
        self.src_table.configure(yscrollcommand=ysb.set)
        self.src_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        add_context_menu(self.src_table)
        _bind_mousewheel(self.src_table)

        # Destination
        dest_frame = ttk.LabelFrame(self, text="الوجهة")
        dest_frame.pack(fill=tk.X, padx=8, pady=4)
        drow = ttk.Frame(dest_frame)
        drow.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(drow, text="مخزن الوجهة:").pack(side=tk.LEFT)
        self.dest_wh = ttk.Combobox(drow, values=WAREHOUSE_NUMBER_DISPLAY_VALUES, width=14, state="readonly")
        self.dest_wh.pack(side=tk.LEFT, padx=4)
        ttk.Label(drow, text="عبوة الوجهة:").pack(side=tk.LEFT)
        self.dest_pkg = ttk.Entry(drow, width=8)
        self.dest_pkg.pack(side=tk.LEFT, padx=4)
        ttk.Label(drow, text="الكمية:").pack(side=tk.LEFT)
        self.transfer_qty = ttk.Entry(drow, width=8)
        self.transfer_qty.pack(side=tk.LEFT, padx=4)
        _bt = ttk.Button(drow, text="تحويل", command=self._do_transfer); _bt.pack(side=tk.LEFT, padx=8)
        ToolTip(_bt, "تحويل الكمية المحددة إلى الوجهة")

    def _load_source(self):
        filters = {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "color": self.f_color.get() or None,
            "size": self.f_size.get() or None,
            "warehouse_no": self.f_wh.get() or None,
            "package_no": self.f_pkg.get().strip() or None,
        }
        try:
            rows = self.db.current_inventory(filters)
        except Exception as ex:
            messagebox.showerror("فشل البحث", str(ex), parent=self)
            return
        self.src_table.delete(*self.src_table.get_children())
        for r in rows:
            self.src_table.insert("", tk.END, iid=str(r["id"]), values=(
                r["id"], r["item_type"], r["school"], r["color"], r["size"],
                r["warehouse_no"], r["package_no"],
                ("✓" if int(r.get("has_badge") or 0) else ""),
                f"{float(r['unit_price']):.2f}", r["count"]
            ))
        apply_zebra_tags(self.src_table)

    def _do_transfer(self):
        sel = self.src_table.selection()
        if not sel:
            messagebox.showwarning("لم يتم التحديد", "اختر صف مخزون أولاً.", parent=self)
            return
        stock_id = int(sel[0])
        try:
            dest_wh = int(warehouse_numeric_value(self.dest_wh.get()))
            dest_pkg = int(self.dest_pkg.get().strip())
            qty = int(self.transfer_qty.get().strip())
            if dest_wh < 1 or dest_pkg < 1 or qty < 1:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("بيانات غير صالحة",
                                 "أدخل مخزن وعبوة وكمية صحيحة (>= 1).", parent=self)
            return
        if not messagebox.askyesno("تأكيد التحويل",
                                   f"تحويل {qty} وحدة إلى المخزن {dest_wh} / العبوة {dest_pkg}؟",
                                   parent=self):
            return
        try:
            self.db.transfer_stock(stock_id, qty, dest_wh, dest_pkg)
            show_toast(self, f"تم التحويل بنجاح ({qty} وحدة)")
            self._load_source()
        except Exception as ex:
            messagebox.showerror("فشل التحويل", str(ex), parent=self)

# ------------------- Stock Audit Window -------------------

class StockAuditWindow(tk.Toplevel):
    """Physical stock count and adjustment."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("جرد المخزون")
        self.geometry("1050x620")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  جرد المخزون", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._stock_rows: List[Dict[str, Any]] = []
        self._build()

    def _build(self):
        # Filters
        filters = ttk.LabelFrame(self, text="تصفية الجرد")
        filters.pack(fill=tk.X, padx=8, pady=8)
        frow = ttk.Frame(filters)
        frow.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(frow, text="المخزن:").pack(side=tk.LEFT)
        self.f_wh = ttk.Combobox(frow, values=["", "1", "2", "3", "4"], width=5, state="readonly")
        self.f_wh.pack(side=tk.LEFT, padx=4)
        ttk.Label(frow, text="العبوة:").pack(side=tk.LEFT)
        self.f_pkg = ttk.Entry(frow, width=8)
        self.f_pkg.pack(side=tk.LEFT, padx=4)
        self.f_type = LabeledCombobox(frow, "النوع", self.db, "item_type")
        self.f_type.pack(side=tk.LEFT, padx=4)
        self.f_school = LabeledCombobox(frow, "المدرسة", self.db, "school")
        self.f_school.pack(side=tk.LEFT, padx=4)
        _bl = ttk.Button(frow, text="تحميل المخزون", command=self._load_stock); _bl.pack(side=tk.LEFT, padx=8)
        ToolTip(_bl, "تحميل الأصناف المطابقة للجرد")

        # Audit table
        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "wh", "pkg",
                     "expected", "actual", "diff"),
            show="headings", height=16, selectmode="browse")
        for col, txt, w in [
            ("id", "ID", 50), ("type", "النوع", 110), ("school", "المدرسة", 110),
            ("color", "اللون", 70), ("size", "المقاس", 60), ("wh", "المخزن", 50),
            ("pkg", "العبوة", 50), ("expected", "الكمية المتوقعة", 100),
            ("actual", "الكمية الفعلية", 100), ("diff", "الفرق", 70),
        ]:
            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=ysb.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        add_context_menu(self.table)
        _bind_mousewheel(self.table)
        self.table.tag_configure("surplus", background="#dcfce7")
        self.table.tag_configure("deficit", background="#fee2e2")

        # Edit actual qty
        edit_frame = ttk.Frame(self)
        edit_frame.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(edit_frame, text="الكمية الفعلية للبند المحدد:").pack(side=tk.LEFT)
        self._actual_var = tk.StringVar(value="0")
        self._actual_entry = ttk.Entry(edit_frame, textvariable=self._actual_var, width=8)
        self._actual_entry.pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="تعيين", command=self._set_actual).pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="مطابق (الكل)", command=self._mark_all_matched).pack(side=tk.LEFT, padx=4)

        # Action buttons
        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(btns, text="تطبيق التسويات", command=self._apply).pack(side=tk.RIGHT, padx=4)

    def _load_stock(self):
        filters = {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "warehouse_no": self.f_wh.get() or None,
            "package_no": self.f_pkg.get().strip() or None,
        }
        try:
            rows = self.db.current_inventory(filters)
        except Exception as ex:
            messagebox.showerror("فشل التحميل", str(ex), parent=self)
            return
        self._stock_rows = list(rows)
        self.table.delete(*self.table.get_children())
        for r in self._stock_rows:
            count = int(r["count"])
            self.table.insert("", tk.END, iid=str(r["id"]), values=(
                r["id"], r["item_type"], r["school"], r["color"], r["size"],
                r["warehouse_no"], r["package_no"], count, count, 0
            ))
        apply_zebra_tags(self.table)

    def _set_actual(self):
        sel = self.table.selection()
        if not sel:
            return
        try:
            actual = int(self._actual_var.get())
            if actual < 0:
                raise ValueError
        except ValueError:
            return
        vals = list(self.table.item(sel[0], "values"))
        expected = int(vals[7])
        diff = actual - expected
        vals[8] = actual
        vals[9] = diff
        tag = "surplus" if diff > 0 else ("deficit" if diff < 0 else "")
        self.table.item(sel[0], values=vals, tags=(tag,) if tag else ())

    def _mark_all_matched(self):
        for child in self.table.get_children():
            vals = list(self.table.item(child, "values"))
            vals[8] = vals[7]
            vals[9] = 0
            self.table.item(child, values=vals, tags=())

    def _apply(self):
        adjustments = []
        for child in self.table.get_children():
            vals = self.table.item(child, "values")
            expected = int(vals[7])
            actual = int(vals[8])
            if actual != expected:
                adjustments.append({
                    "stock_id": int(vals[0]),
                    "expected": expected,
                    "actual": actual,
                })
        if not adjustments:
            show_toast(self, "لا توجد تسويات للتطبيق", bg="#f59e0b")
            return
        if not messagebox.askyesno("تأكيد التسوية",
                                   f"سيتم تطبيق {len(adjustments)} تسوية على المخزون.\nهل أنت متأكد؟",
                                   parent=self):
            return
        try:
            count = self.db.apply_stock_adjustments(adjustments)
            show_toast(self, f"تم تطبيق {count} تسوية بنجاح")
            self._load_stock()
        except Exception as ex:
            messagebox.showerror("فشل التسوية", str(ex), parent=self)

# ------------------- Branch Inventory Queue -------------------

class BranchInventoryQueueWindow(tk.Toplevel):
    """Review branch returns/transfers before deciding where they go."""

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("منتجات الفروع غير المعالجة")
        self.geometry("1220x650")
        self.configure(bg=_UI["BG"])
        hdr = tk.Frame(self, bg=_UI["ACCENT"], height=44)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  منتجات الفروع غير المعالجة", bg=_UI["ACCENT"], fg="#FFFFFF",
                 font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()
        self._refresh()

    def _build(self):
        top = ttk.LabelFrame(self, text="المعالجة")
        top.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(top, text="الحالة:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self._status = tk.StringVar(value="PENDING")
        ttk.Combobox(top, textvariable=self._status,
                     values=["PENDING", "ASSIGNED", "DISCARDED", "REROUTED", ""],
                     state="readonly", width=16).grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="مخزن:").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self._wh_var = tk.StringVar(value="1")
        ttk.Entry(top, textvariable=self._wh_var, width=8).grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="عبوة:").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self._pkg_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._pkg_var, width=10).grid(row=0, column=5, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="إلى فرع:").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self._target_var = tk.StringVar(value="")
        self._target_cb = ttk.Combobox(top, textvariable=self._target_var, state="readonly", width=18)
        self._target_cb.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="ملاحظة:").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self._note_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._note_var, width=54).grid(
            row=1, column=3, columnspan=3, sticky="ew", padx=6, pady=6
        )
        top.columnconfigure(5, weight=1)

        btns = ttk.Frame(top)
        btns.grid(row=0, column=6, rowspan=2, sticky="ns", padx=8)
        ttk.Button(btns, text="تحديث", command=self._refresh).pack(side=tk.TOP, fill=tk.X, pady=2)
        ttk.Button(btns, text="إسناد للمخزون", command=self._assign_selected).pack(side=tk.TOP, fill=tk.X, pady=2)
        ttk.Button(btns, text="تحويل إلى فرع", command=self._reroute_selected).pack(side=tk.TOP, fill=tk.X, pady=2)
        ttk.Button(btns, text="إتلاف / معيب", command=self._discard_selected).pack(side=tk.TOP, fill=tk.X, pady=2)

        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        cols = ("id", "kind", "source", "target", "type", "school", "color", "size", "qty", "price", "status", "note")
        self._tree = ttk.Treeview(wrap, columns=cols, show="headings")
        for col, txt, w in [
            ("id", "#", 55), ("kind", "النوع", 80), ("source", "المرسل", 100), ("target", "الهدف", 100),
            ("type", "الصنف", 150), ("school", "المدرسة", 130), ("color", "اللون", 110), ("size", "المقاس", 70),
            ("qty", "الكمية", 65), ("price", "السعر", 80), ("status", "الحالة", 100), ("note", "ملاحظة", 220),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self._tree.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        _bind_mousewheel(self._tree)
        add_context_menu(self._tree)
        self._tree.bind("<<TreeviewSelect>>", lambda _e: self._load_selected_defaults())

        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, bg=_UI["BG"], fg=_UI["TEXT_SEC"]).pack(fill=tk.X, padx=8, pady=(0, 8))

    def _selected_queue_id(self) -> Optional[int]:
        sel = self._tree.selection()
        if not sel:
            return None
        return int(sel[0])

    def _load_selected_defaults(self):
        queue_id = self._selected_queue_id()
        row = self.db.get_branch_inventory_queue_item(queue_id) if queue_id is not None else None
        names = self.db.list_known_pos_device_names()
        if row:
            names = [n for n in names if n != row.get("source_device")]
        ui_names = [branch_display_name(n) for n in names]
        ui_to_dev = {branch_display_name(n): n for n in names}
        self._target_ui_to_dev = ui_to_dev
        self._target_cb["values"] = ui_names

        try:
            wh = int((self._wh_var.get() or "1").strip())
        except Exception:
            wh = 1
        if not (self._pkg_var.get() or "").strip():
            info = self.db.package_numbers_summary(wh)
            self._pkg_var.set(str(info.get("next") or 1))

        req = branch_display_name(row.get("requested_target_device")) if row else ""
        if row and req in ui_names:
            self._target_var.set(req)
        elif self._target_var.get() not in ui_names:
            self._target_var.set(ui_names[0] if ui_names else "")

    def _refresh(self):
        rows = self.db.list_branch_inventory_queue(self._status.get() or None)
        self._tree.delete(*self._tree.get_children())
        for row in rows:
            kind = "مرتجع" if row.get("queue_kind") == "RETURN" else "تحويل"
            self._tree.insert(
                "", tk.END, iid=str(row["id"]),
                values=(
                    row["id"], kind, branch_display_name(row.get("source_device") or ""),
                    branch_display_name(row.get("requested_target_device") or ""), row.get("item_type") or "",
                    row.get("school") or "", row.get("color") or "", row.get("size") or "",
                    row.get("qty") or 0, f"{float(row.get('unit_price') or 0):.2f}",
                    row.get("status") or "", row.get("note") or "",
                ),
            )
        apply_zebra_tags(self._tree)
        self._status_var.set(f"عدد العناصر: {len(rows)}")
        self._load_selected_defaults()

    def _assign_selected(self):
        queue_id = self._selected_queue_id()
        if queue_id is None:
            messagebox.showwarning("تنبيه", "اختر عنصرًا من القائمة أولاً.", parent=self)
            return
        try:
            w = int((self._wh_var.get() or "").strip())
            p = int((self._pkg_var.get() or "").strip())
        except Exception:
            messagebox.showerror("بيانات غير صالحة", "أدخل رقم مخزن وعبوة صحيحين.", parent=self)
            return
        try:
            self.db.assign_branch_inventory_queue_item(queue_id, w, p, self._note_var.get())
            show_toast(self, "تم إسناد العنصر إلى المخزون")
            self._note_var.set("")
            self._pkg_var.set("")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل الإسناد", str(ex), parent=self)

    def _discard_selected(self):
        queue_id = self._selected_queue_id()
        if queue_id is None:
            messagebox.showwarning("تنبيه", "اختر عنصرًا من القائمة أولاً.", parent=self)
            return
        if not messagebox.askyesno("تأكيد", "سيتم إتلاف هذا العنصر ولن يدخل المخزون. هل أنت متأكد؟", parent=self):
            return
        try:
            self.db.discard_branch_inventory_queue_item(queue_id, self._note_var.get())
            show_toast(self, "تم إتلاف / تعليم العنصر كمعيب")
            self._note_var.set("")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل الإتلاف", str(ex), parent=self)

    def _reroute_selected(self):
        queue_id = self._selected_queue_id()
        if queue_id is None:
            messagebox.showwarning("تنبيه", "اختر عنصرًا من القائمة أولاً.", parent=self)
            return
        target_ui = (self._target_var.get() or "").strip()
        target = getattr(self, "_target_ui_to_dev", {}).get(target_ui, target_ui)
        if not target:
            messagebox.showerror("بيانات ناقصة", "اختر الفرع الهدف أولاً.", parent=self)
            return
        try:
            self.db.reroute_branch_inventory_queue_item(queue_id, target, self._note_var.get())
            show_toast(self, f"تم تحويل العنصر إلى الفرع {branch_display_name(target)}")
            self._note_var.set("")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل التحويل", str(ex), parent=self)

# ------------------- Branch Stock View (Phase 3) -------------------

class BranchStockWindow(tk.Toplevel):
    """Warehouse-side view of each POS device's current stock.

    Reads from `pos_stocks_mirror` (filled by POS_STOCK_SNAPSHOT
    events). Shows a dropdown of known POS devices with their snapshot
    timestamps, and a filterable Treeview of the selected POS's stock.
    Includes a "مزامنة الآن" shortcut so the user can refresh without
    leaving the window.
    """

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("مخزون الفروع")
        self.geometry("1000x620")
        self.configure(bg=_UI["BG"])

        header = tk.Frame(self, bg=_UI["ACCENT"], height=44)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(
            header, text="  مخزون الفروع",
            bg=_UI["ACCENT"], fg="#FFFFFF",
            font=_FONTS["h3"],
        ).pack(side=tk.RIGHT, padx=12)

        self._build()
        self._reload_devices()

    # ---- layout ----

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="الفرع:").pack(side=tk.RIGHT, padx=(0, 4))
        self._device_var = tk.StringVar()
        self._device_cb = ttk.Combobox(
            top, textvariable=self._device_var,
            state="readonly", width=28,
        )
        self._device_cb.pack(side=tk.RIGHT, padx=4)
        self._device_cb.bind("<<ComboboxSelected>>", lambda _e: self._reload_stock())

        self._meta_var = tk.StringVar(value="—")
        ttk.Label(top, textvariable=self._meta_var,
                  foreground="#64748b").pack(side=tk.RIGHT, padx=(12, 0))

        ttk.Button(top, text="تحديث",
                   command=self._reload_devices).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="مزامنة الآن",
                   command=self._run_sync_and_reload).pack(side=tk.LEFT, padx=4)

        # Filter row
        filt = ttk.LabelFrame(self, text="تصنيف")
        filt.pack(fill=tk.X, padx=8, pady=(0, 4))

        self._flt_type = LabeledCombobox(filt, "النوع", self.db, "item_type")
        self._flt_school = LabeledCombobox(filt, "المدرسة", self.db, "school")
        self._flt_color = LabeledCombobox(filt, "اللون", self.db, "color")
        self._flt_size = LabeledCombobox(filt, "المقاس", self.db, "size")

        def _constraints(exclude=None):
            d = {
                "item_type": self._flt_type.get() or None,
                "school": self._flt_school.get() or None,
                "color": self._flt_color.get() or None,
                "size": self._flt_size.get() or None,
            }
            if exclude:
                d.pop(exclude, None)
            return d

        self._flt_type.set_supplier(lambda: self._branch_distinct("item_type", _constraints("item_type")))
        self._flt_school.set_supplier(lambda: self._branch_distinct("school", _constraints("school")))
        self._flt_color.set_supplier(lambda: self._branch_distinct("color", _constraints("color")))
        self._flt_size.set_supplier(lambda: self._branch_distinct("size", _constraints("size")))

        for i, w in enumerate((self._flt_type, self._flt_school, self._flt_color, self._flt_size)):
            w.grid(row=0, column=i, padx=6, pady=(6, 0), sticky="ew")
            filt.columnconfigure(i, weight=1)
            for ev in ("<<ComboboxSelected>>", "<KeyRelease>"):
                w.cb.bind(ev, lambda e: (self._refresh_filter_values(), self._apply_filter()), add="+")

        ttk.Label(filt, text="بحث:").grid(row=1, column=3, sticky="e", padx=(6, 4), pady=(6, 6))
        self._filter_var = tk.StringVar()
        ent = ttk.Entry(filt, textvariable=self._filter_var, width=28)
        ent.grid(row=1, column=2, sticky="ew", padx=6, pady=(6, 6))
        self._filter_var.trace_add("write", lambda *_: self._apply_filter())

        btns = ttk.Frame(filt)
        btns.grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 6))
        ttk.Button(btns, text="بحث", command=self._apply_filter).pack(side=tk.LEFT)
        ttk.Button(btns, text="مسح", command=self._clear_filters).pack(side=tk.LEFT, padx=8)

        # Table
        cols = ("item_type", "school", "color", "size", "unit_price", "count", "value")
        headers = {
            "item_type": "النوع", "school": "المدرسة", "color": "اللون",
            "size": "المقاس", "unit_price": "السعر", "count": "الكمية",
            "value": "القيمة",
        }
        widths = {
            "item_type": 120, "school": 110, "color": 90, "size": 70,
            "unit_price": 80, "count": 70, "value": 90,
        }
        self._tree = ttk.Treeview(
            self, columns=cols, show="headings", height=20,
        )
        for c in cols:
            self._tree.heading(c, text=headers[c])
            self._tree.column(c, width=widths[c], anchor="center")
        self._tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        apply_zebra_tags(self._tree)
        _bind_mousewheel(self._tree)

        # Status bar
        self._status_var = tk.StringVar(value="")
        tk.Label(
            self, textvariable=self._status_var,
            bg=_UI.get("BG", "#f1f5f9"),
            fg="#475569", anchor="e",
        ).pack(fill=tk.X, padx=8, pady=(0, 6))

    # ---- data ----

    def _reload_devices(self):
        rows = []
        try:
            rows = self.db.conn.execute(
                "SELECT source_device, snapshot_at, row_count, total_value "
                "FROM pos_stocks_snapshot_meta ORDER BY source_device"
            ).fetchall()
        except Exception:
            rows = []

        # Also include POS devices that are known but have never
        # snapshotted yet, so they appear in the dropdown.
        known = []
        try:
            known = self.db.conn.execute(
                "SELECT device_name FROM known_devices "
                "WHERE role = 'pos' ORDER BY device_name"
            ).fetchall()
        except Exception:
            known = []

        names = []
        self._metas = {}
        for r in rows:
            names.append(r[0])
            self._metas[r[0]] = {
                "snapshot_at": r[1],
                "row_count":   int(r[2] or 0),
                "total_value": float(r[3] or 0.0),
            }
        for k in known:
            if k[0] and k[0] not in names:
                names.append(k[0])

        self._device_cb["values"] = names
        self._device_ui_to_raw = {branch_display_name(n): n for n in names}
        ui_names = [branch_display_name(n) for n in names]
        self._device_cb["values"] = ui_names
        if names:
            current = self._device_var.get()
            if current not in ui_names:
                self._device_var.set(ui_names[0])
            self._reload_stock()
        else:
            self._device_var.set("")
            self._tree.delete(*self._tree.get_children())
            self._meta_var.set("لا توجد فروع مُسجَّلة — شغّل المزامنة أولاً")
            self._status_var.set("")

    def _reload_stock(self):
        pick = (self._device_var.get() or "").strip()
        name = getattr(self, "_device_ui_to_raw", {}).get(pick, pick)
        self._tree.delete(*self._tree.get_children())
        if not name:
            return
        meta = self._metas.get(name)
        if meta:
            self._meta_var.set(
                f"آخر لقطة: {meta['snapshot_at']}  |  "
                f"عدد الصفوف: {meta['row_count']}  |  "
                f"القيمة: {meta['total_value']:.2f}"
            )
        else:
            self._meta_var.set("لا توجد لقطة مخزون بعد لهذا الفرع")

        try:
            rows = self.db.conn.execute(
                """
                SELECT item_type, school, color, size, unit_price, count
                  FROM pos_stocks_mirror
                 WHERE source_device = ?
                 ORDER BY item_type, school, color, size
                """,
                (name,),
            ).fetchall()
        except Exception:
            rows = []

        self._all_rows = [
            (r[0], r[1], r[2], r[3], float(r[4] or 0), int(r[5] or 0))
            for r in rows
        ]
        self._refresh_filter_values()
        self._apply_filter()

    def _branch_distinct(self, field: str, constraints: Dict[str, Any]) -> List[str]:
        idx_map = {"item_type": 0, "school": 1, "color": 2, "size": 3}
        field_idx = idx_map[field]
        values = set()
        for row in getattr(self, "_all_rows", []):
            if constraints.get("item_type") and row[0] != constraints["item_type"]:
                continue
            if constraints.get("school") and row[1] != constraints["school"]:
                continue
            if constraints.get("color") and row[2] != constraints["color"]:
                continue
            if constraints.get("size") and row[3] != constraints["size"]:
                continue
            val = str(row[field_idx] or "").strip()
            if val:
                values.add(val)
        return sorted(values)

    def _refresh_filter_values(self):
        for w in (self._flt_type, self._flt_school, self._flt_color, self._flt_size):
            try:
                w.refresh_values()
            except Exception:
                pass

    def _clear_filters(self):
        for w in (self._flt_type, self._flt_school, self._flt_color, self._flt_size):
            try:
                w.set("")
            except Exception:
                pass
        self._filter_var.set("")
        self._refresh_filter_values()
        self._apply_filter()

    def _apply_filter(self):
        q = (self._filter_var.get() or "").strip().lower()
        item_type = (self._flt_type.get() or "").strip()
        school = (self._flt_school.get() or "").strip()
        color = (self._flt_color.get() or "").strip()
        size = (self._flt_size.get() or "").strip()
        self._tree.delete(*self._tree.get_children())
        shown = 0
        total_qty = 0
        total_val = 0.0
        for it, sc, cl, sz, price, count in self._all_rows:
            if item_type and it != item_type:
                continue
            if school and sc != school:
                continue
            if color and cl != color:
                continue
            if size and sz != size:
                continue
            if q:
                blob = (f"{it} {sc} {cl} {sz}").lower()
                if q not in blob:
                    continue
            value = price * count
            total_qty += count
            total_val += value
            self._tree.insert(
                "", tk.END,
                values=(
                    it, sc, cl, sz, f"{price:.2f}", count, f"{value:.2f}",
                ),
            )
            shown += 1
        self._status_var.set(
            f"يُعرض {shown} صف  |  الكمية: {total_qty}  |  "
            f"القيمة: {total_val:.2f}"
        )

    def _run_sync_and_reload(self):
        # Open the existing sync dialog so the user sees the live log
        # and control buttons rather than blocking this window.
        try:
            import sync_ui
            sync_ui.open_sync_dialog(self, self.db.conn)
        except Exception as ex:
            messagebox.showerror("المزامنة", str(ex), parent=self)
            return
        # When the sync dialog closes, refresh automatically.
        self.after(500, self._reload_devices)


# ------------------- Dashboard Frame -------------------

class DashboardFrame(ttk.Frame):
    """Overview dashboard with stats cards, recent bills, and low-stock alerts."""
    def __init__(self, master, db: SqliteDatabase, app=None):
        super().__init__(master, padding=10)
        self.db = db
        self.app = app
        self._build()
        self._refresh()

    def _build(self):
        # Page title
        title_f = tk.Frame(self, bg=_UI["SURFACE"])
        title_f.pack(fill=tk.X, padx=16, pady=(12, 4))
        tk.Label(title_f, text="لوحة التحكم", bg=_UI["SURFACE"], fg=_UI["TEXT"],
                 font=_FONTS["h1"]).pack(side=tk.RIGHT)
        tk.Label(title_f, text="نظرة عامة على المخزون والمبيعات", bg=_UI["SURFACE"],
                 fg=_UI["TEXT_SEC"], font=_FONTS["small"]).pack(side=tk.RIGHT, padx=(12, 0))
        ttk.Button(title_f, text="تحديث", command=self._refresh).pack(side=tk.LEFT)

        # Stats cards row
        cards_row = tk.Frame(self, bg=_UI["BG"])
        cards_row.pack(fill=tk.X, padx=12, pady=(8, 12))

        self._cards = {}
        card_defs = [
            ("total_items", "إجمالي الأصناف", _UI["BRAND"]),
            ("total_qty", "إجمالي الكميات", _UI["OK"]),
            ("total_value", "إجمالي القيمة", _UI["WARN"]),
            ("total_bills", "عدد الفواتير", "#8B5CF6"),
        ]
        for i, (key, title, color) in enumerate(card_defs):
            cards_row.columnconfigure(i, weight=1)
            outer = tk.Frame(cards_row, bg=_UI["BORDER"])
            outer.grid(row=0, column=i, padx=8, pady=4, sticky="nsew")
            card = tk.Frame(outer, bg=_UI["SURFACE"])
            card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
            # Colored accent strip at top
            tk.Frame(card, bg=color, height=3).pack(fill=tk.X)
            content = tk.Frame(card, bg=_UI["SURFACE"])
            content.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)
            tk.Label(content, text=title, bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
                     font=_FONTS["small"]).pack(anchor="e")
            val_lbl = tk.Label(content, text="0", bg=_UI["SURFACE"], fg=_UI["TEXT"],
                               font=_FONTS["big_num"])
            val_lbl.pack(anchor="e", pady=(4, 0))
            self._cards[key] = val_lbl

        # Two columns: Recent Bills + Low Stock
        bottom = tk.Frame(self, bg=_UI["BG"])
        bottom.pack(fill=tk.BOTH, expand=True, padx=12)

        # Recent bills card
        r_outer = tk.Frame(bottom, bg=_UI["BORDER"])
        r_outer.grid(row=0, column=0, padx=(0, 6), pady=4, sticky="nsew")
        r_card = tk.Frame(r_outer, bg=_UI["SURFACE"])
        r_card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        tk.Label(r_card, text="آخر الفواتير", bg=_UI["SURFACE"], fg=_UI["TEXT"],
                 font=_FONTS["h3"]).pack(anchor="e", padx=12, pady=(10, 4))

        self.recent_tree = ttk.Treeview(r_card, columns=("id", "date", "customer", "total"),
                                         show="headings", height=10)
        for col, txt, w in [("id", "#", 50), ("date", "التاريخ", 140),
                             ("customer", "العميل", 180), ("total", "الإجمالي", 100)]:
            self.recent_tree.heading(col, text=txt)
            self.recent_tree.column(col, width=w, anchor="center")
        self.recent_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.recent_tree)
        _bind_mousewheel(self.recent_tree)

        # Low stock card
        a_outer = tk.Frame(bottom, bg=_UI["BORDER"])
        a_outer.grid(row=0, column=1, padx=(6, 0), pady=4, sticky="nsew")
        a_card = tk.Frame(a_outer, bg=_UI["SURFACE"])
        a_card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        tk.Label(a_card, text="تنبيهات المخزون المنخفض (أقل من 5)", bg=_UI["SURFACE"],
                 fg=_UI["TEXT"], font=_FONTS["h3"]).pack(anchor="e", padx=12, pady=(10, 4))

        self.alert_tree = ttk.Treeview(a_card, columns=("item", "school", "color", "size", "qty"),
                                        show="headings", height=10)
        for col, txt, w in [("item", "النوع", 120), ("school", "المدرسة", 120),
                            ("color", "اللون", 80), ("size", "المقاس", 60), ("qty", "الكمية", 60)]:
            self.alert_tree.heading(col, text=txt)
            self.alert_tree.column(col, width=w, anchor="center")
        self.alert_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.alert_tree)
        _bind_mousewheel(self.alert_tree)

        bottom.columnconfigure(0, weight=1)
        bottom.columnconfigure(1, weight=1)

        # Sync health card (warehouse-central visibility)
        s_outer = tk.Frame(self, bg=_UI["BORDER"])
        s_outer.pack(fill=tk.BOTH, expand=False, padx=12, pady=(8, 4))
        s_card = tk.Frame(s_outer, bg=_UI["SURFACE"])
        s_card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        tk.Label(
            s_card,
            text="صحة المزامنة (الفروع)",
            bg=_UI["SURFACE"],
            fg=_UI["TEXT"],
            font=_FONTS["h3"],
        ).pack(anchor="e", padx=12, pady=(10, 4))
        self._sync_summary_var = tk.StringVar(value="")
        tk.Label(
            s_card,
            textvariable=self._sync_summary_var,
            bg=_UI["SURFACE"],
            fg=_UI["TEXT_SEC"],
            font=_FONTS["small"],
            anchor="e",
            justify="right",
        ).pack(fill=tk.X, padx=12, pady=(0, 4))
        act = ttk.Frame(s_card)
        act.pack(fill=tk.X, padx=8, pady=(0, 6))
        ttk.Button(
            act,
            text="مزامنة الآن",
            command=(lambda: self.app._open_sync_dialog()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="طابور الفروع",
            command=(lambda: self.app._open_branch_inventory_queue()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="مرآة حجوزات الفروع",
            command=(lambda: self.app._open_pos_reservations_mirror()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        self.sync_tree = ttk.Treeview(
            s_card,
            columns=("branch", "status", "last_snapshot", "age_min", "rows", "value", "sync_errors"),
            show="headings",
            height=6,
        )
        for col, txt, w in [
            ("branch", "الفرع", 180),
            ("status", "الحالة", 100),
            ("last_snapshot", "آخر لقطة", 190),
            ("age_min", "عمر اللقطة (د)", 110),
            ("rows", "الصفوف", 90),
            ("value", "القيمة", 120),
            ("sync_errors", "أخطاء مزامنة", 120),
        ]:
            self.sync_tree.heading(col, text=txt)
            self.sync_tree.column(col, width=w, anchor="center")
        self.sync_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.sync_tree)
        _bind_mousewheel(self.sync_tree)

        ex_outer = tk.Frame(self, bg=_UI["BORDER"])
        ex_outer.pack(fill=tk.BOTH, expand=False, padx=12, pady=(4, 4))
        ex_card = tk.Frame(ex_outer, bg=_UI["SURFACE"])
        ex_card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        tk.Label(
            ex_card,
            text="استثناءات المزامنة (أعلى الأسباب)",
            bg=_UI["SURFACE"],
            fg=_UI["TEXT"],
            font=_FONTS["h3"],
        ).pack(anchor="e", padx=12, pady=(10, 4))
        self.sync_ex_tree = ttk.Treeview(
            ex_card,
            columns=("etype", "count", "last_error"),
            show="headings",
            height=4,
        )
        for col, txt, w in [
            ("etype", "نوع الحدث", 220),
            ("count", "عدد الأخطاء", 100),
            ("last_error", "آخر خطأ", 760),
        ]:
            self.sync_ex_tree.heading(col, text=txt)
            self.sync_ex_tree.column(col, width=w, anchor="center")
        self.sync_ex_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.sync_ex_tree)
        _bind_mousewheel(self.sync_ex_tree)

    def _refresh(self):
        try:
            cur = self.db.conn.cursor()

            cur.execute("SELECT COUNT(DISTINCT item_type || '|' || school || '|' || color || '|' || size) FROM stocks WHERE count > 0")
            self._cards["total_items"].config(text=str(cur.fetchone()[0] or 0))

            cur.execute("SELECT COALESCE(SUM(count), 0) FROM stocks WHERE count > 0")
            self._cards["total_qty"].config(text=f"{cur.fetchone()[0]:,}")

            cur.execute("SELECT COALESCE(SUM(count * unit_price), 0) FROM stocks WHERE count > 0")
            val = cur.fetchone()[0] or 0
            self._cards["total_value"].config(text=f"{val:,.2f}")

            cur.execute("SELECT COUNT(*) FROM bills WHERE COALESCE(status,'CONFIRMED')='CONFIRMED'")
            self._cards["total_bills"].config(text=str(cur.fetchone()[0] or 0))

            # Recent bills
            self.recent_tree.delete(*self.recent_tree.get_children())
            cur.execute("SELECT id, created_at, customer, total FROM bills WHERE COALESCE(status,'CONFIRMED')='CONFIRMED' ORDER BY id DESC LIMIT 10")
            for r in cur.fetchall():
                self.recent_tree.insert("", tk.END, values=(r[0], r[1], r[2] or "", f"{r[3]:.2f}"))
            apply_zebra_tags(self.recent_tree)

            # Low stock alerts
            self.alert_tree.delete(*self.alert_tree.get_children())
            cur.execute("""
                SELECT item_type, school, color, size, SUM(count) as total_count
                FROM stocks WHERE count > 0
                GROUP BY item_type, school, color, size
                HAVING total_count < 5
                ORDER BY total_count ASC
                LIMIT 20
            """)
            for r in cur.fetchall():
                self.alert_tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3], r[4]))
            apply_zebra_tags(self.alert_tree)

            # Sync health summary for POS mirrors
            self.sync_tree.delete(*self.sync_tree.get_children())
            q = """
                SELECT
                    COALESCE(kd.device_name, pm.source_device) AS branch_name,
                    pm.snapshot_at,
                    pm.row_count,
                    pm.total_value
                FROM pos_stocks_snapshot_meta pm
                LEFT JOIN known_devices kd
                    ON kd.device_name = pm.source_device
                    OR kd.device_uuid = pm.source_device
                ORDER BY branch_name
            """
            rows = cur.execute(q).fetchall()
            now = datetime.now()
            ok_count = 0
            warn_count = 0
            critical_count = 0
            for r in rows:
                snap = str(r[1] or "")
                age_min: Optional[int] = None
                if snap:
                    try:
                        dt = datetime.fromisoformat(snap.replace("Z", "+00:00"))
                        age_min = max(0, int((datetime.now(dt.tzinfo) - dt).total_seconds() // 60))
                    except Exception:
                        age_min = None
                if age_min is None:
                    status = "غير معروف"
                    warn_count += 1
                elif age_min >= 60:
                    status = "حرج"
                    critical_count += 1
                elif age_min >= 15:
                    status = "تحذير"
                    warn_count += 1
                else:
                    status = "جيد"
                    ok_count += 1
                self.sync_tree.insert(
                    "",
                    tk.END,
                    values=(
                        branch_display_name(r[0] or ""),
                        status,
                        snap,
                        age_min if age_min is not None else "",
                        int(r[2] or 0),
                        f"{float(r[3] or 0.0):,.2f}",
                        "—",
                    ),
                )
            apply_zebra_tags(self.sync_tree)

            # Global sync backlog/error counters from local sync tables
            outbox_pending = 0
            inbox_errors = 0
            queue_pending = 0
            try:
                outbox_pending = int(cur.execute("SELECT COUNT(*) FROM sync_outbox WHERE status='pending'").fetchone()[0] or 0)
                inbox_errors = int(cur.execute("SELECT COUNT(*) FROM sync_inbox WHERE apply_status='error'").fetchone()[0] or 0)
                queue_pending = int(cur.execute("SELECT COUNT(*) FROM branch_inventory_queue WHERE status='PENDING'").fetchone()[0] or 0)
            except Exception:
                pass
            self._sync_summary_var.set(
                f"فروع سليمة: {ok_count}  |  تحذير: {warn_count}  |  حرج: {critical_count}  |  "
                f"أحداث صادرة قيد الانتظار: {outbox_pending}  |  أخطاء تطبيق وارد: {inbox_errors}  |  "
                f"طابور الفرع (معلّق): {queue_pending}"
            )

            # Top sync apply error causes
            self.sync_ex_tree.delete(*self.sync_ex_tree.get_children())
            try:
                ex_rows = cur.execute(
                    """
                    SELECT event_type, COUNT(*) AS c, MAX(COALESCE(apply_error,'')) AS last_error
                    FROM sync_inbox
                    WHERE apply_status='error'
                    GROUP BY event_type
                    ORDER BY c DESC, event_type
                    LIMIT 8
                    """
                ).fetchall()
            except Exception:
                ex_rows = []
            for er in ex_rows:
                self.sync_ex_tree.insert(
                    "",
                    tk.END,
                    values=(
                        er[0] or "",
                        int(er[1] or 0),
                        str(er[2] or "")[:350],
                    ),
                )
            apply_zebra_tags(self.sync_ex_tree)

            cur.close()
        except Exception:
            pass


class SyncDiagnosticsFrame(ttk.Frame):
    """Dedicated sync diagnostics tab (branch health + top exceptions)."""
    def __init__(self, master, db: SqliteDatabase, app=None):
        super().__init__(master, padding=10)
        self.db = db
        self.app = app
        self._build()
        self._refresh()

    def _build(self):
        ttk.Label(self, text="تشخيص المزامنة", font=_FONTS["h1"]).pack(anchor="w", pady=(0, 8))

        self._sync_summary_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._sync_summary_var).pack(fill=tk.X, pady=(0, 6))

        act = ttk.Frame(self)
        act.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            act,
            text="مزامنة الآن",
            command=(lambda: self.app._open_sync_dialog()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="طابور الفروع",
            command=(lambda: self.app._open_branch_inventory_queue()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="مرآة حجوزات الفروع",
            command=(lambda: self.app._open_pos_reservations_mirror()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)

        health_fr = ttk.LabelFrame(self, text="صحة الفروع")
        health_fr.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.sync_tree = ttk.Treeview(
            health_fr,
            columns=("branch", "status", "last_snapshot", "age_min", "rows", "value", "sync_errors"),
            show="headings",
            height=8,
        )
        for col, txt, w in [
            ("branch", "الفرع", 180),
            ("status", "الحالة", 100),
            ("last_snapshot", "آخر لقطة", 190),
            ("age_min", "عمر اللقطة (د)", 110),
            ("rows", "الصفوف", 90),
            ("value", "القيمة", 120),
            ("sync_errors", "أخطاء مزامنة", 120),
        ]:
            self.sync_tree.heading(col, text=txt)
            self.sync_tree.column(col, width=w, anchor="center")
        self.sync_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.sync_tree)
        _bind_mousewheel(self.sync_tree)

        ex_fr = ttk.LabelFrame(self, text="استثناءات المزامنة (أعلى الأسباب)")
        ex_fr.pack(fill=tk.BOTH, expand=False)
        self.sync_ex_tree = ttk.Treeview(
            ex_fr,
            columns=("etype", "count", "last_error"),
            show="headings",
            height=6,
        )
        for col, txt, w in [
            ("etype", "نوع الحدث", 220),
            ("count", "عدد الأخطاء", 100),
            ("last_error", "آخر خطأ", 760),
        ]:
            self.sync_ex_tree.heading(col, text=txt)
            self.sync_ex_tree.column(col, width=w, anchor="center")
        self.sync_ex_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.sync_ex_tree)
        _bind_mousewheel(self.sync_ex_tree)

    def _refresh(self):
        try:
            cur = self.db.conn.cursor()

            self.sync_tree.delete(*self.sync_tree.get_children())
            q = """
                SELECT
                    COALESCE(kd.device_name, pm.source_device) AS branch_name,
                    pm.snapshot_at,
                    pm.row_count,
                    pm.total_value
                FROM pos_stocks_snapshot_meta pm
                LEFT JOIN known_devices kd
                    ON kd.device_name = pm.source_device
                    OR kd.device_uuid = pm.source_device
                ORDER BY branch_name
            """
            rows = cur.execute(q).fetchall()
            ok_count = 0
            warn_count = 0
            critical_count = 0
            for r in rows:
                snap = str(r[1] or "")
                age_min: Optional[int] = None
                if snap:
                    try:
                        dt = datetime.fromisoformat(snap.replace("Z", "+00:00"))
                        age_min = max(0, int((datetime.now(dt.tzinfo) - dt).total_seconds() // 60))
                    except Exception:
                        age_min = None
                if age_min is None:
                    status = "غير معروف"
                    warn_count += 1
                elif age_min >= 60:
                    status = "حرج"
                    critical_count += 1
                elif age_min >= 15:
                    status = "تحذير"
                    warn_count += 1
                else:
                    status = "جيد"
                    ok_count += 1
                self.sync_tree.insert(
                    "",
                    tk.END,
                    values=(
                        branch_display_name(r[0] or ""),
                        status,
                        snap,
                        age_min if age_min is not None else "",
                        int(r[2] or 0),
                        f"{float(r[3] or 0.0):,.2f}",
                        "—",
                    ),
                )
            apply_zebra_tags(self.sync_tree)

            outbox_pending = 0
            inbox_errors = 0
            queue_pending = 0
            try:
                outbox_pending = int(cur.execute("SELECT COUNT(*) FROM sync_outbox WHERE status='pending'").fetchone()[0] or 0)
                inbox_errors = int(cur.execute("SELECT COUNT(*) FROM sync_inbox WHERE apply_status='error'").fetchone()[0] or 0)
                queue_pending = int(cur.execute("SELECT COUNT(*) FROM branch_inventory_queue WHERE status='PENDING'").fetchone()[0] or 0)
            except Exception:
                pass
            self._sync_summary_var.set(
                f"فروع سليمة: {ok_count}  |  تحذير: {warn_count}  |  حرج: {critical_count}  |  "
                f"أحداث صادرة قيد الانتظار: {outbox_pending}  |  أخطاء تطبيق وارد: {inbox_errors}  |  "
                f"طابور الفرع (معلّق): {queue_pending}"
            )

            self.sync_ex_tree.delete(*self.sync_ex_tree.get_children())
            try:
                ex_rows = cur.execute(
                    """
                    SELECT event_type, COUNT(*) AS c, MAX(COALESCE(apply_error,'')) AS last_error
                    FROM sync_inbox
                    WHERE apply_status='error'
                    GROUP BY event_type
                    ORDER BY c DESC, event_type
                    LIMIT 8
                    """
                ).fetchall()
            except Exception:
                ex_rows = []
            for er in ex_rows:
                self.sync_ex_tree.insert(
                    "",
                    tk.END,
                    values=(
                        er[0] or "",
                        int(er[1] or 0),
                        str(er[2] or "")[:350],
                    ),
                )
            apply_zebra_tags(self.sync_ex_tree)
            cur.close()
        except Exception:
            pass


# ------------------- Statistics Frame -------------------

class StatisticsFrame(ttk.Frame):
    """Statistics and reports tab with date-range filtering."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        self._build()
        self._refresh()

    def _build(self):
        ttk.Label(self, text="الإحصائيات والتقارير", font=_FONTS["h1"]).pack(anchor="w", pady=(0, 12))

        # Date range filters (stays outside the scroll area)
        filters = ttk.Frame(self)
        filters.pack(fill=tk.X, pady=(0, 8))

        self.df = DateField(filters, "من")
        self.df.pack(side=tk.LEFT, padx=(0, 8))
        self.dt = DateField(filters, "إلى")
        self.dt.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(filters, text="تحديث", command=self._refresh).pack(side=tk.LEFT, padx=8)

        # Scrollable container for all stats sections
        canvas = tk.Canvas(self, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        inner = ttk.Frame(canvas)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_configure(e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_configure)

        def _on_canvas_configure(e):
            canvas.itemconfig(inner_id, width=e.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        _bind_mousewheel(inner, canvas)

        # ---- Summary cards ----
        summary = ttk.LabelFrame(inner, text="ملخص الفترة")
        summary.pack(fill=tk.X, pady=(0, 8), padx=4)

        self._stats = {}
        stat_defs = [
            ("bills_count", "عدد الفواتير"),
            ("bills_total", "إجمالي المبيعات"),
            ("items_in", "كميات واردة"),
            ("items_out", "كميات منصرفة"),
        ]
        for i, (key, label) in enumerate(stat_defs):
            ttk.Label(summary, text=label + ":").grid(row=0, column=i*2, padx=6, pady=8, sticky="e")
            val = ttk.Label(summary, text="0", font=("Segoe UI", 12, "bold"))
            val.grid(row=0, column=i*2+1, padx=(0, 16), pady=8, sticky="w")
            self._stats[key] = val
            summary.columnconfigure(i*2+1, weight=1)

        # ---- Top selling items ----
        top_frame = ttk.LabelFrame(inner, text="أكثر الأصناف مبيعاً")
        top_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        self.top_tree = ttk.Treeview(top_frame,
            columns=("item", "school", "color", "qty", "revenue"),
            show="headings", height=10)
        for col, txt, w in [("item", "النوع", 160), ("school", "المدرسة", 160),
                            ("color", "اللون", 120), ("qty", "الكمية المباعة", 110),
                            ("revenue", "الإيراد", 130)]:
            self.top_tree.heading(col, text=txt)
            self.top_tree.column(col, width=w, anchor="center")
        ysb2 = ttk.Scrollbar(top_frame, orient="vertical", command=self.top_tree.yview)
        self.top_tree.configure(yscrollcommand=ysb2.set)
        self.top_tree.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4, pady=4)
        ysb2.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        add_context_menu(self.top_tree)
        _bind_mousewheel(self.top_tree)

        # ---- Sales Trends ----
        trends_frame = ttk.LabelFrame(inner, text="اتجاهات المبيعات")
        trends_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        self._trend_mode = tk.StringVar(value="daily")
        mode_row = ttk.Frame(trends_frame)
        mode_row.pack(fill=tk.X, padx=4, pady=4)
        for val, lbl in [("daily", "يومي"), ("weekly", "أسبوعي"), ("monthly", "شهري")]:
            ttk.Radiobutton(mode_row, text=lbl, variable=self._trend_mode,
                            value=val, command=self._refresh).pack(side=tk.LEFT, padx=4)

        self.trend_tree = ttk.Treeview(trends_frame,
            columns=("period", "bills_count", "total_sales", "items_sold"),
            show="headings", height=8)
        for col, txt, w in [
            ("period", "الفترة", 140), ("bills_count", "عدد الفواتير", 100),
            ("total_sales", "إجمالي المبيعات", 120), ("items_sold", "الأصناف المباعة", 100),
        ]:
            self.trend_tree.heading(col, text=txt)
            self.trend_tree.column(col, width=w, anchor="center")
        self.trend_tree.pack(fill=tk.X, expand=True, padx=4, pady=4)
        add_context_menu(self.trend_tree)
        _bind_mousewheel(self.trend_tree)

        # ---- Dead Stock / Slow Movers ----
        dead_frame = ttk.LabelFrame(inner, text="أصناف راكدة (بدون حركة بيع)")
        dead_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        days_row = ttk.Frame(dead_frame)
        days_row.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(days_row, text="بدون بيع لمدة (أيام):").pack(side=tk.LEFT)
        self._dead_days_var = tk.StringVar(value="30")
        ttk.Entry(days_row, textvariable=self._dead_days_var, width=6).pack(side=tk.LEFT, padx=4)
        ttk.Button(days_row, text="تحديث", command=self._refresh).pack(side=tk.LEFT, padx=4)

        self.dead_tree = ttk.Treeview(dead_frame,
            columns=("type", "school", "color", "size", "qty", "value", "last_sale"),
            show="headings", height=8)
        for col, txt, w in [
            ("type", "النوع", 120), ("school", "المدرسة", 120), ("color", "اللون", 80),
            ("size", "المقاس", 60), ("qty", "الكمية", 60), ("value", "القيمة", 90),
            ("last_sale", "آخر بيع", 120),
        ]:
            self.dead_tree.heading(col, text=txt)
            self.dead_tree.column(col, width=w, anchor="center")
        self.dead_tree.pack(fill=tk.X, expand=True, padx=4, pady=4)
        add_context_menu(self.dead_tree)
        _bind_mousewheel(self.dead_tree)

        # ---- Inventory Valuation by Warehouse ----
        val_frame = ttk.LabelFrame(inner, text="تقييم المخزون حسب المخزن")
        val_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        self.val_tree = ttk.Treeview(val_frame,
            columns=("wh", "items_count", "total_qty", "total_value"),
            show="headings", height=5)
        for col, txt, w in [
            ("wh", "المخزن", 80), ("items_count", "عدد الأصناف", 100),
            ("total_qty", "إجمالي الكميات", 100), ("total_value", "إجمالي القيمة", 120),
        ]:
            self.val_tree.heading(col, text=txt)
            self.val_tree.column(col, width=w, anchor="center")
        self.val_tree.pack(fill=tk.X, expand=True, padx=4, pady=4)
        add_context_menu(self.val_tree)
        _bind_mousewheel(self.val_tree)

    def _refresh(self):
        try:
            date_from = self.df.get() or None
            date_to = self.dt.get() or None

            cur = self.db.conn.cursor()

            # Bills stats (confirmed only)
            where = ["COALESCE(status,'CONFIRMED')='CONFIRMED'"]
            args: List[Any] = []
            if date_from:
                where.append("date(created_at) >= date(?)")
                args.append(date_from)
            if date_to:
                where.append("date(created_at) <= date(?)")
                args.append(date_to)

            cur.execute(f"SELECT COUNT(*), COALESCE(SUM(total), 0) FROM bills WHERE {' AND '.join(where)}", args)
            r = cur.fetchone()
            self._stats["bills_count"].config(text=str(r[0] or 0))
            self._stats["bills_total"].config(text=f"{r[1] or 0:,.2f}")

            # Movements IN/OUT
            mwhere = ["1=1"]
            margs: List[Any] = []
            if date_from:
                mwhere.append("date(ts) >= date(?)")
                margs.append(date_from)
            if date_to:
                mwhere.append("date(ts) <= date(?)")
                margs.append(date_to)

            cur.execute(f"SELECT COALESCE(SUM(qty), 0) FROM movements WHERE direction='IN' AND {' AND '.join(mwhere)}", margs)
            self._stats["items_in"].config(text=f"{cur.fetchone()[0] or 0:,}")

            cur.execute(f"SELECT COALESCE(SUM(qty), 0) FROM movements WHERE direction IN ('OUT','OUT_FACTORY') AND {' AND '.join(mwhere)}", margs)
            self._stats["items_out"].config(text=f"{cur.fetchone()[0] or 0:,}")

            # Top items
            self.top_tree.delete(*self.top_tree.get_children())
            bi_where = ["COALESCE(b.status,'CONFIRMED')='CONFIRMED'"]
            bi_args: List[Any] = []
            if date_from:
                bi_where.append("date(b.created_at) >= date(?)")
                bi_args.append(date_from)
            if date_to:
                bi_where.append("date(b.created_at) <= date(?)")
                bi_args.append(date_to)

            cur.execute(f"""
                SELECT bi.item_type, bi.school, bi.color,
                       SUM(bi.qty) as total_qty, SUM(bi.line_total) as total_revenue
                FROM bill_items bi
                JOIN bills b ON bi.bill_id = b.id
                WHERE {' AND '.join(bi_where)}
                GROUP BY bi.item_type, bi.school, bi.color
                ORDER BY total_qty DESC
                LIMIT 20
            """, bi_args)

            for r in cur.fetchall():
                self.top_tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3], f"{r[4]:,.2f}"))
            apply_zebra_tags(self.top_tree)

            # ---- Sales Trends ----
            self.trend_tree.delete(*self.trend_tree.get_children())
            mode = self._trend_mode.get()
            if mode == "daily":
                group_expr = "date(b.created_at)"
            elif mode == "weekly":
                group_expr = "strftime('%Y-W%W', b.created_at)"
            else:
                group_expr = "strftime('%Y-%m', b.created_at)"

            trend_where = ["COALESCE(b.status,'CONFIRMED')='CONFIRMED'"]
            trend_args: List[Any] = []
            if date_from:
                trend_where.append("date(b.created_at) >= date(?)")
                trend_args.append(date_from)
            if date_to:
                trend_where.append("date(b.created_at) <= date(?)")
                trend_args.append(date_to)

            cur.execute(f"""
                SELECT {group_expr} AS period,
                       COUNT(DISTINCT b.id) AS bills_count,
                       COALESCE(SUM(b.total), 0) AS total_sales,
                       COALESCE(SUM(bi.qty), 0) AS items_sold
                FROM bills b
                LEFT JOIN bill_items bi ON bi.bill_id = b.id
                WHERE {' AND '.join(trend_where)}
                GROUP BY period
                ORDER BY period DESC
                LIMIT 30
            """, trend_args)
            for r in cur.fetchall():
                self.trend_tree.insert("", tk.END, values=(
                    r[0], r[1], f"{r[2]:,.2f}", r[3]
                ))
            apply_zebra_tags(self.trend_tree)

            # ---- Dead Stock / Slow Movers ----
            self.dead_tree.delete(*self.dead_tree.get_children())
            try:
                days = int(self._dead_days_var.get() or 30)
            except ValueError:
                days = 30

            cur.execute(f"""
                SELECT s.item_type, s.school, s.color, s.size,
                       SUM(s.count) AS qty,
                       SUM(s.count * s.unit_price) AS value,
                       (SELECT MAX(m.ts) FROM movements m
                        WHERE m.item_type = s.item_type AND m.school = s.school
                          AND m.color = s.color AND m.size = s.size
                          AND m.direction IN ('OUT', 'OUT_FACTORY')) AS last_sale
                FROM stocks s
                WHERE s.count > 0
                GROUP BY s.item_type, s.school, s.color, s.size
                HAVING last_sale IS NULL OR date(last_sale) < date('now', '-{days} days')
                ORDER BY value DESC
                LIMIT 20
            """)
            for r in cur.fetchall():
                self.dead_tree.insert("", tk.END, values=(
                    r[0], r[1], r[2], r[3], r[4], f"{r[5]:,.2f}",
                    r[6] or "لا يوجد"
                ))
            apply_zebra_tags(self.dead_tree)

            # ---- Inventory Valuation by Warehouse ----
            self.val_tree.delete(*self.val_tree.get_children())
            cur.execute("""
                SELECT warehouse_no,
                       COUNT(DISTINCT item_type || '|' || school || '|' || color || '|' || size) AS items_count,
                       SUM(count) AS total_qty,
                       SUM(count * unit_price) AS total_value
                FROM stocks WHERE count > 0
                GROUP BY warehouse_no
                ORDER BY warehouse_no
            """)
            for r in cur.fetchall():
                self.val_tree.insert("", tk.END, values=(
                    r[0], r[1], f"{r[2]:,}", f"{r[3]:,.2f}"
                ))
            apply_zebra_tags(self.val_tree)

            cur.close()
        except Exception:
            pass


# ------------------- Branch / POS reporting (warehouse) -------------------

class BranchBillsSyncLogWindow(tk.Toplevel):
    """Shipments to branches (warehouse bills) + synced inbound queue."""

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master)
        self.db = db
        self.title("سجل فواتير الفروع والوارد المتزامن")
        self.geometry("1080x620")
        self._build()

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tab_b = ttk.Frame(nb)
        nb.add(tab_b, text="شحنات صادرة (فواتير فرع)")
        top_b = ttk.Frame(tab_b)
        top_b.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(top_b, text="تحديث", command=lambda: self._fill_branch_bills()).pack(side=tk.RIGHT)
        wrap_b = ttk.Frame(tab_b)
        wrap_b.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        cols_b = ("id", "created_at", "customer", "total", "status")
        self._tree_b = ttk.Treeview(wrap_b, columns=cols_b, show="headings", height=18)
        for col, txt, w in [
            ("id", "#", 70),
            ("created_at", "التاريخ", 170),
            ("customer", "الفرع / العميل", 260),
            ("total", "الإجمالي", 100),
            ("status", "الحالة", 90),
        ]:
            self._tree_b.heading(col, text=txt)
            self._tree_b.column(col, width=w, anchor="center")
        sb_b = ttk.Scrollbar(wrap_b, orient="vertical", command=self._tree_b.yview)
        self._tree_b.configure(yscrollcommand=sb_b.set)
        self._tree_b.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_b.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree_b)

        tab_q = ttk.Frame(nb)
        nb.add(tab_q, text="وارد متزامن (طابور الفروع)")
        top_q = ttk.Frame(tab_q)
        top_q.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(top_q, text="تحديث", command=lambda: self._fill_queue()).pack(side=tk.RIGHT)
        wrap_q = ttk.Frame(tab_q)
        wrap_q.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        cols_q = ("id", "kind", "source", "target", "type", "school", "color", "size", "qty", "status", "created_at")
        self._tree_q = ttk.Treeview(wrap_q, columns=cols_q, show="headings", height=18)
        for col, txt, w in [
            ("id", "#", 55),
            ("kind", "النوع", 80),
            ("source", "المرسل", 110),
            ("target", "الهدف", 110),
            ("type", "الصنف", 140),
            ("school", "المدرسة", 120),
            ("color", "اللون", 80),
            ("size", "المقاس", 70),
            ("qty", "الكمية", 60),
            ("status", "الحالة", 100),
            ("created_at", "التاريخ", 150),
        ]:
            self._tree_q.heading(col, text=txt)
            self._tree_q.column(col, width=w, anchor="center")
        sb_q = ttk.Scrollbar(wrap_q, orient="vertical", command=self._tree_q.yview)
        self._tree_q.configure(yscrollcommand=sb_q.set)
        self._tree_q.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_q.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree_q)

        self._fill_branch_bills()
        self._fill_queue()

    def _fill_branch_bills(self):
        self._tree_b.delete(*self._tree_b.get_children())
        status_map = {"DRAFT": "مسودة", "CONFIRMED": "مؤكدة", "VOID": "ملغاة"}
        for b in self.db.list_bills():
            if b.get("bill_kind") != "BRANCH_SHIPMENT":
                continue
            st = status_map.get(b.get("status", "CONFIRMED"), b.get("status", ""))
            self._tree_b.insert(
                "", tk.END,
                values=(
                    b["id"],
                    b.get("created_at") or "",
                    b.get("customer") or "",
                    f"{float(b.get('total') or 0):.2f}",
                    st,
                ),
            )
        apply_zebra_tags(self._tree_b)

    def _fill_queue(self):
        self._tree_q.delete(*self._tree_q.get_children())
        for row in self.db.list_branch_inventory_queue(None):
            kind = "مرتجع" if row.get("queue_kind") == "RETURN" else "تحويل"
            self._tree_q.insert(
                "", tk.END,
                values=(
                    row.get("id"),
                    kind,
                    row.get("source_device") or "",
                    row.get("requested_target_device") or "",
                    row.get("item_type") or "",
                    row.get("school") or "",
                    row.get("color") or "",
                    row.get("size") or "",
                    row.get("qty") or 0,
                    row.get("status") or "",
                    row.get("created_at") or "",
                ),
            )
        apply_zebra_tags(self._tree_q)


class PosReservationsMirrorWindow(tk.Toplevel):
    """Reserved products per POS device (from sync mirror)."""

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master)
        self.db = db
        self.title("حجوزات الفروع (مرآة من المزامنة)")
        self.geometry("1220x600")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="نقطة البيع:").pack(side=tk.RIGHT, padx=4)
        pick = self.db.list_pos_reservations_mirror_device_picklist()
        self._dev = tk.StringVar(value="")
        self._dev_cb = ttk.Combobox(
            top, textvariable=self._dev, values=pick, width=34, state="readonly",
        )
        self._dev_cb.pack(side=tk.RIGHT)
        self._dev_cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh())
        self._active_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top, text="المعلقة فقط", variable=self._active_only,
            command=self._refresh,
        ).pack(side=tk.RIGHT, padx=12)
        ttk.Button(top, text="تحديث", command=self._refresh).pack(side=tk.LEFT)

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        tab_d = ttk.Frame(nb)
        nb.add(tab_d, text="تفصيلي (كل حجز)")
        ttk.Label(
            tab_d,
            text=(
                "كل سطر هنا = سطر حجز واحد في المزامنة (معرّفات مختلفة تظهر منفصلة). "
                "لرؤية إجمالي الكمية لنفس المنتج (نوع + مدرسة + لون + مقاس) بعد جمع كل الأسطر، "
                "استخدم تبويب «مجمّع حسب المنتج» (يفتح افتراضياً)."
            ),
            wraplength=1100,
        ).pack(anchor="w", padx=6, pady=(4, 2))
        wrap = ttk.Frame(tab_d)
        wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        cols = (
            "device", "key", "customer", "type", "school", "color", "size",
            "qty", "price", "total", "paid", "status", "updated",
        )
        self._tree = ttk.Treeview(wrap, columns=cols, show="headings", height=18)
        heads = [
            ("device", "الجهاز", 110),
            ("key", "المعرّف", 180),
            ("customer", "العميل", 120),
            ("type", "النوع", 110),
            ("school", "المدرسة", 110),
            ("color", "اللون", 70),
            ("size", "المقاس", 55),
            ("qty", "الكمية", 50),
            ("price", "السعر", 65),
            ("total", "الإجمالي", 75),
            ("paid", "المدفوع", 70),
            ("status", "الحالة", 85),
            ("updated", "آخر تحديث", 140),
        ]
        for col, txt, w in heads:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=ysb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree)

        tab_a = ttk.Frame(nb)
        nb.add(tab_a, text="مجمّع حسب المنتج")
        hint = (
            "تُجمّع كل الأسطر ذات نفس (النوع، المدرسة، اللون، المقاس). "
            "عند اختيار نقطة بيع يُحسب المجموع لهذا الفرع فقط (بما يتوافق مع الاسم أو المعرّف المخزّن). "
            "عند ترك «نقطة البيع» فارغاً يُجمع عبر كل الفروع."
        )
        ttk.Label(tab_a, text=hint, wraplength=1100).pack(anchor="w", padx=6, pady=(4, 2))
        agg_wrap = ttk.Frame(tab_a)
        agg_wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        acols = (
            "type", "school", "color", "size", "qty", "lines", "branches",
            "avg_price", "sum_total", "sum_paid", "updated",
        )
        self._tree_agg = ttk.Treeview(agg_wrap, columns=acols, show="headings", height=16)
        ahead = [
            ("type", "النوع", 120),
            ("school", "المدرسة", 120),
            ("color", "اللون", 80),
            ("size", "المقاس", 60),
            ("qty", "إجمالي الكمية", 90),
            ("lines", "عدد الحجوزات", 95),
            ("branches", "عدد الفروع", 85),
            ("avg_price", "سعر وسيط", 80),
            ("sum_total", "إجمالي المبالغ", 95),
            ("sum_paid", "إجمالي المدفوع", 95),
            ("updated", "آخر تحديث", 130),
        ]
        for col, txt, w in ahead:
            self._tree_agg.heading(col, text=txt)
            self._tree_agg.column(col, width=w, anchor="center")
        asb = ttk.Scrollbar(agg_wrap, orient="vertical", command=self._tree_agg.yview)
        self._tree_agg.configure(yscrollcommand=asb.set)
        self._tree_agg.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        asb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree_agg)

        nb.select(tab_a)
        self._refresh()

    def _refresh(self) -> None:
        self._refresh_detail()
        self._refresh_agg()

    def _refresh_detail(self) -> None:
        self._tree.delete(*self._tree.get_children())
        dev = (self._dev.get() or "").strip() or None
        rows = self.db.list_pos_reservations_mirror(
            source_device=dev,
            active_only=bool(self._active_only.get()),
        )
        for r in rows:
            dev_show = self.db.display_name_for_sync_source(r.get("source_device"))
            self._tree.insert(
                "", tk.END,
                values=(
                    dev_show,
                    r.get("reservation_key") or "",
                    r.get("customer") or "",
                    r.get("item_type") or "",
                    r.get("school") or "",
                    r.get("color") or "",
                    r.get("size") or "",
                    r.get("qty") or 0,
                    f"{float(r.get('unit_price') or 0):.2f}",
                    f"{float(r.get('total_amount') or 0):.2f}",
                    f"{float(r.get('paid_amount') or 0):.2f}",
                    r.get("status") or "",
                    r.get("updated_at") or "",
                ),
            )
        apply_zebra_tags(self._tree)

    def _refresh_agg(self) -> None:
        self._tree_agg.delete(*self._tree_agg.get_children())
        dev = (self._dev.get() or "").strip() or None
        rows = self.db.list_pos_reservations_mirror_aggregated(
            source_device=dev,
            active_only=bool(self._active_only.get()),
        )
        for r in rows:
            self._tree_agg.insert(
                "", tk.END,
                values=(
                    r.get("item_type") or "",
                    r.get("school") or "",
                    r.get("color") or "",
                    r.get("size") or "",
                    int(r.get("agg_qty") or 0),
                    int(r.get("reservation_line_count") or 0),
                    int(r.get("pos_device_count") or 0),
                    f"{float(r.get('avg_unit_price') or 0):.2f}",
                    f"{float(r.get('sum_total_amount') or 0):.2f}",
                    f"{float(r.get('sum_paid_amount') or 0):.2f}",
                    r.get("last_updated") or "",
                ),
            )
        apply_zebra_tags(self._tree_agg)


class PosBranchFinancialWindow(tk.Toplevel):
    """Day-by-day POS cashflow totals from synced events."""

    _CAT_AR = {
        "sale": "مبيعات",
        "return_bill": "مرتجعات",
        "void_bill": "إلغاء فواتير",
        "exchange_net": "استبدال (صافي)",
        "reservation_downpayment": "عربون حجز",
        "reservation_payment": "دفعات حجز",
        "reservation_collect": "تحصيل عند التسليم",
    }

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master)
        self.db = db
        self.title("التدفقات المالية للفروع (يومي)")
        self.geometry("1120x580")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="من تاريخ:").pack(side=tk.RIGHT, padx=4)
        self._df = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._df, width=12).pack(side=tk.RIGHT)
        ttk.Label(top, text="إلى:").pack(side=tk.RIGHT, padx=4)
        self._dt = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._dt, width=12).pack(side=tk.RIGHT)
        ttk.Label(top, text="الجهاز:").pack(side=tk.RIGHT, padx=4)
        names = self.db.list_pos_reservations_mirror_device_picklist()
        self._dev = tk.StringVar(value="")
        ttk.Combobox(top, textvariable=self._dev, values=names, width=30, state="readonly").pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="تحديث", command=self._refresh_summary).pack(side=tk.LEFT)

        mid = ttk.Panedwindow(self, orient=tk.VERTICAL)
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        sum_frame = ttk.Frame(mid)
        mid.add(sum_frame, weight=2)
        cols = (
            "day", "sales", "returns", "voids", "exch", "rdep", "rpay", "rcoll", "net",
        )
        self._sum_tree = ttk.Treeview(sum_frame, columns=cols, show="headings", height=10)
        for col, txt, w in [
            ("day", "اليوم", 100),
            ("sales", "مبيعات", 90),
            ("returns", "مرتجعات", 90),
            ("voids", "إلغاء", 80),
            ("exch", "استبدال", 90),
            ("rdep", "عربون", 80),
            ("rpay", "دفعات حجز", 90),
            ("rcoll", "تحصيل تسليم", 95),
            ("net", "الصافي", 100),
        ]:
            self._sum_tree.heading(col, text=txt)
            self._sum_tree.column(col, width=w, anchor="center")
        ssb = ttk.Scrollbar(sum_frame, orient="vertical", command=self._sum_tree.yview)
        self._sum_tree.configure(yscrollcommand=ssb.set)
        self._sum_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ssb.pack(side=tk.RIGHT, fill=tk.Y)
        self._sum_tree.bind("<<TreeviewSelect>>", lambda _e: self._load_detail())
        _bind_mousewheel(self._sum_tree)

        det_fr = ttk.LabelFrame(mid, text="تفاصيل اليوم المحدد")
        mid.add(det_fr, weight=3)
        dcols = ("cat", "amount", "etype", "eid", "meta")
        self._det_tree = ttk.Treeview(det_fr, columns=dcols, show="headings", height=14)
        for col, txt, w in [
            ("cat", "البند", 140),
            ("amount", "المبلغ", 90),
            ("etype", "نوع الحدث", 160),
            ("eid", "مرجع", 220),
            ("meta", "ملاحظة", 320),
        ]:
            self._det_tree.heading(col, text=txt)
            self._det_tree.column(col, width=w, anchor="center")
        dsb = ttk.Scrollbar(det_fr, orient="vertical", command=self._det_tree.yview)
        self._det_tree.configure(yscrollcommand=dsb.set)
        self._det_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        dsb.pack(side=tk.RIGHT, fill=tk.Y, pady=4)
        _bind_mousewheel(self._det_tree)

        self._refresh_summary()

    def _selected_day(self) -> Optional[str]:
        sel = self._sum_tree.selection()
        if not sel:
            return None
        vals = self._sum_tree.item(sel[0], "values")
        return str(vals[0]) if vals else None

    def _refresh_summary(self):
        self._sum_tree.delete(*self._sum_tree.get_children())
        self._det_tree.delete(*self._det_tree.get_children())
        dev = (self._dev.get() or "").strip() or None
        df = (self._df.get() or "").strip() or None
        dt = (self._dt.get() or "").strip() or None
        rows = self.db.list_pos_financial_summary_by_day(dev, df, dt)
        for r in rows:
            self._sum_tree.insert(
                "", tk.END, iid=str(r.get("day") or ""),
                values=(
                    r.get("day") or "",
                    f"{float(r.get('sales_amt') or 0):.2f}",
                    f"{float(r.get('returns_amt') or 0):.2f}",
                    f"{float(r.get('voids_amt') or 0):.2f}",
                    f"{float(r.get('exchange_amt') or 0):.2f}",
                    f"{float(r.get('res_dep_amt') or 0):.2f}",
                    f"{float(r.get('res_pay_amt') or 0):.2f}",
                    f"{float(r.get('res_coll_amt') or 0):.2f}",
                    f"{float(r.get('net_amt') or 0):.2f}",
                ),
            )
        apply_zebra_tags(self._sum_tree)

    def _load_detail(self):
        day = self._selected_day()
        self._det_tree.delete(*self._det_tree.get_children())
        if not day:
            return
        dev = (self._dev.get() or "").strip() or None
        for r in self.db.list_pos_financial_ledger_detail(day, dev):
            cat = self._CAT_AR.get(str(r.get("category") or ""), r.get("category") or "")
            meta = r.get("meta_json") or ""
            self._det_tree.insert(
                "", tk.END,
                values=(
                    cat,
                    f"{float(r.get('amount') or 0):.2f}",
                    r.get("event_type") or "",
                    r.get("event_uuid") or "",
                    meta[:400],
                ),
            )
        apply_zebra_tags(self._det_tree)


class BranchCycleSummaryWindow(tk.Toplevel):
    """Per-branch cycle reconciliation: shipments vs cash vs current stock value."""

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master)
        self.db = db
        self.title("ملخص دورة الفروع")
        self.geometry("1180x600")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="من تاريخ:").pack(side=tk.RIGHT, padx=4)
        self._df = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._df, width=12).pack(side=tk.RIGHT)
        ttk.Label(top, text="إلى:").pack(side=tk.RIGHT, padx=4)
        self._dt = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._dt, width=12).pack(side=tk.RIGHT)
        ttk.Button(top, text="تحديث", command=self._refresh).pack(side=tk.LEFT)

        cols = (
            "branch", "bills", "ship_qty", "ship_value", "cash_net",
            "stock_qty", "stock_value", "gap",
        )
        self._tree = ttk.Treeview(self, columns=cols, show="headings", height=19)
        for col, txt, w in [
            ("branch", "الفرع", 160),
            ("bills", "عدد فواتير الشحن", 110),
            ("ship_qty", "كمية الشحنات", 95),
            ("ship_value", "قيمة الشحنات", 120),
            ("cash_net", "صافي المتحصل", 110),
            ("stock_qty", "كمية المخزون الحالي", 120),
            ("stock_value", "قيمة المخزون الحالي", 130),
            ("gap", "فجوة الدورة", 120),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=ysb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree)

        self._sum = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._sum).pack(fill=tk.X, padx=10, pady=(0, 8))

        self._refresh()

    def _refresh(self):
        self._tree.delete(*self._tree.get_children())
        df = (self._df.get() or "").strip() or None
        dt = (self._dt.get() or "").strip() or None
        rows = self.db.list_branch_cycle_reconciliation(df, dt)
        t_ship = 0.0
        t_cash = 0.0
        t_stock = 0.0
        t_gap = 0.0
        for r in rows:
            t_ship += float(r.get("shipment_value") or 0.0)
            t_cash += float(r.get("cash_net") or 0.0)
            t_stock += float(r.get("stock_value") or 0.0)
            t_gap += float(r.get("cycle_gap") or 0.0)
            self._tree.insert(
                "",
                tk.END,
                values=(
                    r.get("branch_name") or "",
                    int(r.get("shipment_bills_count") or 0),
                    int(r.get("shipment_qty") or 0),
                    f"{float(r.get('shipment_value') or 0.0):.2f}",
                    f"{float(r.get('cash_net') or 0.0):.2f}",
                    int(r.get("stock_qty") or 0),
                    f"{float(r.get('stock_value') or 0.0):.2f}",
                    f"{float(r.get('cycle_gap') or 0.0):.2f}",
                ),
            )
        apply_zebra_tags(self._tree)
        self._sum.set(
            f"إجمالي قيمة الشحنات: {t_ship:.2f}  |  إجمالي صافي المتحصل: {t_cash:.2f}  |  "
            f"إجمالي قيمة المخزون الحالي: {t_stock:.2f}  |  الفجوة الكلية: {t_gap:.2f}"
        )


# ------------------- App -------------------

class WarehouseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("إدارة المخازن")
        self.geometry("1320x760")
        self.db = SqliteDatabase(DB_PATH)
        self._dark_mode = False

        # Apply palette BEFORE creating widgets so ttk picks it up
        self._apply_colorful_theme(True)

        self._build()
        self._bind_shortcuts()
        self._tick_clock()
        try:
            import sync_periodic

            sync_periodic.attach_periodic_sync(self, self.db.path)
        except Exception:
            pass

    def _bind_shortcuts(self):
        """Global keyboard shortcuts."""
        self.bind("<F1>", lambda e: self._switch_tab("income"))
        self.bind("<F2>", lambda e: self._switch_tab("outcome"))
        self.bind("<F3>", lambda e: self._open_inventory())
        self.bind("<F4>", lambda e: self._open_bills_history())
        self.bind("<F5>", lambda e: self._refresh_current_tab())
        self.bind("<F9>", lambda e: self._switch_tab("dashboard"))
        self.bind("<F11>", lambda e: self._switch_tab("sync_diagnostics"))
        self.bind("<F10>", lambda e: self._switch_tab("statistics"))
        self.bind("<Control-i>", lambda e: self._open_inventory())
        self.bind("<Control-p>", lambda e: self._open_bills_history())

    def _tick_clock(self):
        """Update the status bar and header clocks every second."""
        try:
            now = datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
            self._clock_var.set(now)
            self._header_clock_var.set(now)
        except Exception:
            pass
        self.after(1000, self._tick_clock)

    def _apply_dark_mode(self, enable: bool):
        """Toggle dark mode colors."""
        s = ttk.Style(self)
        if enable:
            BG      = "#0F172A"
            SURFACE = "#1E293B"
            SURFACE2= "#334155"
            TEXT    = "#F1F5F9"
            TEXT_SEC= "#CBD5E1"
            TEXT_DIM= "#64748B"
            BRAND   = "#60A5FA"
            BRAND_H = "#93C5FD"
            ROW     = "#1E293B"
            SELBG   = "#334155"
            EDGE    = "#475569"

            try: self.configure(bg=BG)
            except Exception: pass

            s.configure("TFrame", background=SURFACE)
            s.configure("TLabelframe", background=SURFACE, bordercolor=EDGE)
            s.configure("TLabelframe.Label", background=SURFACE, foreground=TEXT_SEC)
            s.configure("TLabel", background=SURFACE, foreground=TEXT)
            s.configure("TButton", background=BRAND, foreground=BG,
                        bordercolor=BRAND, padding=[12, 6])
            s.map("TButton",
                background=[("active", BRAND_H), ("pressed", BRAND_H)])
            for sty in ("TEntry", "TCombobox", "TSpinbox"):
                s.configure(sty, fieldbackground=SURFACE2, background=SURFACE2,
                            foreground=TEXT, bordercolor=EDGE,
                            lightcolor=EDGE, darkcolor=EDGE)
                s.map(sty, fieldbackground=[("focus", SURFACE2)])
            s.configure("TCheckbutton", background=SURFACE, foreground=TEXT)
            s.map("TCheckbutton", background=[("active", SURFACE)])
            s.configure("TRadiobutton", background=SURFACE, foreground=TEXT)
            s.map("TRadiobutton", background=[("active", SURFACE)])
            s.configure("TNotebook", background=BG, bordercolor=EDGE)
            s.configure("TNotebook.Tab", background=SURFACE, foreground=TEXT_DIM,
                        padding=[14, 6])
            s.map("TNotebook.Tab",
                background=[("selected", BG)], foreground=[("selected", TEXT)])
            s.configure("Treeview", background=ROW, fieldbackground=ROW,
                        foreground=TEXT, bordercolor=EDGE, rowheight=26)
            s.configure("Treeview.Heading", background=SURFACE2,
                        foreground=TEXT_SEC, bordercolor=EDGE)
            s.map("Treeview",
                background=[("selected", SELBG)], foreground=[("selected", TEXT)])
        else:
            self._apply_colorful_theme(self._colorful_var.get())

    def _switch_tab(self, name):
        """Switch to a named tab."""
        tab_map = {
            "dashboard": 0,
            "outcome": 1,
            "income": 2,
            "sync_diagnostics": 3,
            "statistics": 4,
        }
        idx = tab_map.get(name)
        if idx is not None and hasattr(self, "notebook"):
            try:
                self.notebook.select(idx)
            except Exception:
                pass

    def _refresh_current_tab(self):
        """Refresh the currently active tab."""
        try:
            idx = self.notebook.index(self.notebook.select())
            if idx == 0 and hasattr(self, "_dashboard"):
                self._dashboard._refresh()
            elif idx == 3 and hasattr(self, "_sync_diagnostics"):
                self._sync_diagnostics._refresh()
            elif idx == 4 and hasattr(self, "_statistics"):
                self._statistics._refresh()
        except Exception:
            pass

    def _open_income(self):
        """Switch to the Income tab instead of opening a separate window."""
        self._switch_tab("income")

    def _apply_colorful_theme(self, enable: bool = True):
        s = ttk.Style(self)

        base = "clam"
        try:
            s.theme_use(base)
        except Exception:
            pass

        if not enable:
            try:
                if "vista" in s.theme_names():
                    s.theme_use("vista")
            except Exception:
                pass
            return

        # -------- Slate Frost palette --------
        BG      = _UI["BG"]
        SURFACE = _UI["SURFACE"]
        SURFACE2= _UI["SURFACE2"]
        BORDER  = _UI["BORDER"]
        ACCENT  = _UI["ACCENT"]
        ACCENT_H= _UI["ACCENT_H"]
        BRAND   = _UI["BRAND"]
        TEXT    = _UI["TEXT"]
        TEXT_SEC= _UI["TEXT_SEC"]
        TEXT_DIM= _UI["TEXT_DIM"]
        SEL_BG  = _UI["SEL_BG"]
        SEL_FG  = _UI["SEL_FG"]

        try: self.configure(bg=BG)
        except Exception: pass

        # -------- base widgets --------
        s.configure("TFrame",  background=SURFACE)
        s.configure("TLabelframe", background=SURFACE, bordercolor=BORDER, relief="solid", borderwidth=1)
        s.configure("TLabelframe.Label", background=SURFACE, foreground=TEXT_SEC,
                    font=("Segoe UI", 9, "bold"))

        s.configure("TLabel",  background=SURFACE, foreground=TEXT)
        s.configure("TButton", background=ACCENT, foreground="white",
                    bordercolor=ACCENT, focusthickness=0, padding=[8, 4],
                    font=("Segoe UI", 9, "bold"))
        s.map("TButton",
            background=[("active", ACCENT_H), ("pressed", ACCENT_H)],
            bordercolor=[("focus", BRAND)])

        # entries / combos — white field, blue focus border
        for sty in ("TEntry", "TCombobox", "TSpinbox"):
            s.configure(sty, fieldbackground=SURFACE, background=SURFACE,
                        foreground=TEXT, bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER,
                        arrowcolor=TEXT_SEC, padding=[8, 6])
            s.map(sty,
                fieldbackground=[("focus", SURFACE), ("readonly", SURFACE2)],
                bordercolor=[("focus", BRAND)],
                lightcolor=[("focus", BRAND)],
                darkcolor=[("focus", BRAND)])

        # checkbutton / radiobutton
        s.configure("TCheckbutton", background=SURFACE, foreground=TEXT)
        s.map("TCheckbutton", background=[("active", SURFACE)])
        s.configure("TRadiobutton", background=SURFACE, foreground=TEXT)
        s.map("TRadiobutton", background=[("active", SURFACE)])

        # notebook tabs
        s.configure("TNotebook", background=BG, bordercolor=BORDER)
        s.configure("TNotebook.Tab", background=SURFACE2, foreground=TEXT_DIM,
                    padding=[10, 4], font=("Segoe UI", 9))
        s.map("TNotebook.Tab",
            background=[("selected", SURFACE)],
            foreground=[("selected", TEXT)])

        # -------- Treeview (tables) --------
        s.configure("Treeview",
                    background=SURFACE, fieldbackground=SURFACE, foreground=TEXT,
                    bordercolor=BORDER, rowheight=26,
                    font=("Segoe UI", 9))
        s.configure("Treeview.Heading",
                    background=SURFACE2, foreground=TEXT_SEC, bordercolor=BORDER,
                    font=("Segoe UI", 8, "bold"), padding=[6, 4])
        s.map("Treeview",
            background=[("selected", SEL_BG)],
            foreground=[("selected", SEL_FG)])

        # -------- PanedWindow --------
        s.configure("TPanedwindow", background=BG)
        s.configure("Sash", sashthickness=4, gripcount=0)

    def _build(self):
        # ======== MENU BAR ========
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        appearance = tk.Menu(menubar, tearoff=False)
        self._colorful_var = tk.BooleanVar(value=True)
        self._dark_var = tk.BooleanVar(value=False)

        def _toggle_theme():
            self._apply_colorful_theme(self._colorful_var.get())
        def _toggle_dark():
            self._dark_mode = self._dark_var.get()
            self._apply_dark_mode(self._dark_mode)

        appearance.add_checkbutton(label="مظهر ملوّن", variable=self._colorful_var, command=_toggle_theme)
        appearance.add_checkbutton(label="الوضع الليلي", variable=self._dark_var, command=_toggle_dark)
        menubar.add_cascade(label="المظهر", menu=appearance)

        inv = tk.Menu(menubar, tearoff=False)
        inv.add_command(label="عرض المخزون (جدول)      Ctrl+I", command=self._open_inventory)
        inv.add_command(label="سجل الحركات", command=self._open_movements)
        inv.add_command(label="تحويل بين المخازن", command=self._open_transfer)
        inv.add_command(label="جرد المخزون", command=self._open_audit)
        inv.add_separator()
        inv.add_command(label="مخزون الفروع", command=self._open_branch_stock)
        inv.add_command(label="منتجات الفروع غير المعالجة", command=self._open_branch_inventory_queue)
        menubar.add_cascade(label="المخزون", menu=inv)

        bills = tk.Menu(menubar, tearoff=False)
        bills.add_command(label="السجل      Ctrl+P", command=self._open_bills_history)
        bills.add_command(label="شحنات الفروع والوارد المتزامن…", command=self._open_branch_bills_sync_log)
        menubar.add_cascade(label="الفواتير", menu=bills)

        admin_menu = tk.Menu(menubar, tearoff=False)
        admin_menu.add_command(label="إعدادات المدير…", command=self._open_admin)
        menubar.add_cascade(label="المدير", menu=admin_menu)

        shortcuts_menu = tk.Menu(menubar, tearoff=False)
        shortcuts_menu.add_command(label="F1  - الوارد", state="disabled")
        shortcuts_menu.add_command(label="F2  - المنصرف", state="disabled")
        shortcuts_menu.add_command(label="F3  - المخزون", state="disabled")
        shortcuts_menu.add_command(label="F4  - الفواتير", state="disabled")
        shortcuts_menu.add_command(label="F5  - تحديث", state="disabled")
        shortcuts_menu.add_command(label="F9  - لوحة التحكم", state="disabled")
        shortcuts_menu.add_command(label="F11 - تشخيص المزامنة", state="disabled")
        shortcuts_menu.add_command(label="F10 - الإحصائيات", state="disabled")
        menubar.add_cascade(label="اختصارات", menu=shortcuts_menu)

        sync_menu = tk.Menu(menubar, tearoff=False)
        sync_menu.add_command(label="مزامنة الآن…", command=self._open_sync_dialog)
        sync_menu.add_command(label="إعدادات المزامنة…", command=self._open_sync_setup)
        sync_menu.add_separator()
        sync_menu.add_command(label="سجل تعديلات أسعار الفروع…", command=self._open_price_sync_audit)
        sync_menu.add_separator()
        sync_menu.add_command(label="سجل شحنات الفروع والوارد المتزامن…", command=self._open_branch_bills_sync_log)
        sync_menu.add_command(label="حجوزات الفروع (مرآة)…", command=self._open_pos_reservations_mirror)
        sync_menu.add_command(label="التدفقات المالية للفروع (يومي)…", command=self._open_pos_financial_by_day)
        sync_menu.add_command(label="ملخص دورة الفروع…", command=self._open_branch_cycle_summary)
        sync_menu.add_separator()
        sync_menu.add_command(label="تشخيص المزامنة (F11)", command=lambda: self._switch_tab("sync_diagnostics"))
        menubar.add_cascade(label="المزامنة", menu=sync_menu)

        # ======== HEADER BAR (Premium dark slate) ========
        _HBG = _UI["ACCENT"]       # #0F172A
        _HBG2 = _UI["ACCENT_H"]    # #1E293B

        header = tk.Frame(self, bg=_HBG, height=56)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        # Brand area (right for RTL)
        brand = tk.Frame(header, bg=_HBG)
        brand.pack(side=tk.RIGHT, padx=16)
        tk.Label(brand, text="إدارة المخازن", bg=_HBG, fg="#FFFFFF",
                 font=("Segoe UI", 12, "bold")).pack(side=tk.RIGHT)
        tk.Label(brand, text=" PRO ", bg=_UI["BRAND"], fg="#FFFFFF",
                 font=("Segoe UI", 7, "bold"), padx=4, pady=1).pack(side=tk.RIGHT, padx=(0, 8))

        # Separator line
        tk.Frame(header, bg=_HBG2, width=1).pack(side=tk.RIGHT, fill=tk.Y, padx=8, pady=10)

        # Navigation pills (replace notebook tab visuals)
        nav = tk.Frame(header, bg=_HBG)
        nav.pack(side=tk.RIGHT, padx=8)
        self._nav_btns = []
        nav_items = [
            ("لوحة التحكم", 0),
            ("المنصرف", 1),
            ("الوارد", 2),
            ("تشخيص المزامنة", 3),
            ("الإحصائيات", 4),
        ]
        for text, idx in nav_items:
            b = tk.Button(nav, text=f"  {text}  ", bg=_HBG2, fg=_UI["TEXT_DIM"],
                          font=("Segoe UI", 10), bd=0, padx=12, pady=6,
                          cursor="hand2", activebackground="#334155",
                          activeforeground="#FFFFFF",
                          command=lambda i=idx: self._select_nav(i))
            b.pack(side=tk.RIGHT, padx=3)
            _add_hover(b, "#334155", _HBG2, "#F1F5F9", _UI["TEXT_DIM"])
            self._nav_btns.append(b)

        # Utility buttons (left side for RTL)
        util = tk.Frame(header, bg=_HBG)
        util.pack(side=tk.LEFT, padx=12)

        self._header_clock_var = tk.StringVar(value="")
        tk.Label(util, textvariable=self._header_clock_var, bg=_HBG, fg="#64748B",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=8)

        _ghost_hdr = {"bg": _HBG, "fg": "#CBD5E1", "bd": 0, "font": ("Segoe UI", 9),
                      "padx": 8, "pady": 3, "cursor": "hand2",
                      "activebackground": _HBG2, "activeforeground": "#FFFFFF"}
        for txt, cmd, tip in [
            ("المخزون", self._open_inventory, "F3 / Ctrl+I"),
            ("الفواتير", self._open_bills_history, "F4 / Ctrl+P"),
            ("الحركات", self._open_movements, "سجل الحركات"),
            ("المدير", self._open_admin, "إعدادات المدير"),
            ("تحويل", self._open_transfer, "تحويل بين المخازن"),
            ("جرد", self._open_audit, "جرد المخزون"),
        ]:
            b = tk.Button(util, text=txt, command=cmd, **_ghost_hdr)
            b.pack(side=tk.LEFT, padx=2)
            _add_hover(b, _HBG2, _HBG, "#FFFFFF", "#CBD5E1")
            ToolTip(b, tip)

        # ======== TABBED NOTEBOOK ========
        self.notebook = ttk.Notebook(self)
        try:
            s = ttk.Style(self)
            s.layout("App.Navless.TNotebook.Tab", [])
            self.notebook.configure(style="App.Navless.TNotebook")
        except Exception:
            pass
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        self._dashboard = DashboardFrame(self.notebook, self.db, app=self)
        self.notebook.add(self._dashboard, text="  لوحة التحكم  (F9)  ")

        self._outcome = OutcomeFrame(self.notebook, self.db)
        self.notebook.add(self._outcome, text="  المنصرف  (F2)  ")

        self._income = IncomeFrame(self.notebook, self.db)
        self.notebook.add(self._income, text="  الوارد  (F1)  ")

        self._sync_diagnostics = SyncDiagnosticsFrame(self.notebook, self.db, app=self)
        self.notebook.add(self._sync_diagnostics, text="  تشخيص المزامنة  (F11)  ")

        self._statistics = StatisticsFrame(self.notebook, self.db)
        self.notebook.add(self._statistics, text="  الإحصائيات  (F10)  ")

        self.notebook.select(1)
        self._highlight_nav(1)

        def _on_tab_change(event):
            try:
                idx = self.notebook.index(self.notebook.select())
                self._highlight_nav(idx)
                if idx == 0:
                    self._dashboard._refresh()
                elif idx == 3:
                    self._sync_diagnostics._refresh()
                elif idx == 4:
                    self._statistics._refresh()
            except Exception:
                pass
        self.notebook.bind("<<NotebookTabChanged>>", _on_tab_change)

        # ======== STATUS BAR (minimal) ========
        status = tk.Frame(self, bg=_UI["SURFACE"], height=28)
        status.pack(fill=tk.X, side=tk.BOTTOM)
        status.pack_propagate(False)
        tk.Frame(status, bg=_UI["BORDER"], height=1).pack(fill=tk.X, side=tk.TOP)

        self._clock_var = tk.StringVar(value="")
        tk.Label(status, textvariable=self._clock_var, bg=_UI["SURFACE"],
                 fg=_UI["TEXT_DIM"], font=_FONTS["caption"]).pack(side=tk.LEFT, padx=8)

        tk.Label(status, text=f"DB: {DB_PATH}", bg=_UI["SURFACE"],
                 fg=_UI["TEXT_DIM"], font=_FONTS["caption"]).pack(side=tk.LEFT, padx=16)

        self._status_msg = tk.StringVar(value="جاهز")
        tk.Label(status, textvariable=self._status_msg, bg=_UI["SURFACE"],
                 fg=_UI["TEXT_DIM"], font=_FONTS["caption"]).pack(side=tk.RIGHT, padx=8)



    def _select_nav(self, idx):
        """Select a tab via the header navigation pills."""
        try:
            self.notebook.select(idx)
            self._highlight_nav(idx)
        except Exception:
            pass

    def _highlight_nav(self, active_idx):
        """Highlight the active nav pill button."""
        for i, btn in enumerate(self._nav_btns):
            if i == active_idx:
                btn.config(bg=_UI["BRAND"], fg="#FFFFFF")
                # Override hover for active button
                btn.unbind("<Enter>")
                btn.unbind("<Leave>")
            else:
                btn.config(bg=_UI["ACCENT_H"], fg=_UI["TEXT_DIM"])
                _add_hover(btn, "#334155", _UI["ACCENT_H"], "#F1F5F9", _UI["TEXT_DIM"])

    def _open_admin(self):
        AdminWindow(self, self.db)

    def _open_sync_dialog(self):
        try:
            import sync_ui
            sync_ui.open_sync_dialog(self, self.db.conn)
        except Exception as e:
            messagebox.showerror("المزامنة", f"تعذّر فتح نافذة المزامنة:\n{e}", parent=self)

    def _open_branch_stock(self):
        try:
            BranchStockWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("مخزون الفروع", f"تعذّر فتح نافذة مخزون الفروع:\n{e}", parent=self)

    def _open_branch_inventory_queue(self):
        try:
            BranchInventoryQueueWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("منتجات الفروع غير المعالجة", f"تعذّر فتح نافذة المنتجات غير المعالجة:\n{e}", parent=self)

    def _open_sync_setup(self):
        try:
            import sync_ui
            sync_ui.open_sync_setup(self, self.db.conn)
        except Exception as e:
            messagebox.showerror("المزامنة", f"تعذّر فتح إعدادات المزامنة:\n{e}", parent=self)

    def _open_price_sync_audit(self):
        try:
            open_price_sync_audit_dialog(self, self.db)
        except Exception as e:
            messagebox.showerror("سجل أسعار الفروع", f"تعذّر فتح السجل:\n{e}", parent=self)

    def _open_branch_bills_sync_log(self):
        try:
            BranchBillsSyncLogWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("سجل الفروع", f"تعذّر فتح النافذة:\n{e}", parent=self)

    def _open_pos_reservations_mirror(self):
        try:
            PosReservationsMirrorWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("حجوزات الفروع", f"تعذّر فتح النافذة:\n{e}", parent=self)

    def _open_pos_financial_by_day(self):
        try:
            PosBranchFinancialWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("التدفقات المالية", f"تعذّر فتح النافذة:\n{e}", parent=self)

    def _open_branch_cycle_summary(self):
        try:
            BranchCycleSummaryWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("ملخص دورة الفروع", f"تعذّر فتح النافذة:\n{e}", parent=self)

    def _open_inventory(self):
        InventoryWindow(self, self.db)

    def _open_bills_history(self):
        BillsHistoryWindow(self, self.db)

    def _open_movements(self):
        MovementsWindow(self, self.db)

    def _open_transfer(self):
        TransferWindow(self, self.db)

    def _open_audit(self):
        StockAuditWindow(self, self.db)


def main():
    app = WarehouseApp()
    app.mainloop()


if __name__ == "__main__":
    main()
