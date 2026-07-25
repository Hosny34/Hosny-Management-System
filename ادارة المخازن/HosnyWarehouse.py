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
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from dataclasses import dataclass
from datetime import datetime, date   # +date for calendar
import calendar                       # ADD
from typing import Any, Dict, List, Optional, Sequence, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

_DIGIT_TRANSLATION = str.maketrans({
    "\u0660": "0", "\u0661": "1", "\u0662": "2", "\u0663": "3", "\u0664": "4",
    "\u0665": "5", "\u0666": "6", "\u0667": "7", "\u0668": "8", "\u0669": "9",
    "\u06f0": "0", "\u06f1": "1", "\u06f2": "2", "\u06f3": "3", "\u06f4": "4",
    "\u06f5": "5", "\u06f6": "6", "\u06f7": "7", "\u06f8": "8", "\u06f9": "9",
})


def western_digits(value: Any) -> str:
    return ("" if value is None else str(value)).translate(_DIGIT_TRANSLATION)


_LTR_MARK = "\u200e"
_RTL_MARK = "\u200f"
_NUMBER_RUN_RE = re.compile(r"(?<!\u200e)([0-9](?:[0-9.,:/\\\- ]*[0-9])?)(?!\u200e)")


def _strip_digit_marks(value: Any) -> str:
    return western_digits(value).replace(_LTR_MARK, "").replace(_RTL_MARK, "")


def western_digits_for_display(value: Any) -> str:
    text = _strip_digit_marks(value)
    return _NUMBER_RUN_RE.sub(lambda m: _LTR_MARK + m.group(1) + _LTR_MARK, text)


def _westernize_value(value: Any) -> Any:
    if value is None:
        return value
    if isinstance(value, (list, tuple)):
        return type(value)(_westernize_value(v) for v in value)
    return western_digits_for_display(value)


def _westernize_options(options: Dict[str, Any]) -> None:
    for key in ("text", "value", "values"):
        if key in options:
            options[key] = _westernize_value(options[key])


def _install_western_digit_tk_patch() -> None:
    """Keep Tk widgets from showing Arabic-Indic/Persian digits."""
    if getattr(tk, "_hosny_western_digits_patch", False):
        return
    tk._hosny_western_digits_patch = True

    original_stringvar_set = tk.StringVar.set
    original_stringvar_get = tk.StringVar.get

    def stringvar_set(self, value):
        return original_stringvar_set(self, _westernize_value(value))

    def stringvar_get(self):
        return _strip_digit_marks(original_stringvar_get(self))

    tk.StringVar.set = stringvar_set
    tk.StringVar.get = stringvar_get

    def patch_entry_class(cls):
        original_get = cls.get

        def get(self):
            return _strip_digit_marks(original_get(self))

        cls.get = get

    for entry_cls in (tk.Entry, ttk.Entry, ttk.Combobox):
        patch_entry_class(entry_cls)

    def patch_widget_class(cls):
        original_init = cls.__init__
        original_configure = cls.configure

        def __init__(self, *args, **kwargs):
            _westernize_options(kwargs)
            original_init(self, *args, **kwargs)

        def configure(self, cnf=None, **kwargs):
            if isinstance(cnf, dict):
                cnf = dict(cnf)
                _westernize_options(cnf)
            _westernize_options(kwargs)
            return original_configure(self, cnf, **kwargs)

        cls.__init__ = __init__
        cls.configure = configure
        cls.config = configure

    for widget_cls in (
        tk.Label, tk.Button, tk.LabelFrame, tk.Checkbutton, tk.Radiobutton,
        ttk.Label, ttk.Button, ttk.LabelFrame, ttk.Checkbutton, ttk.Radiobutton,
        ttk.Combobox,
    ):
        patch_widget_class(widget_cls)

    original_title = tk.Wm.title

    def title(self, string=None):
        if string is not None:
            string = western_digits(string)
        return original_title(self, string)

    tk.Wm.title = title

    original_heading = ttk.Treeview.heading
    original_insert = ttk.Treeview.insert
    original_item = ttk.Treeview.item
    original_set = ttk.Treeview.set

    def heading(self, column, option=None, **kwargs):
        _westernize_options(kwargs)
        return original_heading(self, column, option, **kwargs)

    def insert(self, parent, index, iid=None, **kwargs):
        _westernize_options(kwargs)
        return original_insert(self, parent, index, iid, **kwargs)

    def item(self, item, option=None, **kwargs):
        _westernize_options(kwargs)
        return original_item(self, item, option, **kwargs)

    def set_value(self, item, column=None, value=None):
        if value is not None:
            value = _westernize_value(value)
        return original_set(self, item, column, value)

    ttk.Treeview.heading = heading
    ttk.Treeview.insert = insert
    ttk.Treeview.item = item
    ttk.Treeview.set = set_value


_install_western_digit_tk_patch()

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
    "POS-ZAY": "ÙØ±Ø¹ Ø²Ø§ÙŠØ¯",
    "POS-OCT": "ÙØ±Ø¹ Ø§ÙƒØªÙˆØ¨Ø±",
    "POS-BAH": "ÙØ±Ø¹ Ø¨Ù‡ØªÙŠÙ…",
    "POS-CEN": "ÙØ±Ø¹ Ø§Ù„Ø³Ù†ØªØ±",
    "POS-OBO": "ÙØ±Ø¹ Ø§Ù„Ø¹Ø¨ÙˆØ±",
    "POS-GESR": "ÙØ±Ø¹ Ø¬Ø³Ø± Ø§Ù„Ø³ÙˆÙŠØ³",
}
BRANCH_DEVICE_BY_UI_NAME = {v: k for k, v in BRANCH_UI_NAME_BY_DEVICE.items()}
WAREHOUSE_NUMBER_VALUES = ["1", "2", "3", "4", "5", "6", "7"]
WAREHOUSE_NUMBER_LABELS = {
    "1": "Ù…Ø®Ø²Ù† 1",
    "2": "Ù…Ø®Ø²Ù† 2",
    "3": "Ù…Ø®Ø²Ù† 3",
    "4": "Ù…Ø®Ø²Ù† 4",
    "5": "Ù…Ø®Ø²Ù† Ø²Ø§ÙŠØ¯",
    "6": "Ù…Ø®Ø²Ù† Ø§ÙƒØªÙˆØ¨Ø±",
    "7": "Ù…Ø®Ø²Ù† Ø§Ù„Ø¹Ø¨ÙˆØ±",
}
WAREHOUSE_NUMBER_DISPLAY_VALUES = [WAREHOUSE_NUMBER_LABELS[v] for v in WAREHOUSE_NUMBER_VALUES]


def normalize_branch_customer_name(value: Any) -> Optional[str]:
    raw = str(value or "").strip()
    if not raw:
        return None
    if raw.startswith("ÙØ±Ø¹:"):
        raw = raw.split(":", 1)[1].strip()
    if raw in DEFAULT_BRANCH_POS_NAMES:
        return raw
    return BRANCH_DEVICE_BY_UI_NAME.get(raw)


def canonical_branch_device_name(value: Any, known_devices: Optional[Sequence[str]] = None) -> Optional[str]:
    """Resolve any branch-facing label to canonical POS device name.

    Accepts internal names (POS-*), Arabic UI names, and mixed labels like:
    "ÙØ±Ø¹: ÙØ±Ø¹ Ø²Ø§ÙŠØ¯ (POS-ZAY)".
    """
    raw = str(value or "").strip()
    if not raw:
        return None
    direct = normalize_branch_customer_name(raw)
    if direct:
        return direct
    if raw.startswith("ÙØ±Ø¹:"):
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


def fmt_local_ts(value: Any, empty: str = "â€”") -> str:
    raw = str(value or "").strip()
    if not raw:
        return empty
    try:
        txt = raw[:-1] + "+00:00" if raw.endswith("Z") else raw
        dt = datetime.fromisoformat(txt)
        if dt.tzinfo is None:
            return western_digits(dt.strftime("%Y-%m-%d %H:%M:%S"))
        return western_digits(dt.astimezone().strftime("%Y-%m-%d %H:%M:%S"))
    except Exception:
        return western_digits(raw.replace("T", " "))


def branch_customer_label(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    if raw.startswith("ÙØ±Ø¹:"):
        raw = raw.split(":", 1)[1].strip()
    return f"ÙØ±Ø¹: {branch_display_name(raw)}"


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


def format_money(value: Any) -> str:
    raw = _strip_digit_marks(value).strip()
    if not raw:
        raw = "0"
    try:
        dec = Decimal(raw)
    except (InvalidOperation, ValueError, TypeError):
        try:
            dec = Decimal(str(float(value or 0)))
        except Exception:
            return "0"
    dec = dec.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return western_digits(str(int(dec)))


def parse_float_text(value: Any, default: Optional[float] = None) -> Optional[float]:
    raw = _strip_digit_marks(value).strip()
    if not raw:
        return default
    raw = raw.replace(",", "")
    try:
        return float(raw)
    except (TypeError, ValueError):
        return default


def parse_int_text(value: Any, default: Optional[int] = None) -> Optional[int]:
    raw = _strip_digit_marks(value).strip()
    if not raw:
        return default
    raw = raw.replace(",", "")
    try:
        return int(raw)
    except (TypeError, ValueError):
        return default

ALLOWED_NUMERIC_RANGES = {
    (0, 24): [str(i) for i in range(0, 26, 2)],
    (0, 16): [str(i) for i in range(0, 18, 2)],
    (6, 22): [str(i) for i in range(6, 24, 2)],
    (14, 28): [str(i) for i in range(14, 30, 2)],
    (18, 30): [str(i) for i in range(18, 32, 2)],
    (0, 9):  [str(i) for i in range(0, 10, 1)],
    (32, 62): [str(i) for i in range(32, 63,2)],

}
NUMERIC_RANGE_LABELS = [
    f"{a} â†’ {b}" for (a, b) in ALLOWED_NUMERIC_RANGES.keys()
]
DEFAULT_NUMERIC_SIZE_RANGE = (0, 24)
DEFAULT_NUMERIC_SIZE_RANGE_LABEL = f"{DEFAULT_NUMERIC_SIZE_RANGE[0]} â†’ {DEFAULT_NUMERIC_SIZE_RANGE[1]}"
DEFAULT_SIZE_PROFILE = (
    DEFAULT_NUMERIC_SIZE_RANGE[0],
    DEFAULT_NUMERIC_SIZE_RANGE[1],
    None,
    None,
    0,
)

ALPHA_SIZES = ["S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL"]
PREFERRED_WAREHOUSE_ITEM_ORDER = {
    "ØªÙŠØ´ÙŠØ±Øª ØµÙŠÙÙŠ": 0,
    "ØªÙŠØ´ÙŠØ±Øª Ø®Ø±ÙŠÙÙŠ": 1,
    "ØªÙŠØ´ÙŠØ±Øª Ø´ØªÙˆÙŠ": 2,
    "Ø´Ø±ÙˆØ§Ù„": 3,
    "Ø¬Ø§ÙƒÙŠØª": 4,
    "ØªÙŠØ´ÙŠØ±Øª Ø±ÙŠØ§Ø¶ÙŠ": 5,
    "ØªÙŠØ´ÙŠØ±Øª Ø±ÙŠØ§Ø¶ÙŠ ÙƒÙ…": 6,
    "ØªÙŠØ´ÙŠØ±Øª Ø±ÙŠØ§Ø¶ÙŠ Ø´ØªÙˆÙŠ": 7,
    "Ø´Ø±ÙˆØ§Ù„ Ø±ÙŠØ§Ø¶ÙŠ": 8,
    "Ø¬Ø§ÙƒÙŠØª Ø±ÙŠØ§Ø¶ÙŠ": 9,
}


def _normalize_item_order_label(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "").strip())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.split())


def warehouse_item_priority(item_type: Any) -> int:
    item_clean = _normalize_item_order_label(item_type)
    return PREFERRED_WAREHOUSE_ITEM_ORDER.get(
        item_clean,
        len(PREFERRED_WAREHOUSE_ITEM_ORDER),
    )


def warehouse_item_sort_key(item_type: Any, color: Any = "") -> Tuple[Any, ...]:
    item_clean = _normalize_item_order_label(item_type)
    color_clean = str(color or "").strip()
    return (
        warehouse_item_priority(item_clean),
        item_clean.lower(),
        color_clean.lower(),
    )


def warehouse_size_sort_key(size: Any) -> Tuple[Any, ...]:
    size_clean = _normalize_size_label(str(size or ""))
    if size_clean.isdigit():
        return (0, int(size_clean), "")
    alpha_rank = {label: idx for idx, label in enumerate(ALPHA_SIZES, start=1)}
    return (1, alpha_rank.get(size_clean, 999), size_clean.lower())


def format_weight_kg(grams: Any) -> str:
    try:
        grams_val = float(grams or 0.0)
    except (TypeError, ValueError):
        grams_val = 0.0
    kg_val = grams_val / 1000.0
    text = f"{kg_val:.3f}".rstrip("0").rstrip(".")
    if not text:
        text = "0"
    return f"{text} ÙƒØ¬Ù…"


def merged_numeric_size_labels_from_profile(
    r1s: Optional[int],
    r1e: Optional[int],
    r2s: Optional[int],
    r2e: Optional[int],
) -> List[str]:
    """Union of predefined numeric columns for range 1 and range 2 (no duplicate sizes).

    Overlapping presets such as (6â†’22) and (14â†’28) become one sorted list 6â€¦28.
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
    return western_digits(datetime.now().isoformat(timespec="seconds"))


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
            show_toast(parent, f"ØªÙ… ÙØªØ­ Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ù„Ù„Ø·Ø¨Ø§Ø¹Ø© â€” Ø¹Ø¯Ø¯ Ø§Ù„Ù†Ø³Ø®: {copies}")
    except Exception as ex:
        if parent is not None:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©", f"{ex}", parent=parent)

def save_bill_as_html(path: str, bill: Dict[str, Any], items: List[Dict[str, Any]]) -> None:
    """Generate a single HTML with 2 pages: customer copy + warehouse copy."""

    def _fmtf(x: Any) -> str:
        try:
            return f"{format_money(float(x))}"
        except Exception:
            return "0"

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
    bill_total = f"{format_money(float(bill['total']))}"

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>ÙØ§ØªÙˆØ±Ø© #{bill['id']}</title>
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
  <span class="copy-label">Ù†Ø³Ø®Ø© Ø§Ù„Ø¹Ù…ÙŠÙ„</span>
  <h1>ÙØ§ØªÙˆØ±Ø© Ø±Ù‚Ù…: {bill['id']}</h1>
  <div>Ø§Ù„Ø¹Ù…ÙŠÙ„: {customer_name}</div>
  <div>Ø§Ù„ØªØ§Ø±ÙŠØ®: {bill_date}</div>
  <div>Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ: {bill_total}</div>
  <table>
    <thead><tr>
      <th>Ø§Ù„Ù†ÙˆØ¹</th><th>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©</th><th>Ø§Ù„Ù„ÙˆÙ†</th><th>Ø§Ù„Ù…Ù‚Ø§Ø³</th>
      <th>Ø§Ù„Ø³Ø¹Ø±</th><th>Ø§Ù„ÙƒÙ…ÙŠØ©</th><th>Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ</th>
    </tr></thead>
    <tbody>{rows1}</tbody>
  </table>
</div>

<!-- ===== PAGE 2: Warehouse Copy ===== -->
<div class="page page-break">
  <span class="copy-label" style="background:#059669">Ù†Ø³Ø®Ø© Ø§Ù„Ù…Ø®Ø²Ù†</span>
  <h1>ÙØ§ØªÙˆØ±Ø© Ø±Ù‚Ù…: {bill['id']}</h1>
  <div>Ø§Ù„Ø¹Ù…ÙŠÙ„: {customer_name}</div>
  <div>Ø§Ù„ØªØ§Ø±ÙŠØ®: {bill_date}</div>
  <div>Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ: {bill_total}</div>
  <table>
    <thead><tr>
      <th>Ø§Ù„Ù†ÙˆØ¹</th><th>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©</th><th>Ø§Ù„Ù„ÙˆÙ†</th><th>Ø§Ù„Ù…Ù‚Ø§Ø³</th>
      <th>Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù†</th><th>Ø±Ù‚Ù… Ø§Ù„Ø·Ø±Ø¯</th>
      <th>Ø§Ù„Ø³Ø¹Ø±</th><th>Ø§Ù„ÙƒÙ…ÙŠØ©</th><th>Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ</th>
    </tr></thead>
    <tbody>{rows2}</tbody>
  </table>
</div>

</body>
</html>"""

    with open(path, "w", encoding="utf-8") as f:
        f.write(html)

def _html(s: str) -> str:
    return western_digits(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

# Normalize size text for matching (handles Arabic/Persian digits & case)
_AR_DIGITS = _DIGIT_TRANSLATION
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
        self.repair_stock_integrity()


    def repair_stock_integrity(self) -> Dict[str, int]:
        """Repair unsafe stock shapes that can corrupt billing and reports."""
        fixed_negative = 0
        merged_rows = 0
        with self.conn:
            negative_rows = self.conn.execute(
                "SELECT * FROM stocks WHERE COALESCE(count,0) < 0"
            ).fetchall()
            for s in negative_rows:
                qty = abs(int(s["count"] or 0))
                self.conn.execute("UPDATE stocks SET count=0 WHERE id=?", (int(s["id"]),))
                self.conn.execute(
                    """INSERT INTO movements
                    (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,
                     warehouse_no,package_no,unit_price,has_badge)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        now_iso(), "ADJUST_IN", int(s["id"]), qty,
                        "Automatic repair: negative stock clamped to zero",
                        None, s["item_type"], s["school"], s["color"], s["size"],
                        int(s["warehouse_no"]), int(s["package_no"]),
                        float(s["unit_price"]), int(s["has_badge"] or 0),
                    ),
                )
                fixed_negative += 1

            duplicate_groups = self.conn.execute(
                """
                SELECT item_type,school,color,size,warehouse_no,package_no,unit_price,
                       COALESCE(has_badge,0) AS has_badge
                FROM stocks
                GROUP BY item_type,school,color,size,warehouse_no,package_no,unit_price,COALESCE(has_badge,0)
                HAVING COUNT(*) > 1
                """
            ).fetchall()
            for g in duplicate_groups:
                rows = self.conn.execute(
                    """
                    SELECT id,COALESCE(count,0) AS count
                    FROM stocks
                    WHERE item_type=? AND school=? AND color=? AND size=?
                      AND warehouse_no=? AND package_no=? AND unit_price=? AND COALESCE(has_badge,0)=?
                    ORDER BY id ASC
                    """,
                    (
                        g["item_type"], g["school"], g["color"], g["size"],
                        int(g["warehouse_no"]), int(g["package_no"]),
                        float(g["unit_price"]), int(g["has_badge"] or 0),
                    ),
                ).fetchall()
                if len(rows) <= 1:
                    continue
                keeper_id = int(rows[0]["id"])
                duplicate_ids = [int(r["id"]) for r in rows[1:]]
                total_count = sum(max(0, int(r["count"] or 0)) for r in rows)
                placeholders = ",".join("?" for _ in duplicate_ids)
                if duplicate_ids:
                    self.conn.execute(
                        f"UPDATE movements SET stock_id=? WHERE stock_id IN ({placeholders})",
                        [keeper_id, *duplicate_ids],
                    )
                    self.conn.execute(
                        f"DELETE FROM stocks WHERE id IN ({placeholders})",
                        duplicate_ids,
                    )
                    merged_rows += len(duplicate_ids)
                self.conn.execute(
                    "UPDATE stocks SET count=? WHERE id=?",
                    (int(total_count), keeper_id),
                )
        return {"negative_fixed": fixed_negative, "duplicates_merged": merged_rows}


    def add_or_update_stock_row(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
        warehouse_no: int,
        package_no: int,
        unit_price: float,
        qty: int,
        has_badge: int = 0,
    ) -> int:
        """Add stock to an exact row, creating it only when it does not exist."""
        item_type = str(item_type or "").strip()
        school = str(school or "").strip()
        color = str(color or "").strip()
        size = str(size or "").strip()
        qty = int(qty or 0)
        if qty < 0:
            raise ValueError("Stock quantity cannot be negative")
        row = self.conn.execute(
            """
            SELECT id,COALESCE(count,0) AS count
            FROM stocks
            WHERE item_type=? AND school=? AND color=? AND size=?
              AND warehouse_no=? AND package_no=? AND unit_price=? AND COALESCE(has_badge,0)=?
            ORDER BY id ASC
            LIMIT 1
            """,
            (
                item_type, school, color, size,
                int(warehouse_no), int(package_no), float(unit_price), int(has_badge or 0),
            ),
        ).fetchone()
        if row:
            self.conn.execute(
                "UPDATE stocks SET count = count + ? WHERE id = ?",
                (qty, int(row["id"])),
            )
            return int(row["id"])
        cur = self.conn.execute(
            """
            INSERT INTO stocks(item_type,school,color,size,warehouse_no,package_no,unit_price,count,has_badge)
            VALUES(?,?,?,?,?,?,?,?,?)
            """,
            (
                item_type, school, color, size,
                int(warehouse_no), int(package_no), float(unit_price), qty, int(has_badge or 0),
            ),
        )
        return int(cur.lastrowid)


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
        without error â€” we must never break a business operation.

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
                    # Legacy outbox schema (no target_scope column) â€” fall back
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

    def _record_sync_event_or_raise(
        self,
        event_type: str,
        payload: Dict[str, Any],
        *,
        target_scope: Optional[str] = None,
    ) -> str:
        event_uuid = self._record_sync_event(event_type, payload, target_scope=target_scope)
        if not event_uuid:
            raise RuntimeError(f"Failed to record required sync event: {event_type}")
        return str(event_uuid)

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
            CREATE TABLE IF NOT EXISTS price_profiles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                notes TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS price_profile_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                profile_id INTEGER NOT NULL,
                item_type TEXT NOT NULL,
                size TEXT NOT NULL,
                price REAL NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(profile_id, item_type, size)
            );
            CREATE INDEX IF NOT EXISTS idx_price_profile_lines_profile
            ON price_profile_lines(profile_id, item_type, size);
            CREATE TABLE IF NOT EXISTS price_profile_assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                profile_id INTEGER NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(item_type, school, color)
            );
            CREATE INDEX IF NOT EXISTS idx_price_profile_assignments_specs
            ON price_profile_assignments(item_type, school, color);

            CREATE TABLE IF NOT EXISTS fabric_weights (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_type TEXT NOT NULL,
                size TEXT NOT NULL,
                weight_grams REAL NOT NULL,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(item_type, size)
            );
            CREATE INDEX IF NOT EXISTS idx_fabric_weights_specs
            ON fabric_weights(item_type, size);

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

            CREATE TABLE IF NOT EXISTS stock_audit_reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                source TEXT,
                note TEXT,
                bucket_key TEXT,
                filters_json TEXT,
                line_count INTEGER NOT NULL DEFAULT 0,
                total_diff INTEGER NOT NULL DEFAULT 0,
                total_value REAL NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_stock_audit_reports_created
            ON stock_audit_reports(created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_stock_audit_reports_bucket
            ON stock_audit_reports(source, bucket_key);
            CREATE TABLE IF NOT EXISTS stock_audit_report_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_id INTEGER NOT NULL,
                stock_id INTEGER,
                item_type TEXT,
                school TEXT,
                color TEXT,
                size TEXT,
                warehouse_no TEXT,
                package_no TEXT,
                expected INTEGER NOT NULL,
                actual INTEGER NOT NULL,
                diff INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                diff_value REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_stock_audit_report_lines_report
            ON stock_audit_report_lines(report_id);

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

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(stock_audit_reports)")}
            if "bucket_key" not in cols:
                self.conn.execute(
                    "ALTER TABLE stock_audit_reports ADD COLUMN bucket_key TEXT"
                )
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_stock_audit_reports_bucket ON stock_audit_reports(source, bucket_key)"
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
            if "bill_type" not in cols:
                self.conn.execute(
                    "ALTER TABLE bills ADD COLUMN bill_type TEXT NOT NULL DEFAULT 'SALE'"
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

        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS shipment_receipt_reviews (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sync_event_uuid TEXT NOT NULL UNIQUE,
                shipment_uuid TEXT NOT NULL,
                source_device TEXT NOT NULL,
                payload_json TEXT NOT NULL,
                has_diff INTEGER NOT NULL DEFAULT 0,
                note TEXT,
                status TEXT NOT NULL DEFAULT 'PENDING', -- PENDING | ACCEPTED | REJECTED
                decision_note TEXT,
                decided_at TEXT,
                created_at TEXT NOT NULL,
                shown_at TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_shipment_receipt_reviews_status
                ON shipment_receipt_reviews(status, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_shipment_receipt_reviews_shown
                ON shipment_receipt_reviews(shown_at, created_at DESC);
        """)

        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS branch_cash_receipts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                branch_device TEXT NOT NULL,
                amount REAL NOT NULL,
                received_at TEXT NOT NULL,
                note TEXT,
                created_at TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_branch_cash_receipts_branch_date
                ON branch_cash_receipts(branch_device, received_at DESC);
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
            return DEFAULT_SIZE_PROFILE
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

    # ------------------- Price Profiles -------------------
    def list_price_profiles(self) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            """
            SELECT
                p.id,
                p.name,
                p.notes,
                p.created_at,
                p.updated_at,
                COALESCE((
                    SELECT COUNT(*)
                    FROM price_profile_assignments a
                    WHERE a.profile_id = p.id
                ), 0) AS assignment_count
            FROM price_profiles p
            ORDER BY LOWER(TRIM(p.name)), p.id
            """
        )
        return [dict(r) for r in cur.fetchall()]

    def create_price_profile(self, name: str, notes: str = "") -> int:
        clean_name = str(name or "").strip()
        if not clean_name:
            raise ValueError("Ø§Ø³Ù… Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø± Ù…Ø·Ù„ÙˆØ¨.")
        with self.conn:
            cur = self.conn.execute(
                """
                INSERT INTO price_profiles(name, notes, created_at, updated_at)
                VALUES (?, ?, datetime('now'), datetime('now'))
                """,
                (clean_name, str(notes or "").strip() or None),
            )
            return int(cur.lastrowid)

    def rename_price_profile(self, profile_id: int, name: str, notes: str = "") -> None:
        clean_name = str(name or "").strip()
        if not clean_name:
            raise ValueError("Ø§Ø³Ù… Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø± Ù…Ø·Ù„ÙˆØ¨.")
        with self.conn:
            cur = self.conn.execute(
                """
                UPDATE price_profiles
                SET name=?, notes=?, updated_at=datetime('now')
                WHERE id=?
                """,
                (clean_name, str(notes or "").strip() or None, int(profile_id)),
            )
            if int(cur.rowcount or 0) <= 0:
                raise ValueError("Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø± ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯.")

    def delete_price_profile(self, profile_id: int) -> None:
        used = self.conn.execute(
            "SELECT COUNT(*) AS c FROM price_profile_assignments WHERE profile_id=?",
            (int(profile_id),),
        ).fetchone()
        if int((used["c"] if used else 0) or 0) > 0:
            raise ValueError("Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø­Ø°Ù Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ù„Ø£Ù†Ù‡ Ù…Ø³ØªØ®Ø¯Ù… ÙÙŠ Ø±Ø¨Ø· Ø£ØµÙ†Ø§Ù Ø§Ù„Ù…Ø®Ø²ÙˆÙ†.")
        with self.conn:
            self.conn.execute("DELETE FROM price_profile_lines WHERE profile_id=?", (int(profile_id),))
            cur = self.conn.execute("DELETE FROM price_profiles WHERE id=?", (int(profile_id),))
            if int(cur.rowcount or 0) <= 0:
                raise ValueError("Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø± ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯.")

    def get_price_profile(self, profile_id: int) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            "SELECT id, name, notes, created_at, updated_at FROM price_profiles WHERE id=?",
            (int(profile_id),),
        ).fetchone()
        return None if row is None else dict(row)

    def list_price_profile_item_types(self, profile_id: int) -> List[str]:
        cur = self.conn.execute(
            """
            SELECT DISTINCT TRIM(item_type) AS item_type
            FROM price_profile_lines
            WHERE profile_id=?
            ORDER BY LOWER(TRIM(item_type))
            """,
            (int(profile_id),),
        )
        return [str(r["item_type"] or "").strip() for r in cur.fetchall() if str(r["item_type"] or "").strip()]

    def list_sizes_for_price_profile_item(self, item_type: str) -> List[str]:
        item = str(item_type or "").strip()
        if not item:
            return []
        sizes: Dict[str, str] = {}
        has_size_profile = False

        cur = self.conn.execute(
            """
            SELECT num_start_1, num_end_1, num_start_2, num_end_2, has_alpha
            FROM size_profiles
            WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
            """,
            (item,),
        )
        for row in cur.fetchall():
            has_size_profile = True
            sizes_from_profile = merged_numeric_size_labels_from_profile(
                row["num_start_1"],
                row["num_end_1"],
                row["num_start_2"],
                row["num_end_2"],
            )
            for size in sizes_from_profile:
                sizes.setdefault(size.casefold(), size)
            if int(row["has_alpha"] or 0):
                for size in ALPHA_SIZES:
                    sizes.setdefault(size.casefold(), size)

        for table_name in ("stocks", "bill_items", "price_profile_lines"):
            cur = self.conn.execute(
                f"""
                SELECT DISTINCT TRIM(size) AS size
                FROM {table_name}
                WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                  AND TRIM(size) <> ''
                """,
                (item,),
            )
            for row in cur.fetchall():
                size = str(row["size"] or "").strip()
                if size:
                    sizes.setdefault(size.casefold(), size)

        if not has_size_profile and any(str(size).strip().isdigit() for size in sizes.values()):
            for size in merged_numeric_size_labels_from_profile(*DEFAULT_SIZE_PROFILE[:4]):
                sizes.setdefault(size.casefold(), size)

        if not sizes:
            for size in merged_numeric_size_labels_from_profile(*DEFAULT_SIZE_PROFILE[:4]):
                sizes.setdefault(size.casefold(), size)

        values = list(sizes.values())
        values.sort(key=warehouse_size_sort_key)
        return values

    def list_price_profile_lines(self, profile_id: int, item_type: Optional[str] = None) -> List[Dict[str, Any]]:
        where = ["profile_id=?"]
        args: List[Any] = [int(profile_id)]
        item = str(item_type or "").strip()
        if item:
            where.append("LOWER(TRIM(item_type)) = LOWER(TRIM(?))")
            args.append(item)
        cur = self.conn.execute(
            f"""
            SELECT id, profile_id, item_type, size, price, created_at, updated_at
            FROM price_profile_lines
            WHERE {' AND '.join(where)}
            ORDER BY LOWER(TRIM(item_type)), LOWER(TRIM(size))
            """,
            tuple(args),
        )
        rows = [dict(r) for r in cur.fetchall()]
        rows.sort(key=lambda r: ((r.get("item_type") or "").casefold(), warehouse_size_sort_key(r.get("size"))))
        return rows

    def replace_price_profile_item_prices(
        self,
        profile_id: int,
        item_type: str,
        rows: Sequence[Dict[str, Any]],
    ) -> None:
        item = str(item_type or "").strip()
        if not item:
            raise ValueError("Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù Ù…Ø·Ù„ÙˆØ¨ Ø¯Ø§Ø®Ù„ Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±.")
        cleaned: Dict[str, float] = {}
        for row in rows or []:
            size = str((row or {}).get("size") or "").strip()
            if not size:
                continue
            try:
                price = float((row or {}).get("price"))
            except (TypeError, ValueError):
                continue
            if price < 0:
                raise ValueError("Ø§Ù„Ø³Ø¹Ø± ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ† ØµÙØ± Ø£Ùˆ Ø£ÙƒØ¨Ø±.")
            cleaned[size] = float(price)

        with self.conn:
            self.conn.execute(
                """
                DELETE FROM price_profile_lines
                WHERE profile_id=? AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                """,
                (int(profile_id), item),
            )
            for size, price in cleaned.items():
                self.conn.execute(
                    """
                    INSERT INTO price_profile_lines(
                        profile_id, item_type, size, price, created_at, updated_at
                    )
                    VALUES (?, ?, ?, ?, datetime('now'), datetime('now'))
                    """,
                    (int(profile_id), item, size, float(price)),
                )
            self.conn.execute(
                "UPDATE price_profiles SET updated_at=datetime('now') WHERE id=?",
                (int(profile_id),),
            )

    def get_price_profile_assignment(self, item_type: str, school: str, color: str) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            """
            SELECT
                a.id,
                a.item_type,
                a.school,
                a.color,
                a.profile_id,
                a.updated_at,
                p.name AS profile_name
            FROM price_profile_assignments a
            JOIN price_profiles p ON p.id = a.profile_id
            WHERE LOWER(TRIM(a.item_type)) = LOWER(TRIM(?))
              AND LOWER(TRIM(a.school)) = LOWER(TRIM(?))
              AND LOWER(TRIM(a.color)) = LOWER(TRIM(?))
            """,
            (item_type, school, color),
        ).fetchone()
        return None if row is None else dict(row)

    def assign_price_profile(self, item_type: str, school: str, color: str, profile_id: int) -> None:
        item = str(item_type or "").strip()
        school_txt = str(school or "").strip()
        color_txt = str(color or "").strip()
        if not (item and school_txt and color_txt):
            raise ValueError("Ø§Ù„Ù†ÙˆØ¹ ÙˆØ§Ù„Ù…Ø¯Ø±Ø³Ø© ÙˆØ§Ù„Ù„ÙˆÙ† Ù…Ø·Ù„ÙˆØ¨Ø© Ù„Ø±Ø¨Ø· Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±.")
        if self.get_price_profile(int(profile_id)) is None:
            raise ValueError("Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø± ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯.")
        with self.conn:
            self.conn.execute(
                """
                INSERT INTO price_profile_assignments(
                    item_type, school, color, profile_id, updated_at
                )
                VALUES (?, ?, ?, ?, datetime('now'))
                ON CONFLICT(item_type, school, color)
                DO UPDATE SET
                    profile_id = excluded.profile_id,
                    updated_at = datetime('now')
                """,
                (item, school_txt, color_txt, int(profile_id)),
            )

    def clear_price_profile_assignment(self, item_type: str, school: str, color: str) -> int:
        with self.conn:
            cur = self.conn.execute(
                """
                DELETE FROM price_profile_assignments
                WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                  AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                  AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                """,
                (item_type, school, color),
            )
            return int(cur.rowcount or 0)

    def get_price_profile_price(
        self,
        profile_id: int,
        item_type: str,
        size: str,
    ) -> Optional[float]:
        item_key = str(item_type or "").strip()
        size_key = _normalize_size_label(_strip_digit_marks(size))
        row = self.conn.execute(
            """
            SELECT price
            FROM price_profile_lines
            WHERE profile_id=?
              AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
              AND LOWER(TRIM(size)) = LOWER(TRIM(?))
            LIMIT 1
            """,
            (int(profile_id), item_key, size_key),
        ).fetchone()
        if row is None or row["price"] is None:
            return None
        try:
            return float(row["price"])
        except (TypeError, ValueError):
            return None

    def apply_price_profile_to_stock(
        self,
        profile_id: int,
        targets: Sequence[Tuple[str, str, str]],
        *,
        stock_ids: Optional[Sequence[int]] = None,
        note: str = "Price profile applied",
    ) -> Dict[str, int]:
        clean_targets = [
            (str(item or "").strip(), str(school or "").strip(), str(color or "").strip())
            for item, school, color in (targets or [])
        ]
        clean_targets = [(item, school, color) for item, school, color in clean_targets if item and school and color]
        if not clean_targets:
            return {"updated": 0, "skipped": 0, "checked": 0}

        where: List[str] = []
        args: List[Any] = []

        ids = [int(x) for x in (stock_ids or []) if x not in (None, "")]
        if ids:
            placeholders = ",".join("?" for _ in ids)
            where.append(f"id IN ({placeholders})")
            args.extend(ids)

        target_clauses: List[str] = []
        for item, school, color in clean_targets:
            target_clauses.append(
                "(LOWER(TRIM(item_type)) = LOWER(?) AND LOWER(TRIM(school)) = LOWER(?) AND LOWER(TRIM(color)) = LOWER(?))"
            )
            args.extend([item, school, color])
        where.append("(" + " OR ".join(target_clauses) + ")")

        cur = self.conn.execute(
            f"SELECT id, item_type, school, color, size, unit_price FROM stocks WHERE {' AND '.join(where)}",
            tuple(args),
        )
        rows = [dict(r) for r in cur.fetchall()]
        updated = 0
        skipped = 0
        for row in rows:
            price = self.get_price_profile_price(profile_id, row.get("item_type"), row.get("size"))
            if price is None:
                skipped += 1
                continue
            try:
                old_price = float(row.get("unit_price") or 0.0)
            except (TypeError, ValueError):
                old_price = None
            if old_price is not None and abs(old_price - float(price)) < 0.000001:
                skipped += 1
                continue
            updated += self.update_prices(
                {"id": int(row["id"])},
                float(price),
                note=note,
                price_sync_mode="all-pos",
            )
        return {"updated": int(updated), "skipped": int(skipped), "checked": len(rows)}

    def apply_price_profile_item_to_assigned_stock(
        self,
        profile_id: int,
        item_type: str,
        note: str = "Price profile prices updated",
    ) -> Dict[str, int]:
        item = str(item_type or "").strip()
        if not item:
            return {"updated": 0, "skipped": 0, "checked": 0}
        rows = self.conn.execute(
            """
            SELECT item_type, school, color
            FROM price_profile_assignments
            WHERE profile_id=?
              AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
            ORDER BY LOWER(TRIM(school)), LOWER(TRIM(color))
            """,
            (int(profile_id), item),
        ).fetchall()
        targets = [
            (
                str(r["item_type"] or "").strip(),
                str(r["school"] or "").strip(),
                str(r["color"] or "").strip(),
            )
            for r in rows
        ]
        if not targets:
            return {"updated": 0, "skipped": 0, "checked": 0}
        return self.apply_price_profile_to_stock(
            int(profile_id),
            targets,
            note=note,
        )

    def resolve_price_profile_id(
        self,
        item_type: str,
        school: str,
        color: str,
        preferred_profile_id: Optional[int] = None,
    ) -> Optional[int]:
        if preferred_profile_id not in (None, "", 0):
            try:
                pid = int(preferred_profile_id)
            except (TypeError, ValueError):
                pid = 0
            if pid > 0 and self.get_price_profile(pid):
                return pid
        assignment = self.get_price_profile_assignment(item_type, school, color)
        if assignment and assignment.get("profile_id") is not None:
            try:
                return int(assignment["profile_id"])
            except (TypeError, ValueError):
                return None
        return None

    # ------------------- Fabric weights -------------------
    def list_fabric_weights(self, item_type: Optional[str] = None) -> List[Dict[str, Any]]:
        where = ["1=1"]
        args: List[Any] = []
        item = str(item_type or "").strip()
        if item:
            where.append("LOWER(TRIM(item_type)) = LOWER(TRIM(?))")
            args.append(item)
        cur = self.conn.execute(
            f"""
            SELECT id, item_type, size, weight_grams, created_at, updated_at
            FROM fabric_weights
            WHERE {' AND '.join(where)}
            ORDER BY LOWER(TRIM(item_type)), LOWER(TRIM(size))
            """,
            tuple(args),
        )
        return [dict(r) for r in cur.fetchall()]

    def get_fabric_weight(self, item_type: str, size: str) -> Optional[float]:
        row = self.conn.execute(
            """
            SELECT weight_grams
            FROM fabric_weights
            WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
              AND LOWER(TRIM(size)) = LOWER(TRIM(?))
            """,
            (item_type, size),
        ).fetchone()
        if row is None or row["weight_grams"] is None:
            return None
        try:
            return float(row["weight_grams"])
        except (TypeError, ValueError):
            return None

    def upsert_fabric_weight(self, item_type: str, size: str, weight_grams: float) -> None:
        item = str(item_type or "").strip()
        size_txt = str(size or "").strip()
        if not item:
            raise ValueError("Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù Ù…Ø·Ù„ÙˆØ¨.")
        if not size_txt:
            raise ValueError("Ø§Ù„Ù…Ù‚Ø§Ø³ Ù…Ø·Ù„ÙˆØ¨.")
        try:
            weight_val = float(weight_grams)
        except (TypeError, ValueError):
            raise ValueError("ÙˆØ²Ù† Ø§Ù„Ù‚Ù…Ø§Ø´ ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ† Ø±Ù‚Ù…Ø§Ù‹ ØµØ§Ù„Ø­Ø§Ù‹.")
        if weight_val <= 0:
            raise ValueError("ÙˆØ²Ù† Ø§Ù„Ù‚Ù…Ø§Ø´ ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ† Ø£ÙƒØ¨Ø± Ù…Ù† ØµÙØ±.")
        with self.conn:
            self.conn.execute(
                """
                INSERT INTO fabric_weights(
                    item_type, size, weight_grams, created_at, updated_at
                )
                VALUES (?, ?, ?, datetime('now'), datetime('now'))
                ON CONFLICT(item_type, size)
                DO UPDATE SET
                    weight_grams = excluded.weight_grams,
                    updated_at = datetime('now')
                """,
                (item, size_txt, weight_val),
            )

    def delete_fabric_weight(self, item_type: str, size: str) -> int:
        with self.conn:
            cur = self.conn.execute(
                """
                DELETE FROM fabric_weights
                WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                  AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                """,
                (item_type, size),
            )
            return int(cur.rowcount or 0)

    def calculate_fabric_requirements(self, filters: Dict[str, Any]) -> Dict[str, Any]:
        movement_rows = self.list_movement_item_totals(filters)
        use_branch_only = bool(str(filters.get("branch_device") or "").strip())
        weight_rows = self.list_fabric_weights()
        weight_map = {
            (
                _normalize_spec_label(r.get("item_type") or "").lower(),
                _normalize_size_label(r.get("size") or "").lower(),
            ): float(r.get("weight_grams") or 0.0)
            for r in weight_rows
        }

        grouped: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
        missing: List[Dict[str, Any]] = []
        detail_rows: List[Dict[str, Any]] = []
        total_requested_qty = 0
        total_weight_grams = 0.0

        for row in movement_rows:
            requested_qty = int(row.get("requested_qty") or 0)
            if requested_qty <= 0:
                continue

            item_type = str(row.get("item_type") or "").strip()
            school = str(row.get("school") or "").strip()
            color = str(row.get("color") or "").strip()
            size = str(row.get("size") or "").strip()
            lookup_key = (
                _normalize_spec_label(item_type).lower(),
                _normalize_size_label(size).lower(),
            )
            weight_grams = weight_map.get(lookup_key)
            if weight_grams is None or weight_grams <= 0:
                missing.append(
                    {
                        "school": school,
                        "item_type": item_type,
                        "color": color,
                        "size": size,
                        "requested_qty": requested_qty,
                    }
                )
                continue

            line_grams = float(requested_qty) * float(weight_grams)
            detail_rows.append(
                {
                    "school": school,
                    "item_type": item_type,
                    "color": color,
                    "size": size,
                    "requested_qty": requested_qty,
                    "weight_grams": float(weight_grams),
                    "line_weight_grams": float(line_grams),
                }
            )

            group_key = (school, item_type, color)
            bucket = grouped.setdefault(
                group_key,
                {
                    "school": school,
                    "item_type": item_type,
                    "color": color,
                    "requested_qty": 0,
                    "total_weight_grams": 0.0,
                },
            )
            bucket["requested_qty"] += int(requested_qty)
            bucket["total_weight_grams"] += float(line_grams)
            total_requested_qty += int(requested_qty)
            total_weight_grams += float(line_grams)

        summary_rows = list(grouped.values())
        summary_rows.sort(
            key=lambda r: (
                (r.get("color") or "").lower(),
                (r.get("school") or "").lower(),
                warehouse_item_priority(r.get("item_type")),
                (r.get("item_type") or "").lower(),
            )
        )
        detail_rows.sort(
            key=lambda r: (
                (r.get("color") or "").lower(),
                (r.get("school") or "").lower(),
                warehouse_item_priority(r.get("item_type")),
                (r.get("item_type") or "").lower(),
                _normalize_size_label(r.get("size") or "").lower(),
            )
        )
        missing.sort(
            key=lambda r: (
                (r.get("color") or "").lower(),
                (r.get("school") or "").lower(),
                warehouse_item_priority(r.get("item_type")),
                (r.get("item_type") or "").lower(),
                _normalize_size_label(r.get("size") or "").lower(),
            )
        )
        return {
            "rows": summary_rows,
            "details": detail_rows,
            "missing": missing,
            "summary": {
                "row_count": len(summary_rows),
                "requested_qty": int(total_requested_qty),
                "total_weight_grams": float(total_weight_grams),
                "missing_count": len(missing),
            },
        }

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

    def get_effective_price(self, item_type, school, color, size, preferred_profile_id: Optional[int] = None):
        """
        Priority:
        1) price profile (explicit or assigned)
        2) exact last price (history)
        3) default price for item
        3) None if unknown
        """
        try:
            profile_id = self.resolve_price_profile_id(
                item_type,
                school,
                color,
                preferred_profile_id=preferred_profile_id,
            )
            if profile_id:
                p = self.get_price_profile_price(profile_id, item_type, size)
                if p is not None:
                    return float(p)
        except Exception:
            pass

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
            raise ValueError("Ø­Ù‚Ù„ ØºÙŠØ± Ù…Ø¯Ø¹ÙˆÙ…")

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


    # AFTER (NEW) â€” list all schools (from stocks or history/bill_items)
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

    def get_school_accounts_report(
        self,
        schools: Sequence[str],
        *,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        cleaned = [str(s or "").strip() for s in schools if str(s or "").strip()]
        if not cleaned:
            return []
        placeholders = ",".join("?" for _ in cleaned)
        where = [
            "COALESCE(b.status,'CONFIRMED')='CONFIRMED'",
            "(b.bill_type='SALE' OR b.bill_type IS NULL)",
            "LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))",
            "LOWER(TRIM(COALESCE(b.customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))",
            "UPPER(TRIM(COALESCE(b.customer,''))) != 'WAREHOUSE'",
            f"LOWER(TRIM(bi.school)) IN ({placeholders})",
        ]
        args: List[Any] = [s.lower() for s in cleaned]
        if date_from:
            where.append("date(b.created_at) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(b.created_at) <= date(?)")
            args.append(date_to)
        cur = self.conn.cursor()
        try:
            cur.execute(
                f"""
                SELECT
                    TRIM(bi.school) AS school,
                    TRIM(bi.item_type) AS item_type,
                    TRIM(bi.color) AS color,
                    TRIM(bi.size) AS size,
                    COALESCE(SUM(bi.qty), 0) AS total_qty,
                    COALESCE(SUM(bi.line_total), 0) AS total_sales
                FROM bill_items bi
                JOIN bills b ON b.id = bi.bill_id
                WHERE {' AND '.join(where)}
                GROUP BY TRIM(bi.school), TRIM(bi.item_type), TRIM(bi.color), TRIM(bi.size)
                ORDER BY LOWER(TRIM(bi.school)), LOWER(TRIM(bi.item_type)), LOWER(TRIM(bi.color)), LOWER(TRIM(bi.size))
                """,
                args,
            )
            return [dict(r) for r in cur.fetchall()]
        finally:
            cur.close()

    # AFTER (NEW) â€” items for a school: distinct (item_type, color) seen anywhere
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

        # AFTER (NEW) â€” sizes for a (school, item_type, color) with current count + last price
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

            rename_events = self._cascade_spec_rename(old_specs, changes)

            self._upsert_history(changes)
        self._emit_spec_rename_sync_events(rename_events)
        self.cleanup_unused_specs()
        return count


    def _cascade_spec_rename(
        self,
        old_specs: Sequence[Any],
        changes: Dict[str, Any],
    ) -> List[Dict[str, Any]]:
        """Propagate a completed spec rename from `stocks` to historical tables.

        For each old (item_type, school, color, size) tuple we just rewrote,
        if no `stocks` row still references that tuple, it was a full rename â€”
        cascade the same field changes to `movements`, `bill_items`, and
        (best-effort) the POS mirror tables so audit views stay consistent.
        """
        if not changes or not old_specs:
            return []
        sets_parts: List[str] = []
        new_vals: List[Any] = []
        for fld in ("item_type", "school", "color", "size"):
            if fld in changes:
                sets_parts.append(f"{fld} = ?")
                new_vals.append(changes[fld])
        if not sets_parts:
            return []
        set_sql = ", ".join(sets_parts)
        rename_events: List[Dict[str, Any]] = []

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
            new_spec = {
                "item_type": str(changes.get("item_type", old_it) or "").strip(),
                "school": str(changes.get("school", old_sc) or "").strip(),
                "color": str(changes.get("color", old_cl) or "").strip(),
                "size": str(changes.get("size", old_sz) or "").strip(),
            }
            old_spec = {
                "item_type": str(old_it or "").strip(),
                "school": str(old_sc or "").strip(),
                "color": str(old_cl or "").strip(),
                "size": str(old_sz or "").strip(),
            }
            if old_spec != new_spec:
                rename_events.append({
                    "old_spec": old_spec,
                    "new_spec": new_spec,
                    "changed_fields": [fld for fld in ("item_type", "school", "color", "size") if old_spec.get(fld) != new_spec.get(fld)],
                })
        return rename_events

    def _emit_spec_rename_sync_events(self, rename_events: Sequence[Dict[str, Any]]) -> None:
        for payload in (rename_events or []):
            try:
                old_spec = payload.get("old_spec") or {}
                new_spec = payload.get("new_spec") or {}
                if not isinstance(old_spec, dict) or not isinstance(new_spec, dict):
                    continue
                if old_spec == new_spec:
                    continue
                self._record_sync_event("SPEC_RENAMED", dict(payload), target_scope="all-pos")
            except Exception:
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
            raise ValueError("Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ© ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ†Ø§ >= 1.")

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

            rename_events = self._cascade_spec_rename(old_specs, changes)

            self._upsert_history(changes)
        self._emit_spec_rename_sync_events(rename_events)
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
                    # dual match: original text AND normalized (Arabicâ†’Latin digits, uppercased)
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

            unit_price_filter = filters.get("unit_price")
            if unit_price_filter not in (None, ""):
                where.append("unit_price = ?")
                args.append(float(unit_price_filter))


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
            raise ValueError("Ø±Ù‚Ù… Ø§Ù„Ø¹Ø¨ÙˆØ© ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ† 1 Ø£Ùˆ Ø£ÙƒØ¨Ø±.")
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
        raise ValueError("Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø§Ù„Ø¥Ø¶Ø§ÙØ©: Ù‡Ø°Ù‡ Ø§Ù„Ø¹Ø¨ÙˆØ© Ù…ØºÙ„Ù‚Ø© ÙˆØ¨Ø¯Ø§Ø®Ù„Ù‡Ø§ Ù…Ø®Ø²ÙˆÙ†.")


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
                raise ValueError("Ù„Ø§ ØªÙˆØ¬Ø¯ Ø­Ø§ÙˆÙŠØ© Ø¨Ù‡Ø°Ø§ Ø§Ù„Ø±Ù‚Ù… Ù„Ø¥ØºÙ„Ø§Ù‚Ù‡Ø§.")
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
            raise ValueError("Ø±Ù‚Ù… Ø§Ù„Ø¹Ø¨ÙˆØ© ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ† 1 Ø£Ùˆ Ø£ÙƒØ¨Ø±.")

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
                raise ValueError("Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø¥Ø¹Ø§Ø¯Ø© Ø§Ù„ÙØªØ­ Ù…Ø¹ ÙˆØ¬ÙˆØ¯ Ù…Ø®Ø²ÙˆÙ† Ø¯Ø§Ø®Ù„ Ø§Ù„Ø¹Ø¨ÙˆØ© (Ø®ÙŠØ§Ø± 'ÙØ§Ø±ØºØ© ÙÙ‚Ø·' Ù…ÙØ¹Ù„).")

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
                WHERE TRIM({field}) <> ''
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
                raise ValueError("Price must be â‰¥ 0")
        else:
            price = self.get_effective_price(item_type, school, color, size)

        if price is None:
            raise ValueError("ÙŠØ¬Ø¨ Ø¥Ø¯Ø®Ø§Ù„ Ø§Ù„Ø³Ø¹Ø± Ù…Ø±Ø© ÙˆØ§Ø­Ø¯Ø© Ø¹Ù„Ù‰ Ø§Ù„Ø£Ù‚Ù„ Ù„Ù‡Ø°Ø§ Ø§Ù„ØµÙ†Ù.")

        # Save default price ONLY if user explicitly typed it
        if user_provided_price:
            self.ensure_default_price(item_type, price)

        self.ensure_package_open(warehouse_no, package_no)

        with self.conn:
            stock_id = self.add_or_update_stock_row(
                item_type.strip(),
                school.strip(),
                color.strip(),
                size.strip(),
                int(warehouse_no),
                int(package_no),
                float(price),
                int(count),
                int(has_badge),
            )

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
            raise ValueError("ÙŠÙ…ÙƒÙ† Ø§Ø®ØªÙŠØ§Ø± Ø£ÙƒØ«Ø± Ù…Ù† Ù‚ÙŠÙ…Ø© Ù„Ø­Ù‚Ù„ ÙˆØ§Ø­Ø¯ ÙÙ‚Ø·ØŒ ÙˆÙ„ÙŠØ³ Ù„Ø£ÙƒØ«Ø± Ù…Ù† Ø­Ù‚Ù„.")

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

    def _stock_candidates_for_bill_line(self, line: Dict[str, Any]) -> List[sqlite3.Row]:
        cur = self.conn.cursor()
        try:
            if line.get("stock_id"):
                cur.execute(
                    "SELECT * FROM stocks WHERE id=? AND count>0 ORDER BY id ASC",
                    (int(line["stock_id"]),),
                )
                return cur.fetchall()

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
            cur.execute(
                f"SELECT * FROM stocks WHERE {' AND '.join(where_parts)} ORDER BY id ASC",
                args,
            )
            return cur.fetchall()
        finally:
            cur.close()

    def available_stock_qty_for_bill_line(self, line: Dict[str, Any]) -> int:
        """Current stock available for one non-factory bill line."""
        return sum(int(r["count"] or 0) for r in self._stock_candidates_for_bill_line(line))

    def validate_stock_available_for_bill_lines(self, bill_lines: List[Dict[str, Any]]) -> None:
        """Preflight stock deductions before a bill/draft can mutate inventory."""
        simulated_counts: Dict[int, int] = {}
        for line in bill_lines:
            if bool(line.get("allow_factory_fill")) or str(line.get("origin") or "").upper() == "FACTORY":
                continue
            try:
                qty_needed = int(line.get("qty") or 0)
            except Exception:
                raise ValueError("Qty must be > 0")
            if qty_needed <= 0:
                raise ValueError("Qty must be > 0")

            candidates = self._stock_candidates_for_bill_line(line)
            available_before = sum(int(r["count"] or 0) for r in candidates)
            remaining = qty_needed
            for s in candidates:
                sid = int(s["id"])
                available = simulated_counts.setdefault(sid, int(s["count"] or 0))
                take = min(available, remaining)
                if take > 0:
                    simulated_counts[sid] = available - take
                    remaining -= take
                if remaining <= 0:
                    break
            if remaining > 0:
                raise ValueError(
                    f"Ù„Ø§ ØªÙˆØ¬Ø¯ ÙƒÙ…ÙŠØ© ÙƒØ§ÙÙŠØ© Ù„Ù„ØµÙ†Ù "
                    f"{line.get('item_type','(Ø§Ù„Ù†ÙˆØ¹?)')} / {line.get('school','(Ø§Ù„Ù…Ø¯Ø±Ø³Ø©?)')} / "
                    f"{line.get('color','(Ø§Ù„Ù„ÙˆÙ†?)')} / {line.get('size','(Ø§Ù„Ù…Ù‚Ø§Ø³?)')} "
                    f"(Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ {qty_needed}ØŒ Ø§Ù„Ù…ØªØ§Ø­ {qty_needed - remaining})."
                )

    def current_inventory(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        where, args = self._filters_where(filters, prefix="s.")
        having_sql = "HAVING SUM(s.count) > 0" if (filters or {}).get("hide_zero") else ""

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
            {having_sql}
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
                    package_no=r["package_no"],  # â† now "93, 121, 137"
                    unit_price=r["unit_price"],
                    count=int(r["count"]),
                    value=float(r["value"]),
                    has_badge=int(r["has_badge"] or 0),
                )
            )

        cur.close()
        rows.sort(
            key=lambda r: (
                warehouse_item_priority(r.get("item_type")),
                (r.get("item_type") or "").lower(),
                (r.get("school") or "").lower(),
                (r.get("color") or "").lower(),
                warehouse_size_sort_key(r.get("size")),
                str(r.get("warehouse_no") or ""),
                str(r.get("package_no") or ""),
            )
        )
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
        if (target_pos or "").strip():
            for line in bill_lines:
                if bool(line.get("allow_factory_fill")):
                    raise ValueError("Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø´Ø­Ù† Ø¨Ù†Ø¯ Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹ Ø¥Ù„Ù‰ ÙØ±Ø¹. Ø§Ø®ØªØ± ÙƒÙ…ÙŠØ© Ù…ØªØ§Ø­Ø© Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø£ÙˆÙ„Ø§Ù‹.")
        self.validate_stock_available_for_bill_lines(bill_lines)

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
                    cands = self._stock_candidates_for_bill_line(line)

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
                        f"Ù„Ø§ ØªÙˆØ¬Ø¯ ÙƒÙ…ÙŠØ© ÙƒØ§ÙÙŠØ© Ù„Ù„ØµÙ†Ù "
                        f"{line.get('item_type','(Ø§Ù„Ù†ÙˆØ¹?)')} / {line.get('school','(Ø§Ù„Ù…Ø¯Ø±Ø³Ø©?)')} / {line.get('size','(Ø§Ù„Ù…Ù‚Ø§Ø³?)')} "
                        f"(Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ {qty_needed}ØŒ Ø§Ù„Ù…ØªØ§Ø­ {qty_needed-remaining})."
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
                self._record_sync_event_or_raise("SALE_CREATED", {
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
            """SELECT id,item_type,school,color,size,warehouse_no,package_no,unit_price,qty,line_total,origin,has_badge
            FROM bill_items WHERE bill_id=?""",
            (int(bill_id),),
        )
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows

    # -------- Bill Lifecycle (Draft / Confirm / Void) --------

    def create_draft_bill(self, customer: str, bill_lines: List[Dict[str, Any]]) -> int:
        """Create a DRAFT bill â€” saves bill data but does NOT deduct stock."""
        if not bill_lines:
            raise ValueError("Bill has no items")
        self.validate_stock_available_for_bill_lines(bill_lines)
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
        """Confirm a DRAFT bill â€” deduct stock and record movements."""
        cur = self.conn.execute(
            "SELECT customer,total,COALESCE(status,'CONFIRMED') AS status FROM bills WHERE id=?", (int(bill_id),))
        row = cur.fetchone()
        if not row:
            raise ValueError("Ø§Ù„ÙØ§ØªÙˆØ±Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø©")
        if row["status"] != "DRAFT":
            raise ValueError("Ù‡Ø°Ù‡ Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ù„ÙŠØ³Øª Ù…Ø³ÙˆØ¯Ø©")
        items = self.list_bill_items(bill_id)
        if not items:
            raise ValueError("Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ù„Ø§ ØªØ­ØªÙˆÙŠ Ø¹Ù„Ù‰ Ø¨Ù†ÙˆØ¯")
        self.validate_stock_available_for_bill_lines(items)
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
                cands = self._stock_candidates_for_bill_line(item)
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
                        f"ÙƒÙ…ÙŠØ© ØºÙŠØ± ÙƒØ§ÙÙŠØ© Ù„Ù„ØµÙ†Ù {item['item_type']} / {item['school']} / {item['size']} "
                        f"(Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ {qty_needed}ØŒ Ø§Ù„Ù…ØªØ§Ø­ {qty_needed - remaining})"
                    )
            self.cleanup_unused_specs()
            self.conn.execute(
                "UPDATE bills SET status='CONFIRMED' WHERE id=?", (int(bill_id),)
            )
            for wh, pkg in affected_packages:
                self.auto_reopen_package_if_empty(wh, pkg)
            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id=?", (int(bill_id),)
            ).fetchone()
            items_payload = [
                {
                    "item_type": item["item_type"],
                    "school": item["school"],
                    "color": item["color"],
                    "size": item["size"],
                    "warehouse_no": int(item.get("warehouse_no") or 0),
                    "package_no": int(item.get("package_no") or 0),
                    "unit_price": float(item["unit_price"]),
                    "qty": int(item["qty"]),
                    "line_total": float(item.get("line_total") or 0.0),
                    "origin": item.get("origin"),
                    "has_badge": int(item.get("has_badge") or 0),
                }
                for item in items
            ]
            branch_target = normalize_branch_customer_name(row["customer"])
            if branch_target:
                shipment_items = [
                    {
                        "item_type": it["item_type"],
                        "school": it["school"],
                        "color": it["color"],
                        "size": it["size"],
                        "unit_price": float(it["unit_price"]),
                        "qty": int(it["qty"]),
                    }
                    for it in items_payload if int(it["qty"]) > 0
                ]
                self._record_branch_shipment_event(
                    shipment_uuid=(bill_uuid_row[0] if bill_uuid_row else None) or "",
                    target_name=branch_target,
                    note=f"draft bill #{bill_id}",
                    lines=shipment_items,
                )
            else:
                self._record_sync_event_or_raise("SALE_CREATED", {
                    "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                    "bill_id": int(bill_id),
                    "customer": (row["customer"] or "").strip() or None,
                    "total": float(row["total"] or 0.0),
                    "items": items_payload,
                })

    def void_bill(self, bill_id: int) -> None:
        """Void a confirmed bill â€” return items to stock."""
        cur = self.conn.execute(
            "SELECT COALESCE(status,'CONFIRMED') AS status FROM bills WHERE id=?", (int(bill_id),))
        row = cur.fetchone()
        if not row:
            raise ValueError("Ø§Ù„ÙØ§ØªÙˆØ±Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø©")
        if row["status"] != "CONFIRMED":
            raise ValueError("ÙŠÙ…ÙƒÙ† Ø¥Ù„ØºØ§Ø¡ Ø§Ù„ÙÙˆØ§ØªÙŠØ± Ø§Ù„Ù…Ø¤ÙƒØ¯Ø© ÙÙ‚Ø·")
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
            self._record_sync_event_or_raise("SALE_VOIDED", {
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
            raise ValueError("Ø§Ù„ÙØ§ØªÙˆØ±Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø©")
        if row["status"] != "DRAFT":
            raise ValueError("ÙŠÙ…ÙƒÙ† Ø­Ø°Ù Ø§Ù„Ù…Ø³ÙˆØ¯Ø§Øª ÙÙ‚Ø·")
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
            raise ValueError("Ø§Ù„ÙØ§ØªÙˆØ±Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø©")
        if row["status"] != "CONFIRMED":
            raise ValueError("Ø§Ù„Ù…Ø±ØªØ¬Ø¹Ø§Øª Ù…ØªØ§Ø­Ø© ÙÙ‚Ø· Ù„Ù„ÙÙˆØ§ØªÙŠØ± Ø§Ù„Ù…Ø¤ÙƒØ¯Ø©")
        if not return_lines:
            raise ValueError("Ù„Ø§ ØªÙˆØ¬Ø¯ Ø£ØµÙ†Ø§Ù Ù„Ù„Ø¥Ø±Ø¬Ø§Ø¹")
        branch_target = normalize_branch_customer_name(row["customer"])
        positive_lines: List[Dict[str, Any]] = []
        requested_by_bill_item: Dict[int, int] = {}
        requested_by_spec: Dict[Tuple[str, str, str, str, int, int, float, int], int] = {}
        for line in return_lines:
            qty = int(line.get("qty") or 0)
            if qty <= 0:
                continue
            positive_lines.append(line)
            bill_item_id = parse_int_text(line.get("bill_item_id"))
            if bill_item_id is not None:
                requested_by_bill_item[bill_item_id] = requested_by_bill_item.get(bill_item_id, 0) + qty
            else:
                key = (
                    str(line.get("item_type") or "").strip(),
                    str(line.get("school") or "").strip(),
                    str(line.get("color") or "").strip(),
                    str(line.get("size") or "").strip(),
                    int(line.get("warehouse_no") or 0),
                    int(line.get("package_no") or 0),
                    float(line.get("unit_price") or 0.0),
                    int(line.get("has_badge") or 0),
                )
                requested_by_spec[key] = requested_by_spec.get(key, 0) + qty
        if not positive_lines:
            raise ValueError("No returnable items were selected")
        for bill_item_id, requested_qty in requested_by_bill_item.items():
            original = self.conn.execute(
                "SELECT qty FROM bill_items WHERE id=? AND bill_id=?",
                (int(bill_item_id), int(bill_id)),
            ).fetchone()
            if not original:
                raise ValueError("Return line does not belong to this bill")
            already = self.conn.execute(
                """
                SELECT COALESCE(SUM(ri.qty),0)
                FROM return_items ri
                JOIN returns r ON r.id = ri.return_id
                WHERE r.bill_id=? AND ri.bill_item_id=?
                """,
                (int(bill_id), int(bill_item_id)),
            ).fetchone()[0] or 0
            if requested_qty > int(original["qty"] or 0) - int(already):
                raise ValueError("Return quantity exceeds the remaining bill quantity")
        for key, requested_qty in requested_by_spec.items():
            it, sc, cl, sz, wh, pkg, price, has_badge = key
            original = self.conn.execute(
                """
                SELECT COALESCE(SUM(qty),0)
                FROM bill_items
                WHERE bill_id=?
                  AND LOWER(TRIM(item_type))=LOWER(TRIM(?))
                  AND LOWER(TRIM(school))=LOWER(TRIM(?))
                  AND LOWER(TRIM(color))=LOWER(TRIM(?))
                  AND LOWER(TRIM(size))=LOWER(TRIM(?))
                  AND warehouse_no=? AND package_no=? AND unit_price=? AND COALESCE(has_badge,0)=?
                """,
                (int(bill_id), it, sc, cl, sz, wh, pkg, price, has_badge),
            ).fetchone()[0] or 0
            already = self.conn.execute(
                """
                SELECT COALESCE(SUM(ri.qty),0)
                FROM return_items ri
                JOIN returns r ON r.id = ri.return_id
                WHERE r.bill_id=?
                  AND LOWER(TRIM(ri.item_type))=LOWER(TRIM(?))
                  AND LOWER(TRIM(ri.school))=LOWER(TRIM(?))
                  AND LOWER(TRIM(ri.color))=LOWER(TRIM(?))
                  AND LOWER(TRIM(ri.size))=LOWER(TRIM(?))
                  AND ri.warehouse_no=? AND ri.package_no=? AND ri.unit_price=? AND COALESCE(ri.has_badge,0)=?
                """,
                (int(bill_id), it, sc, cl, sz, wh, pkg, price, has_badge),
            ).fetchone()[0] or 0
            if requested_qty > int(original) - int(already):
                raise ValueError("Return quantity exceeds the remaining bill quantity")
        return_lines = positive_lines
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
                            (note or "").strip() or f"Ù…Ø±ØªØ¬Ø¹ ÙØ§ØªÙˆØ±Ø© Ø´Ø­Ù† Ø¥Ù„Ù‰ {branch_target}",
                        ),
                    )
                else:
                    if wh > 0 and pkg > 0:
                        self.ensure_package_open(wh, pkg)
                    stock_id = self.add_or_update_stock_row(
                        line["item_type"], line["school"], line["color"],
                        line["size"], wh, pkg, price, qty, has_badge,
                    )
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
            self._record_sync_event_or_raise("SALE_RETURNED", {
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
            raise ValueError("Ø§Ù„ÙƒÙ…ÙŠØ© ÙŠØ¬Ø¨ Ø£Ù† ØªÙƒÙˆÙ† Ø£ÙƒØ¨Ø± Ù…Ù† ØµÙØ±")
        with self.conn:
            cur = self.conn.execute("SELECT * FROM stocks WHERE id=?", (int(stock_id),))
            s = cur.fetchone()
            if not s:
                raise ValueError("ØµÙ Ø§Ù„Ù…Ø®Ø²ÙˆÙ† ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯")
            if int(s["count"]) < qty:
                raise ValueError(f"Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…ØªØ§Ø­Ø© ({s['count']}) Ø£Ù‚Ù„ Ù…Ù† Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ ({qty})")
            src_wh = int(s["warehouse_no"])
            src_pkg = int(s["package_no"])
            if src_wh == dest_warehouse_no and src_pkg == dest_package_no:
                raise ValueError("Ø§Ù„Ù…ØµØ¯Ø± ÙˆØ§Ù„ÙˆØ¬Ù‡Ø© Ù…ØªØ·Ø§Ø¨Ù‚Ø§Ù†")
            self.ensure_package_open(dest_warehouse_no, dest_package_no)
            self.conn.execute(
                "UPDATE stocks SET count = count - ? WHERE id = ?",
                (qty, int(stock_id)))
            dest_stock_id = self.add_or_update_stock_row(
                s["item_type"], s["school"], s["color"], s["size"],
                int(dest_warehouse_no), int(dest_package_no),
                float(s["unit_price"]), qty, int(s["has_badge"] or 0),
            )
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
        return s[:22] + "â€¦" if len(s) > 26 else s

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
            sql += " AND status = 'Ù…Ø¹Ù„Ù‚'"
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

        Multiple reservation lines (same or different Ù…Ø¹Ø±Ù‘Ù) that share the same normalized
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
            sql += " AND status = 'Ù…Ø¹Ù„Ù‚'"
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
                "actual_received": 0.0,
                "stock_qty": 0,
                "stock_value": 0.0,
                "cycle_gap": 0.0,
                "actual_gap": 0.0,
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

        # 2b) Manual actual cash received at warehouse
        for dev in approved:
            out[dev]["actual_received"] = self.sum_branch_cash_receipts(
                branch_device=dev,
                date_from=date_from,
                date_to=date_to,
            )

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
            b["actual_gap"] = float(b["shipment_value"]) - float(b["actual_received"]) - float(b["stock_value"])

        rows = list(out.values())
        rows.sort(key=lambda r: str(r.get("branch_name") or ""))
        return rows

    def add_branch_cash_receipt(
        self,
        branch_device: str,
        amount: float,
        *,
        received_at: Optional[str] = None,
        note: str = "",
    ) -> int:
        dev = canonical_branch_device_name(branch_device, DEFAULT_BRANCH_POS_NAMES) or str(branch_device or "").strip()
        if not dev:
            raise ValueError("Ø§Ø®ØªØ± Ø§Ù„ÙØ±Ø¹ Ø£ÙˆÙ„Ø§Ù‹.")
        try:
            amt = float(amount)
        except (TypeError, ValueError):
            raise ValueError("Ø§Ù„Ù…Ø¨Ù„Øº ØºÙŠØ± ØµØ§Ù„Ø­.")
        if amt < 0:
            raise ValueError("Ø§Ù„Ù…Ø¨Ù„Øº ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ† ØµÙØ±Ø§Ù‹ Ø£Ùˆ Ø£ÙƒØ¨Ø±.")
        dt_txt = str(received_at or "").strip() or now_iso()[:10]
        with self.conn:
            cur = self.conn.execute(
                """
                INSERT INTO branch_cash_receipts(branch_device, amount, received_at, note, created_at)
                VALUES (?, ?, ?, ?, ?)
                """,
                (dev, amt, dt_txt[:10], str(note or "").strip(), now_iso()),
            )
            return int(cur.lastrowid)

    def delete_branch_cash_receipt(self, receipt_id: int) -> None:
        with self.conn:
            self.conn.execute("DELETE FROM branch_cash_receipts WHERE id = ?", (int(receipt_id),))

    def list_branch_cash_receipts(
        self,
        branch_device: Optional[str] = None,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        where = ["1=1"]
        args: List[Any] = []
        dev = canonical_branch_device_name(branch_device, DEFAULT_BRANCH_POS_NAMES) if branch_device else None
        if dev:
            where.append("branch_device = ?")
            args.append(dev)
        if date_from:
            where.append("received_at >= ?")
            args.append(str(date_from)[:10])
        if date_to:
            where.append("received_at <= ?")
            args.append(str(date_to)[:10])
        cur = self.conn.execute(
            f"""
            SELECT id, branch_device, amount, received_at, note, created_at
            FROM branch_cash_receipts
            WHERE {' AND '.join(where)}
            ORDER BY received_at DESC, id DESC
            """,
            tuple(args),
        )
        return [dict(r) for r in cur.fetchall()]

    def sum_branch_cash_receipts(
        self,
        branch_device: Optional[str] = None,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> float:
        where = ["1=1"]
        args: List[Any] = []
        dev = canonical_branch_device_name(branch_device, DEFAULT_BRANCH_POS_NAMES) if branch_device else None
        if dev:
            where.append("branch_device = ?")
            args.append(dev)
        if date_from:
            where.append("received_at >= ?")
            args.append(str(date_from)[:10])
        if date_to:
            where.append("received_at <= ?")
            args.append(str(date_to)[:10])
        row = self.conn.execute(
            f"SELECT COALESCE(SUM(amount), 0) FROM branch_cash_receipts WHERE {' AND '.join(where)}",
            tuple(args),
        ).fetchone()
        return float((row[0] if row else 0.0) or 0.0)

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
            raise ValueError("Ø§Ù„Ø¹Ù†ØµØ± Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯.")
        if (item.get("status") or "").upper() != "PENDING":
            raise ValueError("ØªÙ…Øª Ù…Ø¹Ø§Ù„Ø¬Ø© Ù‡Ø°Ø§ Ø§Ù„Ø¹Ù†ØµØ± Ø¨Ø§Ù„ÙØ¹Ù„.")

        w = int(warehouse_no)
        p = int(package_no)
        if w < 1 or p < 1:
            raise ValueError("Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ© ÙŠØ¬Ø¨ Ø£Ù† ÙŠÙƒÙˆÙ†Ø§ 1 Ø£Ùˆ Ø£ÙƒØ¨Ø±.")

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
            raise ValueError("Ø§Ù„Ø¹Ù†ØµØ± Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯.")
        if (item.get("status") or "").upper() != "PENDING":
            raise ValueError("ØªÙ…Øª Ù…Ø¹Ø§Ù„Ø¬Ø© Ù‡Ø°Ø§ Ø§Ù„Ø¹Ù†ØµØ± Ø¨Ø§Ù„ÙØ¹Ù„.")

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
            raise ValueError("Ø§Ù„Ø¹Ù†ØµØ± Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯.")
        if (item.get("status") or "").upper() != "PENDING":
            raise ValueError("ØªÙ…Øª Ù…Ø¹Ø§Ù„Ø¬Ø© Ù‡Ø°Ø§ Ø§Ù„Ø¹Ù†ØµØ± Ø¨Ø§Ù„ÙØ¹Ù„.")

        target = (target_name or "").strip()
        if not target:
            raise ValueError("Ø§Ø³Ù… Ø§Ù„ÙØ±Ø¹ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ ØºÙŠØ± ØµØ§Ù„Ø­.")
        if target == item.get("source_device"):
            raise ValueError("Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø¥Ø¹Ø§Ø¯Ø© Ø§Ù„ØªÙˆØ¬ÙŠÙ‡ Ø¥Ù„Ù‰ Ù†ÙØ³ Ø§Ù„ÙØ±Ø¹ Ø§Ù„Ù…Ø±Ø³Ù„.")

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
            raise ValueError("Ø§Ø³Ù… Ø§Ù„ÙØ±Ø¹ Ø§Ù„Ù…Ø³ØªÙ‡Ø¯Ù ØºÙŠØ± ØµØ§Ù„Ø­.")

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
            # Legacy outbox schema (no target_scope column) â€” fall back
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

    def get_next_shipment_receipt_review(self) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            """
            SELECT id, sync_event_uuid, shipment_uuid, source_device, payload_json,
                   has_diff, note, created_at
            FROM shipment_receipt_reviews
            WHERE status='PENDING' AND shown_at IS NULL
            ORDER BY id ASC
            LIMIT 1
            """
        ).fetchone()
        return dict(row) if row else None

    def mark_shipment_receipt_review_shown(self, review_id: int) -> None:
        self.conn.execute(
            "UPDATE shipment_receipt_reviews SET shown_at=? WHERE id=?",
            (now_iso(), int(review_id)),
        )

    def list_branch_wrong_bill_counts(self) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            """
            SELECT source_device, COUNT(*) AS wrong_reports
            FROM shipment_receipt_reviews
            WHERE has_diff=1
            GROUP BY source_device
            ORDER BY wrong_reports DESC, source_device ASC
            """
        )
        return [dict(r) for r in cur.fetchall()]

    def get_branch_wrong_bill_count(self, source_device: str) -> int:
        row = self.conn.execute(
            """
            SELECT COUNT(*)
            FROM shipment_receipt_reviews
            WHERE has_diff=1 AND LOWER(TRIM(source_device))=LOWER(TRIM(?))
            """,
            (str(source_device or "").strip(),),
        ).fetchone()
        return int((row[0] if row else 0) or 0)

    def decide_shipment_receipt_review(self, review_id: int, accept: bool, note: str = "") -> Dict[str, Any]:
        row = self.conn.execute(
            """
            SELECT id, shipment_uuid, source_device, payload_json, status
            FROM shipment_receipt_reviews
            WHERE id=?
            """,
            (int(review_id),),
        ).fetchone()
        if not row:
            raise ValueError("Ù…Ø±Ø§Ø¬Ø¹Ø© Ø§Ù„Ø´Ø­Ù†Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø©.")
        if str(row["status"]) != "PENDING":
            raise ValueError("ØªÙ… Ø§ØªØ®Ø§Ø° Ù‚Ø±Ø§Ø± Ù…Ø³Ø¨Ù‚Ø§Ù‹ Ù„Ù‡Ø°Ù‡ Ø§Ù„Ù…Ø±Ø§Ø¬Ø¹Ø©.")

        lines: List[Dict[str, Any]] = []
        try:
            parsed = json.loads(row["payload_json"] or "[]")
            if isinstance(parsed, list):
                lines = [x for x in parsed if isinstance(x, dict)]
        except Exception:
            lines = []
        applied = 0
        if accept:
            needed_by_spec: Dict[Tuple[str, str, str, str], int] = {}
            for ln in lines:
                diff = int(ln.get("diff_qty") or 0)
                if diff <= 0:
                    continue
                it = str(ln.get("item_type") or "").strip()
                sc = str(ln.get("school") or "").strip()
                cl = str(ln.get("color") or "").strip()
                sz = str(ln.get("size") or "").strip()
                if not (it and sc and cl and sz):
                    continue
                key = (it, sc, cl, sz)
                needed_by_spec[key] = needed_by_spec.get(key, 0) + int(diff)
            for (it, sc, cl, sz), needed_qty in needed_by_spec.items():
                available = self.conn.execute(
                    """
                    SELECT COALESCE(SUM(count),0)
                    FROM stocks
                    WHERE count>0
                      AND LOWER(TRIM(item_type))=LOWER(TRIM(?))
                      AND LOWER(TRIM(school))=LOWER(TRIM(?))
                      AND LOWER(TRIM(color))=LOWER(TRIM(?))
                      AND LOWER(TRIM(size))=LOWER(TRIM(?))
                    """,
                    (it, sc, cl, sz),
                ).fetchone()[0] or 0
                if int(available) < int(needed_qty):
                    raise ValueError("Cannot accept shipment review: warehouse stock is not enough for the requested correction")

        with self.conn:
            if accept:
                for ln in lines:
                    diff = int(ln.get("diff_qty") or 0)
                    if diff == 0:
                        continue
                    it = str(ln.get("item_type") or "").strip()
                    sc = str(ln.get("school") or "").strip()
                    cl = str(ln.get("color") or "").strip()
                    sz = str(ln.get("size") or "").strip()
                    price = float(ln.get("unit_price") or 0.0)
                    if not (it and sc and cl and sz):
                        continue

                    if diff < 0:
                        qty = abs(diff)
                        stock_id = self.add_or_update_stock_row(
                            it, sc, cl, sz, 1, 1, price, qty, 0,
                        )
                        self.conn.execute(
                            """INSERT INTO movements
                               (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                               VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",
                            (
                                now_iso(), "ADJUST_IN", int(stock_id), qty,
                                f"Ù‚Ø¨ÙˆÙ„ ÙØ±Ù‚ Ø§Ø³ØªÙ„Ø§Ù… Ø´Ø­Ù†Ø© #{str(row['shipment_uuid'])[:8]} Ù…Ù† {row['source_device']}",
                                None, it, sc, cl, sz, 1, 1, price,
                            ),
                        )
                        applied += 1
                    else:
                        need = int(diff)
                        stocks = self.conn.execute(
                            """
                            SELECT id,count,warehouse_no,package_no,unit_price
                            FROM stocks
                            WHERE count>0
                              AND LOWER(TRIM(item_type))=LOWER(TRIM(?))
                              AND LOWER(TRIM(school))=LOWER(TRIM(?))
                              AND LOWER(TRIM(color))=LOWER(TRIM(?))
                              AND LOWER(TRIM(size))=LOWER(TRIM(?))
                            ORDER BY id ASC
                            """,
                            (it, sc, cl, sz),
                        ).fetchall()
                        for s in stocks:
                            if need <= 0:
                                break
                            take = min(need, int(s["count"] or 0))
                            if take <= 0:
                                continue
                            self.conn.execute("UPDATE stocks SET count=count-? WHERE id=?", (take, int(s["id"])))
                            self.conn.execute(
                                """INSERT INTO movements
                                   (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,warehouse_no,package_no,unit_price,has_badge)
                                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",
                                (
                                    now_iso(), "ADJUST_OUT", int(s["id"]), int(take),
                                    f"Ù‚Ø¨ÙˆÙ„ ÙØ±Ù‚ Ø§Ø³ØªÙ„Ø§Ù… Ø´Ø­Ù†Ø© #{str(row['shipment_uuid'])[:8]} Ù…Ù† {row['source_device']}",
                                    None, it, sc, cl, sz,
                                    int(s["warehouse_no"] or 0), int(s["package_no"] or 0),
                                    float(s["unit_price"] or price),
                                ),
                            )
                            need -= int(take)
                            applied += 1
                        if need > 0:
                            raise ValueError("Cannot accept shipment review: correction was not fully applied")

            self.conn.execute(
                """
                UPDATE shipment_receipt_reviews
                SET status=?, decision_note=?, decided_at=?
                WHERE id=?
                """,
                ("ACCEPTED" if accept else "REJECTED", (note or "").strip(), now_iso(), int(review_id)),
            )
        return {"accepted": bool(accept), "adjustments_applied": applied, "shipment_uuid": row["shipment_uuid"]}

    # -------- Stock Audit --------

    def apply_stock_adjustments(self, adjustments: List[Dict[str, Any]],
                                note: str = "Physical count adjustment") -> int:
        """Apply stock count adjustments from a physical audit."""
        count = 0
        applied_events: List[Dict[str, Any]] = []
        with self.conn:
            for adj in adjustments:
                expected = int(adj["expected"])
                actual = int(adj["actual"])
                diff = actual - expected
                if diff == 0:
                    continue
                stock_id_raw = adj.get("stock_id")
                stock_id = parse_int_text(stock_id_raw)
                s = None
                if stock_id is not None:
                    cur = self.conn.execute("SELECT * FROM stocks WHERE id=?", (stock_id,))
                    s = cur.fetchone()
                if not s:
                    if diff <= 0:
                        continue
                    item_type = str(adj.get("item_type") or "").strip()
                    school = str(adj.get("school") or "").strip()
                    color = str(adj.get("color") or "").strip()
                    size = str(adj.get("size") or "").strip()
                    warehouse_no = parse_int_text(adj.get("warehouse_no"))
                    package_no = parse_int_text(adj.get("package_no"))
                    unit_price = parse_float_text(adj.get("unit_price"), 0.0) or 0.0
                    if not (item_type and school and color and size and warehouse_no is not None and package_no is not None):
                        continue
                    self.ensure_package_open(warehouse_no, package_no)
                    stock_id = self.add_or_update_stock_row(
                        item_type,
                        school,
                        color,
                        size,
                        int(warehouse_no),
                        int(package_no),
                        float(unit_price),
                        0,
                        int(adj.get("has_badge") or 0),
                    )
                    s = self.conn.execute("SELECT * FROM stocks WHERE id=?", (stock_id,)).fetchone()
                if not s or stock_id is None:
                    continue
                direction = "ADJUST_IN" if diff > 0 else "ADJUST_OUT"
                abs_diff = abs(diff)
                current_count = int(s["count"] or 0)
                if actual < 0:
                    raise ValueError("Actual stock count cannot be negative")
                if diff < 0 and abs_diff > current_count:
                    raise ValueError("Stock adjustment would make the item count negative")
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

    def create_stock_audit_report(
        self,
        rows: Sequence[Dict[str, Any]],
        *,
        filters: Optional[Dict[str, Any]] = None,
        source: str = "manual",
        note: str = "",
        created_at: Optional[str] = None,
        bucket_key: Optional[str] = None,
    ) -> Optional[int]:
        report_rows = [dict(r) for r in (rows or []) if int(r.get("diff") or 0) != 0]
        if not report_rows:
            return None
        total_diff = sum(int(r.get("diff") or 0) for r in report_rows)
        total_value = sum(float(r.get("diff_value") or 0.0) for r in report_rows)
        with self.conn:
            cur = self.conn.execute(
                """
                INSERT INTO stock_audit_reports(
                    created_at, source, note, bucket_key, filters_json, line_count, total_diff, total_value
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    str(created_at or now_iso()),
                    str(source or "").strip() or "manual",
                    str(note or "").strip() or None,
                    str(bucket_key or "").strip() or None,
                    json.dumps(filters or {}, ensure_ascii=False, default=str),
                    len(report_rows),
                    int(total_diff),
                    float(total_value),
                ),
            )
            report_id = int(cur.lastrowid)
            for r in report_rows:
                self._insert_stock_audit_report_line(report_id, r)
        return report_id

    def _insert_stock_audit_report_line(self, report_id: int, row: Dict[str, Any]) -> None:
        self.conn.execute(
            """
            INSERT INTO stock_audit_report_lines(
                report_id, stock_id, item_type, school, color, size, warehouse_no, package_no,
                expected, actual, diff, unit_price, diff_value
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                int(report_id),
                row.get("id") or row.get("stock_id"),
                row.get("item_type"),
                row.get("school"),
                row.get("color"),
                row.get("size"),
                str(row.get("warehouse_no") or ""),
                str(row.get("package_no") or ""),
                int(row.get("expected") or 0),
                int(row.get("actual") or 0),
                int(row.get("diff") or 0),
                float(row.get("unit_price") or 0.0),
                float(row.get("diff_value") or 0.0),
            ),
        )

    def append_stock_audit_report_bucket(
        self,
        rows: Sequence[Dict[str, Any]],
        *,
        filters: Optional[Dict[str, Any]] = None,
        source: str = "auto-equalization",
        bucket_key: Optional[str] = None,
        note: str = "",
    ) -> Optional[int]:
        report_rows = [dict(r) for r in (rows or []) if int(r.get("diff") or 0) != 0]
        if not report_rows:
            return None
        bucket = str(bucket_key or now_iso()[:13]).strip()
        total_diff = sum(int(r.get("diff") or 0) for r in report_rows)
        total_value = sum(float(r.get("diff_value") or 0.0) for r in report_rows)
        source_clean = str(source or "").strip() or "auto-equalization"
        with self.conn:
            existing = self.conn.execute(
                """
                SELECT id, filters_json
                FROM stock_audit_reports
                WHERE source=? AND bucket_key=?
                ORDER BY id ASC
                LIMIT 1
                """,
                (source_clean, bucket),
            ).fetchone()
            if existing:
                report_id = int(existing["id"])
                for r in report_rows:
                    self._insert_stock_audit_report_line(report_id, r)
                self.conn.execute(
                    """
                    UPDATE stock_audit_reports
                    SET line_count = line_count + ?,
                        total_diff = total_diff + ?,
                        total_value = total_value + ?
                    WHERE id=?
                    """,
                    (len(report_rows), int(total_diff), float(total_value), report_id),
                )
                return report_id
            return self.create_stock_audit_report(
                report_rows,
                filters=filters,
                source=source_clean,
                note=note or "Auto equalization hourly bucket",
                created_at=f"{bucket}:00:00",
                bucket_key=bucket,
            )

    def normalize_auto_stock_audit_reports_by_hour(self) -> int:
        reports = [
            dict(r)
            for r in self.conn.execute(
                """
                SELECT id, created_at, source, bucket_key
                FROM stock_audit_reports
                WHERE source='auto-equalization'
                ORDER BY created_at ASC, id ASC
                """
            ).fetchall()
        ]
        by_hour: Dict[str, List[Dict[str, Any]]] = {}
        for report in reports:
            hour = str(report.get("bucket_key") or report.get("created_at") or "")[:13]
            if hour:
                by_hour.setdefault(hour, []).append(report)

        merged = 0
        with self.conn:
            for hour, group in by_hour.items():
                if len(group) <= 1 and str(group[0].get("bucket_key") or "") == hour:
                    continue
                bucketed = [r for r in group if str(r.get("bucket_key") or "") == hour]
                keep = bucketed[0] if bucketed else group[0]
                keep_id = int(keep["id"])
                ids = [int(r["id"]) for r in group]
                placeholders = ",".join("?" for _ in ids)
                totals = self.conn.execute(
                    f"""
                    SELECT
                        COUNT(*) AS line_count,
                        COALESCE(SUM(diff), 0) AS total_diff,
                        COALESCE(SUM(diff_value), 0) AS total_value
                    FROM stock_audit_report_lines
                    WHERE report_id IN ({placeholders})
                    """,
                    ids,
                ).fetchone()
                self.conn.execute(
                    f"""
                    UPDATE stock_audit_report_lines
                    SET report_id=?
                    WHERE report_id<>? AND report_id IN ({placeholders})
                    """,
                    [keep_id, keep_id, *ids],
                )
                self.conn.execute(
                    """
                    UPDATE stock_audit_reports
                    SET created_at=?,
                        bucket_key=?,
                        note=?,
                        line_count=?,
                        total_diff=?,
                        total_value=?
                    WHERE id=?
                    """,
                    (
                        f"{hour}:00:00",
                        hour,
                        "Auto equalization hourly bucket",
                        int(totals["line_count"] if totals else 0),
                        int(totals["total_diff"] if totals else 0),
                        float(totals["total_value"] if totals else 0.0),
                        keep_id,
                    ),
                )
                drop_ids = [i for i in ids if i != keep_id]
                if drop_ids:
                    self.conn.execute(
                        f"DELETE FROM stock_audit_reports WHERE id IN ({','.join('?' for _ in drop_ids)})",
                        drop_ids,
                    )
                    merged += len(drop_ids)
        return merged

    def list_stock_audit_reports(self, limit: int = 200) -> List[Dict[str, Any]]:
        self.normalize_auto_stock_audit_reports_by_hour()
        cur = self.conn.execute(
            """
            SELECT id, created_at, source, note, line_count, total_diff, total_value
            FROM stock_audit_reports
            ORDER BY created_at DESC, id DESC
            LIMIT ?
            """,
            (int(limit),),
        )
        return [dict(r) for r in cur.fetchall()]

    def backfill_stock_audit_reports_from_movements(self) -> int:
        existing = self.conn.execute(
            "SELECT COUNT(*) AS c FROM stock_audit_reports WHERE source='legacy-movement'"
        ).fetchone()
        if int((existing["c"] if existing else 0) or 0) > 0:
            return 0

        cur = self.conn.execute(
            """
            SELECT
                id, ts, direction, stock_id, qty, note,
                item_type, school, color, size, warehouse_no, package_no, unit_price
            FROM movements
            WHERE direction IN ('ADJUST_IN','ADJUST_OUT')
              AND note LIKE 'Physical count adjustment%'
            ORDER BY ts ASC, id ASC
            """
        )
        movements = [dict(r) for r in cur.fetchall()]
        if not movements:
            return 0

        def _parse_expected_actual(note: Any) -> Tuple[Optional[int], Optional[int]]:
            match = re.search(r"\(expected\s+(-?\d+),\s*actual\s+(-?\d+)\)", str(note or ""))
            if not match:
                return None, None
            return int(match.group(1)), int(match.group(2))

        def _parse_ts(value: Any) -> datetime:
            raw = str(value or "").strip()
            try:
                return datetime.fromisoformat(raw.replace("Z", "+00:00")).replace(tzinfo=None)
            except Exception:
                return datetime.min

        groups: List[List[Dict[str, Any]]] = []
        for row in movements:
            row_dt = _parse_ts(row.get("ts"))
            if not groups:
                groups.append([row])
                continue
            prev = groups[-1][-1]
            prev_dt = _parse_ts(prev.get("ts"))
            same_note = str(row.get("note") or "").split(" (expected ", 1)[0] == str(prev.get("note") or "").split(" (expected ", 1)[0]
            close = abs((row_dt - prev_dt).total_seconds()) <= 10
            if same_note and close:
                groups[-1].append(row)
            else:
                groups.append([row])

        created = 0
        for group in groups:
            report_rows: List[Dict[str, Any]] = []
            for row in group:
                diff = int(row.get("qty") or 0)
                if row.get("direction") == "ADJUST_OUT":
                    diff = -diff
                expected, actual = _parse_expected_actual(row.get("note"))
                if expected is None or actual is None:
                    actual = 0
                    expected = actual - diff
                unit_price = float(row.get("unit_price") or 0.0)
                report_rows.append({
                    "id": row.get("stock_id"),
                    "item_type": row.get("item_type") or "",
                    "school": row.get("school") or "",
                    "color": row.get("color") or "",
                    "size": row.get("size") or "",
                    "warehouse_no": row.get("warehouse_no") or "",
                    "package_no": row.get("package_no") or "",
                    "expected": int(expected),
                    "actual": int(actual),
                    "diff": int(diff),
                    "unit_price": unit_price,
                    "diff_value": float(diff) * unit_price,
                })
            report_id = self.create_stock_audit_report(
                report_rows,
                source="legacy-movement",
                note="Imported from old physical count movements",
                created_at=str(group[0].get("ts") or now_iso()),
            )
            if report_id:
                created += 1
        return created

    def stock_audit_touched_keys(self) -> set[Tuple[str, str, str, str, str, str]]:
        touched: set[Tuple[str, str, str, str, str, str]] = set()

        def _key(row: sqlite3.Row) -> Tuple[str, str, str, str, str, str]:
            return (
                str(row["item_type"] or "").strip().casefold(),
                str(row["school"] or "").strip().casefold(),
                str(row["color"] or "").strip().casefold(),
                _normalize_size_label(row["size"] or "").casefold(),
                str(row["warehouse_no"] or "").strip(),
                str(row["package_no"] or "").strip(),
            )

        try:
            cur = self.conn.execute(
                """
                SELECT item_type, school, color, size, warehouse_no, package_no
                FROM stock_audit_report_lines
                WHERE diff <> 0
                """
            )
            for row in cur.fetchall():
                touched.add(_key(row))
        except sqlite3.Error:
            pass

        try:
            cur = self.conn.execute(
                """
                SELECT item_type, school, color, size, warehouse_no, package_no
                FROM movements
                WHERE direction IN ('ADJUST_IN','ADJUST_OUT')
                  AND note LIKE 'Physical count adjustment%'
                """
            )
            for row in cur.fetchall():
                touched.add(_key(row))
        except sqlite3.Error:
            pass

        return touched

    def get_stock_audit_report(self, report_id: int) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]]]:
        report = self.conn.execute(
            """
            SELECT id, created_at, source, note, filters_json, line_count, total_diff, total_value
            FROM stock_audit_reports
            WHERE id=?
            """,
            (int(report_id),),
        ).fetchone()
        if report is None:
            return None, []
        cur = self.conn.execute(
            """
            SELECT
                stock_id AS id, item_type, school, color, size, warehouse_no, package_no,
                expected, actual, diff, unit_price, diff_value
            FROM stock_audit_report_lines
            WHERE report_id=?
            ORDER BY id ASC
            """,
            (int(report_id),),
        )
        return dict(report), [dict(r) for r in cur.fetchall()]

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
                    ("âœ“" if int(r.get("has_badge") or 0) else ""),
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
        branch_filter_raw = str(filters.get("branch_device") or "").strip()
        branch_filter = ""
        if branch_filter_raw and branch_filter_raw not in ("*", "all", "ALL"):
            branch_filter = canonical_branch_device_name(branch_filter_raw, DEFAULT_BRANCH_POS_NAMES) or ""
        has_date_filter = bool((filters.get("date_from") or "").strip() or (filters.get("date_to") or "").strip())
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
                    "pos_sold_qty": 0,
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

            selected_sources = []
            if branch_filter:
                row = latest_source_by_branch.get(branch_filter)
                if row:
                    selected_sources = [row[0]]
            else:
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
            sql = """
                SELECT item_type, school, color, size, COALESCE(SUM(qty),0)
                FROM pos_reservations_mirror
                WHERE status = 'Ù…Ø¹Ù„Ù‚'
            """
            args: List[Any] = []
            if branch_filter:
                frag, frag_args = self.resolve_pos_mirror_device_sql_filter(branch_filter)
                sql += frag
                args.extend(frag_args)
            sql += " GROUP BY item_type, school, color, size"
            rows = cur.execute(sql, tuple(args)).fetchall()
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
                         OR (
                                (COALESCE(b.bill_type, 'SALE') = 'SALE' OR b.bill_type IS NULL)
                            AND LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))
                            AND LOWER(TRIM(COALESCE(b.customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))
                            AND UPPER(TRIM(COALESCE(b.customer,''))) != 'WAREHOUSE'
                            )
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
                 OR (
                        (COALESCE(b.bill_type, 'SALE') = 'SALE' OR b.bill_type IS NULL)
                    AND LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))
                    AND LOWER(TRIM(COALESCE(b.customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))
                    AND UPPER(TRIM(COALESCE(b.customer,''))) != 'WAREHOUSE'
                    )
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
        branch_customer = branch_customer_label(branch_filter) if branch_filter else ""
        branch_sql = ""
        branch_args: List[Any] = []
        if branch_filter:
            branch_sql = " AND LOWER(TRIM(COALESCE(b.customer,''))) = LOWER(TRIM(?)) "
            branch_args.append(branch_customer)
        rows = cur.execute(
            f"""
            SELECT bi.item_type, bi.school, bi.color, bi.size, COALESCE(SUM(bi.qty),0)
            FROM bill_items bi
            JOIN bills b ON b.id = bi.bill_id
            WHERE COALESCE(b.status,'CONFIRMED') = 'CONFIRMED'
              AND LOWER(TRIM(COALESCE(b.customer,''))) LIKE LOWER(TRIM('ÙØ±Ø¹:%'))
            {branch_sql}
            {d_sql}
            GROUP BY bi.item_type, bi.school, bi.color, bi.size
            """,
            tuple(branch_args + d_args),
        ).fetchall()
        for r in rows:
            key = _to_key(r)
            _bucket(key)["branch_shipped_qty"] = int(r[4] or 0)

        # POS sold qty from synced POS sale payloads only.
        # Warehouse bills/shipments do not mark an item as sold for this view.
        try:
            sale_sql = """
                SELECT event_type, source_device, payload_json
                FROM sync_inbox
                WHERE event_type IN ('SALE_CREATED','SALE_RETURNED','SALE_VOIDED','SALE_EXCHANGED')
                  AND COALESCE(apply_status,'ok') != 'error'
            """
            sale_args: List[Any] = []
            d_sql, d_args = _date_clause("COALESCE(apply_at, applied_at)")
            sale_sql += d_sql
            sale_args.extend(d_args)
            if branch_filter:
                frag, frag_args = self.resolve_pos_mirror_device_sql_filter(branch_filter)
                sale_sql += frag
                sale_args.extend(frag_args)
            sale_rows = cur.execute(sale_sql, tuple(sale_args)).fetchall()
        except sqlite3.OperationalError:
            sale_rows = []

        def _apply_pos_sale_lines(lines: Any, sign: int) -> None:
            if not isinstance(lines, list):
                return
            for ln in lines:
                if not isinstance(ln, dict):
                    continue
                try:
                    qty = int(ln.get("qty") or 0)
                except Exception:
                    qty = 0
                if qty <= 0:
                    continue
                key = (
                    _normalize_spec_label(ln.get("item_type")),
                    _normalize_spec_label(ln.get("school")),
                    _normalize_spec_label(ln.get("color")),
                    _normalize_size_label(str(ln.get("size") or "")),
                )
                if not all(key):
                    continue
                _bucket(key)["pos_sold_qty"] += int(sign) * qty

        for rr in sale_rows:
            try:
                payload = json.loads(rr["payload_json"] or "{}")
            except Exception:
                payload = {}
            if not isinstance(payload, dict):
                continue
            event_type = str(rr["event_type"] or "").strip()
            if event_type == "SALE_CREATED":
                _apply_pos_sale_lines(payload.get("items") or [], +1)
            elif event_type in ("SALE_RETURNED", "SALE_VOIDED"):
                _apply_pos_sale_lines(payload.get("lines") or payload.get("items") or [], -1)
            elif event_type == "SALE_EXCHANGED":
                _apply_pos_sale_lines(payload.get("take_lines") or [], +1)
                _apply_pos_sale_lines(payload.get("return_lines") or [], -1)

        cur.close()

        out: List[Dict[str, Any]] = []
        for key, m in metrics.items():
            if not _spec_match(key):
                continue
            branch_sold = max(0, int(m["pos_sold_qty"]))
            sold_total = max(0, int(m["pos_sold_qty"]))
            remaining_total = int(m["warehouse_qty"]) + int(m["branch_qty"])
            remaining_for_request = int(m["branch_qty"]) if branch_filter else remaining_total
            requested_qty = max(0, sold_total - remaining_for_request)
            incoming_qty = int(m["branch_shipped_qty"]) if branch_filter else int(m["incoming_qty"])
            if branch_filter and not (
                incoming_qty
                or branch_sold
                or int(m["branch_qty"])
                or int(m["reserved_qty"])
                or requested_qty
            ):
                # When a branch is selected, do not show unrelated warehouse
                # products just because the warehouse has stock or old incoming
                # movements for them.
                continue
            if has_date_filter and not (
                incoming_qty
                or int(m["pos_sold_qty"])
                or int(m["reserved_qty"])
            ):
                # Date filters should narrow the visible product rows to items
                # that actually moved in the selected period.
                continue
            out.append(
                {
                    "item_type": key[0],
                    "school": key[1],
                    "color": key[2],
                    "size": key[3],
                    "incoming_qty": incoming_qty,
                    "sold_branch_qty": int(branch_sold),
                    "sold_warehouse_qty": int(m["sold_warehouse_qty"]),
                    "sold_total_qty": int(sold_total),
                    "reserved_qty": int(m["reserved_qty"]),
                    "remaining_branch_qty": int(m["branch_qty"]),
                    "remaining_warehouse_qty": int(m["warehouse_qty"]),
                    "remaining_total_qty": int(remaining_total),
                    "requested_qty": int(requested_qty),
                }
            )

        out.sort(
            key=lambda r: (
                warehouse_item_priority(r.get("item_type")),
                (r.get("item_type") or "").lower(),
                (r.get("school") or "").lower(),
                (r.get("color") or "").lower(),
                warehouse_size_sort_key(r.get("size")),
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
    - Button â€œâ–¼â€ toggles popup; Entry keeps focus.
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
        self.btn = ttk.Button(row, text="â–¼", width=2, takefocus=0, command=self._on_caret_click)
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
    """Entry + 'ðŸ“…' button; popup calendar; returns YYYY-MM-DD; empty allowed."""
    def __init__(self, master, label: str):
        super().__init__(master)
        ttk.Label(self, text=label).pack(side=tk.TOP, anchor="w")
        row = ttk.Frame(self); row.pack(fill=tk.X)
        self.var = tk.StringVar()
        self.entry = ttk.Entry(row, textvariable=self.var, width=14)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn = ttk.Button(row, text="ðŸ“…", width=3, command=self._open)
        self.btn.pack(side=tk.LEFT, padx=(4,0))
        self._popup: Optional[tk.Toplevel] = None
        self.entry.bind("<Return>", lambda e: self._validate())
        self.entry.bind("<Escape>", lambda e: self._close())
        self.winfo_toplevel().bind_all("<Button-1>", self._handle_global_click, add="+")

    def _open(self):
        if self._popup and self._popup.winfo_exists():
            self._close()
            return
        tp = tk.Toplevel(self)
        self._popup = tp
        tp.wm_overrideredirect(True)
        tp.attributes("-topmost", True)
        x = self.entry.winfo_rootx(); y = self.entry.winfo_rooty() + self.entry.winfo_height()
        tp.geometry(f"+{x}+{y}")
        tp.bind("<Escape>", lambda e: self._close())


        frm = ttk.Frame(tp, padding=6, borderwidth=1, relief="solid"); frm.pack()
        today = date.today()
        try:
            base = datetime.strptime(self.var.get(), "%Y-%m-%d").date()
        except Exception:
            base = today
        self._year = tk.IntVar(value=base.year)
        self._month = tk.IntVar(value=base.month)

        hdr = ttk.Frame(frm); hdr.pack(fill=tk.X)
        ttk.Button(hdr, text="â—€", width=2, command=self._prev_month).pack(side=tk.LEFT)
        self._title = ttk.Label(hdr, font=("", 10, "bold")); self._title.pack(side=tk.LEFT, padx=6)
        ttk.Button(hdr, text="â–¶", width=2, command=self._next_month).pack(side=tk.RIGHT)

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
            messagebox.showerror("ØµÙŠØºØ© Ø§Ù„ØªØ§Ø±ÙŠØ®", "Ø§Ù„Ø±Ø¬Ø§Ø¡ Ø¥Ø¯Ø®Ø§Ù„ Ø§Ù„ØªØ§Ø±ÙŠØ® Ø¨ØµÙŠØºØ© YYYY-MM-DD Ø£Ùˆ Ø§Ø³ØªØ®Ø¯Ù… Ø§Ù„ØªÙ‚ÙˆÙŠÙ….", parent=self)
            self.entry.focus_set()

    def _handle_global_click(self, event):
        popup = self._popup
        if not popup or not popup.winfo_exists():
            return
        widget = getattr(event, "widget", None)
        if widget is None:
            self._close()
            return
        if self._is_descendant(widget, self) or self._is_descendant(widget, popup):
            return
        self._close()

    def _is_descendant(self, widget, ancestor) -> bool:
        current = widget
        while current is not None:
            if current == ancestor:
                return True
            parent_name = current.winfo_parent()
            if not parent_name:
                break
            try:
                current = current.nametowidget(parent_name)
            except Exception:
                break
        return False

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
            ttk.Label(cell, text="Ø§Ù„ÙƒÙ…ÙŠØ©", font=("", 8)).pack()
            v_qty = tk.StringVar()
            ttk.Entry(cell, textvariable=v_qty, width=6).pack()

            # price
            ttk.Label(cell, text="Ø§Ù„Ø³Ø¹Ø±", font=("", 8)).pack()
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
                continue  # â† EMPTY instead of zero

            try:
                qty = int(qty_txt)
            except Exception:
                continue

            if qty <= 0:
                continue  # â† ZERO NOT PRINTED

            price_txt = v_price.get().strip()
            try:
                price = parse_float_text(price_txt)
            except Exception:
                price = None

            out.append({
                "size": size_label,
                "qty": qty,
                "price": price
            })

        return out

    def set_price_for_size(self, size_label: str, price, force: bool = False):
        """Set the price entry for a specific size (only if currently empty)."""
        for sz, (v_qty, v_price, v_label, is_custom) in self._vars.items():
            actual_label = v_label.get().strip() if is_custom else sz
            if actual_label == size_label and (force or not v_price.get().strip()):
                try:
                    v_price.set(f"{format_money(float(price))}")
                except Exception:
                    pass

    def set_price_for_all(self, price, force: bool = False):
        """Set the price entry for ALL sizes (only where currently empty)."""
        for _, (v_qty, v_price, v_label, is_custom) in self._vars.items():
            if force or not v_price.get().strip():
                try:
                    v_price.set(f"{format_money(float(price))}")
                except Exception:
                    pass

# ------------------- Scroll Wheel Utility -------------------

def _bind_mousewheel(widget, scrollable=None):
    """
    Enable hover-to-scroll on any scrollable tkinter widget.
    widget    â€“ the visual element the user hovers over
    scrollable â€“ the target that receives yview_scroll (defaults to widget)
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
    # POS-like light palette for visual parity
    "BG":       "#f8fafc",
    "SURFACE":  "#e2e8f0",
    "SURFACE2": "#f1f5f9",
    "BORDER":   "#cbd5e1",
    "ACCENT":   "#2563eb",
    "ACCENT_H": "#1d4ed8",
    "BRAND":    "#2563eb",
    "BRAND_H":  "#1d4ed8",
    "BRAND_L":  "#bfdbfe",
    "TEXT":     "#0f172a",
    "TEXT_SEC": "#475569",
    "TEXT_DIM": "#94a3b8",
    "OK":       "#16a34a",
    "OK_H":     "#15803d",
    "OK_L":     "#ecfdf5",
    "WARN":     "#f59e0b",
    "WARN_H":   "#d97706",
    "WARN_L":   "#fffbeb",
    "DANGER":   "#dc2626",
    "DANGER_H": "#b91c1c",
    "DANGER_L": "#fef2f2",
    "ROW_EVEN": "#ffffff",
    "ROW_ODD":  "#f1f5f9",
    "SEL_BG":   "#bfdbfe",
    "SEL_FG":   "#0f172a",
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


def apply_zebra_tags(tree, skip_tags: Optional[set] = None):
    """Apply alternating row colors to a treeview."""
    skip_tags = skip_tags or set()
    tree.tag_configure("oddrow", background=_UI["ROW_ODD"])
    tree.tag_configure("evenrow", background=_UI["ROW_EVEN"])
    for i, item in enumerate(tree.get_children("")):
        existing = tuple(tree.item(item, "tags") or ())
        if any(t in skip_tags for t in existing):
            continue
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

    menu.add_command(label="Ù†Ø³Ø® Ø§Ù„Ø®Ù„ÙŠØ©", command=_copy_cell)
    menu.add_command(label="Ù†Ø³Ø® Ø§Ù„ØµÙ", command=_copy_row)

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
        self._income_r1 = tk.StringVar(value=DEFAULT_NUMERIC_SIZE_RANGE_LABEL)
        self._income_r2 = tk.StringVar()
        self._income_has_alpha = tk.BooleanVar(value=False)
        self._income_view_mode = tk.StringVar(value="undefined")
        self._defined_school: Optional[str] = None
        self._defined_item_type: Optional[str] = None
        self._defined_color: Optional[str] = None
        self._defined_size: Optional[str] = None
        self._defined_has_badge = tk.BooleanVar(value=False)
        self._defined_qty_var = tk.StringVar(value="1")
        self._defined_price_var = tk.StringVar(value="")
        self._defined_path_var = tk.StringVar(value="")
        self._income_price_profile_var = tk.StringVar(value="")
        self._defined_price_profile_var = tk.StringVar(value="")
        self._income_price_profile_map: Dict[str, int] = {}
        self._defined_price_profile_map: Dict[str, int] = {}

        self._build()

    def _build(self):
        ttk.Label(self, text="ÙˆØ§Ø±Ø¯ (Ø¥Ø¶Ø§ÙØ© Ø£ØµÙ†Ø§Ù Ø¬Ø¯ÙŠØ¯Ø©)", font=_FONTS["h2"]).pack(anchor="w", pady=(0, 4))

        pkg_frame = ttk.LabelFrame(self, text="Ø­Ø§ÙˆÙŠØ© (Ø­Ø³Ø¨ Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©)")
        pkg_frame.pack(fill=tk.X, pady=(0, 4))

        self.wh  = LabeledStaticCombo(
            pkg_frame, "Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù†",
            values=["", *WAREHOUSE_NUMBER_DISPLAY_VALUES],
            value_map=WAREHOUSE_NUMBER_LABELS,
        )
        self.pkg = LabeledEntry(pkg_frame, "Ø±Ù‚Ù… Ø§Ù„Ø¹Ø¨ÙˆØ©")
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
        ttk.Label(stat_box, text="Ø§Ù„Ø­Ø§Ù„Ø©:").pack(side=tk.LEFT)
        ttk.Label(stat_box, textvariable=self.pkg_status_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(stat_box, text="Ø¹Ø¯Ø¯ ØµÙÙˆÙ Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø¯Ø§Ø®Ù„ Ø§Ù„Ø¹Ø¨ÙˆØ©:").pack(side=tk.LEFT)
        ttk.Label(stat_box, textvariable=self.pkg_count_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(4, 0))

        pkg_frame.columnconfigure(0, weight=1)
        pkg_frame.columnconfigure(1, weight=1)

        view_bar = ttk.Frame(self)
        view_bar.pack(fill=tk.X, pady=(0, 6))
        self._btn_undefined_view = ttk.Button(
            view_bar,
            text="Ø£ØµÙ†Ø§Ù ØºÙŠØ± Ù…Ø¹Ø±ÙØ©",
            command=lambda: self._switch_income_view("undefined"),
        )
        self._btn_undefined_view.pack(side=tk.LEFT, padx=(0, 6))
        self._btn_defined_view = ttk.Button(
            view_bar,
            text="Ø£ØµÙ†Ø§Ù Ù…Ø¹Ø±ÙØ©",
            command=lambda: self._switch_income_view("defined"),
        )
        self._btn_defined_view.pack(side=tk.LEFT)

        self._income_views = ttk.Frame(self)
        self._income_views.pack(fill=tk.BOTH, expand=True)
        self._undefined_view = ttk.Frame(self._income_views)
        self._defined_view = ttk.Frame(self._income_views)

        grid = ttk.LabelFrame(self._undefined_view, text="Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„ØµÙ†Ù")
        grid.pack(fill=tk.BOTH, expand=True)

        # Specs comboboxes (no single 'size' field now)
        self.item_type = LabeledCombobox(grid, "Ø§Ù„Ù†ÙˆØ¹", self.db, "item_type")
        self.school    = LabeledCombobox(grid, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self.color     = LabeledCombobox(grid, "Ø§Ù„Ù„ÙˆÙ†", self.db, "color")
        self._income_price_profile_combo: Optional[ttk.Combobox] = None
        # ---------- Income size ranges (single compact row) ----------
        ranges_box = ttk.Frame(grid)
        ranges_box.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(2, 2))

        ttk.Label(ranges_box, text="Ø§Ù„Ù†Ø·Ø§Ù‚ Ø§Ù„Ø£ÙˆÙ„").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            ranges_box,
            textvariable=self._income_r1,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=14
        ).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(ranges_box, text="Ø§Ù„Ù†Ø·Ø§Ù‚ Ø§Ù„Ø«Ø§Ù†ÙŠ").pack(side=tk.LEFT, padx=(0, 2))
        ttk.Combobox(
            ranges_box,
            textvariable=self._income_r2,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=14
        ).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Checkbutton(
            ranges_box,
            text="Ù…Ù‚Ø§Ø³Ø§Øª Ø­Ø±ÙÙŠØ© (S â†’ 5XL)",
            variable=self._income_has_alpha
        ).pack(side=tk.LEFT)

        # --- Cascading filter suppliers ---
        self.item_type.set_supplier(lambda: self.db.get_distinct_filtered(
            "item_type", self._income_constraints("item_type")) or [])
        self.school.set_supplier(lambda: self.db.get_distinct("school") or [])
        self.color.set_supplier(lambda: self.db.get_distinct_filtered(
            "color", {"school": (self.school.get() or "").strip()} if (self.school.get() or "").strip() else {}) or [])

        # Bind to cascade + auto-load size profile + fill prices when specs change
        def _on_spec_change(e=None):
            self._on_income_filter_changed()
            self._auto_load_size_profile()
            self._sync_income_price_profile_from_specs()
            self._rebuild_sizes_grid()
            self._auto_fill_price_for_grid()
            self.after_idle(self._auto_fill_price_for_grid)

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

        # NEW: badge checkbox â€” inside the sizes_container so it doesn't overlap
        self.has_badge = tk.BooleanVar(value=False)

        # layout of spec fields + sizes grid
        # Row 0: item_type + school
        # Row 1: color
        # Row 2: ranges_box (already placed above)
        # Row 3: sizes_container (canvas â€” expandable)
        # Row 4: badge
        # Row 5: buttons
        self.item_type.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        self.school.grid(   row=0, column=1, padx=6, pady=4, sticky="ew")
        self.color.grid(    row=1, column=0, padx=6, pady=4, sticky="ew")
        profile_row = ttk.Frame(grid)
        profile_row.grid(row=1, column=1, padx=6, pady=4, sticky="ew")
        ttk.Label(profile_row, text="Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±").pack(anchor="w")
        self._income_price_profile_combo = ttk.Combobox(
            profile_row,
            textvariable=self._income_price_profile_var,
            state="readonly",
            values=[],
        )
        self._income_price_profile_combo.pack(fill=tk.X)
        self._income_price_profile_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._auto_fill_price_for_grid(force=True),
            add="+",
        )

        badge_row = ttk.Frame(grid)
        badge_row.grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 2))
        ttk.Checkbutton(badge_row, text="Ø¨Ø§Ø¯Ø¬", variable=self.has_badge).pack(side=tk.LEFT)

        # ensure columns expand equally
        for c in range(2):
            grid.columnconfigure(c, weight=1)

        # Give the sizes row weight so it expands to fill available space
        grid.rowconfigure(3, weight=1)   # sizes_container canvas

        # Put package buttons INSIDE the 'grid' labelframe
        pkg_btns = ttk.Frame(grid)
        pkg_btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 4))

        # keep the same buttons inside it
        _btn_close_pkg = ttk.Button(pkg_btns, text="Ø¥ØºÙ„Ø§Ù‚ Ø§Ù„Ø¹Ø¨ÙˆØ©", command=self._close_current_package); _btn_close_pkg.pack(side=tk.RIGHT)
        ToolTip(_btn_close_pkg, "Ø¥ØºÙ„Ø§Ù‚ Ø§Ù„Ø¹Ø¨ÙˆØ© Ø§Ù„Ø­Ø§Ù„ÙŠØ© ÙˆÙ…Ù†Ø¹ Ø§Ù„Ø¥Ø¶Ø§ÙØ© Ø¥Ù„ÙŠÙ‡Ø§")
        _btn_add = ttk.Button(pkg_btns, text="Ø¥Ø¶Ø§ÙØ©", command=self._on_add); _btn_add.pack(side=tk.LEFT)
        ToolTip(_btn_add, "Ø¥Ø¶Ø§ÙØ© Ø§Ù„ØµÙ†Ù Ø¨Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ø¥Ù„Ù‰ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†")

        btns = ttk.Frame(self._undefined_view)
        btns.pack(side=tk.BOTTOM, fill=tk.X, pady=(12, 0))
        _btn_reset = ttk.Button(btns, text="ØªÙØ±ÙŠØº (ÙŠØ¨Ù‚ÙŠ Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©)", command=self._on_reset_keep_pkg); _btn_reset.pack(side=tk.LEFT, padx=8)
        ToolTip(_btn_reset, "Ù…Ø³Ø­ Ø§Ù„Ø­Ù‚ÙˆÙ„ Ù…Ø¹ Ø§Ù„Ø¥Ø¨Ù‚Ø§Ø¡ Ø¹Ù„Ù‰ Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ©")
        # removed "Ø¥Ø¶Ø§ÙØ© Ù…Ù‚Ø§Ø³Ø§Øª Ù…ØªØ¹Ø¯Ø¯Ø©â€¦" button (inline grid used instead)

        self._build_defined_income_view(self._defined_view)

        # bindings
        self.wh.cb.bind("<<ComboboxSelected>>", lambda *_: (self._refresh_pkg_hints(), self._refresh_pkg_status()))
        self.pkg.var.trace_add("write", lambda *_: self._refresh_pkg_status())

        # initial
        self._refresh_pkg_hints()
        self._refresh_pkg_status()
        self._sync_income_price_profile_from_specs()
        self._sync_defined_price_profile_from_specs()
        self._switch_income_view("undefined")

    def _switch_income_view(self, mode: str) -> None:
        mode = "defined" if mode == "defined" else "undefined"
        self._income_view_mode.set(mode)
        self._undefined_view.pack_forget()
        self._defined_view.pack_forget()
        if mode == "defined":
            self._defined_view.pack(fill=tk.BOTH, expand=True)
            if not self._defined_school:
                self._defined_render_schools()
        else:
            self._undefined_view.pack(fill=tk.BOTH, expand=True)

        for btn, is_active in (
            (self._btn_undefined_view, mode == "undefined"),
            (self._btn_defined_view, mode == "defined"),
        ):
            try:
                btn.state(["disabled"] if is_active else ["!disabled"])
            except Exception:
                pass

    def _build_defined_income_view(self, parent):
        wrap = ttk.LabelFrame(parent, text="Ø¥Ø¶Ø§ÙØ© Ø¥Ù„Ù‰ ØµÙ†Ù Ù…Ø¹Ø±Ù")
        wrap.pack(fill=tk.BOTH, expand=True)
        wrap.columnconfigure(0, weight=3)
        wrap.columnconfigure(1, weight=2)
        wrap.rowconfigure(3, weight=1)

        self._defined_crumb_var = tk.StringVar(value="Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø©")
        ttk.Label(
            wrap,
            text="Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© Ø«Ù… Ø§Ù„Ù†ÙˆØ¹ Ø«Ù… Ø§Ù„Ù„ÙˆÙ† Ø«Ù… Ø§Ù„Ù…Ù‚Ø§Ø³ Ù„Ø¥Ø¶Ø§ÙØ© ÙƒÙ…ÙŠØ© Ø¹Ù„Ù‰ ØµÙ†Ù Ù…ÙˆØ¬ÙˆØ¯ Ø¨Ø§Ù„ÙØ¹Ù„.",
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
        ttk.Label(wrap, textvariable=self._defined_crumb_var, font=_FONTS["h3"]).grid(
            row=1, column=0, sticky="w", padx=8, pady=(0, 6)
        )

        filters = ttk.LabelFrame(wrap, text="ÙÙ„Ø§ØªØ± Ø³Ø±ÙŠØ¹Ø©")
        filters.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        filters.columnconfigure(0, weight=1)
        filters.columnconfigure(1, weight=1)
        filters.columnconfigure(2, weight=1)

        self._defined_filter_school = LabeledCombobox(filters, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self._defined_filter_item = LabeledCombobox(filters, "Ø§Ù„Ù†ÙˆØ¹", self.db, "item_type")
        self._defined_filter_color = LabeledCombobox(filters, "Ø§Ù„Ù„ÙˆÙ†", self.db, "color")
        self._defined_filter_school.grid(row=0, column=0, padx=6, pady=6, sticky="ew")
        self._defined_filter_item.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
        self._defined_filter_color.grid(row=0, column=2, padx=6, pady=6, sticky="ew")

        self._defined_filter_school.set_supplier(
            lambda: self.db.get_distinct_filtered("school", self._defined_filter_constraints("school")) or []
        )
        self._defined_filter_item.set_supplier(
            lambda: self.db.get_distinct_filtered("item_type", self._defined_filter_constraints("item_type")) or []
        )
        self._defined_filter_color.set_supplier(
            lambda: self.db.get_distinct_filtered("color", self._defined_filter_constraints("color")) or []
        )
        for w in (
            self._defined_filter_school.cb,
            self._defined_filter_item.cb,
            self._defined_filter_color.cb,
        ):
            w.bind("<<ComboboxSelected>>", lambda _e: self._defined_on_filter_changed(), add="+")
            w.bind("<KeyRelease>", lambda _e: self._defined_on_filter_changed(), add="+")

        grid_wrap = ttk.Frame(wrap)
        grid_wrap.grid(row=3, column=0, sticky="nsew", padx=(8, 4), pady=(0, 8))
        grid_wrap.columnconfigure(0, weight=1)
        grid_wrap.rowconfigure(0, weight=1)

        self._defined_canvas = tk.Canvas(grid_wrap, highlightthickness=0)
        defined_scroll = ttk.Scrollbar(grid_wrap, orient="vertical", command=self._defined_canvas.yview)
        self._defined_canvas.configure(yscrollcommand=defined_scroll.set)
        self._defined_canvas.grid(row=0, column=0, sticky="nsew")
        defined_scroll.grid(row=0, column=1, sticky="ns")
        self._defined_grid_host = ttk.Frame(self._defined_canvas)
        self._defined_grid_window = self._defined_canvas.create_window((0, 0), window=self._defined_grid_host, anchor="nw")
        self._defined_grid_host.bind("<Configure>", lambda _e=None: self._defined_sync_canvas())
        self._defined_canvas.bind("<Configure>", lambda _e=None: self._defined_sync_canvas())
        _bind_mousewheel(grid_wrap, self._defined_canvas)

        side = ttk.LabelFrame(wrap, text="Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„Ø¥Ø¶Ø§ÙØ©")
        side.grid(row=3, column=1, sticky="nsew", padx=(4, 8), pady=(0, 8))
        side.columnconfigure(1, weight=1)

        ttk.Label(side, text="Ø§Ù„Ø§Ø®ØªÙŠØ§Ø±").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))
        ttk.Label(side, textvariable=self._defined_path_var, wraplength=260, justify="left").grid(
            row=0, column=1, sticky="ew", padx=8, pady=(8, 4)
        )
        ttk.Label(side, text="Ø§Ù„Ø³Ø¹Ø± Ø§Ù„Ø­Ø§Ù„ÙŠ").grid(row=1, column=0, sticky="w", padx=8, pady=4)
        ttk.Label(side, textvariable=self._defined_price_var).grid(row=1, column=1, sticky="w", padx=8, pady=4)
        ttk.Label(side, text="Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±").grid(row=2, column=0, sticky="w", padx=8, pady=4)
        self._defined_price_profile_combo = ttk.Combobox(
            side,
            textvariable=self._defined_price_profile_var,
            state="readonly",
            values=[],
        )
        self._defined_price_profile_combo.grid(row=2, column=1, sticky="ew", padx=8, pady=4)
        self._defined_price_profile_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._defined_refresh_selected_price(),
            add="+",
        )
        ttk.Label(side, text="Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø¶Ø§ÙØ©").grid(row=3, column=0, sticky="w", padx=8, pady=4)
        ttk.Entry(side, textvariable=self._defined_qty_var, width=12).grid(row=3, column=1, sticky="w", padx=8, pady=4)
        ttk.Checkbutton(side, text="Ø¨Ø§Ø¯Ø¬", variable=self._defined_has_badge).grid(
            row=4, column=0, columnspan=2, sticky="w", padx=8, pady=(4, 8)
        )

        action_bar = ttk.Frame(side)
        action_bar.grid(row=5, column=0, columnspan=2, sticky="ew", padx=8, pady=(4, 4))
        ttk.Button(action_bar, text="Ø¥Ø¶Ø§ÙØ©", command=self._on_add_defined).pack(side=tk.LEFT)
        ttk.Button(action_bar, text="ØªÙØ±ÙŠØº Ø§Ù„Ø§Ø®ØªÙŠØ§Ø±", command=self._defined_reset_selection).pack(side=tk.LEFT, padx=(6, 0))

        pkg_actions = ttk.Frame(side)
        pkg_actions.grid(row=6, column=0, columnspan=2, sticky="ew", padx=8, pady=(8, 8))
        ttk.Button(pkg_actions, text="ØªÙØ±ÙŠØº (ÙŠØ¨Ù‚ÙŠ Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©)", command=self._on_reset_keep_pkg).pack(side=tk.LEFT)
        ttk.Button(pkg_actions, text="Ø¥ØºÙ„Ø§Ù‚ Ø§Ù„Ø¹Ø¨ÙˆØ©", command=self._close_current_package).pack(side=tk.RIGHT)

        self._defined_size_rows: Dict[str, Dict[str, Any]] = {}
        self._defined_update_summary()

    def _defined_sync_canvas(self) -> None:
        try:
            self._defined_canvas.configure(scrollregion=self._defined_canvas.bbox("all"))
            self._defined_canvas.itemconfigure(self._defined_grid_window, width=self._defined_canvas.winfo_width())
        except Exception:
            pass

    def _defined_clear_grid(self) -> None:
        for child in self._defined_grid_host.winfo_children():
            child.destroy()
        try:
            self._defined_canvas.yview_moveto(0)
        except Exception:
            pass

    def _defined_make_buttons(self, values: List[str], command, *, cols: int = 4) -> None:
        self._defined_clear_grid()
        if not values:
            ttk.Label(self._defined_grid_host, text="Ù„Ø§ ØªÙˆØ¬Ø¯ Ø¨ÙŠØ§Ù†Ø§Øª Ù…Ø·Ø§Ø¨Ù‚Ø©.").pack(anchor="w", padx=8, pady=8)
            return
        row_frame = None
        for i, value in enumerate(values):
            if i % cols == 0:
                row_frame = ttk.Frame(self._defined_grid_host)
                row_frame.pack(fill=tk.X, padx=4, pady=2)
            ttk.Button(row_frame, text=value, command=lambda v=value: command(v)).pack(
                side=tk.LEFT, padx=3, pady=3, fill=tk.X, expand=True
            )

    def _defined_add_back_button(self, text: str, command) -> None:
        ttk.Button(self._defined_grid_host, text=text, command=command).pack(anchor="w", padx=8, pady=(6, 8))

    def _defined_filter_constraints(self, exclude_field: str) -> Dict[str, Any]:
        constraints: Dict[str, Any] = {}
        school = (self._defined_filter_school.get() or "").strip()
        item_type = (self._defined_filter_item.get() or "").strip()
        color = (self._defined_filter_color.get() or "").strip()
        if exclude_field != "school" and school:
            constraints["school"] = school
        if exclude_field != "item_type" and item_type:
            constraints["item_type"] = item_type
        if exclude_field != "color" and color:
            constraints["color"] = color
        return constraints

    def _defined_active_filters(self) -> Dict[str, Any]:
        constraints: Dict[str, Any] = {}
        school = (self._defined_filter_school.get() or "").strip()
        item_type = (self._defined_filter_item.get() or "").strip()
        color = (self._defined_filter_color.get() or "").strip()
        if school:
            constraints["school"] = school
        if item_type:
            constraints["item_type"] = item_type
        if color:
            constraints["color"] = color
        return constraints

    def _defined_filter_texts(self) -> Dict[str, str]:
        return {
            "school": (self._defined_filter_school.get() or "").strip(),
            "item_type": (self._defined_filter_item.get() or "").strip(),
            "color": (self._defined_filter_color.get() or "").strip(),
        }

    def _defined_filtered_inventory_rows(
        self,
        *,
        school: Optional[str] = None,
        item_type: Optional[str] = None,
        color: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        rows = self.db.current_inventory({}) or []
        texts = self._defined_filter_texts()
        school_exact = (school or "").strip()
        item_exact = (item_type or "").strip()
        color_exact = (color or "").strip()
        out: List[Dict[str, Any]] = []
        for row in rows:
            row_school = str(row.get("school") or "").strip()
            row_item = str(row.get("item_type") or "").strip()
            row_color = str(row.get("color") or "").strip()
            if not (row_school and row_item and row_color):
                continue
            if school_exact and row_school.casefold() != school_exact.casefold():
                continue
            if item_exact and row_item.casefold() != item_exact.casefold():
                continue
            if color_exact and row_color.casefold() != color_exact.casefold():
                continue
            if texts["school"] and texts["school"].casefold() not in row_school.casefold():
                continue
            if texts["item_type"] and texts["item_type"].casefold() not in row_item.casefold():
                continue
            if texts["color"] and texts["color"].casefold() not in row_color.casefold():
                continue
            out.append(row)
        return out

    @staticmethod
    def _defined_unique_values(rows: List[Dict[str, Any]], field: str) -> List[str]:
        seen = set()
        values: List[str] = []
        for row in rows:
            value = str(row.get(field) or "").strip()
            key = value.casefold()
            if value and key not in seen:
                seen.add(key)
                values.append(value)
        values.sort(key=lambda s: s.casefold())
        return values

    def _defined_refresh_filter_combos(self) -> None:
        for w in (
            self._defined_filter_school,
            self._defined_filter_item,
            self._defined_filter_color,
        ):
            try:
                w.refresh_values()
            except Exception:
                pass

    def _defined_on_filter_changed(self) -> None:
        self._defined_refresh_filter_combos()
        school_filter = (self._defined_filter_school.get() or "").strip()
        item_filter = (self._defined_filter_item.get() or "").strip()
        color_filter = (self._defined_filter_color.get() or "").strip()
        if school_filter and self._defined_school and school_filter.casefold() != self._defined_school.casefold():
            self._defined_school = None
            self._defined_item_type = None
            self._defined_color = None
            self._defined_size = None
        elif item_filter and self._defined_item_type and item_filter.casefold() != self._defined_item_type.casefold():
            self._defined_item_type = None
            self._defined_color = None
            self._defined_size = None
        elif color_filter and self._defined_color and color_filter.casefold() != self._defined_color.casefold():
            self._defined_color = None
            self._defined_size = None
        if self._defined_color:
            self._defined_render_sizes()
        elif self._defined_item_type:
            self._defined_select_item(self._defined_item_type)
        elif self._defined_school:
            self._defined_select_school(self._defined_school)
        else:
            self._defined_render_schools()

    def _defined_update_summary(self) -> None:
        parts = []
        if self._defined_school:
            parts.append(self._defined_school)
        if self._defined_item_type:
            parts.append(self._defined_item_type)
        if self._defined_color:
            parts.append(self._defined_color)
        if self._defined_size:
            parts.append(self._defined_size)
        self._defined_path_var.set(" / ".join(parts) if parts else "Ù„Ù… ÙŠØªÙ… Ø§Ø®ØªÙŠØ§Ø± ØµÙ†Ù Ø¨Ø¹Ø¯")

    def _defined_reset_selection(self) -> None:
        self._defined_school = None
        self._defined_item_type = None
        self._defined_color = None
        self._defined_size = None
        self._defined_filter_school.set("")
        self._defined_filter_item.set("")
        self._defined_filter_color.set("")
        self._defined_qty_var.set("1")
        self._defined_price_var.set("")
        self._defined_price_profile_var.set("")
        self._defined_has_badge.set(False)
        self._defined_size_rows = {}
        self._defined_update_summary()
        self._defined_refresh_filter_combos()
        self._defined_render_schools()

    def _defined_render_schools(self) -> None:
        self._defined_school = None
        self._defined_item_type = None
        self._defined_color = None
        self._defined_size = None
        self._defined_price_var.set("")
        self._defined_price_profile_var.set("")
        self._defined_crumb_var.set("Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø©")
        self._defined_update_summary()
        schools = self._defined_unique_values(self._defined_filtered_inventory_rows(), "school")
        self._defined_make_buttons(schools, self._defined_select_school)

    def _defined_select_school(self, school: str) -> None:
        if not school:
            self._defined_render_schools()
            return
        self._defined_school = school
        self._defined_item_type = None
        self._defined_color = None
        self._defined_size = None
        self._defined_price_var.set("")
        self._defined_crumb_var.set(f"{school} â† Ø§Ø®ØªØ± Ø§Ù„Ù†ÙˆØ¹")
        self._defined_update_summary()
        items = self._defined_unique_values(
            self._defined_filtered_inventory_rows(school=school),
            "item_type",
        )
        self._defined_make_buttons(items, self._defined_select_item)
        self._defined_add_back_button("Ø±Ø¬ÙˆØ¹ Ø¥Ù„Ù‰ Ø§Ù„Ù…Ø¯Ø§Ø±Ø³", self._defined_render_schools)

    def _defined_select_item(self, item_type: str) -> None:
        if not item_type:
            self._defined_select_school(self._defined_school or "")
            return
        self._defined_item_type = item_type
        self._defined_color = None
        self._defined_size = None
        self._defined_price_var.set("")
        self._defined_crumb_var.set(f"{self._defined_school} / {item_type} â† Ø§Ø®ØªØ± Ø§Ù„Ù„ÙˆÙ†")
        self._defined_update_summary()
        colors = self._defined_unique_values(
            self._defined_filtered_inventory_rows(
                school=self._defined_school,
                item_type=item_type,
            ),
            "color",
        )
        self._defined_make_buttons(colors, self._defined_select_color)
        self._defined_add_back_button("Ø±Ø¬ÙˆØ¹ Ø¥Ù„Ù‰ Ø§Ù„Ø£Ù†ÙˆØ§Ø¹", lambda: self._defined_select_school(self._defined_school or ""))

    def _defined_select_color(self, color: str) -> None:
        if not color:
            self._defined_select_item(self._defined_item_type or "")
            return
        self._defined_color = color
        self._defined_size = None
        self._defined_price_var.set("")
        self._defined_crumb_var.set(
            f"{self._defined_school} / {self._defined_item_type} / {color} â† Ø§Ø®ØªØ± Ø§Ù„Ù…Ù‚Ø§Ø³"
        )
        self._defined_update_summary()
        self._defined_render_sizes()

    def _defined_render_sizes(self) -> None:
        school = self._defined_school or ""
        item_type = self._defined_item_type or ""
        color = self._defined_color or ""
        self._sync_defined_price_profile_from_specs()
        size_rows = self._defined_collect_size_rows(school, item_type, color)
        self._defined_size_rows = {str(r.get("size") or "").strip(): r for r in size_rows}
        labels = []
        for row in size_rows:
            size = str(row.get("size") or "").strip()
            if not size:
                continue
            labels.append(f"{size} ({int(row.get('count') or 0)})")
        self._defined_make_buttons(labels, self._defined_select_size)
        self._defined_add_back_button("Ø±Ø¬ÙˆØ¹ Ø¥Ù„Ù‰ Ø§Ù„Ø£Ù„ÙˆØ§Ù†", lambda: self._defined_select_item(self._defined_item_type or ""))

    def _defined_collect_size_rows(self, school: str, item_type: str, color: str) -> List[Dict[str, Any]]:
        sizes: List[str] = []
        profile = self.db.get_size_profile(item_type, school, color)
        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            sizes.extend(merged_numeric_size_labels_from_profile(r1s, r1e, r2s, r2e))
            if has_alpha:
                sizes.extend(ALPHA_SIZES)
        else:
            seen = set()
            try:
                for row in self.db.current_inventory({
                    "school": school,
                    "item_type": item_type,
                    "color": color,
                }):
                    size = str(row.get("size") or "").strip()
                    if size and size not in seen:
                        seen.add(size)
                        sizes.append(size)
            except Exception:
                pass

        rows: List[Dict[str, Any]] = []
        for size in sizes:
            try:
                rows.append(self.db._size_row(school, item_type, color, size))
            except Exception:
                rows.append({"size": size, "count": 0, "last_price": None})
        return rows

    def _defined_select_size(self, label: str) -> None:
        size = label.split(" (", 1)[0].strip()
        self._defined_size = size
        self._defined_update_summary()
        self._defined_refresh_selected_price()
        return
        row = self._defined_size_rows.get(size, {})
        price = row.get("last_price")
        if price is None:
            try:
                price = self.db.get_effective_price(
                    self._defined_item_type or "",
                    self._defined_school or "",
                    self._defined_color or "",
                    size,
                )
            except Exception:
                price = None
        if price is None:
            self._defined_price_var.set("ØºÙŠØ± Ù…Ø¹Ø±ÙˆÙ")
        else:
            try:
                self._defined_price_var.set(f"{format_money(float(price))}")
            except Exception:
                self._defined_price_var.set(str(price))

    def _defined_refresh_selected_price(self) -> None:
        size = self._defined_size or ""
        if not size:
            self._defined_price_var.set("")
            return
        price = None
        try:
            price = self.db.get_effective_price(
                self._defined_item_type or "",
                self._defined_school or "",
                self._defined_color or "",
                size,
                preferred_profile_id=self._selected_defined_price_profile_id(),
            )
        except Exception:
            price = None
        if price is None:
            row = self._defined_size_rows.get(size, {})
            price = row.get("last_price")
        if price is None:
            self._defined_price_var.set("ØºÙŠØ± Ù…Ø¹Ø±ÙˆÙ")
            return
        try:
            self._defined_price_var.set(f"{format_money(float(price))}")
        except Exception:
            self._defined_price_var.set(str(price))

    def _on_add_defined(self) -> None:
        w = self._parse_int_or_none(self.wh.get())
        p = self._parse_int_or_none(self.pkg.get())
        if not (w and p and w >= 1 and p >= 1):
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", "Ø£Ø¯Ø®Ù„ Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ© Ø¨Ø´ÙƒÙ„ ØµØ­ÙŠØ­ (>= 1).")
            return
        if not all([self._defined_school, self._defined_item_type, self._defined_color, self._defined_size]):
            messagebox.showwarning("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", "Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© ÙˆØ§Ù„Ù†ÙˆØ¹ ÙˆØ§Ù„Ù„ÙˆÙ† ÙˆØ§Ù„Ù…Ù‚Ø§Ø³ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        qty = parse_int_text(self._defined_qty_var.get(), 0) or 0
        if qty <= 0:
            messagebox.showwarning("ÙƒÙ…ÙŠØ© ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ø£Ø¯Ø®Ù„ ÙƒÙ…ÙŠØ© Ø£ÙƒØ¨Ø± Ù…Ù† ØµÙØ±.", parent=self)
            return

        price_txt = (self._defined_price_var.get() or "").strip()
        price = None
        if price_txt and price_txt != "ØºÙŠØ± Ù…Ø¹Ø±ÙˆÙ":
            try:
                price = parse_float_text(price_txt)
            except Exception:
                price = None

        try:
            self.db.ensure_package_open(w, p)
            self.db.add_stock(
                item_type=self._defined_item_type or "",
                school=self._defined_school or "",
                color=self._defined_color or "",
                size=self._defined_size or "",
                warehouse_no=w,
                package_no=p,
                unit_price=price,
                count=qty,
                has_badge=1 if self._defined_has_badge.get() else 0,
            )
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¥Ø¶Ø§ÙØ©", str(ex), parent=self)
            return

        self._defined_qty_var.set("1")
        self._defined_has_badge.set(False)
        self._refresh_pkg_hints()
        self._refresh_pkg_status()
        self._defined_render_sizes()
        if self._defined_size:
            self._defined_select_size(self._defined_size)
        show_toast(self, "ØªÙ…Øª Ø¥Ø¶Ø§ÙØ© Ø§Ù„ÙƒÙ…ÙŠØ© Ø¨Ù†Ø¬Ø§Ø­")

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
        # Type can be new, so do not clear it just because no stock rows match yet.
        if sc:
            valid = self.db.get_distinct("school")
            if sc not in valid:
                self.school.set("")
        if cl:
            valid = self.db.get_distinct_filtered(
                "color",
                {"school": sc} if sc else {},
            )
            if cl not in valid:
                self.color.set("")
        # Refresh all combo dropdown lists
        self.item_type.refresh_values()
        self.school.refresh_values()
        self.color.refresh_values()

    def _set_default_income_profile(self) -> None:
        self._income_r1.set(DEFAULT_NUMERIC_SIZE_RANGE_LABEL)
        self._income_r2.set("")
        self._income_has_alpha.set(False)

    def _refresh_price_profile_combo(
        self,
        combo: Optional[ttk.Combobox],
        value_var: tk.StringVar,
        target_map: Dict[str, int],
        preferred_profile_id: Optional[int] = None,
    ) -> None:
        target_map.clear()
        values = [""]
        for profile in self.db.list_price_profiles():
            label = str(profile.get("name") or "").strip()
            if not label:
                continue
            target_map[label] = int(profile["id"])
            values.append(label)
        if combo is not None:
            combo.configure(values=values)

        selected_label = ""
        if preferred_profile_id:
            for label, pid in target_map.items():
                if pid == int(preferred_profile_id):
                    selected_label = label
                    break
        elif (value_var.get() or "").strip() in target_map:
            selected_label = (value_var.get() or "").strip()
        value_var.set(selected_label)

    def _selected_income_price_profile_id(self) -> Optional[int]:
        return self._income_price_profile_map.get((self._income_price_profile_var.get() or "").strip())

    def _selected_defined_price_profile_id(self) -> Optional[int]:
        return self._defined_price_profile_map.get((self._defined_price_profile_var.get() or "").strip())

    def _sync_income_price_profile_from_specs(self) -> None:
        preferred_profile_id = self.db.resolve_price_profile_id(
            (self.item_type.get() or "").strip(),
            (self.school.get() or "").strip(),
            (self.color.get() or "").strip(),
        )
        self._refresh_price_profile_combo(
            self._income_price_profile_combo,
            self._income_price_profile_var,
            self._income_price_profile_map,
            preferred_profile_id=preferred_profile_id,
        )

    def _sync_defined_price_profile_from_specs(self) -> None:
        preferred_profile_id = self.db.resolve_price_profile_id(
            self._defined_item_type or "",
            self._defined_school or "",
            self._defined_color or "",
        )
        self._refresh_price_profile_combo(
            self._defined_price_profile_combo,
            self._defined_price_profile_var,
            self._defined_price_profile_map,
            preferred_profile_id=preferred_profile_id,
        )

    def _current_income_profile_values(self) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int], bool]:
        def _parse_range_label(v: str) -> Tuple[Optional[int], Optional[int]]:
            v = (v or "").strip()
            if not v:
                return None, None
            a, b = v.split("â†’")
            return int(a.strip()), int(b.strip())

        r1s, r1e = _parse_range_label(self._income_r1.get())
        r2s, r2e = _parse_range_label(self._income_r2.get())
        return r1s, r1e, r2s, r2e, bool(self._income_has_alpha.get())

    def _auto_load_size_profile(self):
        """Auto-set range dropdowns from saved size_profile when item+school+color are all set."""
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        if not (it and sc and cl):
            return
        try:
            profile = self.db.get_size_profile(it, sc, cl)
            r1s, r1e, r2s, r2e, has_alpha = profile
            # Set range 1
            if r1s is not None and r1e is not None:
                label1 = f"{r1s} â†’ {r1e}"
                if label1 in NUMERIC_RANGE_LABELS:
                    self._income_r1.set(label1)
                else:
                    self._income_r1.set(DEFAULT_NUMERIC_SIZE_RANGE_LABEL)
            else:
                self._income_r1.set(DEFAULT_NUMERIC_SIZE_RANGE_LABEL)
            # Set range 2
            if r2s is not None and r2e is not None:
                label2 = f"{r2s} â†’ {r2e}"
                if label2 in NUMERIC_RANGE_LABELS:
                    self._income_r2.set(label2)
                else:
                    self._income_r2.set("")
            else:
                self._income_r2.set("")
            # Set alpha
            self._income_has_alpha.set(bool(has_alpha))
        except Exception:
            self._set_default_income_profile()

    def _rebuild_sizes_grid(self):
        try:
            self.sizes_grid.destroy()
        except Exception:
            pass

        numeric_ranges = []

        def _parse_range(label: str):
            if not label:
                return None
            a, b = label.split("â†’")
            return int(a.strip()), int(b.strip())

        r1 = _parse_range(self._income_r1.get())
        r2 = _parse_range(self._income_r2.get())
        sizes = merged_numeric_size_labels_from_profile(
            r1[0] if r1 else None,
            r1[1] if r1 else None,
            r2[0] if r2 else None,
            r2[1] if r2 else None,
        )

        numeric_ranges = []

        # alpha sizes
        if self._income_has_alpha.get():
            sizes.extend(ALPHA_SIZES)

        self.sizes_grid = SizesGrid(self._sizes_inner, sizes)
        self.sizes_grid.pack(fill="both", expand=False)

        self._sizes_inner.update_idletasks()



    def _auto_fill_price_for_grid(self, force: bool = False):
        """Fill empty price cells from profile first, then historical/default prices."""
        if not self.sizes_grid:
            return
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        if not (it and sc and cl):
            return
        try:
            preferred_profile_id = self._selected_income_price_profile_id()
            found_any = False
            for sz in self.sizes_grid.fixed_sizes:
                p = self.db.get_effective_price(
                    it,
                    sc,
                    cl,
                    sz,
                    preferred_profile_id=preferred_profile_id,
                )
                if p is not None:
                    self.sizes_grid.set_price_for_size(sz, p, force=force)
                    found_any = True
            if not found_any and self.sizes_grid.fixed_sizes:
                first = self.sizes_grid.fixed_sizes[0]
                p = self.db.get_effective_price(
                    it,
                    sc,
                    cl,
                    first,
                    preferred_profile_id=preferred_profile_id,
                )
                if p is not None:
                    self.sizes_grid.set_price_for_all(p, force=force)
        except Exception:
            pass

    def _parse_int_or_none(self, s: str) -> Optional[int]:
        s = warehouse_numeric_value(s)
        s = (s or "").strip()
        if not s:
            return None
        try:
            v = parse_int_text(s)
            if v is None:
                return None
            return v
        except Exception:
            return None

    def _refresh_pkg_hints(self):
        w = self._parse_int_or_none(self.wh.get())
        if not w or w < 1:
            self.pkg_hint_var.set("Ø£Ø¯Ø®Ù„ Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† Ø£ÙˆÙ„Ø§Ù‹.")
            return
        try:
            info = self.db.package_numbers_summary(w)
        except Exception as ex:
            self.pkg_hint_var.set(str(ex))
            return
        if info["free"]:
            free_txt = "ØŒ ".join(map(str, info["free"][:20]))
            more = "" if len(info["free"]) <= 20 else " â€¦"
            hint = f"Ø§Ù„Ø£Ø±Ù‚Ø§Ù… Ø§Ù„Ù…ØªØ§Ø­Ø© (ÙØ§Ø±ØºØ©): {free_txt}{more}"
        else:
            hint = f"Ù„Ø§ ØªÙˆØ¬Ø¯ ÙØ±Ø§ØºØ§Øª. Ø§Ù„Ø±Ù‚Ù… Ø§Ù„ØªØ§Ù„ÙŠ: {info['next']}"
        if info.get("open"):
            hint += f" | Ø­Ø§ÙˆÙŠØ§Øª Ù…ÙØªÙˆØ­Ø©: {', '.join(map(str, info['open'][:20]))}"
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
            self.pkg_status_var.set("ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø© (Ø³ÙŠØªÙ… ÙØªØ­Ù‡Ø§ ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹ Ø¹Ù†Ø¯ Ø£ÙˆÙ„ Ø¥Ø¶Ø§ÙØ©)")
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
            messagebox.showwarning("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", "Ø£Ø¯Ø®Ù„ Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ© Ø£ÙˆÙ„Ø§Ù‹.")
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯", f"Ø¥ØºÙ„Ø§Ù‚ Ø§Ù„Ø¹Ø¨ÙˆØ© {p} ÙÙŠ Ø§Ù„Ù…Ø®Ø²Ù† {w} ØŸ\nÙ„Ù† ÙŠÙØ³Ù…Ø­ Ø¨Ø£ÙŠ Ø¥Ø¶Ø§ÙØ§Øª Ù„Ø§Ø­Ù‚Ø§Ù‹."):
            return
        try:
            self.db.close_package(w, p)
            show_toast(self, "ØªÙ… Ø¥ØºÙ„Ø§Ù‚ Ø§Ù„Ø¹Ø¨ÙˆØ© Ø¨Ù†Ø¬Ø§Ø­")
            self._on_reset_keep_pkg()
            self._refresh_pkg_hints()
            try:
                self.item_type.cb.focus_set()
            except Exception:
                pass
        except Exception as ex:
            messagebox.showerror("Ø®Ø·Ø£", str(ex))

    def _on_reset_keep_pkg(self):
        for w in (self.item_type, self.school, self.color):
            w.set("")
        self._set_default_income_profile()

        # clear sizes grid entries and badge
        if self.sizes_grid:
            self.sizes_grid.destroy()
            self.sizes_grid = None

        self._sizes_inner.update_idletasks()

        self.has_badge.set(False)
        self._defined_school = None
        self._defined_item_type = None
        self._defined_color = None
        self._defined_size = None
        self._defined_qty_var.set("1")
        self._defined_price_var.set("")
        self._defined_has_badge.set(False)
        self._defined_size_rows = {}
        self._defined_update_summary()
        try:
            self._defined_render_schools()
        except Exception:
            pass

        self._refresh_pkg_status()

    def _on_add(self):
        w = self._parse_int_or_none(self.wh.get())
        p = self._parse_int_or_none(self.pkg.get())
        if not (w and p and w >= 1 and p >= 1):
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", "Ø£Ø¯Ø®Ù„ Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ© Ø¨Ø´ÙƒÙ„ ØµØ­ÙŠØ­ (>= 1).")
            return
        try:
            item_type = self.item_type.get() or self._err("Ø§Ù„Ù†ÙˆØ¹ Ù…Ø·Ù„ÙˆØ¨")
            school = self.school.get() or self._err("Ø§Ù„Ù…Ø¯Ø±Ø³Ø© Ù…Ø·Ù„ÙˆØ¨Ø©")
            color = self.color.get() or self._err("Ø§Ù„Ù„ÙˆÙ† Ù…Ø·Ù„ÙˆØ¨")
        except RuntimeError as ex:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", str(ex))
            return

        rows = self.sizes_grid.get_rows()
        # keep only rows with qty > 0
        rows = [r for r in rows if int(r.get("qty") or 0) > 0]
        if not rows:
            messagebox.showwarning("ÙØ§Ø±Øº", "Ù„Ù… ØªÙØ¯Ø®Ù„ Ø£ÙŠ ÙƒÙ…ÙŠØ§Øª ÙÙŠ Ø´Ø¨ÙƒØ© Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª.", parent=self)
            return

        try:
            r1s, r1e, r2s, r2e, has_alpha = self._current_income_profile_values()
            self.db.upsert_size_profile(
                item_type,
                school,
                color,
                r1_start=r1s,
                r1_end=r1e,
                r2_start=r2s,
                r2_end=r2e,
                has_alpha=has_alpha,
            )
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø­ÙØ¸ Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª", str(ex), parent=self)
            return

        # ensure package open
        try:
            self.db.ensure_package_open(w, p)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¥Ø¶Ø§ÙØ©", str(ex))
            return

        added = 0
        errors: List[str] = []
        preferred_profile_id = self._selected_income_price_profile_id()
        for r in rows:
            size = str(r["size"]).strip()
            qty = int(r["qty"])
            price = r["price"]

            if price is None:
                price = self.db.get_effective_price(
                    item_type,
                    school,
                    color,
                    size,
                    preferred_profile_id=preferred_profile_id,
                )

            if price is None:
                errors.append(f"{size}: Ø§Ù„Ø³Ø¹Ø± ØºÙŠØ± Ù…Ø¹Ø±ÙˆÙ (Ø£Ø¯Ø®Ù„Ù‡ Ù…Ø±Ø© ÙˆØ§Ø­Ø¯Ø© Ø¹Ù„Ù‰ Ø§Ù„Ø£Ù‚Ù„).")
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
            msg_parts.append(f"ØªÙ…Øª Ø¥Ø¶Ø§ÙØ© {added} ØµÙÙ‹Ø§ Ù„Ù„Ù…Ø®Ø²ÙˆÙ†.")
        if errors:
            msg_parts.append("Ø£Ø®Ø·Ø§Ø¡:\n" + "\n".join(errors))
        if not msg_parts:
            show_toast(self, "Ù„Ù… ØªÙØ¬Ø± Ø£ÙŠ ØªØºÙŠÙŠØ±Ø§Øª", bg="#f59e0b")
        else:
            show_toast(self, "\n".join(msg_parts), duration=3500)

        if added:
            # clear specs
            self.item_type.set("")
            self.school.set("")
            self.color.set("")

            # clear income-only ranges
            self._set_default_income_profile()

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
        self._sel_warehouse_no: Optional[int] = None
        self._sel_school: Optional[str] = None
        self._sel_item: Optional[str] = None
        self._sel_color: Optional[str] = None
        self._sel_size: Optional[str]  = None

        self._sizes_cache: List[Dict[str, Any]] = []
        self._build()

    # ---- Premium button palettes ----
    _BTN_PRODUCT = {
        "bg": _UI["ACCENT"], "fg": "#FFFFFF", "font": ("Segoe UI", 9, "bold"),
        "bd": 0, "relief": "flat", "padx": 12, "pady": 6, "cursor": "hand2",
        "activebackground": _UI["ACCENT_H"], "activeforeground": "#FFFFFF",
        "highlightbackground": _UI["BORDER"], "highlightthickness": 1,
    }
    _BTN_BACK = {
        "bg": _UI["SURFACE"], "fg": _UI["TEXT"], "font": ("Segoe UI", 9),
        "bd": 0, "cursor": "hand2", "padx": 8, "pady": 4,
        "highlightbackground": _UI["BORDER"], "highlightthickness": 1,
        "activebackground": _UI["BORDER"], "activeforeground": _UI["TEXT"],
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
        # Match POS startup split: selector side slightly wider.
        hsplit.add(left, weight=6)

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
            ("Ø§Ù„Ù…Ø®Ø²Ù†", "_filter_warehouse", 14),
            ("Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", "_filter_school", 18),
            ("Ø§Ù„Ù†ÙˆØ¹", "_filter_type", 16),
            ("Ø§Ù„Ù„ÙˆÙ†", "_filter_color", 12),
        ]):
            pill = tk.Frame(qf_inner, bg=_UI["SURFACE2"])
            pill.pack(side=tk.RIGHT, padx=6)
            tk.Label(pill, text=label_text, bg=_UI["SURFACE2"], fg=_UI["TEXT_SEC"],
                     font=_FONTS["small"]).pack(side=tk.RIGHT, padx=(0, 4))
            cb = ttk.Combobox(pill, width=width, state="readonly")
            cb.pack(side=tk.RIGHT)
            setattr(self, attr, cb)
        self._filter_warehouse["values"] = ["", *WAREHOUSE_NUMBER_DISPLAY_VALUES]
        for child in qf_inner.winfo_children():
            child.destroy()

        filter_specs = [
            ("Ø§Ù„Ù…Ø®Ø²Ù†", "_filter_warehouse", "warehouse_no", 14),
            ("Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", "_filter_school", "school", 18),
            ("Ø§Ù„Ù†ÙˆØ¹", "_filter_type", "item_type", 16),
            ("Ø§Ù„Ù„ÙˆÙ†", "_filter_color", "color", 12),
        ]
        for label_text, attr, field, width in filter_specs:
            pill = tk.Frame(qf_inner, bg=_UI["SURFACE2"])
            pill.pack(side=tk.RIGHT, padx=6)
            widget = LabeledCombobox(pill, label_text, self.db, field, width=width)
            widget.pack(side=tk.RIGHT)
            setattr(self, attr, widget)

        self._filter_warehouse.set_supplier(lambda: list(WAREHOUSE_NUMBER_DISPLAY_VALUES))
        self._filter_school.set_supplier(
            lambda: self.db.get_distinct_filtered("school", self._outcome_filter_constraints("school")) or []
        )
        self._filter_type.set_supplier(
            lambda: self.db.get_distinct_filtered("item_type", self._outcome_filter_constraints("item_type")) or []
        )
        self._filter_color.set_supplier(
            lambda: self.db.get_distinct_filtered("color", self._outcome_filter_constraints("color")) or []
        )

        clear_f = tk.Button(qf_inner, text="Ù…Ø³Ø­", command=self._clear_quick_filters,
                            bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"], font=_FONTS["small"],
                            bd=0, padx=12, pady=3, cursor="hand2",
                            highlightbackground=_UI["BORDER"], highlightthickness=1,
                            activebackground=_UI["SURFACE2"])
        clear_f.pack(side=tk.LEFT, padx=4)
        _add_hover(clear_f, _UI["SURFACE2"], _UI["SURFACE"])

        for field_name, widget in (
            ("warehouse_no", self._filter_warehouse),
            ("school", self._filter_school),
            ("item_type", self._filter_type),
            ("color", self._filter_color),
        ):
            widget.cb.bind(
                "<<ComboboxSelected>>",
                lambda _e, changed=field_name: self._on_filter_changed(changed),
                add="+",
            )
            widget.cb.bind(
                "<KeyRelease>",
                lambda _e, changed=field_name: self._on_filter_changed(changed),
                add="+",
            )
        self._refresh_filter_combos()

        # -- Breadcrumb --
        crumb = tk.Frame(left, bg=_UI["BG"])
        crumb.pack(fill=tk.X, padx=14, pady=(4, 2))
        self._crumb_var = tk.StringVar(value="Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø©")
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
        self._fav_frame = ttk.LabelFrame(left, text="Ø§Ù„Ø£ÙƒØ«Ø± Ù…Ø¨ÙŠØ¹Ø§Ù‹")
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
        tk.Label(act_inner, text="Ø§Ù„ÙƒÙ…ÙŠØ©:", bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
                 font=_FONTS["small"]).pack(side=tk.LEFT, padx=(0, 4))
        self.qty_var = tk.StringVar(value="1")
        _qm = tk.Button(act_inner, text="-",
                         command=lambda: self.qty_var.set(str(max(1, (parse_int_text(self.qty_var.get(), 1) or 1) - 1))),
                         bg=_UI["SURFACE"], fg=_UI["TEXT"], font=("Segoe UI", 11, "bold"),
                         bd=0, width=3, cursor="hand2",
                         highlightbackground=_UI["BORDER"], highlightthickness=1,
                         activebackground=_UI["SURFACE2"])
        _qm.pack(side=tk.LEFT)
        _add_hover(_qm, _UI["SURFACE2"], _UI["SURFACE"])
        ttk.Entry(act_inner, textvariable=self.qty_var, width=5, justify="center",
                  font=_FONTS["body"]).pack(side=tk.LEFT, padx=2)
        _qp = tk.Button(act_inner, text="+",
                         command=lambda: self.qty_var.set(str((parse_int_text(self.qty_var.get(), 1) or 1) + 1)),
                         bg=_UI["SURFACE"], fg=_UI["TEXT"], font=("Segoe UI", 11, "bold"),
                         bd=0, width=3, cursor="hand2",
                         highlightbackground=_UI["BORDER"], highlightthickness=1,
                         activebackground=_UI["SURFACE2"])
        _qp.pack(side=tk.LEFT, padx=(0, 16))
        _add_hover(_qp, _UI["SURFACE2"], _UI["SURFACE"])

        # Price
        tk.Label(act_inner, text="Ø§Ù„Ø³Ø¹Ø±:", bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
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
        ttk.Checkbutton(act_inner, text="Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹", variable=self.instant_mode).pack(side=tk.LEFT, padx=(0, 8))

        # Add to bill button (brand blue)
        _add_btn = tk.Button(act_inner, text="Ø£Ø¶Ù Ø¥Ù„Ù‰ Ø§Ù„ÙØ§ØªÙˆØ±Ø©", command=self._add_current_selection,
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
                    # Start close to the operator-tuned warehouse ratio.
                    hsplit.sashpos(0, int(w * 0.60))
                    hsplit.unbind("<Map>")
            except Exception:
                pass
        hsplit.bind("<Map>", _set_sash)

        # -- Customer --
        cust_frame = tk.Frame(right, bg=_UI["SURFACE"])
        cust_frame.pack(fill=tk.X, padx=10, pady=(10, 6))
        tk.Label(cust_frame, text="Ø§Ù„Ø¹Ù…ÙŠÙ„:", bg=_UI["SURFACE"], fg=_UI["TEXT_SEC"],
                 font=_FONTS["body"]).pack(side=tk.RIGHT, padx=(4, 0))
        self.customer = LabeledCombobox(cust_frame, "", self.db, "item_type")

        def _customer_supplier():
            # Branch POS devices appear at the top with a "ÙØ±Ø¹: " prefix
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
            ("type", "Ø§Ù„Ù†ÙˆØ¹", 100), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 100), ("color", "Ø§Ù„Ù„ÙˆÙ†", 70),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 55), ("price", "Ø§Ù„Ø³Ø¹Ø±", 65), ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 50),
            ("total", "Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ", 75),
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

        _del = tk.Button(bar, text="Ø­Ø°Ù Ø§Ù„Ø³Ø·Ø±", command=self._remove_bill_line,
                         bg=_UI["SURFACE"], fg=_UI["DANGER"], font=("Segoe UI", 9, "bold"),
                         bd=0, padx=12, pady=4, cursor="hand2",
                         highlightbackground=_UI["DANGER_L"], highlightthickness=1,
                         activebackground=_UI["DANGER_L"])
        _del.pack(side=tk.LEFT, padx=(0, 4))
        _add_hover(_del, _UI["DANGER_L"], _UI["SURFACE"])

        _clr = tk.Button(bar, text="ØªÙØ±ÙŠØº", command=self._clear_bill,
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
        tk.Label(total_inner, text="Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ", bg=_UI["OK_L"], fg="#065F46",
                 font=_FONTS["body"]).pack(side=tk.RIGHT, padx=(0, 12))
        self.total_var = tk.StringVar(value="0")
        tk.Label(total_inner, textvariable=self.total_var, bg=_UI["OK_L"], fg="#065F46",
                 font=_FONTS["price"]).pack(side=tk.RIGHT)

        # -- Confirm button (full width, premium green) --
        _conf = tk.Button(right, text="ØªØ£ÙƒÙŠØ¯ Ø§Ù„ÙØ§ØªÙˆØ±Ø© / Ø§Ù„Ø­Ø¬Ø²  (F8)", command=self._finalize_bill,
                          bg=_UI["OK"], fg="white", font=_FONTS["btn_lg"],
                          bd=0, pady=14, cursor="hand2",
                          activebackground=_UI["OK_H"], activeforeground="white")
        _conf.pack(fill=tk.X, padx=10, pady=(0, 4))
        _add_hover(_conf, _UI["OK_H"], _UI["OK"])

        # Draft button
        ttk.Button(right, text="Ø­ÙØ¸ ÙƒÙ…Ø³ÙˆØ¯Ø©", command=self._save_as_draft)\
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
            tk.Label(self._grid_host, text="Ù„Ø§ ØªÙˆØ¬Ø¯ Ø¨ÙŠØ§Ù†Ø§Øª.", bg=_UI["BG"],
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
            # Keep hover transitions aligned with the actual base button style.
            base_bg = self._BTN_PRODUCT.get("bg", _UI["ACCENT"])
            base_fg = self._BTN_PRODUCT.get("fg", "#FFFFFF")
            hover_bg = self._BTN_PRODUCT.get("activebackground", _UI["ACCENT_H"])
            hover_fg = self._BTN_PRODUCT.get("activeforeground", "#FFFFFF")
            _add_hover(btn, hover_bg, base_bg, hover_fg, base_fg)

    # ---------------- Cascading Quick Filters ----------------
    def _warehouse_no_from_filter(self) -> Optional[int]:
        raw = warehouse_numeric_value(self._filter_warehouse.get())
        raw = (raw or "").strip()
        if not raw:
            return None
        try:
            return int(raw)
        except Exception:
            return None

    @staticmethod
    def _warehouse_display_for_no(warehouse_no: Optional[int]) -> str:
        if warehouse_no is None:
            return ""
        return WAREHOUSE_NUMBER_LABELS.get(str(int(warehouse_no)), str(int(warehouse_no)))

    def _exact_filter_text(self, widget) -> Optional[str]:
        text = (widget.get() or "").strip()
        if not text:
            return None
        values = getattr(widget, "_all_values", None) or []
        return text if text in values else None

    def _exact_warehouse_no_from_filter(self) -> Optional[int]:
        raw_no = self._warehouse_no_from_filter()
        if raw_no is None:
            return None
        text = (self._filter_warehouse.get() or "").strip()
        if not text:
            return None
        expected_label = self._warehouse_display_for_no(raw_no)
        if text in {expected_label, str(int(raw_no))}:
            return raw_no
        return None

    def _outcome_filter_constraints(self, exclude: Optional[str] = None) -> Dict[str, Any]:
        constraints: Dict[str, Any] = {}
        warehouse_no = None if exclude == "warehouse_no" else self._exact_warehouse_no_from_filter()
        school = None if exclude == "school" else self._exact_filter_text(self._filter_school)
        item_type = None if exclude == "item_type" else self._exact_filter_text(self._filter_type)
        color = None if exclude == "color" else self._exact_filter_text(self._filter_color)

        if warehouse_no is not None:
            constraints["warehouse_no"] = warehouse_no
        if school:
            constraints["school"] = school
        if item_type:
            constraints["item_type"] = item_type
        if color:
            constraints["color"] = color
        return constraints

    def _refresh_filter_combos(self):
        """Refresh outcome filter menus using only confirmed exact selections."""
        try:
            self._filter_warehouse.refresh_values()
            self._filter_school.refresh_values()
            self._filter_type.refresh_values()
            self._filter_color.refresh_values()
        except Exception:
            pass

    def _on_filter_changed(self, changed_field: str):
        """Called when any quick filter changes, including typed narrowing input."""
        warehouse_no = self._exact_warehouse_no_from_filter()
        school = self._exact_filter_text(self._filter_school)
        item_type = self._exact_filter_text(self._filter_type)
        color = self._exact_filter_text(self._filter_color)

        # Invalidate any selection that is no longer valid given the others
        # Check school against current item_type + color
        if school:
            c = {}
            if warehouse_no is not None:
                c["warehouse_no"] = warehouse_no
            if item_type: c["item_type"] = item_type
            if color: c["color"] = color
            if school not in self.db.get_distinct_filtered("school", c):
                self._filter_school.set("")
                school = None

        # Check item_type against current school + color
        if item_type:
            c = {}
            if warehouse_no is not None:
                c["warehouse_no"] = warehouse_no
            if school: c["school"] = school
            if color: c["color"] = color
            if item_type not in self.db.get_distinct_filtered("item_type", c):
                self._filter_type.set("")
                item_type = None

        # Check color against current school + item_type
        if color:
            c = {}
            if warehouse_no is not None:
                c["warehouse_no"] = warehouse_no
            if school: c["school"] = school
            if item_type: c["item_type"] = item_type
            if color not in self.db.get_distinct_filtered("color", c):
                self._filter_color.set("")
                color = None

        self._refresh_filter_combos()
        self._sel_warehouse_no = warehouse_no
        self._sel_school = school
        self._sel_item = item_type
        self._sel_color = color
        self._sel_size = None
        self._price_user_edited = False

        # Render: School â†’ Item â†’ Color â†’ Size
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
        self._filter_warehouse.set("")
        self._filter_school.set("")
        self._filter_type.set("")
        self._filter_color.set("")
        self._sel_warehouse_no = None
        self._sel_school = None
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._price_user_edited = False
        self._refresh_filter_combos()
        self._render_schools()

    def _sync_filters_to_combos(self):
        """Sync internal selection state to filter comboboxes."""
        self._filter_warehouse.set(self._warehouse_display_for_no(self._sel_warehouse_no))
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
            self._crumb_var.set("Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø©")
        else:
            self._crumb_var.set("Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© (Ø¨Ø¹Ø¯ ØªØ¶ÙŠÙŠÙ‚ Ø§Ù„Ù†ØªØ§Ø¦Ø¬)")
        self._clear_grid()
        try:
            constraints: Dict[str, Any] = {}
            if self._sel_warehouse_no is not None:
                constraints["warehouse_no"] = self._sel_warehouse_no
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
        self._crumb_var.set(f"Ù†ØªØ§Ø¦Ø¬ Ø§Ù„Ø¨Ø­Ø« â€” Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© ({len(names)})")
        self._clear_grid()
        if not names:
            self._crumb_var.set("Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø©")
            self._bind_grid_scroll()
            return
        self._mk_grid_buttons(names, self._select_school, cols=4)
        tk.Button(
            self._grid_host,
            text="â—€ Ø¹Ø±Ø¶ ÙƒÙ„ Ø§Ù„Ù…Ø¯Ø§Ø±Ø³",
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
            inv_constraints: Dict[str, Any] = {}
            if self._sel_warehouse_no is not None:
                inv_constraints["warehouse_no"] = self._sel_warehouse_no
            inv_rows = self.db.current_inventory(inv_constraints)
            matches: List[Tuple[str, str]] = []
            for r in inv_rows:
                school = str(r.get("school") or "").strip()
                item_type = str(r.get("item_type") or "").strip()
                if school and item_type and q in item_type.lower():
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
                    show_toast(self, f"ØªÙ… Ø§Ù„Ø§Ù†ØªÙ‚Ø§Ù„ Ø¥Ù„Ù‰: {chosen_item} / {schools_for_item[0]}")
                else:
                    self._sel_school = None
                    self._sync_filters_to_combos()
                    self._render_schools()
                    show_toast(self, f"ØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ {chosen_item} ÙÙŠ Ø£ÙƒØ«Ø± Ù…Ù† Ù…Ø¯Ø±Ø³Ø©")
                return
        except Exception:
            pass
        # Then try matching school name(s) by substring (all matches, not only the first).
        try:
            school_constraints: Dict[str, Any] = {}
            if self._sel_warehouse_no is not None:
                school_constraints["warehouse_no"] = self._sel_warehouse_no
            schools = sorted({
                r["school"] for r in self.db.current_inventory(school_constraints) if r.get("school")
            })
            matching_schools = [s for s in schools if q in (s or "").lower()]
            if not matching_schools:
                show_toast(self, "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ù†ØªØ§Ø¦Ø¬", bg="#dc2626")
                return
            if len(matching_schools) == 1:
                self._select_school(matching_schools[0])
                show_toast(self, f"ØªÙ… Ø§Ù„Ø§Ù†ØªÙ‚Ø§Ù„ Ø¥Ù„Ù‰: {matching_schools[0]}")
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
                f"ÙŠÙˆØ¬Ø¯ {len(matching_schools)} Ù…Ø¯Ø±Ø³Ø© Ù…Ø·Ø§Ø¨Ù‚Ø© Ù„Ù„Ø¨Ø­Ø« â€” Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© Ù…Ù† Ø§Ù„Ù‚Ø§Ø¦Ù…Ø©",
            )
            return
        except Exception:
            pass
        show_toast(self, "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ù†ØªØ§Ø¦Ø¬", bg="#dc2626")

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
                ttk.Label(self._fav_inner, text="Ù„Ø§ ØªÙˆØ¬Ø¯ Ø¨ÙŠØ§Ù†Ø§Øª Ø¨Ø¹Ø¯").pack(side=tk.LEFT)
                return
            for r in rows:
                label = f"{r[0]} / {r[1]} / {r[2]}"
                btn = ttk.Button(self._fav_inner, text=label,
                                 command=lambda it=r[0], sc=r[1], cl=r[2]: self._jump_to_favorite(it, sc, cl))
                btn.pack(side=tk.RIGHT, padx=3, pady=2)
                ToolTip(btn, f"Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø¨Ø§Ø¹Ø©: {r[3]}")
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
        show_toast(self, f"ØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯: {item_type} / {school} / {color}")

    def _render_items(self):
        """Show item types for the selected school (or all if none selected)."""
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._clear_grid()

        if self._sel_school:
            self._crumb_var.set(f"Ø§Ù„Ù…Ø¯Ø±Ø³Ø©: {self._sel_school}  âŸ¶  Ø§Ø®ØªØ± Ø§Ù„Ù†ÙˆØ¹")
            try:
                items = sorted({
                    str(r.get("item_type") or "").strip()
                    for r in self.db.current_inventory({
                        "warehouse_no": self._sel_warehouse_no,
                        "school": self._sel_school,
                    })
                    if str(r.get("item_type") or "").strip()
                })
            except Exception:
                items = []
            self._mk_grid_buttons(items, self._select_item, cols=4)
            # Back to schools
            tk.Button(self._grid_host, text="â—€ Ø±Ø¬ÙˆØ¹ Ø¥Ù„Ù‰ Ø§Ù„Ù…Ø¯Ø§Ø±Ø³", command=self._render_schools,
                      **self._BTN_GRAY).pack(anchor="w", padx=4, pady=4)
            self._bind_grid_scroll()
        else:
            self._crumb_var.set("Ø§Ø®ØªØ± Ø§Ù„Ù†ÙˆØ¹")
            try:
                items = sorted({
                    r["item_type"]
                    for r in self.db.current_inventory(
                        {"warehouse_no": self._sel_warehouse_no} if self._sel_warehouse_no is not None else {}
                    )
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
        self._crumb_var.set(f"{self._sel_school}  âŸ¶  {self._sel_item}  âŸ¶  Ø§Ø®ØªØ± Ø§Ù„Ù„ÙˆÙ†")
        self._clear_grid()
        rows = self.db.current_inventory({
            "warehouse_no": self._sel_warehouse_no,
            "school": self._sel_school,
            "item_type": self._sel_item,
        })
        colors = sorted({str(r.get("color") or "").strip() for r in rows if str(r.get("color") or "").strip()})
        self._mk_grid_buttons(colors, self._select_color, cols=4)
        # Back to items
        tk.Button(self._grid_host, text="â—€ Ø±Ø¬ÙˆØ¹ Ø¥Ù„Ù‰ Ø§Ù„Ø£Ù†ÙˆØ§Ø¹", command=self._render_items,
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
        """Qty already on the draft bill that will deduct warehouse stock (excludes Ù…ØµÙ†Ø¹ Ù…Ø¨Ø§Ø´Ø± lines)."""
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
            f"{self._sel_school}  âŸ¶  {self._sel_item}  âŸ¶  {self._sel_color}  âŸ¶  Ø§Ø®ØªØ± Ø§Ù„Ù…Ù‚Ø§Ø³"
        )
        self._clear_grid()

        # Get sizes with counts/prices
        raw_sizes = self._get_sizes_for_bill(
            self._sel_school or "",
            self._sel_item or "",
            self._sel_color or "",
            self._sel_warehouse_no,
        )

        # Build cache with counts/prices if stock exists
        stock_rows = {}
        try:
            for r in self.db.current_inventory({
                "warehouse_no": self._sel_warehouse_no,
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
            ToolTip(btn, f"Ù†Ù‚Ø±: Ø§Ø®ØªÙŠØ§Ø± | Ù†Ù‚Ø± Ù…Ø²Ø¯ÙˆØ¬: Ø¥Ø¶Ø§ÙØ© Ø³Ø±ÙŠØ¹Ø© +1")

            # save button reference
            self._size_btns[label] = (btn, (cnt == 0))  # second value tracks zero-count


        # Back link
        nav_row = tk.Frame(self._grid_host, bg=_UI["BG"])
        nav_row.pack(anchor="w", padx=4, pady=4)
        tk.Button(nav_row, text="â—€ Ø±Ø¬ÙˆØ¹ Ø¥Ù„Ù‰ Ø§Ù„Ø£Ù„ÙˆØ§Ù†", command=self._render_colors,
                  **self._BTN_GRAY).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(nav_row, text="Add all to bill", command=self._add_all_visible_sizes_to_bill,
                  **self._BTN_PRODUCT).pack(side=tk.LEFT)

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
                self.price_var.set(f"{format_money(float(computed))}")
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
            show_toast(self, "Ù„Ø§ ÙŠÙˆØ¬Ø¯ Ø³Ø¹Ø± Ù…Ø­Ø¯Ø¯ - Ø§Ø®ØªØ± Ø§Ù„Ø³Ø¹Ø± Ø£ÙˆÙ„Ø§Ù‹", bg="#f59e0b", fg="#000")

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

            # 1) exact match (history) â€” return as-is (preserve original value)
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
                # nearest price derived from history â€” treat as a calculated fallback and ceil to next 10
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
                self.price_var.set(f"{format_money(float(computed))}")
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
            messagebox.showwarning("Ø§Ø®ØªØ± Ø£ÙˆÙ„Ù‹Ø§", "Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© Ø«Ù… Ø§Ù„Ù†ÙˆØ¹ Ø«Ù… Ø§Ù„Ù„ÙˆÙ† Ø«Ù… Ø§Ù„Ù…Ù‚Ø§Ø³.")
            return
        # qty
        try:
            qty = parse_int_text(self.qty_var.get(), 0) or 0
            if qty <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("ÙƒÙ…ÙŠØ© ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ø§Ù„ÙƒÙ…ÙŠØ© ÙŠØ¬Ø¨ Ø£Ù† ØªÙƒÙˆÙ† Ø¹Ø¯Ø¯Ù‹Ø§ ØµØ­ÙŠØ­Ù‹Ø§ Ù…ÙˆØ¬Ø¨Ù‹Ø§.")
            return
        # price
        try:
            price = parse_float_text(self.price_var.get())
            if price is None:
                raise ValueError
            if price < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("Ø³Ø¹Ø± ØºÙŠØ± ØµØ§Ù„Ø­", "Ø£Ø¯Ø®Ù„ Ø³Ø¹Ø±Ù‹Ø§ Ø±Ù‚Ù…ÙŠÙ‹Ø§ ØºÙŠØ± Ø³Ø§Ù„Ø¨.")
            return

        user_set_flag = bool(getattr(self, "_price_user_edited", False))
        allow_factory = bool(self.instant_mode.get())
        if allow_factory:
            try:
                customer_raw = (self.customer.get() or "").strip()
                known_pos = set(self.db.list_known_pos_device_names() or [])
                if canonical_branch_device_name(customer_raw, known_pos) or customer_raw in known_pos:
                    messagebox.showwarning(
                        "ØºÙŠØ± Ù…ØªØ§Ø­ Ù„Ù„Ø´Ø­Ù†",
                        "Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø¥Ø¶Ø§ÙØ© Ø¨Ù†Ø¯ Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹ Ø¥Ù„Ù‰ Ø´Ø­Ù†Ø© ÙØ±Ø¹. Ø§Ø®ØªØ± ÙƒÙ…ÙŠØ© Ù…ØªØ§Ø­Ø© Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø£ÙˆÙ„Ø§Ù‹.",
                        parent=self,
                    )
                    return
            except Exception:
                pass
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
        if self._sel_warehouse_no is not None:
            line["warehouse_no"] = int(self._sel_warehouse_no)

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
                "warehouse_no": self._sel_warehouse_no,
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

            try:
                self.db.validate_stock_available_for_bill_lines(self.bill_lines + [line])
            except Exception as ex:
                messagebox.showwarning("ÙƒÙ…ÙŠØ© ØºÙŠØ± Ù…ØªØ§Ø­Ø©", str(ex), parent=self)
                self._refresh_size_grid_if_current(preserve_size=self._sel_size)
                return

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

    def _add_all_visible_sizes_to_bill(self):
        if not all([self._sel_school, self._sel_item, self._sel_color]):
            messagebox.showwarning("Ø§Ø®ØªØ± Ø£ÙˆÙ„Ù‹Ø§", "Ø§Ø®ØªØ± Ø§Ù„Ù…Ø¯Ø±Ø³Ø© Ø«Ù… Ø§Ù„Ù†ÙˆØ¹ Ø«Ù… Ø§Ù„Ù„ÙˆÙ†.")
            return
        if not self._sizes_cache:
            show_toast(self, "Ù„Ø§ ØªÙˆØ¬Ø¯ Ù…Ù‚Ø§Ø³Ø§Øª Ù„Ù„Ø¥Ø¶Ø§ÙØ©", bg="#f59e0b", fg="#000")
            return

        allow_factory = bool(self.instant_mode.get())
        if allow_factory:
            try:
                customer_raw = (self.customer.get() or "").strip()
                known_pos = set(self.db.list_known_pos_device_names() or [])
                if canonical_branch_device_name(customer_raw, known_pos) or customer_raw in known_pos:
                    messagebox.showwarning(
                        "ØºÙŠØ± Ù…ØªØ§Ø­ Ù„Ù„Ø´Ø­Ù†",
                        "Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø¥Ø¶Ø§ÙØ© Ø¨Ù†ÙˆØ¯ Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹ Ø¥Ù„Ù‰ Ø´Ø­Ù†Ø© ÙØ±Ø¹. Ø§Ø®ØªØ± ÙƒÙ…ÙŠØ§Øª Ù…ØªØ§Ø­Ø© Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø£ÙˆÙ„Ø§Ù‹.",
                        parent=self,
                    )
                    return
            except Exception:
                pass

        try:
            factory_qty = parse_int_text(self.qty_var.get(), 1) or 1
            if factory_qty <= 0:
                raise ValueError
        except Exception:
            messagebox.showerror("ÙƒÙ…ÙŠØ© ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ø§Ù„ÙƒÙ…ÙŠØ© ÙŠØ¬Ø¨ Ø£Ù† ØªÙƒÙˆÙ† Ø¹Ø¯Ø¯Ù‹Ø§ ØµØ­ÙŠØ­Ù‹Ø§ Ù…ÙˆØ¬Ø¨Ù‹Ø§.", parent=self)
            return

        original_size = self._sel_size
        added = 0
        skipped_zero = 0
        skipped_price = 0
        for row in list(self._sizes_cache):
            size = str(row.get("size") or "").strip()
            if not size:
                continue
            count = int(row.get("count") or 0)
            if not allow_factory and count <= 0:
                skipped_zero += 1
                continue

            price = self._compute_price_for_size(size)
            if price is None:
                skipped_price += 1
                continue

            self._sel_size = size
            self.price_var.set(format_money(float(price)))
            self.qty_var.set(str(factory_qty if allow_factory else count))
            before = len(self.bill_lines)
            before_qty = sum(int(ln.get("qty") or 0) for ln in self.bill_lines)
            self._add_current_selection()
            after_qty = sum(int(ln.get("qty") or 0) for ln in self.bill_lines)
            if len(self.bill_lines) > before or after_qty > before_qty:
                added += 1

        self._sel_size = original_size
        self.qty_var.set("1")
        self._price_user_edited = False
        self._sync_bill_table()
        self._refresh_size_grid_if_current(preserve_size=original_size)

        if added <= 0:
            messagebox.showwarning("Ù„Ù… ØªØªÙ… Ø§Ù„Ø¥Ø¶Ø§ÙØ©", "Ù„Ø§ ØªÙˆØ¬Ø¯ Ù…Ù‚Ø§Ø³Ø§Øª ØµØ§Ù„Ø­Ø© Ù„Ù„Ø¥Ø¶Ø§ÙØ©.", parent=self)
            return
        note = f"ØªÙ…Øª Ø¥Ø¶Ø§ÙØ© {added} Ù…Ù‚Ø§Ø³"
        if skipped_zero and not allow_factory:
            note += f" - ØªÙ… ØªØ¬Ø§ÙˆØ² {skipped_zero} Ø¨Ø¯ÙˆÙ† Ù…Ø®Ø²ÙˆÙ†"
        if skipped_price:
            note += f" - Ø¨Ø¯ÙˆÙ† Ø³Ø¹Ø±: {skipped_price}"
        show_toast(self, note)

    def _remove_bill_line(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = parse_int_text(sel[0], -1)
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
        idx = parse_int_text(sel[0], -1)
        if 0 <= idx < len(self.bill_lines):
            next_lines = [dict(ln) for ln in self.bill_lines]
            next_lines[idx]["qty"] = int(next_lines[idx]["qty"]) + 1
            try:
                self.db.validate_stock_available_for_bill_lines(next_lines)
            except Exception as ex:
                messagebox.showwarning("ÙƒÙ…ÙŠØ© ØºÙŠØ± Ù…ØªØ§Ø­Ø©", str(ex), parent=self)
                self._refresh_size_grid_if_current(preserve_size=self._sel_size)
                return
            self.bill_lines[idx]["qty"] = next_lines[idx]["qty"]
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _decrement_bill_line(self):
        """Decrease qty of selected bill line by 1 (min 1)."""
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = parse_int_text(sel[0], -1)
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
                        f"{format_money(float(ln['unit_price']))}", ln["qty"], f"{format_money(line_total)}")
            )
        self.total_var.set(f"{format_money(total)}")
        apply_zebra_tags(self.bill_table)


    # ---------------- Finalize / print ----------------
    def _finalize_bill(self):
        if not self.bill_lines:
            show_toast(self, "Ø£Ø¶Ù Ø¨Ù†Ø¯Ù‹Ø§ ÙˆØ§Ø­Ø¯Ù‹Ø§ Ø¹Ù„Ù‰ Ø§Ù„Ø£Ù‚Ù„", bg="#f59e0b", fg="#000")
            return
        # Customer is required
        customer_raw = (self.customer.get() or "").strip()
        if not customer_raw:
            show_toast(self, "ÙŠØ±Ø¬Ù‰ Ø§Ø®ØªÙŠØ§Ø± Ø§Ø³Ù… Ø¹Ù…ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹", bg="#dc2626")
            return

        # Phase 3: "ÙØ±Ø¹: <branch ui name>" or raw device name â†’ shipment.
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

        if target_pos and any(bool(ln.get("allow_factory_fill")) for ln in self.bill_lines):
            messagebox.showwarning(
                "ØºÙŠØ± Ù…ØªØ§Ø­ Ù„Ù„Ø´Ø­Ù†",
                "Ù„Ø§ ÙŠÙ…ÙƒÙ† ØªØ£ÙƒÙŠØ¯ Ø´Ø­Ù†Ø© ÙØ±Ø¹ ØªØ­ØªÙˆÙŠ Ø¹Ù„Ù‰ Ø¨Ù†Ø¯ Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹. Ø§Ø­Ø°Ù Ù‡Ø°Ø§ Ø§Ù„Ø¨Ù†Ø¯ ÙˆØ§Ø®ØªØ± ÙƒÙ…ÙŠØ© Ù…ØªØ§Ø­Ø© Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†.",
                parent=self,
            )
            return

        # ===== CONFIRMATION OVERLAY =====
        total = self.total_var.get()
        n_items = len(self.bill_lines)

        overlay = tk.Toplevel(self)
        overlay.title("ØªØ£ÙƒÙŠØ¯ Ø§Ù„ÙØ§ØªÙˆØ±Ø©" if not target_pos else "ØªØ£ÙƒÙŠØ¯ Ø´Ø­Ù†Ø© Ø¥Ù„Ù‰ ÙØ±Ø¹")
        overlay.transient(self.winfo_toplevel())
        overlay.grab_set()
        overlay.resizable(False, False)

        frm = ttk.Frame(overlay, padding=20)
        frm.pack(fill=tk.BOTH, expand=True)

        header_text = "Ù…Ø±Ø§Ø¬Ø¹Ø© Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ù‚Ø¨Ù„ Ø§Ù„Ø­ÙØ¸" if not target_pos else "Ù…Ø±Ø§Ø¬Ø¹Ø© Ø´Ø­Ù†Ø© Ø§Ù„ÙØ±Ø¹ Ù‚Ø¨Ù„ Ø§Ù„Ø­ÙØ¸"
        ttk.Label(frm, text=header_text, font=("Segoe UI", 13, "bold")).pack(pady=(0, 12))
        if target_pos:
            ttk.Label(frm, text=f"Ø§Ù„ÙØ±Ø¹ Ø§Ù„Ù…ÙØ³ØªÙ„ÙÙ…: {target_pos}",
                      font=("Segoe UI", 11, "bold"),
                      foreground="#0ea5e9").pack(anchor="w", padx=8)
        else:
            ttk.Label(frm, text=f"Ø§Ù„Ø¹Ù…ÙŠÙ„: {customer}", font=("Segoe UI", 11)).pack(anchor="w", padx=8)
        ttk.Label(frm, text=f"Ø¹Ø¯Ø¯ Ø§Ù„Ø¨Ù†ÙˆØ¯: {n_items}", font=("Segoe UI", 11)).pack(anchor="w", padx=8)
        ttk.Label(frm, text=f"Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ: {total}", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(4, 12))

        # Summary of items
        summary = ttk.Treeview(frm, columns=("item", "size", "qty", "price"), show="headings", height=min(6, n_items))
        for col, txt, w in [("item", "Ø§Ù„ØµÙ†Ù", 180), ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 70), ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 60), ("price", "Ø§Ù„Ø³Ø¹Ø±", 80)]:
            summary.heading(col, text=txt)
            summary.column(col, width=w, anchor="center")
        for ln in self.bill_lines:
            summary.insert("", tk.END, values=(
                f"{ln['item_type']} / {ln['school']}",
                ln["size"], ln["qty"], f"{format_money(float(ln['unit_price']))}"
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
                messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­ÙØ¸", str(ex))
                return
            self._clear_bill()
            self._refresh_favorites()
            if print_after:
                self._direct_print_bill(bill_id)
                show_toast(self, f"Ø§Ù„ÙØ§ØªÙˆØ±Ø© #{bill_id} ØªÙ… Ø§Ù„Ø­ÙØ¸ ÙˆØ§Ù„Ø·Ø¨Ø§Ø¹Ø©")
            elif target_pos:
                show_toast(self, f"ØªÙ… Ø¥Ù†Ø´Ø§Ø¡ Ø´Ø­Ù†Ø© Ø§Ù„ÙØ±Ø¹ #{bill_id} Ø¥Ù„Ù‰ {target_pos}")
                try:
                    import sync_ui
                    sync_ui.run_sync_now(self.winfo_toplevel(), self.db.conn, reason=f"Ø´Ø­Ù†Ø© Ø¥Ù„Ù‰ {target_pos}")
                except Exception:
                    pass
            else:
                show_toast(self, f"ØªÙ… Ø­ÙØ¸ Ø§Ù„ÙØ§ØªÙˆØ±Ø© #{bill_id}")

        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=overlay.destroy).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø­ÙØ¸ Ø¨Ø¯ÙˆÙ† Ø·Ø¨Ø§Ø¹Ø©", command=lambda: _do_save(False)).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="Ø­ÙØ¸ ÙˆØ·Ø¨Ø§Ø¹Ø©", command=lambda: _do_save(True)).pack(side=tk.RIGHT, padx=4)

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
            show_toast(self, "Ø£Ø¶Ù Ø¨Ù†Ø¯ ÙˆØ§Ø­Ø¯ Ø¹Ù„Ù‰ Ø§Ù„Ø£Ù‚Ù„", bg="#f59e0b", fg="#000")
            return
        customer = (self.customer.get() or "").strip()
        if not customer:
            show_toast(self, "ÙŠØ±Ø¬Ù‰ Ø§Ø®ØªÙŠØ§Ø± Ø§Ø³Ù… Ø¹Ù…ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹", bg="#dc2626")
            return
        try:
            bill_id = self.db.create_draft_bill(customer, self.bill_lines)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­ÙØ¸", str(ex))
            return
        self._clear_bill()
        show_toast(self, f"ØªÙ… Ø­ÙØ¸ Ø§Ù„Ù…Ø³ÙˆØ¯Ø© #{bill_id}")

    def _print_bill(self, bill_id: int):
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©", "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ø§Ù„ÙØ§ØªÙˆØ±Ø©.")
            return
        tmp_dir = tempfile.gettempdir()
        path = os.path.join(tmp_dir, f"bill_{bill_id}.html")
        save_bill_as_html(path, bill, items)
        webbrowser.open_new_tab(f"file:///{path.replace(os.sep, '/')}")

    def _get_sizes_for_bill(self, school: str, item: str, color: str, warehouse_no: Optional[int] = None) -> List[str]:
        """
        Priority:
        1) Available sizes from the selected warehouse stock
        2) Size profile (numeric + alpha), appended only after real stock sizes
        """
        sizes: List[str] = []
        seen = set()
        try:
            filters: Dict[str, Any] = {
                "school": school,
                "item_type": item,
                "color": color,
            }
            if warehouse_no is not None:
                filters["warehouse_no"] = int(warehouse_no)
            rows = self.db.current_inventory(filters)
            for r in rows:
                sz = str(r.get("size") or "").strip()
                key = _normalize_size_label(sz).casefold()
                if sz and key not in seen and int(r.get("count") or 0) > 0:
                    seen.add(key)
                    sizes.append(sz)
        except Exception:
            pass

        try:
            profile = self.db.get_size_profile(item, school, color)
        except Exception:
            profile = None
        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            profile_sizes = list(merged_numeric_size_labels_from_profile(r1s, r1e, r2s, r2e))
            if has_alpha:
                profile_sizes.extend(ALPHA_SIZES)
            for sz in profile_sizes:
                key = _normalize_size_label(sz).casefold()
                if sz and key not in seen:
                    seen.add(key)
                    sizes.append(sz)

        return sizes


    def _direct_print_bill(self, bill_id: int, copies: int = 2):
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©", "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ø§Ù„ÙØ§ØªÙˆØ±Ø©.")
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
        self.title("Ø¥Ø¶Ø§ÙØ© Ø¨Ù†Ø¯ Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹")
        self.transient(master)
        self.grab_set()
        self.geometry("520x420")
        self.resizable(False, False)

        preset = preset or {}
        frm = ttk.Frame(self, padding=10); frm.pack(fill=tk.BOTH, expand=True)

        # Specs (combos for consistency with the app)
        self.t = LabeledCombobox(frm, "Ø§Ù„Ù†ÙˆØ¹",   db, "item_type"); self.t.set(preset.get("item_type",""))
        self.s = LabeledCombobox(frm, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", db, "school");    self.s.set(preset.get("school",""))
        self.c = LabeledCombobox(frm, "Ø§Ù„Ù„ÙˆÙ†",   db, "color");     self.c.set(preset.get("color",""))
        self.z = LabeledCombobox(frm, "Ø§Ù„Ù…Ù‚Ø§Ø³",  db, "size");      self.z.set(preset.get("size",""))
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
        ttk.Label(row_wp, text="Ø§Ù„Ù…Ø®Ø²Ù†:").pack(side=tk.LEFT)
        ttk.Entry(row_wp, textvariable=self.whv, width=8).pack(side=tk.LEFT, padx=(4,10))
        ttk.Label(row_wp, text="Ø§Ù„Ø¹Ø¨ÙˆØ©:").pack(side=tk.LEFT)
        ttk.Entry(row_wp, textvariable=self.pkv, width=10).pack(side=tk.LEFT, padx=(4,0))

        # Price/Qty + badge
        self.qv = tk.StringVar(value=preset.get("qty","1"))
        self.pv = tk.StringVar(value=preset.get("unit_price",""))
        self.badge = tk.BooleanVar(value=bool(int(preset.get("has_badge","0") or 0)))

        grid2 = ttk.Frame(frm); grid2.grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        ttk.Label(grid2, text="Ø§Ù„ÙƒÙ…ÙŠØ©:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(grid2, textvariable=self.qv, width=12).grid(row=0, column=1, sticky="w")
        ttk.Label(grid2, text="Ø³Ø¹Ø± Ø§Ù„ÙˆØ­Ø¯Ø©:").grid(row=0, column=2, sticky="e", padx=12, pady=4)
        ttk.Entry(grid2, textvariable=self.pv, width=12).grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(grid2, text="Ø¨Ø§Ø¯Ø¬", variable=self.badge).grid(row=0, column=4, sticky="w", padx=(12,0))

        # Buttons
        btns = ttk.Frame(frm); btns.grid(row=4, column=0, columnspan=2, sticky="e", padx=6, pady=(10,0))
        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ø¥Ø¶Ø§ÙØ©", command=self._ok).pack(side=tk.RIGHT, padx=8)

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
                self.pv.set(f"{format_money(p)}")
        except Exception:
            # silent: don't block user flow
            pass

    def _ok(self):
        try:
            item_type = self.t.get() or self._err("Ø§Ù„Ù†ÙˆØ¹ Ù…Ø·Ù„ÙˆØ¨")
            school    = self.s.get() or self._err("Ø§Ù„Ù…Ø¯Ø±Ø³Ø© Ù…Ø·Ù„ÙˆØ¨Ø©")
            color     = self.c.get() or self._err("Ø§Ù„Ù„ÙˆÙ† Ù…Ø·Ù„ÙˆØ¨")
            size      = self.z.get() or self._err("Ø§Ù„Ù…Ù‚Ø§Ø³ Ù…Ø·Ù„ÙˆØ¨")
            qty       = parse_int_text(self.qv.get(), 0) or 0;   assert qty > 0
            price     = parse_float_text(self.pv.get());         assert price is not None and price >= 0.0
            wh        = parse_int_text(self.whv.get(), 0) or 0
            pkg       = parse_int_text(self.pkv.get(), 0) or 0
        except AssertionError:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "ØªØ­Ù‚Ù‚ Ù…Ù† Ø§Ù„ÙƒÙ…ÙŠØ© (>0) ÙˆØ§Ù„Ø³Ø¹Ø± (>=0).", parent=self); return
        except Exception as ex:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", str(ex), parent=self); return

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
    dlg.title("Ø³Ø¬Ù„ ØªØ¹Ø¯ÙŠÙ„Ø§Øª Ø£Ø³Ø¹Ø§Ø± Ø§Ù„ÙØ±ÙˆØ¹")
    try:
        dlg.transient(parent.winfo_toplevel())
    except Exception:
        pass
    dlg.grab_set()

    top = ttk.Frame(dlg, padding=10)
    top.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        top,
        text="ÙŠØ¹Ø±Ø¶ Ù‡Ø°Ø§ Ø§Ù„Ø³Ø¬Ù„ Ù‚Ø±Ø§Ø±Ø§Øª Â«ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±Â» Ø§Ù„ØªÙŠ ØªÙ…Øª Ù…Ø²Ø§Ù…Ù†ØªÙ‡Ø§ Ù…Ø¹ ÙØ±ÙˆØ¹ Ø§Ù„Ø¨ÙŠØ¹ (POS) â€” Ø¨Ù…Ø§ ÙÙŠÙ‡Ø§ Ø£Ø­Ø¯Ø§Ø« Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ø¹Ù†Ø¯ Ø¥Ø¯Ø®Ø§Ù„ Ø³Ø¹Ø± ÙŠØ¯ÙˆÙŠ.",
        wraplength=860,
    ).pack(anchor="w", pady=(0, 8))

    # Treeview must live inside `grid` â€” `top` already uses pack(); mixing grid+pack on one parent raises TclError.
    grid = ttk.Frame(top)
    grid.pack(fill=tk.BOTH, expand=True)

    cols = ("ts", "price", "mode", "targets", "uuids", "filters", "note")
    tv = ttk.Treeview(grid, columns=cols, show="headings", height=16)
    for c, t, w in [
        ("ts", "Ø§Ù„ÙˆÙ‚Øª", 170),
        ("price", "Ø§Ù„Ø³Ø¹Ø±", 90),
        ("mode", "ÙˆØ¶Ø¹ Ø§Ù„Ø³Ø¬Ù„", 120),
        ("targets", "Ø§Ù„Ù†Ø·Ø§Ù‚/Ø§Ù„ÙØ±ÙˆØ¹", 260),
        ("uuids", "Ù…Ø¹Ø±Ù‘ÙØ§Øª Ø§Ù„Ø­Ø¯Ø«", 260),
        ("filters", "Ø¹ÙˆØ§Ù…Ù„ Ø§Ù„ØªØµÙÙŠØ©", 320),
        ("note", "Ù…Ù„Ø§Ø­Ø¸Ø©", 220),
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
        return s[: max_len - 1] + "â€¦"

    def _reload():
        tv.delete(*tv.get_children())
        rows = db.list_price_sync_audit(limit=limit)
        for r in rows:
            tv.insert(
                "",
                tk.END,
                values=(
                    str(r.get("created_at") or ""),
                    f"{format_money(float(r.get('new_price') or 0))}",
                    str(r.get("sync_mode") or ""),
                    _short_json(str(r.get("targets_json") or ""), 260),
                    _short_json(str(r.get("event_uuids_json") or ""), 260),
                    _short_json(str(r.get("filters_json") or ""), 320),
                    str(r.get("note") or ""),
                ),
            )

    btns = ttk.Frame(top)
    btns.pack(fill=tk.X, pady=(8, 0))
    ttk.Button(btns, text="ØªØ­Ø¯ÙŠØ«", command=_reload).pack(side=tk.RIGHT)
    ttk.Button(btns, text="Ø¥ØºÙ„Ø§Ù‚", command=dlg.destroy).pack(side=tk.RIGHT, padx=(0, 8))

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
        self.title("Ø§Ù„Ù…Ø®Ø²ÙˆÙ†")
        self.geometry("1000x560")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)

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
        self.show_zero_var = tk.BooleanVar(value=True)

        self._build()

    # ---------- UI ----------

    def _build(self):
        filters = ttk.LabelFrame(self, text="ØªØµÙ†ÙŠÙ")
        filters.pack(fill=tk.X, padx=8, pady=8)

        # Controls
        self.f_type   = LabeledCombobox(filters, "Ø§Ù„Ù†ÙˆØ¹",    self.db, "item_type")
        self.f_school = LabeledCombobox(filters, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self.f_color  = LabeledCombobox(filters, "Ø§Ù„Ù„ÙˆÙ†",    self.db, "color")
        self.f_size   = LabeledCombobox(filters, "Ø§Ù„Ù…Ù‚Ø§Ø³",   self.db, "size")
        self.f_wh     = LabeledStaticCombo(
            filters, "Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù†",
            values=["", *WAREHOUSE_NUMBER_DISPLAY_VALUES],
            value_map=WAREHOUSE_NUMBER_LABELS,
        )
        self.f_pkg    = LabeledEntry(filters, "Ø±Ù‚Ù… Ø§Ù„Ø¹Ø¨ÙˆØ©")

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

            # Add "Ø§Ø®ØªÙŠØ§Ø± Ù…ØªØ¹Ø¯Ø¯â€¦" button under each filter
            from functools import partial  # put this at the top of the file once

            def _safe_open(field):
                try:
                    self._open_multi_dialog(field)
                except Exception as ex:
                    messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=self)

            btn = ttk.Button(filters, text="Ø§Ø®ØªÙŠØ§Ø± Ù…ØªØ¹Ø¯Ø¯â€¦", command=partial(_safe_open, fld))
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
        _b1 = ttk.Button(btns, text="Ø¨Ø­Ø«", command=self._refresh); _b1.pack(side=tk.LEFT)
        ToolTip(_b1, "ØªØ·Ø¨ÙŠÙ‚ Ø§Ù„ÙÙ„Ø§ØªØ± ÙˆØ¹Ø±Ø¶ Ø§Ù„Ù†ØªØ§Ø¦Ø¬")
        _b2 = ttk.Button(btns, text="Ù…Ø³Ø­", command=self._clear_all); _b2.pack(side=tk.LEFT, padx=8)
        ToolTip(_b2, "Ù…Ø³Ø­ Ø¬Ù…ÙŠØ¹ Ø§Ù„ÙÙ„Ø§ØªØ±")
        _b3 = ttk.Button(btns, text="ØªØµØ¯ÙŠØ± Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„", command=self._export_excel); _b3.pack(side=tk.LEFT, padx=8)
        ToolTip(_b3, "ØªØµØ¯ÙŠØ± Ø§Ù„Ù†ØªØ§Ø¦Ø¬ Ø§Ù„Ø­Ø§Ù„ÙŠØ© Ø¥Ù„Ù‰ Ù…Ù„Ù Ø¥ÙƒØ³Ù„")
        _b4 = ttk.Button(btns, text="Ø·Ø¨Ø§Ø¹Ø© Ø¬Ø¯Ø§ÙˆÙ„ Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª", command=self._print_size_sheets); _b4.pack(side=tk.LEFT, padx=8)
        ToolTip(_b4, "Ø·Ø¨Ø§Ø¹Ø© Ø¬Ø¯Ø§ÙˆÙ„ Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª Ù„ÙƒÙ„ Ù†ÙˆØ¹/Ù…Ø¯Ø±Ø³Ø©/Ù„ÙˆÙ†")
        _b5 = ttk.Button(btns, text="ØªØ¹Ø¯ÙŠÙ„ Ù†Ø·Ø§Ù‚Ø§Øª Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øªâ€¦", command=self._edit_size_ranges_dialog); _b5.pack(side=tk.LEFT, padx=8)
        ToolTip(_b5, "ØªØ¹Ø¯ÙŠÙ„ Ù†Ø·Ø§Ù‚Ø§Øª Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª Ø§Ù„Ù…Ø¹Ø±ÙˆØ¶Ø© ÙÙŠ Ø§Ù„Ø¬Ø¯Ø§ÙˆÙ„")
        _b6 = ttk.Button(btns, text="ØªØ¹ÙŠÙŠÙ† Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø³Ø¹Ø±â€¦", command=self._edit_price_profile_assignment_dialog); _b6.pack(side=tk.LEFT, padx=8)
        ToolTip(_b6, "Ø±Ø¨Ø· Ø§Ù„Ù†ÙˆØ¹/Ø§Ù„Ù…Ø¯Ø±Ø³Ø©/Ø§Ù„Ù„ÙˆÙ† Ø§Ù„Ø­Ø§Ù„ÙŠÙŠÙ† Ø¨Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø³Ø¹Ø± Ù‚Ø§Ø¨Ù„ Ù„Ø¥Ø¹Ø§Ø¯Ø© Ø§Ù„Ø§Ø³ØªØ®Ø¯Ø§Ù…")
        _cz = ttk.Checkbutton(btns, text="Ø¥Ø¸Ù‡Ø§Ø± Ø§Ù„ÙƒÙ…ÙŠØ§Øª Ø§Ù„ØµÙØ±ÙŠØ©", variable=self.show_zero_var, command=self._refresh)
        _cz.pack(side=tk.LEFT, padx=8)
        ToolTip(_cz, "Ø¥Ø¸Ù‡Ø§Ø± Ø£Ùˆ Ø¥Ø®ÙØ§Ø¡ Ø§Ù„Ø£ØµÙ†Ø§Ù Ø§Ù„ØªÙŠ ÙƒÙ…ÙŠØªÙ‡Ø§ ØµÙØ±")

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
            ("id","Ø§Ù„Ù…Ø¹Ø±Ù‘Ù",60), ("type","Ø§Ù„Ù†ÙˆØ¹",140), ("school","Ø§Ù„Ù…Ø¯Ø±Ø³Ø©",160),
            ("color","Ø§Ù„Ù„ÙˆÙ†",80), ("size","Ø§Ù„Ù…Ù‚Ø§Ø³",70), ("wh","Ø§Ù„Ù…Ø®Ø²Ù†",50),
            ("pkg","Ø§Ù„Ø¹Ø¨ÙˆØ©",60), ("badge","Ø¨Ø§Ø¯Ø¬",60), ("price","Ø§Ù„Ø³Ø¹Ø±",80), ("count","Ø§Ù„ÙƒÙ…ÙŠØ©",70), ("value","Ø§Ù„Ù‚ÙŠÙ…Ø©",90),
        ]:

            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        self.table.tag_configure("zero_stock", background="#f3f4f6", foreground="#6b7280")
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
        self.sum_val = tk.StringVar(value="0")
        ttk.Label(bar, text="Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙƒÙ…ÙŠØ©:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_qty, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(bar, text="Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚ÙŠÙ…Ø©:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_val, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        _ba = ttk.Button(bar, text="Ø³Ø¬Ù„ Ø£Ø³Ø¹Ø§Ø± Ø§Ù„ÙØ±ÙˆØ¹â€¦", command=self._open_price_branch_audit)
        _ba.pack(side=tk.RIGHT)
        ToolTip(_ba, "Ø¹Ø±Ø¶ Ø³Ø¬Ù„ Ù…Ø²Ø§Ù…Ù†Ø© ØªØ¹Ø¯ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø± Ø¥Ù„Ù‰ ÙØ±ÙˆØ¹ Ø§Ù„Ø¨ÙŠØ¹ (POS)")
        _bp = ttk.Button(bar, text="ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±â€¦", command=self._edit_price_dialog); _bp.pack(side=tk.RIGHT, padx=(8, 0))
        ToolTip(_bp, "ØªØ¹Ø¯ÙŠÙ„ Ø³Ø¹Ø± Ø§Ù„Ø£ØµÙ†Ø§Ù Ø§Ù„Ù…Ø·Ø§Ø¨Ù‚Ø© Ù„Ù„ÙÙ„Ø§ØªØ±")
        _bs = ttk.Button(bar, text="ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ù…ÙˆØ§ØµÙØ§Øªâ€¦", command=self._edit_specs_dialog); _bs.pack(side=tk.RIGHT, padx=(8, 0))
        ToolTip(_bs, "ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©/Ø§Ù„Ø¨Ø§Ø¯Ø¬ Ù„Ù„Ø£ØµÙ†Ø§Ù Ø§Ù„Ù…Ø·Ø§Ø¨Ù‚Ø©")
        _bd = ttk.Button(bar, text="Ø­Ø°Ù Ø§Ù„Ù…Ø­Ø¯Ø¯â€¦", command=self._remove_selected_dialog); _bd.pack(side=tk.RIGHT, padx=(8, 0))
        ToolTip(_bd, "Ø­Ø°Ù Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†")

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
                "ØªØ­Ø¯ÙŠØ¯ ØºÙŠØ± ØµØ§Ù„Ø­",
                "ÙŠØ¬Ø¨ Ø£Ù† ØªÙƒÙˆÙ† Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ù…Ù† Ù†ÙØ³ (Ø§Ù„Ù†ÙˆØ¹ØŒ Ø§Ù„Ù…Ø¯Ø±Ø³Ø©ØŒ Ø§Ù„Ù„ÙˆÙ†).",
                parent=self
            )
            return None

        return keys.pop()

    def _get_selected_price_profile_targets(self):
        """
        Price profiles can be assigned to mixed item types/colors, but
        the selected rows must belong to one school.
        """
        sel = self.table.selection()
        if not sel:
            return None

        schools = set()
        targets = set()
        for iid in sel:
            vals = self.table.item(iid, "values")
            if not vals:
                continue
            item_type = str(vals[1] or "").strip()
            school = str(vals[2] or "").strip()
            color = str(vals[3] or "").strip()
            if school:
                schools.add(school)
            if item_type and school and color:
                targets.add((item_type, school, color))

        if not targets:
            return None

        if len(schools) != 1:
            messagebox.showwarning(
                "ØªØ­Ø¯ÙŠØ¯ ØºÙŠØ± ØµØ§Ù„Ø­",
                "ÙŠØ¬Ø¨ Ø£Ù† ØªÙƒÙˆÙ† Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ù…Ù† Ù†ÙØ³ Ø§Ù„Ù…Ø¯Ø±Ø³Ø©.",
                parent=self
            )
            return None

        return sorted(targets, key=lambda x: (x[1].casefold(), x[0].casefold(), x[2].casefold()))

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
                    "Ø­Ø¯Ø¯ Ø§Ù„ØµÙ†Ù",
                    "Ø­Ø¯Ø¯ ØµÙÙˆÙÙ‹Ø§ Ù…Ù† Ø§Ù„Ø¬Ø¯ÙˆÙ„ Ø£Ùˆ Ø§Ø®ØªØ± (Ø§Ù„Ù†ÙˆØ¹ØŒ Ø§Ù„Ù…Ø¯Ø±Ø³Ø©ØŒ Ø§Ù„Ù„ÙˆÙ†) Ø£ÙˆÙ„Ø§Ù‹.",
                    parent=self
                )
                return


        # Fetch existing profile (if any)

        dlg = tk.Toplevel(self)
        dlg.title("ØªØ¹Ø¯ÙŠÙ„ Ù†Ø·Ø§Ù‚Ø§Øª Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª")
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

        ttk.Label(frm, text="Ø§Ù„Ù†Ø·Ø§Ù‚ Ø§Ù„Ø£ÙˆÙ„").grid(row=0, column=0, sticky="w")

        r1_var = tk.StringVar()
        r1_combo = ttk.Combobox(
            frm,
            textvariable=r1_var,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=12,
        )
        r1_combo.grid(row=0, column=1, columnspan=2, sticky="w")


        ttk.Label(frm, text="Ø§Ù„Ù†Ø·Ø§Ù‚ Ø§Ù„Ø«Ø§Ù†ÙŠ (Ø§Ø®ØªÙŠØ§Ø±ÙŠ)").grid(row=1, column=0, sticky="w")

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
                r1_var.set(f"{r1s} â†’ {r1e}")

            if r2s is not None and r2e is not None:
                r2_var.set(f"{r2s} â†’ {r2e}")

            has_alpha_var.set(bool(has_alpha))

        ttk.Checkbutton(
            frm,
            text="ØªÙØ¹ÙŠÙ„ Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª Ø¨Ø§Ù„Ø­Ø±ÙˆÙ (S / M / L ...)",
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
            a, b = v.split("â†’")
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
                show_toast(dlg, "ØªÙ… Ø­ÙØ¸ Ù†Ø·Ø§Ù‚Ø§Øª Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª")
                dlg.destroy()
            except Exception as ex:
                messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=dlg)


        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ø­ÙØ¸", command=on_save).pack(side=tk.RIGHT, padx=6)

    def _edit_price_profile_assignment_dialog(self):
        has_selection = bool(self.table.selection())
        selected_stock_ids: List[int] = []
        selected_inventory_filters: List[Dict[str, Any]] = []
        for iid in self.table.selection():
            vals = self.table.item(iid, "values")
            stock_id = parse_int_text(vals[0]) if vals else None
            if stock_id is not None:
                selected_stock_ids.append(stock_id)
            if vals:
                selected_inventory_filters.append({
                    "id": stock_id,
                    "item_type": str(vals[1] or "").strip(),
                    "school": str(vals[2] or "").strip(),
                    "color": str(vals[3] or "").strip(),
                    "size": _normalize_size_label(_strip_digit_marks(vals[4])),
                    "warehouse_no": warehouse_numeric_value(vals[5]),
                    "package_no": str(vals[6] or "").strip(),
                })
        targets = self._get_selected_price_profile_targets()
        if has_selection and not targets:
            return
        picked = targets[0] if targets else None
        if picked:
            item_type, school, color = picked
        else:
            item_type = (self.f_type.get() or "").strip()
            school = (self.f_school.get() or "").strip()
            color = (self.f_color.get() or "").strip()
            if not (item_type and school and color):
                messagebox.showwarning(
                    "Ø­Ø¯Ø¯ Ø§Ù„ØµÙ†Ù",
                    "Ø­Ø¯Ø¯ ØµÙÙˆÙØ§Ù‹ Ù…Ù† Ø§Ù„Ø¬Ø¯ÙˆÙ„ Ø£Ùˆ Ø§Ø®ØªØ± (Ø§Ù„Ù†ÙˆØ¹ØŒ Ø§Ù„Ù…Ø¯Ø±Ø³Ø©ØŒ Ø§Ù„Ù„ÙˆÙ†) Ø£ÙˆÙ„Ø§Ù‹.",
                    parent=self,
                )
                return
            targets = [(item_type, school, color)]

        profiles = self.db.list_price_profiles()
        if not profiles:
            if messagebox.askyesno(
                "Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±",
                "Ù„Ø§ ØªÙˆØ¬Ø¯ Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø£Ø³Ø¹Ø§Ø± Ø¨Ø¹Ø¯. Ù‡Ù„ ØªØ±ÙŠØ¯ ÙØªØ­ Ù†Ø§ÙØ°Ø© Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø± Ø§Ù„Ø¢Ù†ØŸ",
                parent=self,
            ):
                PriceProfileManagerWindow(self, self.db)
            return

        current_names = set()
        for t_item, t_school, t_color in targets:
            current = self.db.get_price_profile_assignment(t_item, t_school, t_color) or {}
            current_name = str(current.get("profile_name") or "").strip()
            if current_name:
                current_names.add(current_name)
        profile_names = [""] + [str(p.get("name") or "").strip() for p in profiles if str(p.get("name") or "").strip()]
        profile_map = {str(p.get("name") or "").strip(): int(p["id"]) for p in profiles if str(p.get("name") or "").strip()}

        dlg = tk.Toplevel(self)
        dlg.title("ØªØ¹ÙŠÙŠÙ† Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø³Ø¹Ø±")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        target_text = f"{item_type} / {school} / {color}"
        if len(targets) > 1:
            target_text = f"{school} - {len(targets)} Ù†ÙˆØ¹/Ù„ÙˆÙ† Ù…Ø­Ø¯Ø¯"
        ttk.Label(frm, text=target_text, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 10))

        selected_name = tk.StringVar(value=next(iter(current_names)) if len(current_names) == 1 else "")
        ttk.Label(frm, text="Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±").pack(anchor="w")
        combo = ttk.Combobox(frm, textvariable=selected_name, state="readonly", values=profile_names)
        combo.pack(fill=tk.X, pady=(4, 8))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(8, 0))

        def _save():
            label = (selected_name.get() or "").strip()
            profile_id = profile_map.get(label)
            if not profile_id:
                messagebox.showwarning("Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„", "Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø³Ø¹Ø± Ù‚Ø¨Ù„ Ø§Ù„Ø­ÙØ¸.", parent=dlg)
                return
            try:
                for t_item, t_school, t_color in targets:
                    self.db.assign_price_profile(t_item, t_school, t_color, profile_id)
                if selected_inventory_filters:
                    updated = 0
                    skipped = 0
                    for flt in selected_inventory_filters:
                        price = self.db.get_price_profile_price(profile_id, flt.get("item_type"), flt.get("size"))
                        if price is None:
                            skipped += 1
                            continue
                        update_filter = {"id": int(flt["id"])} if flt.get("id") else {
                            "item_type": flt.get("item_type"),
                            "school": flt.get("school"),
                            "color": flt.get("color"),
                            "size": flt.get("size"),
                            "warehouse_no": flt.get("warehouse_no"),
                            "package_no": flt.get("package_no"),
                        }
                        updated += self.db.update_prices(
                            update_filter,
                            float(price),
                            note=f"Price profile applied: {label}",
                        )
                    result = {"updated": updated, "skipped": skipped, "checked": len(selected_inventory_filters)}
                else:
                    result = self.db.apply_price_profile_to_stock(
                        profile_id,
                        targets,
                        stock_ids=selected_stock_ids if selected_stock_ids else None,
                        note=f"Price profile applied: {label}",
                    )
                show_toast(
                    dlg,
                    f"ØªÙ… Ø­ÙØ¸ Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„ ÙˆØªØ­Ø¯ÙŠØ« {result.get('updated', 0)} ØµÙØŒ ÙˆØªØ®Ø·ÙŠ {result.get('skipped', 0)}",
                    bg="#16a34a" if int(result.get("updated", 0) or 0) else "#f59e0b",
                )
                dlg.destroy()
                self._refresh()
            except Exception as ex:
                messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=dlg)

        def _clear():
            try:
                for t_item, t_school, t_color in targets:
                    self.db.clear_price_profile_assignment(t_item, t_school, t_color)
                show_toast(dlg, f"ØªÙ… Ù…Ø³Ø­ Ø§Ù„Ø±Ø¨Ø· Ù…Ù† {len(targets)} Ù†ÙˆØ¹/Ù„ÙˆÙ†")
                dlg.destroy()
            except Exception as ex:
                messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=dlg)

        ttk.Button(btns, text="Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øªâ€¦", command=lambda: PriceProfileManagerWindow(dlg, self.db)).pack(side=tk.LEFT)
        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ø­ÙØ¸", command=_save).pack(side=tk.RIGHT, padx=6)
        ttk.Button(btns, text="Ù…Ø³Ø­ Ø§Ù„Ø±Ø¨Ø·", command=_clear).pack(side=tk.RIGHT, padx=6)

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
        dlg = MultiSelectDialog(self, title="Ø§Ø®ØªÙŠØ§Ø± Ù…ØªØ¹Ø¯Ø¯", values=values, preselected=preselected)
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
            btn.configure(text=f"Ø§Ø®ØªÙŠØ§Ø± Ù…ØªØ¹Ø¯Ø¯â€¦ ({len(self.multi[field])})" if self.multi[field] else "Ø§Ø®ØªÙŠØ§Ø± Ù…ØªØ¹Ø¯Ø¯â€¦")

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
            messagebox.showerror("ÙØ´Ù„", str(ex), parent=self)
            return

        if not rows:
            show_toast(self, "Ù„Ø§ ØªÙˆØ¬Ø¯ ØµÙÙˆÙ Ù…Ø·Ø§Ø¨Ù‚Ø© Ù„Ù„Ø·Ø¨Ø§Ø¹Ø©", bg="#f59e0b")
            return
        from collections import defaultdict, OrderedDict

        def _item_print_sort_key(item_type: str, color: str):
            return warehouse_item_sort_key(item_type, color)

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

        for sch, item_groups in sorted(school_groups.items(), key=lambda kv: kv[0].casefold()):
            ordered_item_groups = sorted(
                item_groups.items(),
                key=lambda kv: _item_print_sort_key(kv[0][0], kv[0][1]),
            )

            for (t, clr), items in ordered_item_groups:


                size_counts = defaultdict(int)
                used_sizes = set()

                for r in items:
                    sz = _normalize_size_label(r.get("size") or "")
                    size_counts[sz] += int(r.get("count") or 0)
                    used_sizes.add(sz)

                profile = self.db.get_size_profile(t, sch, clr)
                numeric_tables, alpha_labels = build_size_ranges_from_profile(profile)

                # ðŸ”´ FALLBACK: no profile â†’ use actual available sizes
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
                    <span>Ø§Ù„Ù†ÙˆØ¹: {_html(t)}</span>
                    <span>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©: {_html(sch)}</span>
                    <span>Ø§Ù„Ù„ÙˆÙ†: {_html(clr)}</span>
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
                    <div style="margin-top:6px;font-weight:600">Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª Ø¨Ø§Ù„Ø­Ø±ÙˆÙ</div>
                    """ + build_table(alpha_labels))

                tables_html.append(f'<section class="sheet">{head}{"".join(tables)}</section>')



            html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<title>Ø¬Ø¯Ø§ÙˆÙ„ Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª</title>
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
            f["hide_zero"] = not bool(self.show_zero_var.get())
            return f

        # No multi: build from single-value controls
        return {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "color": self.f_color.get() or None,
            "size": self.f_size.get() or None,
            "warehouse_no": (self.f_wh.get() or None),
            "package_no": (self.f_pkg.get() or None),
            "hide_zero": not bool(self.show_zero_var.get()),
        }

    def _refresh(self):
        try:
            rows = self.db.current_inventory(self._filters())
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¨Ø­Ø«", str(ex), parent=self)
            return

        self.table.delete(*self.table.get_children())
        total_qty = 0
        total_value = 0.0
        for r in rows:
            tags = ("zero_stock",) if int(r.get("count") or 0) == 0 else ()
            self.table.insert(
                "", tk.END,
                values=(r["id"], r["item_type"], r["school"], r["color"], r["size"],
                        r["warehouse_no"], r["package_no"],
                        ("âœ“" if int(r.get("has_badge") or 0) else ""),
                        f"{format_money(float(r['unit_price']))}",
                        r["count"], f"{format_money(float(r['value']))}"),
                tags=tags,
            )

            total_qty += int(r["count"])
            total_value += float(r["value"])
        self.sum_qty.set(str(total_qty))
        self.sum_val.set(f"{format_money(total_value)}")
        apply_zebra_tags(self.table, skip_tags={"zero_stock"})

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
            b.configure(text="Ø§Ø®ØªÙŠØ§Ø± Ù…ØªØ¹Ø¯Ø¯â€¦")

        # Re-enable everything
        self._enforce_single_active_field()
        self._refresh()

    # ---------- Existing actions kept as-is ----------

    def _export_excel(self):
        try:
            db_rows = self.db.current_inventory(self._filters())
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex), parent=self)
            return

        path = filedialog.asksaveasfilename(
            title="ØªØµØ¯ÙŠØ± Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        try:
            self.db.export_inventory_excel(path, db_rows)
            show_toast(self, f"ØªÙ… Ø­ÙØ¸ Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„ Ø¨Ù†Ø¬Ø§Ø­")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex), parent=self)

    def _remove_selected_dialog(self):
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning(
                "Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯",
                "Ø§Ø®ØªØ± ØµÙÙ‹Ø§ ÙˆØ§Ø­Ø¯Ù‹Ø§ Ø£Ùˆ Ø£ÙƒØ«Ø± Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø£ÙˆÙ„Ù‹Ø§.",
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
            title = "Ø­Ø°Ù Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† (ÙŠØªØ·Ù„Ø¨ ÙƒÙ„Ù…Ø© Ù…Ø±ÙˆØ±)"
        else:
            label = f"Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø­Ø¯Ø¯Ø©: {len(rows)}"
            available = sum(int(r[9]) for r in rows)
            title = "Ø­Ø°Ù Ù…ØªØ¹Ø¯Ø¯ Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† (ÙŠØªØ·Ù„Ø¨ ÙƒÙ„Ù…Ø© Ù…Ø±ÙˆØ±)"

        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=label, font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

        ttk.Label(frm, text=f"Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…ØªÙˆÙØ±: {available}")\
            .grid(row=1, column=0, columnspan=2, sticky="w")

        ttk.Label(frm, text="Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ Ø­Ø°ÙÙ‡Ø§ (ÙØ§Ø±Øº = Ø§Ù„ÙƒÙ„):")\
            .grid(row=2, column=0, sticky="e", padx=4, pady=8)

        qty_var = tk.StringVar()
        ttk.Entry(frm, textvariable=qty_var, width=12)\
            .grid(row=2, column=1, sticky="w", padx=4, pady=8)

        ttk.Label(frm, text="ÙƒÙ„Ù…Ø© Ø§Ù„Ù…Ø±ÙˆØ±:")\
            .grid(row=3, column=0, sticky="e", padx=4, pady=4)

        pw_var = tk.StringVar()
        ttk.Entry(frm, textvariable=pw_var, show="*")\
            .grid(row=3, column=1, sticky="w", padx=4, pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))

        def on_ok():
            if not self.db.verify_admin_password(pw_var.get()):
                messagebox.showerror("Ù…Ø±ÙÙˆØ¶", "ÙƒÙ„Ù…Ø© Ø§Ù„Ù…Ø±ÙˆØ± ØºÙŠØ± ØµØ­ÙŠØ­Ø©.", parent=dlg)
                return

            qtext = qty_var.get().strip()
            qty = None
            if qtext:
                qty = parse_int_text(qtext)
                if qty is None or qty <= 0:
                    messagebox.showerror(
                        "ÙƒÙ…ÙŠØ© ØºÙŠØ± ØµØ§Ù„Ø­Ø©",
                        "Ø£Ø¯Ø®Ù„ Ø¹Ø¯Ø¯Ù‹Ø§ Ù…ÙˆØ¬Ø¨Ù‹Ø§ Ø£Ùˆ Ø§ØªØ±ÙƒÙ‡ ÙØ§Ø±ØºÙ‹Ø§ Ù„Ù„ÙƒÙ„.",
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
                show_toast(self, f"ØªÙ… Ø­Ø°Ù {total_removed} ÙˆØ­Ø¯Ø© Ù…Ù† {len(ids)} ØµÙ(ÙˆÙ)")
                self._refresh()

            except Exception as ex:
                messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=dlg)

        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ø­Ø°Ù", command=on_ok).pack(side=tk.RIGHT, padx=6)

    # file: ui/edit_price_dialog.py

    def _edit_price_dialog(self):
        row = 0
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ØµÙÙ‹Ø§ Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø£ÙˆÙ„Ù‹Ø§.", parent=self)
            return

        rows = [self.table.item(i, "values") for i in sel]
        first = rows[0]

        current_price = parse_float_text(first[8], 0.0) or 0.0
        multi = len(rows) > 1

        dlg = tk.Toplevel(self)
        dlg.title("ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        # ---- Header ----
        if multi:
            ttk.Label(
                frm,
                text=f"Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø­Ø¯Ø¯Ø©: {len(rows)}",
                font=("Segoe UI", 10, "bold")
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 2))
            row += 1
            ttk.Label(
                frm,
                text="ÙŠÙ…ÙƒÙ†Ùƒ ØªØ­Ø¯ÙŠØ¯ ØµÙÙˆÙ Ø¨Ù†ÙØ³ Ø§Ù„Ù†ÙˆØ¹ ÙˆØ§Ù„Ù…Ø¯Ø±Ø³Ø© ÙˆÙ…Ù‚Ø§Ø³Ø§Øª Ù…Ø®ØªÙ„ÙØ©Ø› Ø³ÙŠØªÙ… ØªØ­Ø¯ÙŠØ« ÙƒÙ„ ØµÙ ÙˆÙ…Ø²Ø§Ù…Ù†Ø© Ø§Ù„ÙØ±ÙˆØ¹ Ù„ÙƒÙ„ Ù…Ù‚Ø§Ø³ Ø¹Ù„Ù‰ Ø­Ø¯Ø© Ø¹Ù†Ø¯ ØªÙØ¹ÙŠÙ„ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©.",
                wraplength=520,
                foreground="#64748B",
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
        else:
            ttk.Label(
                frm,
                text=f"Ø§Ù„ØµÙ†Ù: {first[1]} / {first[2]} / {first[3]} / {first[4]}"
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1

            ttk.Label(
                frm,
                text=f"Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©: {first[5]} / {first[6]}"
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1

            ttk.Label(
                frm,
                text=f"Ø§Ù„Ø³Ø¹Ø± Ø§Ù„Ø­Ø§Ù„ÙŠ: {format_money(current_price)}"
            ).grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 10))
            row += 1

        # ---- New price ----
        ttk.Label(frm, text="Ø§Ù„Ø³Ø¹Ø± Ø§Ù„Ø¬Ø¯ÙŠØ¯:").grid(row=row, column=0, sticky="e", padx=4, pady=6)
        price_var = tk.StringVar(value=f"{format_money(current_price)}")
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

        sync_box = ttk.LabelFrame(frm, text="Ù…Ø²Ø§Ù…Ù†Ø© Ø³Ø¹Ø± Ø§Ù„ÙØ±ÙˆØ¹ (POS)")
        sync_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 10))
        pos_sync_var = tk.StringVar(value="all-pos")

        ttk.Radiobutton(
            sync_box,
            text="ÙƒÙ„ ÙØ±ÙˆØ¹ Ø§Ù„Ø¨ÙŠØ¹",
            variable=pos_sync_var,
            value="all-pos",
        ).pack(anchor="w", padx=8, pady=(6, 2))

        ttk.Radiobutton(
            sync_box,
            text="ÙØ±ÙˆØ¹ Ù…Ø­Ø¯Ø¯Ø©â€¦",
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
            text="Ù„Ø§ (ØªØ­Ø¯ÙŠØ« Ø§Ù„Ù…Ø®Ø²Ù† ÙÙ‚Ø·)",
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
            scope_box = ttk.LabelFrame(frm, text="Ø§Ù„Ù†Ø·Ø§Ù‚")
            scope_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 10))

            ttk.Radiobutton(
                scope_box,
                text="Ø§Ù„ØµÙ Ø§Ù„Ù…Ø­Ø¯Ø¯ ÙÙ‚Ø·",
                variable=scope_var,
                value="row"
            ).pack(anchor="w", padx=8, pady=4)

            ttk.Radiobutton(
                scope_box,
                text="ÙƒÙ„ Ø§Ù„ØµÙÙˆÙ Ø¨Ù†ÙØ³ (Ø§Ù„Ù†ÙˆØ¹/Ø§Ù„Ù…Ø¯Ø±Ø³Ø©/Ø§Ù„Ù…Ù‚Ø§Ø³) Ø¯Ø§Ø®Ù„ Ù†ÙØ³ Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©",
                variable=scope_var,
                value="same_pkg"
            ).pack(anchor="w", padx=8, pady=4)

            row += 1

        # ---- OK handler ----
        def on_ok():
            def _single_int_or_none(value: Any) -> Optional[int]:
                txt = _strip_digit_marks(value).strip()
                if not txt or "," in txt:
                    return None
                try:
                    return int(txt)
                except Exception:
                    return None

            def _visible_row_filter(values: Sequence[Any]) -> Dict[str, Any]:
                row_id = _single_int_or_none(values[0])
                if row_id is not None:
                    return {"id": row_id}
                return {
                    "item_type": values[1],
                    "school": values[2],
                    "color": values[3],
                    "size": values[4],
                    "warehouse_no": _single_int_or_none(values[5]),
                    "package_no": _single_int_or_none(values[6]),
                    "unit_price": parse_float_text(values[8], 0.0) or 0.0,
                }

            try:
                new_price = parse_float_text(price_var.get())
                if new_price < 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("Ù‚ÙŠÙ…Ø© ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ø£Ø¯Ø®Ù„ Ø³Ø¹Ø±Ù‹Ø§ Ø±Ù‚Ù…ÙŠÙ‹Ø§ ØºÙŠØ± Ø³Ø§Ù„Ø¨.", parent=dlg)
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
                            "Ø§Ø®ØªØ± Ø§Ù„ÙØ±ÙˆØ¹",
                            "Ø§Ø®ØªØ± ÙØ±Ø¹Ù‹Ø§ ÙˆØ§Ø­Ø¯Ù‹Ø§ Ø¹Ù„Ù‰ Ø§Ù„Ø£Ù‚Ù„ Ù…Ù† Ù‚Ø§Ø¦Ù…Ø© ÙØ±ÙˆØ¹ Ø§Ù„Ø¨ÙŠØ¹ØŒ Ø£Ùˆ ØºÙŠÙ‘Ø± Ø®ÙŠØ§Ø± Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© Ø¥Ù„Ù‰ Â«ÙƒÙ„ Ø§Ù„ÙØ±ÙˆØ¹Â».",
                            parent=dlg,
                        )
                        return

                emit_sync = sync_mode != "none"
                sync_devs = selected_pos if sync_mode == "selected-pos" else None

                if multi:
                    # Always update the exact selected inventory rows.
                    # The inventory table can show grouped package text like
                    # "70, 163, 166, 219", so rebuilding a shared warehouse/package
                    # filter from the visible cells is not reliable.
                    for vals in rows:
                        updated_total += self.db.update_prices(
                            _visible_row_filter(vals),
                            new_price,
                            note="Price update (multi-selection)",
                            price_sync_mode=sync_mode,
                            price_sync_pos_devices=sync_devs,
                            emit_price_sync=emit_sync,
                        )
                else:
                    if scope_var.get() == "row":
                        updated_total = self.db.update_prices(
                            _visible_row_filter(first),
                            new_price,
                            note="Price update (single row)",
                            price_sync_mode=sync_mode,
                            price_sync_pos_devices=sync_devs,
                            emit_price_sync=emit_sync,
                        )
                    else:
                        pkg_value = _single_int_or_none(first[6])
                        if pkg_value is None:
                            messagebox.showwarning(
                                "Ù†Ø·Ø§Ù‚ ØºÙŠØ± ØµØ§Ù„Ø­",
                                "Ù‡Ø°Ø§ Ø§Ù„ØµÙ ÙŠØ¹Ø±Ø¶ Ø£ÙƒØ«Ø± Ù…Ù† Ø¹Ø¨ÙˆØ© Ù…Ø¬Ù…Ø¹Ø©ØŒ Ù„Ø°Ù„Ùƒ Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø§Ø³ØªØ®Ø¯Ø§Ù… Ø®ÙŠØ§Ø± Â«Ù†ÙØ³ Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©Â» Ù…Ø¹Ù‡. Ø§Ø³ØªØ®Ø¯Ù… Â«Ø§Ù„ØµÙ Ø§Ù„Ù…Ø­Ø¯Ø¯ ÙÙ‚Ø·Â».",
                                parent=dlg,
                            )
                            return
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
                    show_toast(self, "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ ØµÙÙˆÙ Ù…Ø·Ø§Ø¨Ù‚Ø©", bg="#f59e0b")
                else:
                    show_toast(self, f"ØªÙ… ØªØ­Ø¯ÙŠØ« Ø§Ù„Ø³Ø¹Ø± ÙÙŠ {updated_total} ØµÙ(ÙˆÙ)")
                self._refresh()

            except Exception as ex:
                messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ­Ø¯ÙŠØ«", str(ex), parent=dlg)

        # ---- Buttons (ONCE, ALWAYS) ----
        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))

        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ø­ÙØ¸", command=on_ok).pack(side=tk.RIGHT, padx=6)

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
                stock_id = parse_int_text(vals[0])
                if stock_id is None:
                    raise ValueError(f"Ù…Ø¹Ø±Ù ØµÙ ØºÙŠØ± ØµØ§Ù„Ø­: {vals[0]}")
                ids.append(stock_id)
            scope_text = f"Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø­Ø¯Ø¯Ø©: {len(ids)}"
            scope_mode = "ids"
        else:
            # fall back to package filters
            wh_txt = (self.f_wh.get() or "").strip()
            pkg_txt = (self.f_pkg.get() or "").strip()
            if not (wh_txt and pkg_txt and wh_txt.isdigit() and pkg_txt.isdigit()):
                messagebox.showwarning(
                    "Ø­Ø¯Ø¯ Ø§Ù„Ù†Ø·Ø§Ù‚",
                    "Ø§Ø®ØªØ± ØµÙÙˆÙÙ‹Ø§ Ù…Ù† Ø§Ù„Ø¬Ø¯ÙˆÙ„ Ø£Ùˆ Ø£Ø¯Ø®Ù„ (Ø§Ù„Ù…Ø®Ø²Ù†/Ø§Ù„Ø¹Ø¨ÙˆØ©) ÙÙŠ Ø£Ø¹Ù„Ù‰ Ø§Ù„Ø´Ø§Ø´Ø© Ù„ØªØ·Ø¨ÙŠÙ‚ Ø§Ù„ØªØ¹Ø¯ÙŠÙ„ Ø¹Ù„Ù‰ Ø§Ù„Ø¹Ø¨ÙˆØ© ÙƒÙ„Ù‡Ø§.",
                    parent=self,
                )
                return
            w = int(wh_txt); p = int(pkg_txt)
            scope_text = f"Ø§Ù„Ù†Ø·Ø§Ù‚: Ø§Ù„Ù…Ø®Ø²Ù† {w} / Ø§Ù„Ø¹Ø¨ÙˆØ© {p}"
            scope_mode = "pkg"
            ids = [w, p]  # stash just to reuse locals (won't be used as ids)

        # --- Dialog UI ---
        dlg = tk.Toplevel(self)
        dlg.title("ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ù…ÙˆØ§ØµÙØ§Øª")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=scope_text, font=("", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,8))
        ttk.Label(frm, text="Ø§ØªØ±Ùƒ Ø§Ù„Ø­Ù‚Ù„ ÙØ§Ø±ØºÙ‹Ø§ Ø¥Ø°Ø§ ÙƒÙ†Øª Ù„Ø§ ØªØ±ÙŠØ¯ ØªØºÙŠÙŠØ±Ù‡.").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,10))

        it_var = tk.StringVar()
        sc_var = tk.StringVar()
        cl_var = tk.StringVar()
        sz_var = tk.StringVar()

        ttk.Label(frm, text="Ø§Ù„Ù†ÙˆØ¹ (Ø¬Ø¯ÙŠØ¯):").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=it_var, width=28).grid(row=2, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Ø§Ù„Ù…Ø¯Ø±Ø³Ø© (Ø¬Ø¯ÙŠØ¯):").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=sc_var, width=28).grid(row=3, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Ø§Ù„Ù„ÙˆÙ† (Ø¬Ø¯ÙŠØ¯):").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=cl_var, width=28).grid(row=4, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(frm, text="Ø§Ù„Ù…Ù‚Ø§Ø³ (Ø¬Ø¯ÙŠØ¯):").grid(row=5, column=0, sticky="e", padx=6, pady=4)
        ttk.Entry(frm, textvariable=sz_var, width=28).grid(row=5, column=1, sticky="w", padx=6, pady=4)

        btns = ttk.Frame(frm); btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=dlg.destroy).pack(side=tk.RIGHT)

        def on_ok():
            # Gather non-empty changes
            changes = {}
            if it_var.get().strip(): changes["item_type"] = it_var.get().strip()
            if sc_var.get().strip(): changes["school"]    = sc_var.get().strip()
            if cl_var.get().strip(): changes["color"]     = cl_var.get().strip()
            if sz_var.get().strip(): changes["size"]      = sz_var.get().strip()

            if not changes:
                messagebox.showwarning("Ù„Ø§ ØªØºÙŠÙŠØ±Ø§Øª", "Ù„Ù… ØªÙØ¯Ø®Ù„ Ø£ÙŠ Ù‚ÙŠÙ… Ø¬Ø¯ÙŠØ¯Ø©.", parent=dlg)
                return

            try:
                if scope_mode == "ids":
                    updated = self.db.update_specs_by_ids(ids, **changes)
                else:
                    w, p = parse_int_text(self.f_wh.get()), parse_int_text(self.f_pkg.get())
                    if w is None or p is None:
                        raise ValueError("Invalid warehouse or package number")
                    updated = self.db.update_specs_in_package(w, p, **changes)

                dlg.destroy()
                if updated == 0:
                    show_toast(self, "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ ØµÙÙˆÙ Ù…Ø·Ø§Ø¨Ù‚Ø©", bg="#f59e0b")
                else:
                    show_toast(self, f"ØªÙ… ØªØ¹Ø¯ÙŠÙ„ Ø§Ù„Ù…ÙˆØ§ØµÙØ§Øª ÙÙŠ {updated} ØµÙ(ÙˆÙ)")
                self._refresh()
            except Exception as ex:
                messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ­Ø¯ÙŠØ«", str(ex), parent=dlg)

        ttk.Button(btns, text="Ø­ÙØ¸", command=on_ok).pack(side=tk.RIGHT, padx=6)

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
        ttk.Label(frow, text="Ø¨Ø­Ø«:").pack(side=tk.LEFT)
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
        ttk.Button(btns, text="Ù…Ø³Ø­ Ø§Ù„ØªØ­Ø¯ÙŠØ¯", command=self._clear_sel).pack(side=tk.LEFT)
        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=self._on_cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ù…ÙˆØ§ÙÙ‚", command=self._on_ok).pack(side=tk.RIGHT, padx=6)

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

    def show(self) -> Optional[List[str]]:
        """Backward-compatible alias for callers that still use show()."""
        return self.run()


# ------------------- Bills History Window -------------------

class BillsHistoryWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("Ø³Ø¬Ù„ Ø§Ù„ÙÙˆØ§ØªÙŠØ±")
        self.geometry("1000x560")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  Ø³Ø¬Ù„ Ø§Ù„ÙÙˆØ§ØªÙŠØ±", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="Ø³Ø¬Ù„ Ø§Ù„ÙÙˆØ§ØªÙŠØ±", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)
        _br = ttk.Button(top, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh); _br.pack(side=tk.RIGHT)
        ToolTip(_br, "ØªØ­Ø¯ÙŠØ« Ù‚Ø§Ø¦Ù…Ø© Ø§Ù„ÙÙˆØ§ØªÙŠØ±")
        _bo = ttk.Button(top, text="ÙØªØ­", command=self._print_selected); _bo.pack(side=tk.RIGHT, padx=8)
        ToolTip(_bo, "ÙØªØ­ Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ù„Ù„Ø·Ø¨Ø§Ø¹Ø©")
        _be = ttk.Button(top, text="ØªØµØ¯ÙŠØ± Ø§Ù„Ù…Ø­Ø¯Ø¯ Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„", command=self._export_selected); _be.pack(side=tk.RIGHT)
        ToolTip(_be, "ØªØµØ¯ÙŠØ± Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ø¥Ù„Ù‰ Ù…Ù„Ù Ø¥ÙƒØ³Ù„")

        # Bill lifecycle buttons
        _bc = ttk.Button(top, text="ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ù…Ø³ÙˆØ¯Ø©", command=self._confirm_draft); _bc.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bc, "ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ù…Ø³ÙˆØ¯Ø© ÙˆØ®ØµÙ… Ø§Ù„ÙƒÙ…ÙŠØ§Øª Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†")
        _bv = ttk.Button(top, text="Ø¥Ù„ØºØ§Ø¡ Ø§Ù„ÙØ§ØªÙˆØ±Ø©", command=self._void_bill); _bv.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bv, "Ø¥Ù„ØºØ§Ø¡ ÙØ§ØªÙˆØ±Ø© Ù…Ø¤ÙƒØ¯Ø© ÙˆØ¥Ø±Ø¬Ø§Ø¹ Ø§Ù„ÙƒÙ…ÙŠØ§Øª")
        _bdel = ttk.Button(top, text="Ø­Ø°Ù Ø§Ù„Ù…Ø³ÙˆØ¯Ø©", command=self._delete_draft); _bdel.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bdel, "Ø­Ø°Ù Ø§Ù„Ù…Ø³ÙˆØ¯Ø© Ù†Ù‡Ø§Ø¦ÙŠØ§Ù‹")
        _bret = ttk.Button(top, text="Ù…Ø±ØªØ¬Ø¹", command=self._process_return); _bret.pack(side=tk.RIGHT, padx=4)
        ToolTip(_bret, "Ù…Ø¹Ø§Ù„Ø¬Ø© Ù…Ø±ØªØ¬Ø¹ Ù„ÙØ§ØªÙˆØ±Ø© Ù…Ø¤ÙƒØ¯Ø©")

        bills_wrap = ttk.Frame(self)
        bills_wrap.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 6))
        self.bills_table = ttk.Treeview(
            bills_wrap, columns=("id", "created_at", "kind", "customer", "total", "status"), show="headings", height=10
        )
        for col, txt, w in [("id","Ø§Ù„Ù…Ø¹Ø±Ù‘Ù",80), ("created_at","Ø§Ù„ØªØ§Ø±ÙŠØ®",180), ("kind","Ø§Ù„Ù†ÙˆØ¹",120), ("customer","Ø§Ù„Ø¹Ù…ÙŠÙ„",200), ("total","Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ",120), ("status","Ø§Ù„Ø­Ø§Ù„Ø©",90)]:
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

        ttk.Label(self, text="Ø¨Ù†ÙˆØ¯ Ø§Ù„ÙØ§ØªÙˆØ±Ø©").pack(anchor="w", padx=8)

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
            ("type","Ø§Ù„Ù†ÙˆØ¹",140), ("school","Ø§Ù„Ù…Ø¯Ø±Ø³Ø©",160), ("color","Ø§Ù„Ù„ÙˆÙ†",80), ("size","Ø§Ù„Ù…Ù‚Ø§Ø³",70),
            ("origin","Ø§Ù„Ù…ØµØ¯Ø±",90), ("badge","Ø¨Ø§Ø¯Ø¬",60), ("wh","Ø§Ù„Ù…Ø®Ø²Ù†",50), ("pkg","Ø§Ù„Ø¹Ø¨ÙˆØ©",60), ("price","Ø§Ù„Ø³Ø¹Ø±",80),
            ("qty","Ø§Ù„ÙƒÙ…ÙŠØ©",60), ("total","Ø¥Ø¬Ù…Ø§Ù„ÙŠ ",100),
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
        status_map = {"DRAFT": "Ù…Ø³ÙˆØ¯Ø©", "CONFIRMED": "Ù…Ø¤ÙƒØ¯Ø©", "VOID": "Ù…Ù„ØºØ§Ø©"}
        for b in self.db.list_bills():
            status_txt = status_map.get(b.get("status", "CONFIRMED"), b.get("status", ""))
            kind_txt = "Ø´Ø­Ù† ÙØ±Ø¹" if b.get("bill_kind") == "BRANCH_SHIPMENT" else "ÙØ§ØªÙˆØ±Ø©"
            self.bills_table.insert(
                "", tk.END, iid=str(b["id"]),
                values=(b["id"], fmt_local_ts(b["created_at"], ""), kind_txt, b.get("customer") or "",
                        f"{format_money(float(b['total']))}", status_txt)
            )
        apply_zebra_tags(self.bills_table)
        # Color-code by status
        self.bills_table.tag_configure("draft", background="#fef3c7")
        self.bills_table.tag_configure("void", background="#fee2e2")
        for child in self.bills_table.get_children():
            vals = self.bills_table.item(child, "values")
            if len(vals) >= 5:
                if vals[4] == "Ù…Ø³ÙˆØ¯Ø©":
                    self.bills_table.item(child, tags=("draft",))
                elif vals[4] == "Ù…Ù„ØºØ§Ø©":
                    self.bills_table.item(child, tags=("void",))
        self.items_table.delete(*self.items_table.get_children())

    def _confirm_draft(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯", "Ù‡Ù„ ØªØ±ÙŠØ¯ ØªØ£ÙƒÙŠØ¯ Ù‡Ø°Ù‡ Ø§Ù„Ù…Ø³ÙˆØ¯Ø© ÙˆØ®ØµÙ… Ø§Ù„ÙƒÙ…ÙŠØ§Øª Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†ØŸ", parent=self):
            return
        try:
            self.db.confirm_draft_bill(bill_id)
            show_toast(self, f"ØªÙ… ØªØ£ÙƒÙŠØ¯ Ø§Ù„ÙØ§ØªÙˆØ±Ø© #{bill_id}")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ£ÙƒÙŠØ¯", str(ex), parent=self)

    def _void_bill(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø¥Ù„ØºØ§Ø¡", "Ù‡Ù„ ØªØ±ÙŠØ¯ Ø¥Ù„ØºØ§Ø¡ Ù‡Ø°Ù‡ Ø§Ù„ÙØ§ØªÙˆØ±Ø© ÙˆØ¥Ø±Ø¬Ø§Ø¹ Ø§Ù„ÙƒÙ…ÙŠØ§ØªØŸ", parent=self):
            return
        try:
            self.db.void_bill(bill_id)
            show_toast(self, f"ØªÙ… Ø¥Ù„ØºØ§Ø¡ Ø§Ù„ÙØ§ØªÙˆØ±Ø© #{bill_id}")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¥Ù„ØºØ§Ø¡", str(ex), parent=self)

    def _delete_draft(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø­Ø°Ù", "Ù‡Ù„ ØªØ±ÙŠØ¯ Ø­Ø°Ù Ù‡Ø°Ù‡ Ø§Ù„Ù…Ø³ÙˆØ¯Ø© Ù†Ù‡Ø§Ø¦ÙŠØ§Ù‹ØŸ", parent=self):
            return
        try:
            self.db.delete_draft_bill(bill_id)
            show_toast(self, f"ØªÙ… Ø­Ø°Ù Ø§Ù„Ù…Ø³ÙˆØ¯Ø© #{bill_id}")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­Ø°Ù", str(ex), parent=self)

    def _process_return(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        bill = None
        for b in self.db.list_bills():
            if int(b["id"]) == bill_id:
                bill = b
                break
        if not bill or bill.get("status") != "CONFIRMED":
            messagebox.showwarning("ØºÙŠØ± Ù…ØªØ§Ø­", "Ø§Ù„Ù…Ø±ØªØ¬Ø¹Ø§Øª Ù…ØªØ§Ø­Ø© ÙÙ‚Ø· Ù„Ù„ÙÙˆØ§ØªÙŠØ± Ø§Ù„Ù…Ø¤ÙƒØ¯Ø©.", parent=self)
            return
        ReturnDialog(self, self.db, bill_id, on_done=self._refresh)

    def _get_selected_bill_id(self) -> Optional[int]:
        sel = self.bills_table.selection()
        if not sel:
            return None
        return parse_int_text(sel[0])

    def _load_items(self):
        bill_id = self._get_selected_bill_id()
        self.items_table.delete(*self.items_table.get_children())
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ù‹Ø§.")
            return
        items = self.db.list_bill_items(bill_id)
        for ln in items:
            origin_txt = "Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†" if ln.get("origin") == "STOCK" else ("Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹" if ln.get("origin") == "FACTORY" else "")
            wh_txt  = "" if (ln.get("warehouse_no") in (None, "", 0, "0")) else ln.get("warehouse_no")
            pkg_txt = "" if (ln.get("package_no")  in (None, "", 0, "0")) else ln.get("package_no")
            self.items_table.insert(
                "", tk.END,
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"], origin_txt,
                        "âœ“" if int(ln.get("has_badge") or 0) else "",
                        wh_txt, pkg_txt, f"{format_money(float(ln['unit_price']))}",
                        ln["qty"], f"{format_money(float(ln['line_total']))}")
            )
        apply_zebra_tags(self.items_table)

    def _export_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ù‹Ø§.")
            return
        items = self.db.list_bill_items(bill_id)
        if not items:
            messagebox.showwarning("ÙØ§Ø±Øº", "Ù„Ø§ ØªØ­ØªÙˆÙŠ Ù‡Ø°Ù‡ Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ø¹Ù„Ù‰ Ø¨Ù†ÙˆØ¯.")
            return
        path = filedialog.asksaveasfilename(
            title="ØªØµØ¯ÙŠØ± Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"bill_{bill_id}.xlsx",
        )
        if not path:
            return
        headers = ["type", "school", "color", "size", "origin", "has_badge", "warehouse_no", "package_no", "unit_price", "qty", "line_total"]
        def _origin_txt(o: Optional[str]) -> str:
            return "Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†" if o == "STOCK" else ("Ù…Ù† Ø§Ù„Ù…ØµÙ†Ø¹" if o == "FACTORY" else "")
        rows = [
            [
                ln["item_type"], ln["school"], ln["color"], ln["size"],
                _origin_txt(ln.get("origin")), ("âœ“" if int(ln.get("has_badge") or 0) else ""),
                ln["warehouse_no"], ln["package_no"],
                float(ln["unit_price"]), int(ln["qty"]), float(ln["line_total"]),
            ]
            for ln in items
        ]
        try:
            export_to_excel(path, headers, rows)
            show_toast(self, "ØªÙ… ØªØµØ¯ÙŠØ± Ø§Ù„ÙØ§ØªÙˆØ±Ø© Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„ Ø¨Ù†Ø¬Ø§Ø­")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex))

    def _print_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙØ§ØªÙˆØ±Ø© Ø£ÙˆÙ„Ù‹Ø§.")
            return
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == bill_id)
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©", "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ø§Ù„ÙØ§ØªÙˆØ±Ø©.")
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
        self.title(f"Ù…Ø±ØªØ¬Ø¹ - ÙØ§ØªÙˆØ±Ø© #{bill_id}")
        self.geometry("800x500")
        self.transient(master)
        self.grab_set()
        self._items = self.db.list_bill_items(bill_id)
        self._build()

    def _build(self):
        ttk.Label(self, text=f"Ø§Ø®ØªØ± Ø§Ù„Ø£ØµÙ†Ø§Ù ÙˆØ§Ù„ÙƒÙ…ÙŠØ§Øª Ø§Ù„Ù…Ø±ØªØ¬Ø¹Ø© Ù…Ù† Ø§Ù„ÙØ§ØªÙˆØ±Ø© #{self.bill_id}",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=8)

        cols = ("idx", "type", "school", "color", "size", "qty_orig", "return_qty")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        for col, txt, w in [
            ("idx", "#", 40), ("type", "Ø§Ù„Ù†ÙˆØ¹", 140), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 140),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 80), ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 70),
            ("qty_orig", "Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ø£ØµÙ„ÙŠØ©", 90), ("return_qty", "ÙƒÙ…ÙŠØ© Ø§Ù„Ø¥Ø±Ø¬Ø§Ø¹", 90)
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
        ttk.Label(edit_frame, text="ÙƒÙ…ÙŠØ© Ø§Ù„Ø¥Ø±Ø¬Ø§Ø¹ Ù„Ù„Ø¨Ù†Ø¯ Ø§Ù„Ù…Ø­Ø¯Ø¯:").pack(side=tk.LEFT)
        self._ret_var = tk.StringVar(value="0")
        self._ret_entry = ttk.Entry(edit_frame, textvariable=self._ret_var, width=8)
        self._ret_entry.pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="ØªØ¹ÙŠÙŠÙ†", command=self._set_return_qty).pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="Ø¥Ø±Ø¬Ø§Ø¹ Ø§Ù„ÙƒÙ„", command=self._return_all).pack(side=tk.LEFT, padx=4)

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=self.destroy).pack(side=tk.LEFT)
        ttk.Button(btns, text="ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ù…Ø±ØªØ¬Ø¹", command=self._confirm).pack(side=tk.RIGHT)

    def _set_return_qty(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = parse_int_text(sel[0])
        if idx is None:
            return
        try:
            ret_qty = parse_int_text(self._ret_var.get())
            if ret_qty is None:
                raise ValueError
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
            ret_qty = parse_int_text(vals[6], 0) or 0
            if ret_qty <= 0:
                continue
            row_no = parse_int_text(vals[0])
            if row_no is None:
                continue
            idx = row_no - 1
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
            messagebox.showwarning("Ù„Ø§ Ø´ÙŠØ¡", "Ù„Ù… ÙŠØªÙ… ØªØ­Ø¯ÙŠØ¯ Ø£ÙŠ ÙƒÙ…ÙŠØ§Øª Ù„Ù„Ø¥Ø±Ø¬Ø§Ø¹.", parent=self)
            return
        try:
            ret_id = self.db.process_return(self.bill_id, return_lines)
            show_toast(self, f"ØªÙ… Ø§Ù„Ù…Ø±ØªØ¬Ø¹ #{ret_id} Ø¨Ù†Ø¬Ø§Ø­")
            self.destroy()
            if self.on_done:
                self.on_done()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ù…Ø±ØªØ¬Ø¹", str(ex), parent=self)

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
        elif d == "OUT":
            qty_out += q
            val_out += float(q) * up
        elif d == "ADJUST_OUT":
            qty_adj_out += q
        elif d == "RETURN_IN":
            qty_return_in += q
        elif d == "RESERVE":
            qty_reserve += q
            qty_out += q
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
        self.title("Ø³Ø¬Ù„ Ø§Ù„Ø­Ø±ÙƒØ§Øª")
        self.geometry("1180x620")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  Ø³Ø¬Ù„ Ø§Ù„Ø­Ø±ÙƒØ§Øª", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        top = ttk.LabelFrame(self, text="ØªØµÙ†ÙŠÙ")
        top.pack(fill=tk.X, padx=8, pady=8)
        self.ftype  = LabeledCombobox(top, "Ø§Ù„Ù†ÙˆØ¹",   self.db, "item_type");  self.ftype.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        self.fsch   = LabeledCombobox(top, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school");     self.fsch.grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        self.fclr   = LabeledCombobox(top, "Ø§Ù„Ù„ÙˆÙ†",   self.db, "color");      self.fclr.grid(row=0, column=2, padx=6, pady=4, sticky="ew")
        self.fsiz   = LabeledCombobox(top, "Ø§Ù„Ù…Ù‚Ø§Ø³",  self.db, "size");       self.fsiz.grid(row=0, column=3, padx=6, pady=4, sticky="ew")
        branch_vals = [""] + [branch_display_name(n) for n in DEFAULT_BRANCH_POS_NAMES]
        branch_map = {n: branch_display_name(n) for n in DEFAULT_BRANCH_POS_NAMES}
        self.fbranch = LabeledStaticCombo(top, "Ø§Ù„ÙØ±Ø¹", branch_vals, value_map=branch_map, width=18)
        self.fbranch.grid(row=0, column=4, padx=6, pady=4, sticky="ew")
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
        for widget in (self.ftype, self.fsch, self.fclr, self.fsiz, self.fbranch):
            for ev in ("<<ComboboxSelected>>", "<KeyRelease>"):
                widget.cb.bind(ev, lambda _e: (_refresh_filter_values(), self._refresh()), add="+")

        self.df = DateField(top, "Ù…Ù† (YYYY-MM-DD)"); self.df.grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.dt = DateField(top, "Ø¥Ù„Ù‰");             self.dt.grid(row=1, column=1, padx=6, pady=4, sticky="w")


        ttk.Label(top, text="Ø¨Ø­Ø«").grid(row=1, column=2, sticky="e", padx=4, pady=4)
        self.txt = ttk.Entry(top); self.txt.grid(row=1, column=3, sticky="ew", padx=6, pady=4)
        for w in (self.df.entry, self.dt.entry):
            w.bind("<Return>", lambda _e: self._refresh(), add="+")
            w.bind("<FocusOut>", lambda _e: self.after_idle(self._refresh), add="+")
        self.txt.bind("<Return>", lambda _e: self._refresh(), add="+")
        top.columnconfigure(3, weight=1)

        btns = ttk.Frame(top)
        btns.grid(row=0, column=5, rowspan=2, sticky="e", padx=6, pady=4)
        _mr = ttk.Button(btns, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh); _mr.pack(side=tk.LEFT)
        ToolTip(_mr, "ØªØ­Ø¯ÙŠØ« Ù‚Ø§Ø¦Ù…Ø© Ø§Ù„Ø­Ø±ÙƒØ§Øª")
        _mc = ttk.Button(btns, text="Ù…Ø³Ø­", command=self._clear); _mc.pack(side=tk.LEFT, padx=6)
        ToolTip(_mc, "Ù…Ø³Ø­ Ø¬Ù…ÙŠØ¹ Ø§Ù„ÙÙ„Ø§ØªØ±")
        _mp = ttk.Button(btns, text="Ø·Ø¨Ø§Ø¹Ø©", command=self._print_report); _mp.pack(side=tk.LEFT, padx=6)
        ToolTip(_mp, "Ø·Ø¨Ø§Ø¹Ø© Ø§Ù„ØªÙ‚Ø±ÙŠØ± Ø§Ù„ØªØ¬Ù…ÙŠØ¹ÙŠ Ø§Ù„Ù…Ø¹Ø±ÙˆØ¶")
        _me = ttk.Button(btns, text="ØªØµØ¯ÙŠØ± Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„", command=self._export); _me.pack(side=tk.LEFT)
        ToolTip(_me, "ØªØµØ¯ÙŠØ± Ø§Ù„Ø­Ø±ÙƒØ§Øª Ø¥Ù„Ù‰ Ù…Ù„Ù Ø¥ÙƒØ³Ù„")

        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        # NEW: add badge column
        self.table = ttk.Treeview(
            table_wrap,
            columns=("type","school","color","size","incoming","sold","remaining","requested","reserved"),
            show="headings",
            height=14,
        )
        for col, txt, w in [
            ("type","Ø§Ù„Ù†ÙˆØ¹",160), ("school","Ø§Ù„Ù…Ø¯Ø±Ø³Ø©",170), ("color","Ø§Ù„Ù„ÙˆÙ†",120), ("size","Ø§Ù„Ù…Ù‚Ø§Ø³",90),
            ("incoming","ÙˆØ§Ø±Ø¯",95), ("sold","Ù…Ø¨Ø§Ø¹",95), ("remaining","Ù…ØªØ¨Ù‚ÙŠ",95), ("requested","Ø§Ù„Ù…Ø·Ù„ÙˆØ¨",95), ("reserved","Ø­Ø¬ÙˆØ²Ø§Øª",90),
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

        sum_fr = ttk.LabelFrame(self, text="Ù…Ù„Ø®Øµ Ø§Ù„Ù†ØªØ§Ø¦Ø¬ Ø§Ù„Ù…Ø¹Ø±ÙˆØ¶Ø© (ØªØ¬Ù…ÙŠØ¹ÙŠ Ø­Ø³Ø¨ Ø§Ù„Ù…Ù†ØªØ¬)")
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
            "branch_device": self.fbranch.get() or None,
            "date_from": self.df.get() or None,           # DateField
            "date_to":   self.dt.get() or None,           # DateField
            "text":      (self.txt.get().strip() or None),
        }


    def _refresh(self):
        try:
            rows = self.db.list_movement_item_totals(self._filters())
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¨Ø­Ø«", str(ex), parent=self)
            self._summary_var.set("")
            return

        self.table.delete(*self.table.get_children())
        branch_selected = str(self.fbranch.get() or "").strip()
        use_branch_only = bool(branch_selected)
        for r in rows:
            sold_qty = int((r.get("sold_branch_qty") if use_branch_only else r.get("sold_total_qty")) or 0)
            remaining_qty = int((r.get("remaining_branch_qty") if use_branch_only else r.get("remaining_total_qty")) or 0)
            requested_qty = int(r.get("requested_qty") or 0)
            self.table.insert(
                "", tk.END,
                values=(
                    r.get("item_type",""), r.get("school",""), r.get("color",""), r.get("size",""),
                    int(r.get("incoming_qty") or 0),
                    sold_qty,
                    remaining_qty,
                    requested_qty,
                    int(r.get("reserved_qty") or 0),
                )
            )
        apply_zebra_tags(self.table)
        s = {
            "n": len(rows),
            "incoming_qty": sum(int(r.get("incoming_qty") or 0) for r in rows),
            "sold_total_qty": sum(int((r.get("sold_branch_qty") if use_branch_only else r.get("sold_total_qty")) or 0) for r in rows),
            "reserved_qty": sum(int(r.get("reserved_qty") or 0) for r in rows),
            "remaining_total_qty": sum(int((r.get("remaining_branch_qty") if use_branch_only else r.get("remaining_total_qty")) or 0) for r in rows),
            "requested_qty": sum(int(r.get("requested_qty") or 0) for r in rows),
        }
        sold_label = "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ù…Ø¨Ø§Ø¹ Ø§Ù„ÙØ±Ø¹" if use_branch_only else "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨Ø§Ø¹"
        rem_label = "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ù…ØªØ¨Ù‚ÙŠ Ø§Ù„ÙØ±Ø¹" if use_branch_only else "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…ØªØ¨Ù‚ÙŠ (Ø§Ù„Ù…Ø®Ø²Ù† + Ø§Ù„ÙØ±ÙˆØ¹)"
        req_label = "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨ Ù„Ù„ÙØ±Ø¹" if use_branch_only else "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨"
        branch_prefix = ""
        if use_branch_only:
            branch_prefix = f"Ø§Ù„ÙØ±Ø¹ Ø§Ù„Ù…Ø­Ø¯Ø¯: {branch_display_name(branch_selected)}  |  "
        self._summary_var.set(
            f"{branch_prefix}"
            f"Ø¹Ø¯Ø¯ Ø§Ù„Ù…Ù†ØªØ¬Ø§Øª Ø§Ù„Ù…Ø¹Ø±ÙˆØ¶Ø©: {s['n']}  |  "
            f"Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙˆØ§Ø±Ø¯: {s['incoming_qty']}  |  "
            f"{sold_label}: {s['sold_total_qty']}  |  "
            f"{rem_label}: {s['remaining_total_qty']}  |  "
            f"{req_label}: {s['requested_qty']}  |  "
            f"Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ø­Ø¬ÙˆØ²Ø§Øª: {s['reserved_qty']}"
        )

    def _clear(self):
        for w in (self.ftype, self.fsch, self.fclr, self.fsiz, self.fbranch):
            w.set("")
        self.df.set("")               # DateField
        self.dt.set("")               # DateField
        self.txt.delete(0, tk.END)
        self._refresh()


    def _export(self):
        rows = self.db.list_movement_item_totals(self._filters())
        if not rows:
            messagebox.showwarning("ÙØ§Ø±Øº", "Ù„Ø§ ØªÙˆØ¬Ø¯ ØµÙÙˆÙ Ù„Ù„ØªØµØ¯ÙŠØ±.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="ØªØµØ¯ÙŠØ± Ø§Ù„Ø­Ø±ÙƒØ§Øª Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"movement_totals_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        headers = [
            "item_type", "school", "color", "size",
            "incoming_qty", "sold_total_qty", "remaining_total_qty", "requested_qty", "reserved_qty",
        ]
        table = [[
            m.get("item_type",""), m.get("school",""), m.get("color",""), m.get("size",""),
            int(m.get("incoming_qty") or 0),
            int((m.get("sold_branch_qty") if str(self.fbranch.get() or "").strip() else m.get("sold_total_qty")) or 0),
            int((m.get("remaining_branch_qty") if str(self.fbranch.get() or "").strip() else m.get("remaining_total_qty")) or 0),
            int(m.get("requested_qty") or 0),
            int(m.get("reserved_qty") or 0),
        ] for m in rows]
        try:
            export_to_excel(path, headers, table)
            show_toast(self, "ØªÙ… Ø­ÙØ¸ Ø§Ù„Ø­Ø±ÙƒØ§Øª Ø¥Ù„Ù‰ Ø¥ÙƒØ³Ù„ Ø¨Ù†Ø¬Ø§Ø­")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex), parent=self)

    def _print_report(self):
        rows = self.db.list_movement_item_totals(self._filters())
        if not rows:
            messagebox.showwarning("ÙØ§Ø±Øº", "Ù„Ø§ ØªÙˆØ¬Ø¯ ØµÙÙˆÙ Ù„Ù„Ø·Ø¨Ø§Ø¹Ø©.", parent=self)
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
            ("Ø§Ù„ÙØ±Ø¹", "branch_device"),
            ("Ø§Ù„Ù†ÙˆØ¹", "item_type"),
            ("Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", "school"),
            ("Ø§Ù„Ù„ÙˆÙ†", "color"),
            ("Ø§Ù„Ù…Ù‚Ø§Ø³", "size"),
            ("Ù…Ù†", "date_from"),
            ("Ø¥Ù„Ù‰", "date_to"),
            ("Ø¨Ø­Ø«", "text"),
        ):
            val = str(f.get(key) or "").strip()
            if val:
                filter_bits.append(f"{label}: {_h(val)}")
        filters_html = " | ".join(filter_bits) if filter_bits else "Ø¨Ø¯ÙˆÙ† ÙÙ„Ø§ØªØ±"

        totals = {
            "incoming": sum(int(r.get("incoming_qty") or 0) for r in rows),
            "sold": sum(int((r.get("sold_branch_qty") if str(self.fbranch.get() or "").strip() else r.get("sold_total_qty")) or 0) for r in rows),
            "remaining": sum(int((r.get("remaining_branch_qty") if str(self.fbranch.get() or "").strip() else r.get("remaining_total_qty")) or 0) for r in rows),
            "requested": sum(int(r.get("requested_qty") or 0) for r in rows),
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
                f"<td class='num'>{int((r.get('sold_branch_qty') if str(self.fbranch.get() or '').strip() else r.get('sold_total_qty')) or 0)}</td>"
                f"<td class='num'>{int((r.get('remaining_branch_qty') if str(self.fbranch.get() or '').strip() else r.get('remaining_total_qty')) or 0)}</td>"
                f"<td class='num'>{int(r.get('requested_qty') or 0)}</td>"
                f"<td class='num'>{int(r.get('reserved_qty') or 0)}</td>"
                "</tr>"
            )

        html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<title>ØªÙ‚Ø±ÙŠØ± Ø­Ø±ÙƒØ© Ø§Ù„Ø£ØµÙ†Ø§Ù - ØªØ¬Ù…ÙŠØ¹ÙŠ</title>
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
    <div class="title">ØªÙ‚Ø±ÙŠØ± Ø­Ø±ÙƒØ© Ø§Ù„Ø£ØµÙ†Ø§Ù (ØªØ¬Ù…ÙŠØ¹ÙŠ)</div>
    <div class="meta">ÙˆÙ‚Øª Ø§Ù„ØªÙ‚Ø±ÙŠØ±: {_h(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>
    <div class="meta">Ø§Ù„ÙÙ„Ø§ØªØ±: {filters_html}</div>
    <div class="summary">
      Ø¹Ø¯Ø¯ Ø§Ù„Ù…Ù†ØªØ¬Ø§Øª: {len(rows)} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙˆØ§Ø±Ø¯: {totals['incoming']} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨Ø§Ø¹: {totals['sold']} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…ØªØ¨Ù‚ÙŠ: {totals['remaining']} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨: {totals['requested']} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ø­Ø¬ÙˆØ²Ø§Øª: {totals['reserved']}
    </div>
    <table>
      <thead>
        <tr>
          <th>Ø§Ù„Ù†ÙˆØ¹</th>
          <th>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©</th>
          <th>Ø§Ù„Ù„ÙˆÙ†</th>
          <th>Ø§Ù„Ù…Ù‚Ø§Ø³</th>
          <th>ÙˆØ§Ø±Ø¯</th>
          <th>Ù…Ø¨Ø§Ø¹</th>
          <th>Ù…ØªØ¨Ù‚ÙŠ</th>
          <th>Ø§Ù„Ù…Ø·Ù„ÙˆØ¨</th>
          <th>Ø­Ø¬ÙˆØ²Ø§Øª</th>
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
        self.title("Ø¥Ø¹Ø¯Ø§Ø¯Ø§Øª Ø§Ù„Ù…Ø¯ÙŠØ±")
        self.geometry("520x380")
        self.resizable(False, False)
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  Ø¥Ø¹Ø¯Ø§Ø¯Ø§Øª Ø§Ù„Ù…Ø¯ÙŠØ±", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        # Password gate at top
        gate = ttk.LabelFrame(self, text="Ø§Ù„ØªØ­Ù‚Ù‚ Ù…Ù† ÙƒÙ„Ù…Ø© Ø§Ù„Ù…Ø±ÙˆØ±")
        gate.pack(fill=tk.X, padx=10, pady=(10, 8))
        ttk.Label(gate, text="ÙƒÙ„Ù…Ø© Ø§Ù„Ù…Ø±ÙˆØ±:").grid(row=0, column=0, sticky="e", padx=6, pady=8)
        self.pw_var = tk.StringVar()
        pw_entry = ttk.Entry(gate, textvariable=self.pw_var, show="*")
        pw_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=8)
        gate.columnconfigure(1, weight=1)

        # Section: Reopen package
        rp = ttk.LabelFrame(self, text="Ø¥Ø¹Ø§Ø¯Ø© ÙØªØ­ Ø­Ø§ÙˆÙŠØ©")
        rp.pack(fill=tk.X, padx=10, pady=(0, 8))

        self.wh_var  = tk.StringVar()
        self.pkg_var = tk.StringVar()
        self.empty_only = tk.BooleanVar(value=False)

        ttk.Label(rp, text="Ø±Ù‚Ù… Ø§Ù„Ù…Ø®Ø²Ù†:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Entry(rp, textvariable=self.wh_var, width=10).grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(rp, text="Ø±Ù‚Ù… Ø§Ù„Ø¹Ø¨ÙˆØ©:").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        ttk.Entry(rp, textvariable=self.pkg_var, width=12).grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Checkbutton(rp, text="Ø¥Ø¹Ø§Ø¯Ø© Ø§Ù„ÙØªØ­ Ø¥Ø°Ø§ ÙƒØ§Ù†Øª ÙØ§Ø±ØºØ© ÙÙ‚Ø·", variable=self.empty_only).grid(
            row=1, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 6)
        )

        act = ttk.Frame(rp); act.grid(row=2, column=0, columnspan=4, sticky="e", padx=6, pady=(6, 6))
        _arp = ttk.Button(act, text="Ø¥Ø¹Ø§Ø¯Ø© ÙØªØ­", command=self._do_reopen); _arp.pack(side=tk.RIGHT)
        ToolTip(_arp, "Ø¥Ø¹Ø§Ø¯Ø© ÙØªØ­ Ø§Ù„Ø¹Ø¨ÙˆØ© Ø§Ù„Ù…ØºÙ„Ù‚Ø© Ù„Ù„Ø¥Ø¶Ø§ÙØ© Ø¹Ù„ÙŠÙ‡Ø§")

        # Future admin tools can go in more frames below...
            # -------- Purge definition (HARD DELETE) --------
        pd = ttk.LabelFrame(self, text="Ø­Ø°Ù ØªØ¹Ø±ÙŠÙ Ù†Ù‡Ø§Ø¦ÙŠ (Ø¥Ø²Ø§Ù„Ø© Ù…Ù† Ø§Ù„Ù†Ø¸Ø§Ù…)")
        pd.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.purge_field = tk.StringVar(value="color")
        self.purge_value = tk.StringVar()

        ttk.Label(pd, text="Ø§Ù„Ø­Ù‚Ù„:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Combobox(
            pd,
            textvariable=self.purge_field,
            values=["item_type", "school", "color", "size"],
            state="readonly",
            width=14
        ).grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(pd, text="Ø§Ù„Ù‚ÙŠÙ…Ø©:").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        ttk.Entry(pd, textvariable=self.purge_value, width=26)\
            .grid(row=0, column=3, sticky="w", padx=6, pady=6)

        _apd = ttk.Button(pd, text="Ø­Ø°Ù Ù†Ù‡Ø§Ø¦ÙŠ", command=self._purge_definition); _apd.grid(row=0, column=4, padx=8)
        ToolTip(_apd, "Ø­Ø°Ù Ø§Ù„Ù‚ÙŠÙ…Ø© Ù†Ù‡Ø§Ø¦ÙŠØ§Ù‹ Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† (Ù„Ø§ ÙŠÙ…ÙƒÙ† Ø§Ù„ØªØ±Ø§Ø¬Ø¹!)")


    def _do_reopen(self):
        # Verify password
        if not self.db.verify_admin_password(self.pw_var.get()):
            messagebox.showerror("Ù…Ø±ÙÙˆØ¶", "ÙƒÙ„Ù…Ø© Ø§Ù„Ù…Ø±ÙˆØ± ØºÙŠØ± ØµØ­ÙŠØ­Ø©.", parent=self)
            return

        # Parse inputs
        try:
            w = parse_int_text(self.wh_var.get(), 0) or 0
            p = parse_int_text(self.pkg_var.get(), 0) or 0
            if w < 1 or p < 1:
                raise ValueError
        except Exception:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ø£Ø¯Ø®Ù„ Ø£Ø±Ù‚Ø§Ù…Ù‹Ø§ ØµØ­ÙŠØ­Ø© Ù„Ù„Ù…Ø®Ø²Ù† ÙˆØ§Ù„Ø¹Ø¨ÙˆØ© (>= 1).", parent=self)
            return

        # Perform action
        try:
            self.db.reopen_package(w, p, require_empty=bool(self.empty_only.get()), note="Admin manual reopen")
            show_toast(self, f"ØªÙ…Øª Ø¥Ø¹Ø§Ø¯Ø© ÙØªØ­ Ø§Ù„Ø¹Ø¨ÙˆØ© {p} ÙÙŠ Ø§Ù„Ù…Ø®Ø²Ù† {w}")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„", str(ex), parent=self)
    def _purge_definition(self):
        # Verify admin password
        if not self.db.verify_admin_password(self.pw_var.get()):
            messagebox.showerror("Ù…Ø±ÙÙˆØ¶", "ÙƒÙ„Ù…Ø© Ø§Ù„Ù…Ø±ÙˆØ± ØºÙŠØ± ØµØ­ÙŠØ­Ø©.", parent=self)
            return

        field = self.purge_field.get()
        value = (self.purge_value.get() or "").strip()

        if not value:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", "Ø£Ø¯Ø®Ù„ Ù‚ÙŠÙ…Ø© Ù„Ù„Ø­Ø°Ù.", parent=self)
            return

        if not messagebox.askyesno(
            "ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø­Ø°Ù Ø§Ù„Ù†Ù‡Ø§Ø¦ÙŠ",
            f"Ø³ÙŠØªÙ… Ø­Ø°Ù '{value}' Ù…Ù† '{field}' Ù†Ù‡Ø§Ø¦ÙŠÙ‹Ø§.\n\n"
            "âš  Ù‡Ø°Ø§ Ø³ÙŠØ­Ø°Ù ÙƒÙ„ Ø§Ù„ØµÙÙˆÙ Ø§Ù„Ù…Ø±ØªØ¨Ø·Ø© Ø¨Ù‡ Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ†.\n\n"
            "Ù‡Ù„ Ø£Ù†Øª Ù…ØªØ£ÙƒØ¯ØŸ",
            parent=self
        ):
            return

        try:
            removed = self.db.purge_definition(field, value)
            if removed == 0:
                show_toast(self, "Ø§Ù„Ù‚ÙŠÙ…Ø© ØºÙŠØ± Ù…ÙˆØ¬ÙˆØ¯Ø© ÙÙŠ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", bg="#f59e0b")
            else:
                show_toast(self, f"ØªÙ… Ø­Ø°Ù {removed} ØµÙ(ÙˆÙ) Ù…Ù† Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ù†Ù‡Ø§Ø¦ÙŠØ§Ù‹")
            self.purge_value.set("")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­Ø°Ù", str(ex), parent=self)



# ------------------- Transfer Window -------------------

class TransferWindow(tk.Toplevel):
    """Select source stock items, choose destination warehouse/package, confirm transfer."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("ØªØ­ÙˆÙŠÙ„ Ø¨ÙŠÙ† Ø§Ù„Ù…Ø®Ø§Ø²Ù†")
        self.geometry("950x560")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  ØªØ­ÙˆÙŠÙ„ Ø¨ÙŠÙ† Ø§Ù„Ù…Ø®Ø§Ø²Ù†", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()

    def _build(self):
        # Source filters
        src_frame = ttk.LabelFrame(self, text="Ø§Ø®ØªØ± Ø§Ù„Ù…ØµØ¯Ø±")
        src_frame.pack(fill=tk.X, padx=8, pady=8)

        row1 = ttk.Frame(src_frame)
        row1.pack(fill=tk.X, padx=4, pady=4)
        self.f_type = LabeledCombobox(row1, "Ø§Ù„Ù†ÙˆØ¹", self.db, "item_type")
        self.f_type.pack(side=tk.LEFT, padx=4)
        self.f_school = LabeledCombobox(row1, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self.f_school.pack(side=tk.LEFT, padx=4)
        self.f_color = LabeledCombobox(row1, "Ø§Ù„Ù„ÙˆÙ†", self.db, "color")
        self.f_color.pack(side=tk.LEFT, padx=4)
        self.f_size = LabeledCombobox(row1, "Ø§Ù„Ù…Ù‚Ø§Ø³", self.db, "size")
        self.f_size.pack(side=tk.LEFT, padx=4)

        row2 = ttk.Frame(src_frame)
        row2.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(row2, text="Ø§Ù„Ù…Ø®Ø²Ù†:").pack(side=tk.LEFT)
        self.f_wh = ttk.Combobox(row2, values=["", "1", "2", "3", "4"], width=5, state="readonly")
        self.f_wh.pack(side=tk.LEFT, padx=4)
        ttk.Label(row2, text="Ø§Ù„Ø¹Ø¨ÙˆØ©:").pack(side=tk.LEFT)
        self.f_pkg = ttk.Entry(row2, width=8)
        self.f_pkg.pack(side=tk.LEFT, padx=4)
        _bsrc = ttk.Button(row2, text="Ø¨Ø­Ø«", command=self._load_source); _bsrc.pack(side=tk.LEFT, padx=8)
        ToolTip(_bsrc, "Ø¹Ø±Ø¶ Ø§Ù„Ø£ØµÙ†Ø§Ù Ø§Ù„Ù…ØªØ§Ø­Ø© Ù„Ù„ØªØ­ÙˆÙŠÙ„")

        # Source table
        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.src_table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "wh", "pkg", "badge", "price", "count"),
            show="headings", height=10)
        for col, txt, w in [
            ("id", "ID", 50), ("type", "Ø§Ù„Ù†ÙˆØ¹", 120), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 120),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 70), ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 60), ("wh", "Ø§Ù„Ù…Ø®Ø²Ù†", 50),
            ("pkg", "Ø§Ù„Ø¹Ø¨ÙˆØ©", 50), ("badge", "Ø¨Ø§Ø¯Ø¬", 40), ("price", "Ø§Ù„Ø³Ø¹Ø±", 70),
            ("count", "Ø§Ù„ÙƒÙ…ÙŠØ©", 60),
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
        dest_frame = ttk.LabelFrame(self, text="Ø§Ù„ÙˆØ¬Ù‡Ø©")
        dest_frame.pack(fill=tk.X, padx=8, pady=4)
        drow = ttk.Frame(dest_frame)
        drow.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(drow, text="Ù…Ø®Ø²Ù† Ø§Ù„ÙˆØ¬Ù‡Ø©:").pack(side=tk.LEFT)
        self.dest_wh = ttk.Combobox(drow, values=WAREHOUSE_NUMBER_DISPLAY_VALUES, width=14, state="readonly")
        self.dest_wh.pack(side=tk.LEFT, padx=4)
        ttk.Label(drow, text="Ø¹Ø¨ÙˆØ© Ø§Ù„ÙˆØ¬Ù‡Ø©:").pack(side=tk.LEFT)
        self.dest_pkg = ttk.Entry(drow, width=8)
        self.dest_pkg.pack(side=tk.LEFT, padx=4)
        ttk.Label(drow, text="Ø§Ù„ÙƒÙ…ÙŠØ©:").pack(side=tk.LEFT)
        self.transfer_qty = ttk.Entry(drow, width=8)
        self.transfer_qty.pack(side=tk.LEFT, padx=4)
        _bt = ttk.Button(drow, text="ØªØ­ÙˆÙŠÙ„", command=self._do_transfer); _bt.pack(side=tk.LEFT, padx=8)
        ToolTip(_bt, "ØªØ­ÙˆÙŠÙ„ Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø­Ø¯Ø¯Ø© Ø¥Ù„Ù‰ Ø§Ù„ÙˆØ¬Ù‡Ø©")

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
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¨Ø­Ø«", str(ex), parent=self)
            return
        self.src_table.delete(*self.src_table.get_children())
        for r in rows:
            self.src_table.insert("", tk.END, iid=str(r["id"]), values=(
                r["id"], r["item_type"], r["school"], r["color"], r["size"],
                r["warehouse_no"], r["package_no"],
                ("âœ“" if int(r.get("has_badge") or 0) else ""),
                f"{format_money(float(r['unit_price']))}", r["count"]
            ))
        apply_zebra_tags(self.src_table)

    def _do_transfer(self):
        sel = self.src_table.selection()
        if not sel:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ØµÙ Ù…Ø®Ø²ÙˆÙ† Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        stock_id = parse_int_text(sel[0])
        if stock_id is None:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ù…Ø¹Ø±Ù ØµÙ Ø§Ù„Ù…Ø®Ø²ÙˆÙ† ØºÙŠØ± ØµØ§Ù„Ø­.", parent=self)
            return
        try:
            dest_wh = int(warehouse_numeric_value(self.dest_wh.get()))
            dest_pkg = parse_int_text(self.dest_pkg.get())
            qty = parse_int_text(self.transfer_qty.get())
            if dest_pkg is None or qty is None:
                raise ValueError
            if dest_wh < 1 or dest_pkg < 1 or qty < 1:
                raise ValueError
        except (ValueError, TypeError):
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª ØºÙŠØ± ØµØ§Ù„Ø­Ø©",
                                 "Ø£Ø¯Ø®Ù„ Ù…Ø®Ø²Ù† ÙˆØ¹Ø¨ÙˆØ© ÙˆÙƒÙ…ÙŠØ© ØµØ­ÙŠØ­Ø© (>= 1).", parent=self)
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯ Ø§Ù„ØªØ­ÙˆÙŠÙ„",
                                   f"ØªØ­ÙˆÙŠÙ„ {qty} ÙˆØ­Ø¯Ø© Ø¥Ù„Ù‰ Ø§Ù„Ù…Ø®Ø²Ù† {dest_wh} / Ø§Ù„Ø¹Ø¨ÙˆØ© {dest_pkg}ØŸ",
                                   parent=self):
            return
        try:
            self.db.transfer_stock(stock_id, qty, dest_wh, dest_pkg)
            show_toast(self, f"ØªÙ… Ø§Ù„ØªØ­ÙˆÙŠÙ„ Ø¨Ù†Ø¬Ø§Ø­ ({qty} ÙˆØ­Ø¯Ø©)")
            self._load_source()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ­ÙˆÙŠÙ„", str(ex), parent=self)

# ------------------- Stock Audit Window -------------------

def _stock_audit_report_totals(rows: Sequence[Dict[str, Any]]) -> Tuple[int, float]:
    return (
        sum(int(r.get("diff") or 0) for r in rows),
        sum(float(r.get("diff_value") or 0.0) for r in rows),
    )


def _export_stock_audit_report_rows(path: str, rows: Sequence[Dict[str, Any]]) -> None:
    headers = [
        "id", "item_type", "school", "color", "size", "warehouse_no", "package_no",
        "expected", "actual", "diff", "unit_price", "diff_value",
    ]
    table = [[
        r.get("id") or "", r.get("item_type") or "", r.get("school") or "",
        r.get("color") or "", r.get("size") or "", r.get("warehouse_no") or "",
        r.get("package_no") or "", int(r.get("expected") or 0),
        int(r.get("actual") or 0), int(r.get("diff") or 0),
        float(r.get("unit_price") or 0.0), float(r.get("diff_value") or 0.0),
    ] for r in rows]
    total_qty, total_value = _stock_audit_report_totals(rows)
    table.append(["", "", "", "", "", "", "", "", "Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ", total_qty, "", total_value])
    export_to_excel(path, headers, table)


def _stock_audit_report_html(rows: Sequence[Dict[str, Any]], title: str = "ØªÙ‚Ø±ÙŠØ± ÙØ±ÙˆÙ‚ Ø§Ù„Ø¬Ø±Ø¯") -> str:
    total_qty, total_value = _stock_audit_report_totals(rows)
    body = []
    for r in rows:
        diff = int(r.get("diff") or 0)
        cls = "surplus" if diff > 0 else "deficit"
        body.append(
            f"<tr class='{cls}'>"
            f"<td>{_html(r.get('item_type') or '')}</td>"
            f"<td>{_html(r.get('school') or '')}</td>"
            f"<td>{_html(r.get('color') or '')}</td>"
            f"<td>{_html(r.get('size') or '')}</td>"
            f"<td>{_html(r.get('warehouse_no') or '')}</td>"
            f"<td>{_html(r.get('package_no') or '')}</td>"
            f"<td>{int(r.get('expected') or 0)}</td>"
            f"<td>{int(r.get('actual') or 0)}</td>"
            f"<td>{diff:+d}</td>"
            f"<td>{format_money(r.get('unit_price') or 0)}</td>"
            f"<td>{format_money(r.get('diff_value') or 0)}</td>"
            f"</tr>"
        )
    return f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<title>{_html(title)}</title>
<style>
@page {{ size: A4 landscape; margin: 10mm; }}
body {{ font-family: "Segoe UI", Tahoma, Arial, sans-serif; direction: rtl; }}
.title {{ font-size: 20px; font-weight: 700; margin-bottom: 4px; }}
.meta {{ color: #334155; margin-bottom: 10px; }}
table {{ border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 12px; }}
th, td {{ border: 1px solid #94a3b8; padding: 6px 4px; text-align: center; }}
th {{ background: #e2e8f0; }}
.surplus td {{ background: #dcfce7; }}
.deficit td {{ background: #fee2e2; }}
tfoot td {{ background: #f8fafc; font-weight: 700; }}
</style>
</head>
<body>
<div class="title">{_html(title)}</div>
<div class="meta">ØªØ§Ø±ÙŠØ® Ø§Ù„ØªÙ‚Ø±ÙŠØ±: {western_digits(datetime.now().strftime('%Y-%m-%d %H:%M:%S'))}</div>
<table>
<thead>
<tr>
<th>Ø§Ù„Ù†ÙˆØ¹</th><th>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©</th><th>Ø§Ù„Ù„ÙˆÙ†</th><th>Ø§Ù„Ù…Ù‚Ø§Ø³</th><th>Ø§Ù„Ù…Ø®Ø²Ù†</th><th>Ø§Ù„Ø¹Ø¨ÙˆØ©</th>
<th>Ø§Ù„Ù…ØªÙˆÙ‚Ø¹</th><th>Ø§Ù„ÙØ¹Ù„ÙŠ</th><th>Ø§Ù„ÙØ±Ù‚</th><th>Ø§Ù„Ø³Ø¹Ø±</th><th>Ù‚ÙŠÙ…Ø© Ø§Ù„ÙØ±Ù‚</th>
</tr>
</thead>
<tbody>
{''.join(body)}
</tbody>
<tfoot>
<tr>
<td colspan="8">Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ</td><td>{total_qty:+d}</td><td></td><td>{format_money(total_value)}</td>
</tr>
</tfoot>
</table>
</body>
</html>
"""


class StockAuditWindow(tk.Toplevel):
    """Physical stock count and adjustment."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("Ø¬Ø±Ø¯ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†")
        self.geometry("1050x620")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44); _wh.pack(fill=tk.X); _wh.pack_propagate(False)
        tk.Label(_wh, text="  Ø¬Ø±Ø¯ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._stock_rows: List[Dict[str, Any]] = []
        self._audit_rows_by_iid: Dict[str, Dict[str, Any]] = {}
        self._history_touched_keys: set[Tuple[str, str, str, str, str, str]] = set()
        self._current_touched_iids: set[str] = set()
        self._build()

    def _build(self):
        # Filters
        filters = ttk.LabelFrame(self, text="ØªØµÙÙŠØ© Ø§Ù„Ø¬Ø±Ø¯")
        filters.pack(fill=tk.X, padx=8, pady=8)
        frow = ttk.Frame(filters)
        frow.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(frow, text="Ø§Ù„Ù…Ø®Ø²Ù†:").pack(side=tk.LEFT)
        self.f_wh = ttk.Combobox(frow, values=["", *WAREHOUSE_NUMBER_DISPLAY_VALUES], width=14, state="readonly")
        self.f_wh.pack(side=tk.LEFT, padx=4)
        ttk.Label(frow, text="Ø§Ù„Ø¹Ø¨ÙˆØ©:").pack(side=tk.LEFT)
        self.f_pkg = ttk.Entry(frow, width=8)
        self.f_pkg.pack(side=tk.LEFT, padx=4)
        self.f_type = LabeledCombobox(frow, "Ø§Ù„Ù†ÙˆØ¹", self.db, "item_type")
        self.f_type.pack(side=tk.LEFT, padx=4)
        self.f_school = LabeledCombobox(frow, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self.f_school.pack(side=tk.LEFT, padx=4)
        self.f_color = LabeledCombobox(frow, "Ø§Ù„Ù„ÙˆÙ†", self.db, "color")
        self.f_color.pack(side=tk.LEFT, padx=4)

        def _constraints(exclude: Optional[str] = None) -> Dict[str, Any]:
            warehouse_no = parse_int_text(warehouse_numeric_value(self.f_wh.get()))
            package_no = parse_int_text(self.f_pkg.get().strip())
            values = {
                "item_type": self.f_type.get() or None,
                "school": self.f_school.get() or None,
                "color": self.f_color.get() or None,
                "warehouse_no": warehouse_no,
                "package_no": package_no,
            }
            if exclude:
                values[exclude] = None
            return values

        self.f_type.set_supplier(lambda: self.db.get_distinct_filtered("item_type", _constraints("item_type")))
        self.f_school.set_supplier(lambda: self.db.get_distinct_filtered("school", _constraints("school")))
        self.f_color.set_supplier(lambda: self.db.get_distinct_filtered("color", _constraints("color")))

        def _refresh_filter_options(*_, clear_invalid_color: bool = False):
            self.f_type.refresh_values()
            self.f_school.refresh_values()
            self.f_color.refresh_values()
            color = self.f_color.get()
            if clear_invalid_color and color and color not in self.db.get_distinct_filtered("color", _constraints("color")):
                self.f_color.set("")

        for widget in (self.f_type, self.f_school):
            widget.cb.bind("<<ComboboxSelected>>", lambda _e: _refresh_filter_options(clear_invalid_color=True), add="+")
            widget.cb.bind("<KeyRelease>", lambda _e: _refresh_filter_options(), add="+")
        self.f_color.cb.bind("<<ComboboxSelected>>", lambda _e: _refresh_filter_options(), add="+")
        self.f_color.cb.bind("<KeyRelease>", lambda _e: self.f_color.refresh_values(), add="+")
        self.f_wh.bind("<<ComboboxSelected>>", lambda _e: _refresh_filter_options(clear_invalid_color=True), add="+")
        self.f_pkg.bind("<KeyRelease>", lambda _e: _refresh_filter_options(clear_invalid_color=True), add="+")
        _bl = ttk.Button(frow, text="ØªØ­Ù…ÙŠÙ„ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", command=self._load_stock); _bl.pack(side=tk.LEFT, padx=8)
        ToolTip(_bl, "ØªØ­Ù…ÙŠÙ„ Ø§Ù„Ø£ØµÙ†Ø§Ù Ø§Ù„Ù…Ø·Ø§Ø¨Ù‚Ø© Ù„Ù„Ø¬Ø±Ø¯")

        # Audit table
        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "wh", "pkg",
                     "expected", "actual", "diff", "price", "diff_value"),
            show="headings", height=16, selectmode="browse")
        for col, txt, w in [
            ("id", "ID", 50), ("type", "Ø§Ù„Ù†ÙˆØ¹", 110), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 110),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 70), ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 60), ("wh", "Ø§Ù„Ù…Ø®Ø²Ù†", 50),
            ("pkg", "Ø§Ù„Ø¹Ø¨ÙˆØ©", 50), ("expected", "Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…ØªÙˆÙ‚Ø¹Ø©", 100),
            ("actual", "Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„ÙØ¹Ù„ÙŠØ©", 100), ("diff", "Ø§Ù„ÙØ±Ù‚", 70),
            ("price", "Ø§Ù„Ø³Ø¹Ø±", 80), ("diff_value", "Ù‚ÙŠÙ…Ø© Ø§Ù„ÙØ±Ù‚", 100),
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
        self.table.tag_configure("current_touched", background="#fef3c7")
        self.table.tag_configure("history_touched", background="#dbeafe")

        # Edit actual qty
        edit_frame = ttk.Frame(self)
        edit_frame.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(edit_frame, text="Ø£Ø²Ø±Ù‚: ØªÙ… Ø¬Ø±Ø¯Ù‡ Ø³Ø§Ø¨Ù‚Ø§Ù‹   Ø£ØµÙØ±: ØªÙ… ØªØ¹Ø¯ÙŠÙ„Ù‡ Ø§Ù„Ø¢Ù†").pack(side=tk.RIGHT, padx=8)
        ttk.Label(edit_frame, text="Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„ÙØ¹Ù„ÙŠØ© Ù„Ù„Ø¨Ù†Ø¯ Ø§Ù„Ù…Ø­Ø¯Ø¯:").pack(side=tk.LEFT)
        self._actual_var = tk.StringVar(value="0")
        self._actual_entry = ttk.Entry(edit_frame, textvariable=self._actual_var, width=8)
        self._actual_entry.pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="ØªØ¹ÙŠÙŠÙ†", command=self._set_actual).pack(side=tk.LEFT, padx=4)
        ttk.Button(edit_frame, text="Ù…Ø·Ø§Ø¨Ù‚ (Ø§Ù„ÙƒÙ„)", command=self._mark_all_matched).pack(side=tk.LEFT, padx=4)

        # Action buttons
        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(btns, text="ØªØ·Ø¨ÙŠÙ‚ Ø§Ù„ØªØ³ÙˆÙŠØ§Øª", command=self._apply).pack(side=tk.RIGHT, padx=4)
        ttk.Button(btns, text="ØªØµØ¯ÙŠØ± ØªÙ‚Ø±ÙŠØ± Ø§Ù„ÙØ±ÙˆÙ‚", command=self._export_conflict_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø·Ø¨Ø§Ø¹Ø© ØªÙ‚Ø±ÙŠØ± Ø§Ù„ÙØ±ÙˆÙ‚", command=self._print_conflict_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø³Ø¬Ù„ ØªÙ‚Ø§Ø±ÙŠØ± Ø§Ù„Ø¬Ø±Ø¯", command=self._open_report_history).pack(side=tk.LEFT, padx=4)

    def _filters(self) -> Dict[str, Any]:
        return {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "color": self.f_color.get() or None,
            "warehouse_no": warehouse_numeric_value(self.f_wh.get()) or None,
            "package_no": self.f_pkg.get().strip() or None,
        }

    def _size_labels_for_specs(self, item_type: str, school: str, color: str) -> List[str]:
        profile = self.db.get_size_profile(item_type, school, color)
        if not profile:
            return []
        r1_start, r1_end, r2_start, r2_end, has_alpha = profile
        labels = merged_numeric_size_labels_from_profile(r1_start, r1_end, r2_start, r2_end)
        if has_alpha:
            labels.extend(ALPHA_SIZES)
        labels = list(dict.fromkeys([str(label).strip() for label in labels if str(label).strip()]))
        labels.sort(key=warehouse_size_sort_key)
        return labels

    def _audit_key_for_row(self, row: Dict[str, Any]) -> Tuple[str, str, str, str, str, str]:
        return (
            str(row.get("item_type") or "").strip().casefold(),
            str(row.get("school") or "").strip().casefold(),
            str(row.get("color") or "").strip().casefold(),
            _normalize_size_label(row.get("size") or "").casefold(),
            str(row.get("warehouse_no") or "").strip(),
            str(row.get("package_no") or "").strip(),
        )

    def _expanded_audit_rows(
        self,
        rows: Sequence[Dict[str, Any]],
        filters: Optional[Dict[str, Any]] = None,
    ) -> List[Dict[str, Any]]:
        expanded: List[Dict[str, Any]] = []
        seen: set[Tuple[str, str, str, str, str, str]] = set()
        groups: Dict[Tuple[str, str, str, str, str], Dict[str, Any]] = {}
        group_sizes: Dict[Tuple[str, str, str, str, str], set[str]] = {}

        for row in rows:
            item_type = str(row.get("item_type") or "").strip()
            school = str(row.get("school") or "").strip()
            color = str(row.get("color") or "").strip()
            size = _normalize_size_label(row.get("size") or "")
            warehouse_no = str(row.get("warehouse_no") or "").strip()
            package_no = str(row.get("package_no") or "").strip()
            unit_price = float(row.get("unit_price") or 0)
            key = (item_type, school, color, warehouse_no, package_no, size)
            seen.add(key)
            expanded.append(dict(row, size=size, _synthetic=False))
            if parse_int_text(warehouse_no) is not None and parse_int_text(package_no) is not None:
                group_key = (item_type, school, color, warehouse_no, package_no)
                groups.setdefault(
                    group_key,
                    {
                        "item_type": item_type,
                        "school": school,
                        "color": color,
                        "warehouse_no": warehouse_no,
                        "package_no": package_no,
                        "unit_price": unit_price,
                        "has_badge": int(row.get("has_badge") or 0),
                    },
                )
                group_sizes.setdefault(group_key, set()).add(size)

        filters = filters or {}
        if not groups:
            item_type = str(filters.get("item_type") or "").strip()
            school = str(filters.get("school") or "").strip()
            color = str(filters.get("color") or "").strip()
            warehouse_no = parse_int_text(filters.get("warehouse_no"))
            package_no = parse_int_text(filters.get("package_no"))
            if item_type and school and color and warehouse_no is not None and package_no is not None:
                group_key = (item_type, school, color, str(warehouse_no), str(package_no))
                groups[group_key] = {
                    "item_type": item_type,
                    "school": school,
                    "color": color,
                    "warehouse_no": str(warehouse_no),
                    "package_no": str(package_no),
                    "unit_price": 0.0,
                    "has_badge": 0,
                }
                group_sizes.setdefault(group_key, set())

        for (item_type, school, color, warehouse_no, package_no), sample in groups.items():
            labels = self._size_labels_for_specs(item_type, school, color)
            if any(size and not size.isdigit() for size in group_sizes.get((item_type, school, color, warehouse_no, package_no), set())):
                labels = list(dict.fromkeys(labels + ALPHA_SIZES))
                labels.sort(key=warehouse_size_sort_key)
            for size in labels:
                key = (item_type, school, color, warehouse_no, package_no, _normalize_size_label(size))
                if key in seen:
                    continue
                unit_price = self.db.get_effective_price(item_type, school, color, size)
                if unit_price is None:
                    unit_price = sample.get("unit_price") or 0
                expanded.append({
                    "id": None,
                    "item_type": item_type,
                    "school": school,
                    "color": color,
                    "size": _normalize_size_label(size),
                    "warehouse_no": warehouse_no,
                    "package_no": package_no,
                    "unit_price": float(unit_price or 0),
                    "count": 0,
                    "value": 0.0,
                    "has_badge": int(sample.get("has_badge") or 0),
                    "_synthetic": True,
                })
                seen.add(key)

        expanded.sort(
            key=lambda r: (
                (r.get("school") or "").casefold(),
                warehouse_item_sort_key(r.get("item_type"), r.get("color")),
                warehouse_size_sort_key(r.get("size")),
                str(r.get("warehouse_no") or ""),
                str(r.get("package_no") or ""),
                1 if r.get("_synthetic") else 0,
            )
        )
        return expanded

    def _load_stock(self):
        filters = self._filters()
        try:
            rows = self.db.current_inventory(filters)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ­Ù…ÙŠÙ„", str(ex), parent=self)
            return
        try:
            self._history_touched_keys = self.db.stock_audit_touched_keys()
        except Exception:
            self._history_touched_keys = set()
        self._current_touched_iids = set()
        self._stock_rows = self._expanded_audit_rows(rows, filters)
        self._audit_rows_by_iid = {}
        self.table.delete(*self.table.get_children())
        for idx, r in enumerate(self._stock_rows):
            count = int(r["count"])
            price = float(r.get("unit_price") or 0)
            iid = str(r["id"]) if r.get("id") not in (None, "") else f"missing-{idx}"
            self._audit_rows_by_iid[iid] = dict(r)
            display_id = r.get("id") if r.get("id") not in (None, "") else ""
            self.table.insert("", tk.END, iid=iid, values=(
                display_id, r["item_type"], r["school"], r["color"], r["size"],
                r["warehouse_no"], r["package_no"], count, count, 0,
                format_money(price), format_money(0),
            ))
        apply_zebra_tags(self.table)
        self._retag_all_audit_rows()

    def _stock_row_by_id(self, stock_id: int) -> Dict[str, Any]:
        for row in self._stock_rows:
            if row.get("id") not in (None, "") and int(row.get("id") or 0) == int(stock_id):
                return row
        return {}

    def _retag_audit_row(self, iid: str):
        vals = self.table.item(iid, "values")
        row = self._audit_rows_by_iid.get(iid, {})
        diff = parse_int_text(vals[9], 0) if vals else 0
        if diff and diff > 0:
            tag = "surplus"
        elif diff and diff < 0:
            tag = "deficit"
        elif iid in self._current_touched_iids:
            tag = "current_touched"
        elif self._audit_key_for_row(row) in self._history_touched_keys:
            tag = "history_touched"
        else:
            siblings = list(self.table.get_children(""))
            idx = siblings.index(iid) if iid in siblings else 0
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
        self.table.item(iid, tags=(tag,))

    def _retag_all_audit_rows(self):
        for iid in self.table.get_children(""):
            self._retag_audit_row(iid)

    def _set_actual(self):
        sel = self.table.selection()
        if not sel:
            return
        iid = sel[0]
        try:
            actual = parse_int_text(self._actual_var.get())
            if actual is None or actual < 0:
                raise ValueError
        except ValueError:
            return
        vals = list(self.table.item(iid, "values"))
        expected = parse_int_text(vals[7], 0) or 0
        diff = actual - expected
        stock_row = self._audit_rows_by_iid.get(iid, {})
        unit_price = float(stock_row.get("unit_price") or parse_float_text(vals[10], 0) or 0)
        vals[8] = actual
        vals[9] = diff
        vals[11] = format_money(diff * unit_price)
        self.table.item(iid, values=vals)
        self._current_touched_iids.add(iid)
        self._retag_audit_row(iid)
        if diff != 0:
            self._auto_apply_row(iid)

    def _mark_all_matched(self):
        for child in self.table.get_children():
            vals = list(self.table.item(child, "values"))
            vals[8] = vals[7]
            vals[9] = 0
            vals[11] = format_money(0)
            self.table.item(child, values=vals)
        self._current_touched_iids = set()
        self._retag_all_audit_rows()

    def _conflict_report_rows(self) -> List[Dict[str, Any]]:
        report_rows: List[Dict[str, Any]] = []
        for child in self.table.get_children():
            vals = self.table.item(child, "values")
            if not vals:
                continue
            stock_row = self._audit_rows_by_iid.get(child, {})
            stock_id = parse_int_text(stock_row.get("id"))
            expected = parse_int_text(vals[7], 0) or 0
            actual = parse_int_text(vals[8], expected)
            if actual is None:
                continue
            diff = actual - expected
            if diff == 0:
                continue
            unit_price = float(stock_row.get("unit_price") or parse_float_text(vals[10], 0) or 0)
            diff_value = float(diff) * unit_price
            report_rows.append({
                "id": stock_id,
                "item_type": vals[1],
                "school": vals[2],
                "color": vals[3],
                "size": vals[4],
                "warehouse_no": vals[5],
                "package_no": vals[6],
                "expected": expected,
                "actual": actual,
                "diff": diff,
                "unit_price": unit_price,
                "diff_value": diff_value,
            })
        return report_rows

    def _save_current_report_snapshot(self, source: str) -> Tuple[Optional[int], List[Dict[str, Any]]]:
        rows = self._conflict_report_rows()
        if not rows:
            return None, []
        report_id = self.db.create_stock_audit_report(
            rows,
            filters=self._filters(),
            source=source,
        )
        return report_id, rows

    def _report_row_for_iid(self, iid: str) -> Optional[Dict[str, Any]]:
        vals = self.table.item(iid, "values")
        if not vals:
            return None
        stock_row = self._audit_rows_by_iid.get(iid, {})
        expected = parse_int_text(vals[7], 0) or 0
        actual = parse_int_text(vals[8], expected)
        if actual is None:
            return None
        diff = actual - expected
        if diff == 0:
            return None
        unit_price = float(stock_row.get("unit_price") or parse_float_text(vals[10], 0) or 0)
        return {
            "id": parse_int_text(stock_row.get("id")),
            "item_type": stock_row.get("item_type") or vals[1],
            "school": stock_row.get("school") or vals[2],
            "color": stock_row.get("color") or vals[3],
            "size": stock_row.get("size") or vals[4],
            "warehouse_no": stock_row.get("warehouse_no") or vals[5],
            "package_no": stock_row.get("package_no") or vals[6],
            "expected": expected,
            "actual": actual,
            "diff": diff,
            "unit_price": unit_price,
            "diff_value": float(diff) * unit_price,
        }

    def _adjustment_for_iid(self, iid: str) -> Optional[Dict[str, Any]]:
        row = self._report_row_for_iid(iid)
        if not row:
            return None
        stock_row = self._audit_rows_by_iid.get(iid, {})
        return {
            "stock_id": row.get("id"),
            "expected": row["expected"],
            "actual": row["actual"],
            "item_type": row["item_type"],
            "school": row["school"],
            "color": row["color"],
            "size": row["size"],
            "warehouse_no": row["warehouse_no"],
            "package_no": row["package_no"],
            "unit_price": row["unit_price"],
            "has_badge": int(stock_row.get("has_badge") or 0),
        }

    def _auto_apply_row(self, iid: str):
        report_row = self._report_row_for_iid(iid)
        adjustment = self._adjustment_for_iid(iid)
        if not report_row or not adjustment:
            return
        try:
            report_id = self.db.append_stock_audit_report_bucket(
                [report_row],
                filters=self._filters(),
                source="auto-equalization",
                bucket_key=now_iso()[:13],
            )
            count = self.db.apply_stock_adjustments([adjustment])
            if count:
                show_toast(self, f"ØªÙ… ØªØ·Ø¨ÙŠÙ‚ Ø§Ù„ØªØ³ÙˆÙŠØ© ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹ ÙˆØ­ÙØ¸ ØªÙ‚Ø±ÙŠØ± #{report_id}")
                self._load_stock()
            else:
                show_toast(self, "Ù„Ù… ÙŠØªÙ… ØªØ·Ø¨ÙŠÙ‚ Ø§Ù„ØªØ³ÙˆÙŠØ© ØªÙ„Ù‚Ø§Ø¦ÙŠØ§Ù‹", bg="#f59e0b")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ³ÙˆÙŠØ© Ø§Ù„ØªÙ„Ù‚Ø§Ø¦ÙŠØ©", str(ex), parent=self)

    def _open_report_history(self):
        StockAuditReportHistoryWindow(self, self.db)

    def _export_conflict_report(self):
        report_id, rows = self._save_current_report_snapshot("export")
        if not rows:
            show_toast(self, "Ù„Ø§ ØªÙˆØ¬Ø¯ ÙØ±ÙˆÙ‚ Ù„Ù„ØªÙ‚Ø±ÙŠØ±", bg="#f59e0b")
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            title="ØªØµØ¯ÙŠØ± ØªÙ‚Ø±ÙŠØ± ÙØ±ÙˆÙ‚ Ø§Ù„Ø¬Ø±Ø¯",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"stock_audit_conflicts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )
        if not path:
            return
        try:
            _export_stock_audit_report_rows(path, rows)
            suffix = f" #{report_id}" if report_id else ""
            show_toast(self, f"ØªÙ… Ø­ÙØ¸ ÙˆØªØµØ¯ÙŠØ± ØªÙ‚Ø±ÙŠØ± ÙØ±ÙˆÙ‚ Ø§Ù„Ø¬Ø±Ø¯{suffix}")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex), parent=self)

    def _print_conflict_report(self):
        report_id, rows = self._save_current_report_snapshot("print")
        if not rows:
            show_toast(self, "Ù„Ø§ ØªÙˆØ¬Ø¯ ÙØ±ÙˆÙ‚ Ù„Ù„Ø·Ø¨Ø§Ø¹Ø©", bg="#f59e0b")
            return
        title = f"ØªÙ‚Ø±ÙŠØ± ÙØ±ÙˆÙ‚ Ø§Ù„Ø¬Ø±Ø¯ #{report_id}" if report_id else "ØªÙ‚Ø±ÙŠØ± ÙØ±ÙˆÙ‚ Ø§Ù„Ø¬Ø±Ø¯"
        html = _stock_audit_report_html(rows, title=title)
        path = os.path.join(
            tempfile.gettempdir(),
            f"stock_audit_conflicts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
        )
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©", str(ex), parent=self)

    def _apply(self):
        snapshot_id, snapshot_rows = self._save_current_report_snapshot("equalization")
        adjustments = []
        for child in self.table.get_children():
            adjustment = self._adjustment_for_iid(child)
            if adjustment:
                adjustments.append(adjustment)
        if not adjustments:
            show_toast(self, "Ù„Ø§ ØªÙˆØ¬Ø¯ ØªØ³ÙˆÙŠØ§Øª Ù„Ù„ØªØ·Ø¨ÙŠÙ‚", bg="#f59e0b")
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯ Ø§Ù„ØªØ³ÙˆÙŠØ©",
                                   f"Ø³ÙŠØªÙ… ØªØ·Ø¨ÙŠÙ‚ {len(adjustments)} ØªØ³ÙˆÙŠØ© Ø¹Ù„Ù‰ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†.\nÙ‡Ù„ Ø£Ù†Øª Ù…ØªØ£ÙƒØ¯ØŸ",
                                   parent=self):
            return
        try:
            count = self.db.apply_stock_adjustments(adjustments)
            suffix = f" ÙˆØ­ÙØ¸ ØªÙ‚Ø±ÙŠØ± #{snapshot_id}" if snapshot_id and snapshot_rows else ""
            show_toast(self, f"ØªÙ… ØªØ·Ø¨ÙŠÙ‚ {count} ØªØ³ÙˆÙŠØ© Ø¨Ù†Ø¬Ø§Ø­{suffix}")
            self._load_stock()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ³ÙˆÙŠØ©", str(ex), parent=self)


class StockAuditReportHistoryWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("Ø³Ø¬Ù„ ØªÙ‚Ø§Ø±ÙŠØ± Ø§Ù„Ø¬Ø±Ø¯")
        self.geometry("980x620")
        self.configure(bg=_UI["BG"])
        self._reports: List[Dict[str, Any]] = []
        self._build()
        self._refresh()

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)
        ttk.Button(top, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="Ø§Ø³ØªÙŠØ±Ø§Ø¯ Ø§Ù„ØªØ³ÙˆÙŠØ§Øª Ø§Ù„Ù‚Ø¯ÙŠÙ…Ø©", command=self._backfill_legacy_reports).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="Ø·Ø¨Ø§Ø¹Ø© Ø§Ù„Ù…Ø­Ø¯Ø¯", command=self._print_selected).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="ØªØµØ¯ÙŠØ± Ø§Ù„Ù…Ø­Ø¯Ø¯", command=self._export_selected).pack(side=tk.LEFT, padx=4)

        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        cols = ("id", "created_at", "source", "lines", "total_diff", "total_value")
        self._tree = ttk.Treeview(wrap, columns=cols, show="headings", height=16)
        for col, txt, w in [
            ("id", "Ø±Ù‚Ù… Ø§Ù„ØªÙ‚Ø±ÙŠØ±", 90),
            ("created_at", "Ø§Ù„ØªØ§Ø±ÙŠØ®", 170),
            ("source", "Ø§Ù„Ø³Ø¨Ø¨", 110),
            ("lines", "Ø¹Ø¯Ø¯ Ø§Ù„ÙØ±ÙˆÙ‚", 100),
            ("total_diff", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙØ±Ù‚", 110),
            ("total_value", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚ÙŠÙ…Ø©", 130),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=ysb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        add_context_menu(self._tree)
        _bind_mousewheel(self._tree)

        detail_wrap = ttk.LabelFrame(self, text="ØªÙØ§ØµÙŠÙ„ Ø§Ù„ØªÙ‚Ø±ÙŠØ± Ø§Ù„Ù…Ø­Ø¯Ø¯")
        detail_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        dcols = ("type", "school", "color", "size", "wh", "pkg", "expected", "actual", "diff", "price", "value")
        self._detail = ttk.Treeview(detail_wrap, columns=dcols, show="headings", height=9)
        for col, txt, w in [
            ("type", "Ø§Ù„Ù†ÙˆØ¹", 120), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 120), ("color", "Ø§Ù„Ù„ÙˆÙ†", 110),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 70), ("wh", "Ø§Ù„Ù…Ø®Ø²Ù†", 70), ("pkg", "Ø§Ù„Ø¹Ø¨ÙˆØ©", 80),
            ("expected", "Ø§Ù„Ù…ØªÙˆÙ‚Ø¹", 80), ("actual", "Ø§Ù„ÙØ¹Ù„ÙŠ", 80), ("diff", "Ø§Ù„ÙØ±Ù‚", 70),
            ("price", "Ø§Ù„Ø³Ø¹Ø±", 80), ("value", "Ù‚ÙŠÙ…Ø© Ø§Ù„ÙØ±Ù‚", 100),
        ]:
            self._detail.heading(col, text=txt)
            self._detail.column(col, width=w, anchor="center")
        dysb = ttk.Scrollbar(detail_wrap, orient="vertical", command=self._detail.yview)
        self._detail.configure(yscrollcommand=dysb.set)
        self._detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        dysb.pack(side=tk.RIGHT, fill=tk.Y)
        self._detail.tag_configure("surplus", background="#dcfce7")
        self._detail.tag_configure("deficit", background="#fee2e2")
        self._tree.bind("<<TreeviewSelect>>", lambda _e: self._load_selected_details())

    def _source_label(self, source: Any) -> str:
        labels = {
            "equalization": "ØªØ³ÙˆÙŠØ©",
            "export": "ØªØµØ¯ÙŠØ±",
            "print": "Ø·Ø¨Ø§Ø¹Ø©",
            "manual": "ÙŠØ¯ÙˆÙŠ",
        }
        return labels.get(str(source or "").strip(), str(source or "").strip())

    def _refresh(self):
        self._reports = self.db.list_stock_audit_reports()
        self._tree.delete(*self._tree.get_children())
        for row in self._reports:
            rid = int(row.get("id") or 0)
            self._tree.insert(
                "",
                tk.END,
                iid=str(rid),
                values=(
                    rid,
                    fmt_local_ts(row.get("created_at") or ""),
                    self._source_label(row.get("source")),
                    int(row.get("line_count") or 0),
                    f"{int(row.get('total_diff') or 0):+d}",
                    format_money(row.get("total_value") or 0),
                ),
            )
        apply_zebra_tags(self._tree)
        self._detail.delete(*self._detail.get_children())

    def _backfill_legacy_reports(self):
        try:
            created = self.db.backfill_stock_audit_reports_from_movements()
            self._refresh()
            if created:
                show_toast(self, f"ØªÙ… Ø§Ø³ØªÙŠØ±Ø§Ø¯ {created} ØªÙ‚Ø±ÙŠØ± Ø¬Ø±Ø¯ Ù‚Ø¯ÙŠÙ…")
            else:
                show_toast(self, "Ù„Ø§ ØªÙˆØ¬Ø¯ ØªØ³ÙˆÙŠØ§Øª Ù‚Ø¯ÙŠÙ…Ø© Ø¬Ø¯ÙŠØ¯Ø© Ù„Ù„Ø§Ø³ØªÙŠØ±Ø§Ø¯", bg="#f59e0b")
        except Exception as ex:
            messagebox.showerror("Ø§Ø³ØªÙŠØ±Ø§Ø¯ Ø§Ù„ØªØ³ÙˆÙŠØ§Øª Ø§Ù„Ù‚Ø¯ÙŠÙ…Ø©", str(ex), parent=self)

    def _selected_report_id(self) -> Optional[int]:
        sel = self._tree.selection()
        if not sel:
            return None
        return parse_int_text(sel[0])

    def _selected_report(self) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]]]:
        rid = self._selected_report_id()
        if rid is None:
            messagebox.showwarning("Ø³Ø¬Ù„ ØªÙ‚Ø§Ø±ÙŠØ± Ø§Ù„Ø¬Ø±Ø¯", "Ø§Ø®ØªØ± ØªÙ‚Ø±ÙŠØ±Ø§Ù‹ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return None, []
        report, rows = self.db.get_stock_audit_report(rid)
        if report is None:
            messagebox.showerror("Ø³Ø¬Ù„ ØªÙ‚Ø§Ø±ÙŠØ± Ø§Ù„Ø¬Ø±Ø¯", "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ø§Ù„ØªÙ‚Ø±ÙŠØ±.", parent=self)
            return None, []
        return report, rows

    def _load_selected_details(self):
        report, rows = self._selected_report()
        if report is None:
            return
        self._detail.delete(*self._detail.get_children())
        for idx, row in enumerate(rows):
            diff = int(row.get("diff") or 0)
            tag = "surplus" if diff > 0 else "deficit"
            self._detail.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(
                    row.get("item_type") or "",
                    row.get("school") or "",
                    row.get("color") or "",
                    row.get("size") or "",
                    row.get("warehouse_no") or "",
                    row.get("package_no") or "",
                    int(row.get("expected") or 0),
                    int(row.get("actual") or 0),
                    f"{diff:+d}",
                    format_money(row.get("unit_price") or 0),
                    format_money(row.get("diff_value") or 0),
                ),
                tags=(tag,),
            )

    def _print_selected(self):
        report, rows = self._selected_report()
        if report is None:
            return
        title = f"ØªÙ‚Ø±ÙŠØ± ÙØ±ÙˆÙ‚ Ø§Ù„Ø¬Ø±Ø¯ #{report.get('id')} - {fmt_local_ts(report.get('created_at') or '')}"
        html = _stock_audit_report_html(rows, title=title)
        path = os.path.join(
            tempfile.gettempdir(),
            f"stock_audit_report_{int(report.get('id') or 0)}.html",
        )
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©", str(ex), parent=self)

    def _export_selected(self):
        report, rows = self._selected_report()
        if report is None:
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            title="ØªØµØ¯ÙŠØ± ØªÙ‚Ø±ÙŠØ± Ø¬Ø±Ø¯ Ù…Ø­ÙÙˆØ¸",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"stock_audit_report_{int(report.get('id') or 0)}.xlsx",
        )
        if not path:
            return
        try:
            _export_stock_audit_report_rows(path, rows)
            show_toast(self, "ØªÙ… ØªØµØ¯ÙŠØ± ØªÙ‚Ø±ÙŠØ± Ø§Ù„Ø¬Ø±Ø¯ Ø§Ù„Ù…Ø­ÙÙˆØ¸")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex), parent=self)

# ------------------- Branch Inventory Queue -------------------

class BranchInventoryQueueWindow(tk.Toplevel):
    """Review branch returns/transfers before deciding where they go."""

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("Ù…Ù†ØªØ¬Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ ØºÙŠØ± Ø§Ù„Ù…Ø¹Ø§Ù„Ø¬Ø©")
        self.geometry("1220x650")
        self.configure(bg=_UI["BG"])
        hdr = tk.Frame(self, bg=_UI["ACCENT"], height=44)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  Ù…Ù†ØªØ¬Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ ØºÙŠØ± Ø§Ù„Ù…Ø¹Ø§Ù„Ø¬Ø©", bg=_UI["ACCENT"], fg="#FFFFFF",
                 font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self._build()
        self._refresh()

    def _build(self):
        top = ttk.LabelFrame(self, text="Ø§Ù„Ù…Ø¹Ø§Ù„Ø¬Ø©")
        top.pack(fill=tk.X, padx=8, pady=8)

        ttk.Label(top, text="Ø§Ù„Ø­Ø§Ù„Ø©:").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self._status = tk.StringVar(value="PENDING")
        ttk.Combobox(top, textvariable=self._status,
                     values=["PENDING", "ASSIGNED", "DISCARDED", "REROUTED", ""],
                     state="readonly", width=16).grid(row=0, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="Ù…Ø®Ø²Ù†:").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self._wh_var = tk.StringVar(value=WAREHOUSE_NUMBER_LABELS.get("1", "1"))
        self._wh_cb = ttk.Combobox(
            top,
            textvariable=self._wh_var,
            values=list(WAREHOUSE_NUMBER_DISPLAY_VALUES),
            state="readonly",
            width=14,
        )
        self._wh_cb.grid(row=0, column=3, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="Ø¹Ø¨ÙˆØ©:").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        self._pkg_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._pkg_var, width=10).grid(row=0, column=5, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="Ø¥Ù„Ù‰ ÙØ±Ø¹:").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self._target_var = tk.StringVar(value="")
        self._target_cb = ttk.Combobox(top, textvariable=self._target_var, state="readonly", width=18)
        self._target_cb.grid(row=1, column=1, sticky="w", padx=6, pady=6)

        ttk.Label(top, text="Ù…Ù„Ø§Ø­Ø¸Ø©:").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self._note_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self._note_var, width=54).grid(
            row=1, column=3, columnspan=3, sticky="ew", padx=6, pady=6
        )
        top.columnconfigure(5, weight=1)

        btns = ttk.Frame(top)
        btns.grid(row=0, column=6, rowspan=2, sticky="ns", padx=8)
        ttk.Button(btns, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.TOP, fill=tk.X, pady=2)
        ttk.Button(btns, text="Ø¥Ø³Ù†Ø§Ø¯ Ù„Ù„Ù…Ø®Ø²ÙˆÙ†", command=self._assign_selected).pack(side=tk.TOP, fill=tk.X, pady=2)
        ttk.Button(btns, text="ØªØ­ÙˆÙŠÙ„ Ø¥Ù„Ù‰ ÙØ±Ø¹", command=self._reroute_selected).pack(side=tk.TOP, fill=tk.X, pady=2)
        ttk.Button(btns, text="Ø¥ØªÙ„Ø§Ù / Ù…Ø¹ÙŠØ¨", command=self._discard_selected).pack(side=tk.TOP, fill=tk.X, pady=2)

        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        cols = ("id", "kind", "source", "target", "type", "school", "color", "size", "qty", "price", "status", "note")
        self._tree = ttk.Treeview(wrap, columns=cols, show="headings")
        for col, txt, w in [
            ("id", "#", 55), ("kind", "Ø§Ù„Ù†ÙˆØ¹", 80), ("source", "Ø§Ù„Ù…Ø±Ø³Ù„", 100), ("target", "Ø§Ù„Ù‡Ø¯Ù", 100),
            ("type", "Ø§Ù„ØµÙ†Ù", 150), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 130), ("color", "Ø§Ù„Ù„ÙˆÙ†", 110), ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 70),
            ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 65), ("price", "Ø§Ù„Ø³Ø¹Ø±", 80), ("status", "Ø§Ù„Ø­Ø§Ù„Ø©", 100), ("note", "Ù…Ù„Ø§Ø­Ø¸Ø©", 220),
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
        self._wh_cb.bind("<<ComboboxSelected>>", lambda _e: self._set_next_package_for_selected_warehouse(force=True))

        self._status_var = tk.StringVar(value="")
        tk.Label(self, textvariable=self._status_var, bg=_UI["BG"], fg=_UI["TEXT_SEC"]).pack(fill=tk.X, padx=8, pady=(0, 8))

    def _selected_queue_id(self) -> Optional[int]:
        sel = self._tree.selection()
        if not sel:
            return None
        return parse_int_text(sel[0])

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

        self._set_next_package_for_selected_warehouse(force=False)

        req = branch_display_name(row.get("requested_target_device")) if row else ""
        if row and req in ui_names:
            self._target_var.set(req)
        elif self._target_var.get() not in ui_names:
            self._target_var.set(ui_names[0] if ui_names else "")

    def _refresh(self):
        rows = self.db.list_branch_inventory_queue(self._status.get() or None)
        self._tree.delete(*self._tree.get_children())
        for row in rows:
            kind = "Ù…Ø±ØªØ¬Ø¹" if row.get("queue_kind") == "RETURN" else "ØªØ­ÙˆÙŠÙ„"
            self._tree.insert(
                "", tk.END, iid=str(row["id"]),
                values=(
                    row["id"], kind, branch_display_name(row.get("source_device") or ""),
                    branch_display_name(row.get("requested_target_device") or ""), row.get("item_type") or "",
                    row.get("school") or "", row.get("color") or "", row.get("size") or "",
                    row.get("qty") or 0, f"{format_money(float(row.get('unit_price') or 0))}",
                    row.get("status") or "", row.get("note") or "",
                ),
            )
        apply_zebra_tags(self._tree)
        self._status_var.set(f"Ø¹Ø¯Ø¯ Ø§Ù„Ø¹Ù†Ø§ØµØ±: {len(rows)}")
        self._load_selected_defaults()

    def _selected_warehouse_no(self) -> int:
        raw = warehouse_numeric_value(self._wh_var.get())
        try:
            w = int((raw or "").strip())
        except Exception:
            w = 1
        return max(1, w)

    def _set_next_package_for_selected_warehouse(self, *, force: bool) -> None:
        if (self._pkg_var.get() or "").strip() and not force:
            return
        try:
            info = self.db.package_numbers_summary(self._selected_warehouse_no())
            self._pkg_var.set(str(info.get("next") or 1))
        except Exception:
            self._pkg_var.set("1")

    def _assign_selected(self):
        queue_id = self._selected_queue_id()
        if queue_id is None:
            messagebox.showwarning("ØªÙ†Ø¨ÙŠÙ‡", "Ø§Ø®ØªØ± Ø¹Ù†ØµØ±Ù‹Ø§ Ù…Ù† Ø§Ù„Ù‚Ø§Ø¦Ù…Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        try:
            w = self._selected_warehouse_no()
            p = parse_int_text(self._pkg_var.get(), 0) or 0
        except Exception:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ø£Ø¯Ø®Ù„ Ø±Ù‚Ù… Ù…Ø®Ø²Ù† ÙˆØ¹Ø¨ÙˆØ© ØµØ­ÙŠØ­ÙŠÙ†.", parent=self)
            return
        try:
            self.db.assign_branch_inventory_queue_item(queue_id, w, p, self._note_var.get())
            show_toast(self, "ØªÙ… Ø¥Ø³Ù†Ø§Ø¯ Ø§Ù„Ø¹Ù†ØµØ± Ø¥Ù„Ù‰ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†")
            self._note_var.set("")
            self._pkg_var.set("")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¥Ø³Ù†Ø§Ø¯", str(ex), parent=self)

    def _discard_selected(self):
        queue_id = self._selected_queue_id()
        if queue_id is None:
            messagebox.showwarning("ØªÙ†Ø¨ÙŠÙ‡", "Ø§Ø®ØªØ± Ø¹Ù†ØµØ±Ù‹Ø§ Ù…Ù† Ø§Ù„Ù‚Ø§Ø¦Ù…Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯", "Ø³ÙŠØªÙ… Ø¥ØªÙ„Ø§Ù Ù‡Ø°Ø§ Ø§Ù„Ø¹Ù†ØµØ± ÙˆÙ„Ù† ÙŠØ¯Ø®Ù„ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†. Ù‡Ù„ Ø£Ù†Øª Ù…ØªØ£ÙƒØ¯ØŸ", parent=self):
            return
        try:
            self.db.discard_branch_inventory_queue_item(queue_id, self._note_var.get())
            show_toast(self, "ØªÙ… Ø¥ØªÙ„Ø§Ù / ØªØ¹Ù„ÙŠÙ… Ø§Ù„Ø¹Ù†ØµØ± ÙƒÙ…Ø¹ÙŠØ¨")
            self._note_var.set("")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø¥ØªÙ„Ø§Ù", str(ex), parent=self)

    def _reroute_selected(self):
        queue_id = self._selected_queue_id()
        if queue_id is None:
            messagebox.showwarning("ØªÙ†Ø¨ÙŠÙ‡", "Ø§Ø®ØªØ± Ø¹Ù†ØµØ±Ù‹Ø§ Ù…Ù† Ø§Ù„Ù‚Ø§Ø¦Ù…Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        target_ui = (self._target_var.get() or "").strip()
        target = getattr(self, "_target_ui_to_dev", {}).get(target_ui, target_ui)
        if not target:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª Ù†Ø§Ù‚ØµØ©", "Ø§Ø®ØªØ± Ø§Ù„ÙØ±Ø¹ Ø§Ù„Ù‡Ø¯Ù Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        try:
            self.db.reroute_branch_inventory_queue_item(queue_id, target, self._note_var.get())
            show_toast(self, f"ØªÙ… ØªØ­ÙˆÙŠÙ„ Ø§Ù„Ø¹Ù†ØµØ± Ø¥Ù„Ù‰ Ø§Ù„ÙØ±Ø¹ {branch_display_name(target)}")
            self._note_var.set("")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØ­ÙˆÙŠÙ„", str(ex), parent=self)

# ------------------- Branch Stock View (Phase 3) -------------------

class BranchStockWindow(tk.Toplevel):
    """Warehouse-side view of each POS device's current stock.

    Reads from `pos_stocks_mirror` (filled by POS_STOCK_SNAPSHOT
    events). Shows a dropdown of known POS devices with their snapshot
    timestamps, and a filterable Treeview of the selected POS's stock.
    Includes a "Ù…Ø²Ø§Ù…Ù†Ø© Ø§Ù„Ø¢Ù†" shortcut so the user can refresh without
    leaving the window.
    """

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("Ù…Ø®Ø²ÙˆÙ† Ø§Ù„ÙØ±ÙˆØ¹")
        self.geometry("1000x620")
        self.configure(bg=_UI["BG"])

        header = tk.Frame(self, bg=_UI["ACCENT"], height=44)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        tk.Label(
            header, text="  Ù…Ø®Ø²ÙˆÙ† Ø§Ù„ÙØ±ÙˆØ¹",
            bg=_UI["ACCENT"], fg="#FFFFFF",
            font=_FONTS["h3"],
        ).pack(side=tk.RIGHT, padx=12)

        self._build()
        self._reload_devices()

    # ---- layout ----

    def _build(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Ø§Ù„ÙØ±Ø¹:").pack(side=tk.RIGHT, padx=(0, 4))
        self._device_var = tk.StringVar()
        self._device_cb = ttk.Combobox(
            top, textvariable=self._device_var,
            state="readonly", width=28,
        )
        self._device_cb.pack(side=tk.RIGHT, padx=4)
        self._device_cb.bind("<<ComboboxSelected>>", lambda _e: self._reload_stock())

        self._meta_var = tk.StringVar(value="â€”")
        ttk.Label(top, textvariable=self._meta_var,
                  foreground="#64748b").pack(side=tk.RIGHT, padx=(12, 0))

        ttk.Button(top, text="ØªØ­Ø¯ÙŠØ«",
                   command=self._reload_devices).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="Ù…Ø²Ø§Ù…Ù†Ø© Ø§Ù„Ø¢Ù†",
                   command=self._run_sync_and_reload).pack(side=tk.LEFT, padx=4)

        # Filter row
        filt = ttk.LabelFrame(self, text="ØªØµÙ†ÙŠÙ")
        filt.pack(fill=tk.X, padx=8, pady=(0, 4))

        self._flt_type = LabeledCombobox(filt, "Ø§Ù„Ù†ÙˆØ¹", self.db, "item_type")
        self._flt_school = LabeledCombobox(filt, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self._flt_color = LabeledCombobox(filt, "Ø§Ù„Ù„ÙˆÙ†", self.db, "color")
        self._flt_size = LabeledCombobox(filt, "Ø§Ù„Ù…Ù‚Ø§Ø³", self.db, "size")

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

        ttk.Label(filt, text="Ø¨Ø­Ø«:").grid(row=1, column=3, sticky="e", padx=(6, 4), pady=(6, 6))
        self._filter_var = tk.StringVar()
        ent = ttk.Entry(filt, textvariable=self._filter_var, width=28)
        ent.grid(row=1, column=2, sticky="ew", padx=6, pady=(6, 6))
        self._filter_var.trace_add("write", lambda *_: self._apply_filter())

        btns = ttk.Frame(filt)
        btns.grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 6))
        ttk.Button(btns, text="Ø¨Ø­Ø«", command=self._apply_filter).pack(side=tk.LEFT)
        ttk.Button(btns, text="Ù…Ø³Ø­", command=self._clear_filters).pack(side=tk.LEFT, padx=8)

        # Table
        cols = ("item_type", "school", "color", "size", "unit_price", "count", "value")
        headers = {
            "item_type": "Ø§Ù„Ù†ÙˆØ¹", "school": "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", "color": "Ø§Ù„Ù„ÙˆÙ†",
            "size": "Ø§Ù„Ù…Ù‚Ø§Ø³", "unit_price": "Ø§Ù„Ø³Ø¹Ø±", "count": "Ø§Ù„ÙƒÙ…ÙŠØ©",
            "value": "Ø§Ù„Ù‚ÙŠÙ…Ø©",
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
            self._meta_var.set("Ù„Ø§ ØªÙˆØ¬Ø¯ ÙØ±ÙˆØ¹ Ù…ÙØ³Ø¬ÙŽÙ‘Ù„Ø© â€” Ø´ØºÙ‘Ù„ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© Ø£ÙˆÙ„Ø§Ù‹")
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
                f"Ø¢Ø®Ø± Ù„Ù‚Ø·Ø©: {fmt_local_ts(meta['snapshot_at'])}  |  "
                f"Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ: {meta['row_count']}  |  "
                f"Ø§Ù„Ù‚ÙŠÙ…Ø©: {format_money(meta['total_value'])}"
            )
        else:
            self._meta_var.set("Ù„Ø§ ØªÙˆØ¬Ø¯ Ù„Ù‚Ø·Ø© Ù…Ø®Ø²ÙˆÙ† Ø¨Ø¹Ø¯ Ù„Ù‡Ø°Ø§ Ø§Ù„ÙØ±Ø¹")

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
                    it, sc, cl, sz, f"{format_money(price)}", count, f"{format_money(value)}",
                ),
            )
            shown += 1
        self._status_var.set(
            f"ÙŠÙØ¹Ø±Ø¶ {shown} ØµÙ  |  Ø§Ù„ÙƒÙ…ÙŠØ©: {total_qty}  |  "
            f"Ø§Ù„Ù‚ÙŠÙ…Ø©: {format_money(total_val)}"
        )

    def _run_sync_and_reload(self):
        # Open the existing sync dialog so the user sees the live log
        # and control buttons rather than blocking this window.
        try:
            import sync_ui
            sync_ui.open_sync_dialog(self, self.db.conn)
        except Exception as ex:
            messagebox.showerror("Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", str(ex), parent=self)
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
        tk.Label(title_f, text="Ù„ÙˆØ­Ø© Ø§Ù„ØªØ­ÙƒÙ…", bg=_UI["SURFACE"], fg=_UI["TEXT"],
                 font=_FONTS["h1"]).pack(side=tk.RIGHT)
        tk.Label(title_f, text="Ù†Ø¸Ø±Ø© Ø¹Ø§Ù…Ø© Ø¹Ù„Ù‰ Ø§Ù„Ù…Ø®Ø²ÙˆÙ† ÙˆØ§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª", bg=_UI["SURFACE"],
                 fg=_UI["TEXT_SEC"], font=_FONTS["small"]).pack(side=tk.RIGHT, padx=(12, 0))
        ttk.Button(title_f, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT)

        # Stats cards row
        cards_row = tk.Frame(self, bg=_UI["BG"])
        cards_row.pack(fill=tk.X, padx=12, pady=(8, 12))

        self._cards = {}
        card_defs = [
            ("total_items", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ø£ØµÙ†Ø§Ù", _UI["BRAND"]),
            ("total_qty", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙƒÙ…ÙŠØ§Øª", _UI["OK"]),
            ("total_value", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚ÙŠÙ…Ø©", _UI["WARN"]),
            ("total_bills", "Ø¹Ø¯Ø¯ Ø§Ù„ÙÙˆØ§ØªÙŠØ±", "#8B5CF6"),
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
        tk.Label(r_card, text="Ø¢Ø®Ø± Ø§Ù„ÙÙˆØ§ØªÙŠØ±", bg=_UI["SURFACE"], fg=_UI["TEXT"],
                 font=_FONTS["h3"]).pack(anchor="e", padx=12, pady=(10, 4))

        self.recent_tree = ttk.Treeview(r_card, columns=("id", "date", "customer", "total"),
                                         show="headings", height=10)
        for col, txt, w in [("id", "#", 50), ("date", "Ø§Ù„ØªØ§Ø±ÙŠØ®", 140),
                             ("customer", "Ø§Ù„Ø¹Ù…ÙŠÙ„", 180), ("total", "Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ", 100)]:
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
        tk.Label(a_card, text="ØªÙ†Ø¨ÙŠÙ‡Ø§Øª Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø§Ù„Ù…Ù†Ø®ÙØ¶ (Ø£Ù‚Ù„ Ù…Ù† 5)", bg=_UI["SURFACE"],
                 fg=_UI["TEXT"], font=_FONTS["h3"]).pack(anchor="e", padx=12, pady=(10, 4))

        self.alert_tree = ttk.Treeview(a_card, columns=("item", "school", "color", "size", "qty"),
                                        show="headings", height=10)
        for col, txt, w in [("item", "Ø§Ù„Ù†ÙˆØ¹", 120), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 120),
                            ("color", "Ø§Ù„Ù„ÙˆÙ†", 80), ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 60), ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 60)]:
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
            text="ØµØ­Ø© Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© (Ø§Ù„ÙØ±ÙˆØ¹)",
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
            text="Ù…Ø²Ø§Ù…Ù†Ø© Ø§Ù„Ø¢Ù†",
            command=(lambda: self.app._open_sync_dialog()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="Ø·Ø§Ø¨ÙˆØ± Ø§Ù„ÙØ±ÙˆØ¹",
            command=(lambda: self.app._open_branch_inventory_queue()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="Ù…Ø±Ø¢Ø© Ø­Ø¬ÙˆØ²Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹",
            command=(lambda: self.app._open_pos_reservations_mirror()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        self.sync_tree = ttk.Treeview(
            s_card,
            columns=("branch", "status", "last_snapshot", "age_min", "rows", "value", "sync_errors"),
            show="headings",
            height=6,
        )
        for col, txt, w in [
            ("branch", "Ø§Ù„ÙØ±Ø¹", 180),
            ("status", "Ø§Ù„Ø­Ø§Ù„Ø©", 100),
            ("last_snapshot", "Ø¢Ø®Ø± Ù„Ù‚Ø·Ø©", 190),
            ("age_min", "Ø¹Ù…Ø± Ø§Ù„Ù„Ù‚Ø·Ø© (Ø¯)", 110),
            ("rows", "Ø§Ù„ØµÙÙˆÙ", 90),
            ("value", "Ø§Ù„Ù‚ÙŠÙ…Ø©", 120),
            ("sync_errors", "Ø£Ø®Ø·Ø§Ø¡ Ù…Ø²Ø§Ù…Ù†Ø©", 120),
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
            text="Ø§Ø³ØªØ«Ù†Ø§Ø¡Ø§Øª Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© (Ø£Ø¹Ù„Ù‰ Ø§Ù„Ø£Ø³Ø¨Ø§Ø¨)",
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
            ("etype", "Ù†ÙˆØ¹ Ø§Ù„Ø­Ø¯Ø«", 220),
            ("count", "Ø¹Ø¯Ø¯ Ø§Ù„Ø£Ø®Ø·Ø§Ø¡", 100),
            ("last_error", "Ø¢Ø®Ø± Ø®Ø·Ø£", 760),
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
            self._cards["total_value"].config(text=f"{format_money(val)}")

            cur.execute("SELECT COUNT(*) FROM bills WHERE COALESCE(status,'CONFIRMED')='CONFIRMED'")
            self._cards["total_bills"].config(text=str(cur.fetchone()[0] or 0))

            # Recent bills
            self.recent_tree.delete(*self.recent_tree.get_children())
            cur.execute("SELECT id, created_at, customer, total FROM bills WHERE COALESCE(status,'CONFIRMED')='CONFIRMED' ORDER BY id DESC LIMIT 10")
            for r in cur.fetchall():
                self.recent_tree.insert("", tk.END, values=(r[0], r[1], r[2] or "", f"{format_money(r[3])}"))
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
                    status = "ØºÙŠØ± Ù…Ø¹Ø±ÙˆÙ"
                    warn_count += 1
                elif age_min >= 60:
                    status = "Ø­Ø±Ø¬"
                    critical_count += 1
                elif age_min >= 15:
                    status = "ØªØ­Ø°ÙŠØ±"
                    warn_count += 1
                else:
                    status = "Ø¬ÙŠØ¯"
                    ok_count += 1
                self.sync_tree.insert(
                    "",
                    tk.END,
                    values=(
                        branch_display_name(r[0] or ""),
                        status,
                        fmt_local_ts(snap, ""),
                        age_min if age_min is not None else "",
                        int(r[2] or 0),
                        f"{format_money(float(r[3] or 0.0))}",
                        "â€”",
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
                f"ÙØ±ÙˆØ¹ Ø³Ù„ÙŠÙ…Ø©: {ok_count}  |  ØªØ­Ø°ÙŠØ±: {warn_count}  |  Ø­Ø±Ø¬: {critical_count}  |  "
                f"Ø£Ø­Ø¯Ø§Ø« ØµØ§Ø¯Ø±Ø© Ù‚ÙŠØ¯ Ø§Ù„Ø§Ù†ØªØ¸Ø§Ø±: {outbox_pending}  |  Ø£Ø®Ø·Ø§Ø¡ ØªØ·Ø¨ÙŠÙ‚ ÙˆØ§Ø±Ø¯: {inbox_errors}  |  "
                f"Ø·Ø§Ø¨ÙˆØ± Ø§Ù„ÙØ±Ø¹ (Ù…Ø¹Ù„Ù‘Ù‚): {queue_pending}"
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
        ttk.Label(self, text="ØªØ´Ø®ÙŠØµ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", font=_FONTS["h1"]).pack(anchor="w", pady=(0, 8))

        self._sync_summary_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._sync_summary_var).pack(fill=tk.X, pady=(0, 6))

        act = ttk.Frame(self)
        act.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            act,
            text="Ù…Ø²Ø§Ù…Ù†Ø© Ø§Ù„Ø¢Ù†",
            command=(lambda: self.app._open_sync_dialog()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="Ø·Ø§Ø¨ÙˆØ± Ø§Ù„ÙØ±ÙˆØ¹",
            command=(lambda: self.app._open_branch_inventory_queue()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)
        ttk.Button(
            act,
            text="Ù…Ø±Ø¢Ø© Ø­Ø¬ÙˆØ²Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹",
            command=(lambda: self.app._open_pos_reservations_mirror()) if self.app else None,
        ).pack(side=tk.LEFT, padx=2)

        health_fr = ttk.LabelFrame(self, text="ØµØ­Ø© Ø§Ù„ÙØ±ÙˆØ¹")
        health_fr.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        self.sync_tree = ttk.Treeview(
            health_fr,
            columns=("branch", "status", "last_snapshot", "age_min", "rows", "value", "sync_errors"),
            show="headings",
            height=8,
        )
        for col, txt, w in [
            ("branch", "Ø§Ù„ÙØ±Ø¹", 180),
            ("status", "Ø§Ù„Ø­Ø§Ù„Ø©", 100),
            ("last_snapshot", "Ø¢Ø®Ø± Ù„Ù‚Ø·Ø©", 190),
            ("age_min", "Ø¹Ù…Ø± Ø§Ù„Ù„Ù‚Ø·Ø© (Ø¯)", 110),
            ("rows", "Ø§Ù„ØµÙÙˆÙ", 90),
            ("value", "Ø§Ù„Ù‚ÙŠÙ…Ø©", 120),
            ("sync_errors", "Ø£Ø®Ø·Ø§Ø¡ Ù…Ø²Ø§Ù…Ù†Ø©", 120),
        ]:
            self.sync_tree.heading(col, text=txt)
            self.sync_tree.column(col, width=w, anchor="center")
        self.sync_tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        add_context_menu(self.sync_tree)
        _bind_mousewheel(self.sync_tree)

        ex_fr = ttk.LabelFrame(self, text="Ø§Ø³ØªØ«Ù†Ø§Ø¡Ø§Øª Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© (Ø£Ø¹Ù„Ù‰ Ø§Ù„Ø£Ø³Ø¨Ø§Ø¨)")
        ex_fr.pack(fill=tk.BOTH, expand=False)
        self.sync_ex_tree = ttk.Treeview(
            ex_fr,
            columns=("etype", "count", "last_error"),
            show="headings",
            height=6,
        )
        for col, txt, w in [
            ("etype", "Ù†ÙˆØ¹ Ø§Ù„Ø­Ø¯Ø«", 220),
            ("count", "Ø¹Ø¯Ø¯ Ø§Ù„Ø£Ø®Ø·Ø§Ø¡", 100),
            ("last_error", "Ø¢Ø®Ø± Ø®Ø·Ø£", 760),
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
                    status = "ØºÙŠØ± Ù…Ø¹Ø±ÙˆÙ"
                    warn_count += 1
                elif age_min >= 60:
                    status = "Ø­Ø±Ø¬"
                    critical_count += 1
                elif age_min >= 15:
                    status = "ØªØ­Ø°ÙŠØ±"
                    warn_count += 1
                else:
                    status = "Ø¬ÙŠØ¯"
                    ok_count += 1
                self.sync_tree.insert(
                    "",
                    tk.END,
                    values=(
                        branch_display_name(r[0] or ""),
                        status,
                        fmt_local_ts(snap, ""),
                        age_min if age_min is not None else "",
                        int(r[2] or 0),
                        f"{format_money(float(r[3] or 0.0))}",
                        "â€”",
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
                f"ÙØ±ÙˆØ¹ Ø³Ù„ÙŠÙ…Ø©: {ok_count}  |  ØªØ­Ø°ÙŠØ±: {warn_count}  |  Ø­Ø±Ø¬: {critical_count}  |  "
                f"Ø£Ø­Ø¯Ø§Ø« ØµØ§Ø¯Ø±Ø© Ù‚ÙŠØ¯ Ø§Ù„Ø§Ù†ØªØ¸Ø§Ø±: {outbox_pending}  |  Ø£Ø®Ø·Ø§Ø¡ ØªØ·Ø¨ÙŠÙ‚ ÙˆØ§Ø±Ø¯: {inbox_errors}  |  "
                f"Ø·Ø§Ø¨ÙˆØ± Ø§Ù„ÙØ±Ø¹ (Ù…Ø¹Ù„Ù‘Ù‚): {queue_pending}"
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
        ttk.Label(self, text="Ø§Ù„Ø¥Ø­ØµØ§Ø¦ÙŠØ§Øª ÙˆØ§Ù„ØªÙ‚Ø§Ø±ÙŠØ±", font=_FONTS["h1"]).pack(anchor="w", pady=(0, 12))

        # Date range filters (stays outside the scroll area)
        filters = ttk.Frame(self)
        filters.pack(fill=tk.X, pady=(0, 8))

        self.df = DateField(filters, "Ù…Ù†")
        self.df.pack(side=tk.LEFT, padx=(0, 8))
        self.dt = DateField(filters, "Ø¥Ù„Ù‰")
        self.dt.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(filters, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT, padx=8)

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
        summary = ttk.LabelFrame(inner, text="Ù…Ù„Ø®Øµ Ø§Ù„ÙØªØ±Ø©")
        summary.pack(fill=tk.X, pady=(0, 8), padx=4)

        self._stats = {}
        stat_defs = [
            ("bills_count", "Ø¹Ø¯Ø¯ Ø§Ù„ÙÙˆØ§ØªÙŠØ±"),
            ("bills_total", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª"),
            ("items_in", "ÙƒÙ…ÙŠØ§Øª ÙˆØ§Ø±Ø¯Ø©"),
            ("items_out", "ÙƒÙ…ÙŠØ§Øª Ù…Ù†ØµØ±ÙØ©"),
        ]
        for i, (key, label) in enumerate(stat_defs):
            ttk.Label(summary, text=label + ":").grid(row=0, column=i*2, padx=6, pady=8, sticky="e")
            val = ttk.Label(summary, text="0", font=("Segoe UI", 12, "bold"))
            val.grid(row=0, column=i*2+1, padx=(0, 16), pady=8, sticky="w")
            self._stats[key] = val
            summary.columnconfigure(i*2+1, weight=1)

        # ---- Top selling items ----
        top_frame = ttk.LabelFrame(inner, text="Ø£ÙƒØ«Ø± Ø§Ù„Ø£ØµÙ†Ø§Ù Ù…Ø¨ÙŠØ¹Ø§Ù‹")
        top_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        self.top_tree = ttk.Treeview(top_frame,
            columns=("item", "school", "color", "qty", "revenue"),
            show="headings", height=10)
        for col, txt, w in [("item", "Ø§Ù„Ù†ÙˆØ¹", 160), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 160),
                            ("color", "Ø§Ù„Ù„ÙˆÙ†", 120), ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø¨Ø§Ø¹Ø©", 110),
                            ("revenue", "Ø§Ù„Ø¥ÙŠØ±Ø§Ø¯", 130)]:
            self.top_tree.heading(col, text=txt)
            self.top_tree.column(col, width=w, anchor="center")
        ysb2 = ttk.Scrollbar(top_frame, orient="vertical", command=self.top_tree.yview)
        self.top_tree.configure(yscrollcommand=ysb2.set)
        self.top_tree.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4, pady=4)
        ysb2.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        add_context_menu(self.top_tree)
        _bind_mousewheel(self.top_tree)

        # ---- Sales Trends ----
        trends_frame = ttk.LabelFrame(inner, text="Ø§ØªØ¬Ø§Ù‡Ø§Øª Ø§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª")
        trends_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        self._trend_mode = tk.StringVar(value="daily")
        mode_row = ttk.Frame(trends_frame)
        mode_row.pack(fill=tk.X, padx=4, pady=4)
        for val, lbl in [("daily", "ÙŠÙˆÙ…ÙŠ"), ("weekly", "Ø£Ø³Ø¨ÙˆØ¹ÙŠ"), ("monthly", "Ø´Ù‡Ø±ÙŠ")]:
            ttk.Radiobutton(mode_row, text=lbl, variable=self._trend_mode,
                            value=val, command=self._refresh).pack(side=tk.LEFT, padx=4)

        self.trend_tree = ttk.Treeview(trends_frame,
            columns=("period", "bills_count", "total_sales", "items_sold"),
            show="headings", height=8)
        for col, txt, w in [
            ("period", "Ø§Ù„ÙØªØ±Ø©", 140), ("bills_count", "Ø¹Ø¯Ø¯ Ø§Ù„ÙÙˆØ§ØªÙŠØ±", 100),
            ("total_sales", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª", 120), ("items_sold", "Ø§Ù„Ø£ØµÙ†Ø§Ù Ø§Ù„Ù…Ø¨Ø§Ø¹Ø©", 100),
        ]:
            self.trend_tree.heading(col, text=txt)
            self.trend_tree.column(col, width=w, anchor="center")
        self.trend_tree.pack(fill=tk.X, expand=True, padx=4, pady=4)
        add_context_menu(self.trend_tree)
        _bind_mousewheel(self.trend_tree)

        # ---- Dead Stock / Slow Movers ----
        dead_frame = ttk.LabelFrame(inner, text="Ø£ØµÙ†Ø§Ù Ø±Ø§ÙƒØ¯Ø© (Ø¨Ø¯ÙˆÙ† Ø­Ø±ÙƒØ© Ø¨ÙŠØ¹)")
        dead_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        days_row = ttk.Frame(dead_frame)
        days_row.pack(fill=tk.X, padx=4, pady=4)
        ttk.Label(days_row, text="Ø¨Ø¯ÙˆÙ† Ø¨ÙŠØ¹ Ù„Ù…Ø¯Ø© (Ø£ÙŠØ§Ù…):").pack(side=tk.LEFT)
        self._dead_days_var = tk.StringVar(value="30")
        ttk.Entry(days_row, textvariable=self._dead_days_var, width=6).pack(side=tk.LEFT, padx=4)
        ttk.Button(days_row, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT, padx=4)

        self.dead_tree = ttk.Treeview(dead_frame,
            columns=("type", "school", "color", "size", "qty", "value", "last_sale"),
            show="headings", height=8)
        for col, txt, w in [
            ("type", "Ø§Ù„Ù†ÙˆØ¹", 120), ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 120), ("color", "Ø§Ù„Ù„ÙˆÙ†", 80),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 60), ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 60), ("value", "Ø§Ù„Ù‚ÙŠÙ…Ø©", 90),
            ("last_sale", "Ø¢Ø®Ø± Ø¨ÙŠØ¹", 120),
        ]:
            self.dead_tree.heading(col, text=txt)
            self.dead_tree.column(col, width=w, anchor="center")
        self.dead_tree.pack(fill=tk.X, expand=True, padx=4, pady=4)
        add_context_menu(self.dead_tree)
        _bind_mousewheel(self.dead_tree)

        # ---- Inventory Valuation by Warehouse ----
        val_frame = ttk.LabelFrame(inner, text="ØªÙ‚ÙŠÙŠÙ… Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø­Ø³Ø¨ Ø§Ù„Ù…Ø®Ø²Ù†")
        val_frame.pack(fill=tk.X, pady=(0, 8), padx=4)

        self.val_tree = ttk.Treeview(val_frame,
            columns=("wh", "items_count", "total_qty", "total_value"),
            show="headings", height=5)
        for col, txt, w in [
            ("wh", "Ø§Ù„Ù…Ø®Ø²Ù†", 80), ("items_count", "Ø¹Ø¯Ø¯ Ø§Ù„Ø£ØµÙ†Ø§Ù", 100),
            ("total_qty", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙƒÙ…ÙŠØ§Øª", 100), ("total_value", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚ÙŠÙ…Ø©", 120),
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
            where = [
                "COALESCE(status,'CONFIRMED')='CONFIRMED'",
                "(COALESCE(bill_type,'SALE')='SALE' OR bill_type IS NULL)",
                "LOWER(TRIM(COALESCE(customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))",
                "LOWER(TRIM(COALESCE(customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))",
                "UPPER(TRIM(COALESCE(customer,''))) != 'WAREHOUSE'",
            ]
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
            self._stats["bills_total"].config(text=f"{format_money(r[1] or 0)}")

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

            cur.execute(f"SELECT COALESCE(SUM(qty), 0) FROM movements WHERE direction IN ('OUT','RESERVE') AND {' AND '.join(mwhere)}", margs)
            self._stats["items_out"].config(text=f"{cur.fetchone()[0] or 0:,}")

            # Top items
            self.top_tree.delete(*self.top_tree.get_children())
            bi_where = [
                "COALESCE(b.status,'CONFIRMED')='CONFIRMED'",
                "(COALESCE(b.bill_type,'SALE')='SALE' OR b.bill_type IS NULL)",
                "LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))",
                "LOWER(TRIM(COALESCE(b.customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))",
                "UPPER(TRIM(COALESCE(b.customer,''))) != 'WAREHOUSE'",
            ]
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
                self.top_tree.insert("", tk.END, values=(r[0], r[1], r[2], r[3], f"{format_money(r[4])}"))
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

            trend_where = [
                "COALESCE(b.status,'CONFIRMED')='CONFIRMED'",
                "(COALESCE(b.bill_type,'SALE')='SALE' OR b.bill_type IS NULL)",
                "LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))",
                "LOWER(TRIM(COALESCE(b.customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))",
                "UPPER(TRIM(COALESCE(b.customer,''))) != 'WAREHOUSE'",
            ]
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
                    r[0], r[1], f"{format_money(r[2])}", r[3]
                ))
            apply_zebra_tags(self.trend_tree)

            # ---- Dead Stock / Slow Movers ----
            self.dead_tree.delete(*self.dead_tree.get_children())
            try:
                days = parse_int_text(self._dead_days_var.get(), 30) or 30
            except Exception:
                days = 30

            cur.execute(f"""
                SELECT s.item_type, s.school, s.color, s.size,
                       SUM(s.count) AS qty,
                       SUM(s.count * s.unit_price) AS value,
                       (SELECT MAX(m.ts) FROM movements m
                        LEFT JOIN bills b ON b.id = m.bill_id
                        WHERE m.item_type = s.item_type AND m.school = s.school
                          AND m.color = s.color AND m.size = s.size
                          AND m.direction IN ('OUT', 'OUT_FACTORY')
                          AND (
                                b.id IS NULL
                             OR (
                                    (COALESCE(b.bill_type,'SALE')='SALE' OR b.bill_type IS NULL)
                                AND LOWER(TRIM(COALESCE(b.customer,''))) NOT LIKE LOWER(TRIM('ÙØ±Ø¹:%'))
                                AND LOWER(TRIM(COALESCE(b.customer,''))) != LOWER(TRIM('Ø§Ù„Ù‰ Ø§Ù„Ù…ØµÙ†Ø¹'))
                                AND UPPER(TRIM(COALESCE(b.customer,''))) != 'WAREHOUSE'
                                )
                          )) AS last_sale
                FROM stocks s
                WHERE s.count > 0
                GROUP BY s.item_type, s.school, s.color, s.size
                HAVING last_sale IS NULL OR date(last_sale) < date('now', '-{days} days')
                ORDER BY value DESC
                LIMIT 20
            """)
            for r in cur.fetchall():
                self.dead_tree.insert("", tk.END, values=(
                    r[0], r[1], r[2], r[3], r[4], f"{format_money(r[5])}",
                    r[6] or "Ù„Ø§ ÙŠÙˆØ¬Ø¯"
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
                    r[0], r[1], f"{r[2]:,}", f"{format_money(r[3])}"
                ))
            apply_zebra_tags(self.val_tree)

            cur.close()
        except Exception:
            pass


class SchoolAccountsFrame(ttk.Frame):
    """School sales report without automatic percentage calculations."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        self._selected_schools: List[str] = []
        self._build()
        self._refresh()

    def _update_selected_schools_label(self) -> None:
        if self._selected_schools:
            self._schools_var.set("Ø§Ù„Ù…Ø¯Ø§Ø±Ø³ Ø§Ù„Ù…Ø®ØªØ§Ø±Ø©: " + " - ".join(self._selected_schools))
        else:
            self._schools_var.set("ÙƒÙ„ Ø§Ù„Ù…Ø¯Ø§Ø±Ø³")

    def _build(self):
        ttk.Label(self, text="Ø­Ø³Ø§Ø¨Ø§Øª Ø§Ù„Ù…Ø¯Ø§Ø±Ø³", font=_FONTS["h1"]).pack(anchor="w", pady=(0, 12))

        filters = ttk.Frame(self)
        filters.pack(fill=tk.X, pady=(0, 8))

        self.df = DateField(filters, "Ù…Ù†")
        self.df.pack(side=tk.LEFT, padx=(0, 8))
        self.dt = DateField(filters, "Ø¥Ù„Ù‰")
        self.dt.pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(filters, text="Ø§Ø®ØªÙŠØ§Ø± Ø§Ù„Ù…Ø¯Ø§Ø±Ø³", command=self._pick_schools).pack(side=tk.LEFT, padx=8)
        ttk.Button(filters, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT, padx=8)
        ttk.Button(filters, text="Ù…Ø³Ø­ Ø§Ù„Ø§Ø®ØªÙŠØ§Ø±", command=self._clear_schools).pack(side=tk.LEFT, padx=8)

        self._schools_var = tk.StringVar(value="ÙƒÙ„ Ø§Ù„Ù…Ø¯Ø§Ø±Ø³")
        ttk.Label(self, textvariable=self._schools_var, foreground=_UI["TEXT_DIM"]).pack(anchor="w", pady=(0, 8))

        summary = ttk.LabelFrame(self, text="Ù…Ù„Ø®Øµ")
        summary.pack(fill=tk.X, pady=(0, 8))
        self._sum_school_count = tk.StringVar(value="0")
        self._sum_qty = tk.StringVar(value="0")
        self._sum_sales = tk.StringVar(value="0")
        for i, (label, var) in enumerate([
            ("Ø¹Ø¯Ø¯ Ø§Ù„Ù…Ø¯Ø§Ø±Ø³:", self._sum_school_count),
            ("Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙƒÙ…ÙŠØ©:", self._sum_qty),
            ("Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª:", self._sum_sales),
        ]):
            ttk.Label(summary, text=label).grid(row=0, column=i * 2, padx=6, pady=8, sticky="e")
            ttk.Label(summary, textvariable=var, font=("Segoe UI", 12, "bold")).grid(
                row=0, column=i * 2 + 1, padx=(0, 16), pady=8, sticky="w"
            )

        table_wrap = ttk.LabelFrame(self, text="ØªÙØ§ØµÙŠÙ„ Ø§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª")
        table_wrap.pack(fill=tk.BOTH, expand=True)
        cols = ("school", "item_type", "color", "size", "qty", "sales_total")
        self._tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=16)
        for col, txt, w in [
            ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 170),
            ("item_type", "Ø§Ù„Ù†ÙˆØ¹", 170),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 110),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 90),
            ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 90),
            ("sales_total", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨ÙŠØ¹Ø§Øª", 130),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=ysb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        ysb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        add_context_menu(self._tree)
        _bind_mousewheel(self._tree)

    def _pick_schools(self):
        values = self.db.list_schools_all()
        dlg = MultiSelectDialog(self, title="Ø§Ø®ØªÙŠØ§Ø± Ø§Ù„Ù…Ø¯Ø§Ø±Ø³", values=values, preselected=self._selected_schools)
        result = dlg.run()
        if result is None:
            return
        self._selected_schools = list(result)
        self._update_selected_schools_label()
        self._refresh()

    def _clear_schools(self):
        self._selected_schools = []
        self._update_selected_schools_label()
        self._refresh()

    def _refresh(self):
        rows = self.db.get_school_accounts_report(
            self._selected_schools or self.db.list_schools_all(),
            date_from=self.df.get() or None,
            date_to=self.dt.get() or None,
        )
        self._tree.delete(*self._tree.get_children())
        total_qty = 0
        total_sales = 0.0
        schools_seen = set()
        for row in rows:
            qty = int(row.get("total_qty") or 0)
            sales_total = float(row.get("total_sales") or 0.0)
            school = str(row.get("school") or "").strip()
            schools_seen.add(school)
            total_qty += qty
            total_sales += sales_total
            self._tree.insert(
                "",
                tk.END,
                values=(
                    school,
                    row.get("item_type") or "",
                    row.get("color") or "",
                    row.get("size") or "",
                    qty,
                    f"{format_money(sales_total)}",
                ),
            )
        apply_zebra_tags(self._tree)
        self._update_selected_schools_label()
        self._sum_school_count.set(str(len(schools_seen)))
        self._sum_qty.set(str(total_qty))
        self._sum_sales.set(f"{format_money(total_sales)}")


class FabricWeightsDialog(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase, on_change=None):
        super().__init__(master)
        self.db = db
        self.on_change = on_change
        self.title("Ø¥Ø¯Ø§Ø±Ø© Ø£ÙˆØ²Ø§Ù† Ø§Ù„Ù‚Ù…Ø§Ø´")
        self.geometry("760x520")
        self.configure(bg=_UI["BG"])
        _wh = tk.Frame(self, bg=_UI["ACCENT"], height=44)
        _wh.pack(fill=tk.X)
        _wh.pack_propagate(False)
        tk.Label(_wh, text="  Ø¥Ø¯Ø§Ø±Ø© Ø£ÙˆØ²Ø§Ù† Ø§Ù„Ù‚Ù…Ø§Ø´", bg=_UI["ACCENT"], fg="#FFFFFF", font=_FONTS["h3"]).pack(side=tk.RIGHT, padx=12)
        self.transient(master)
        self.grab_set()
        self._build()
        self._refresh()

    def _build(self):
        top = ttk.LabelFrame(self, text="Ø¨ÙŠØ§Ù†Ø§Øª Ø§Ù„ÙˆØ²Ù†")
        top.pack(fill=tk.X, padx=10, pady=10)

        self.f_item = LabeledCombobox(top, "Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù", self.db, "item_type")
        self.f_item.grid(row=0, column=0, padx=6, pady=6, sticky="ew")
        self.f_size = LabeledCombobox(top, "Ø§Ù„Ù…Ù‚Ø§Ø³", self.db, "size")
        self.f_size.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
        self.weight_var = tk.StringVar()
        self.f_weight = LabeledEntry(top, "Ø§Ù„ÙˆØ²Ù† Ø¨Ø§Ù„Ø¬Ø±Ø§Ù…")
        self.f_weight.grid(row=0, column=2, padx=6, pady=6, sticky="ew")
        self.f_weight.entry.configure(textvariable=self.weight_var)
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)
        top.columnconfigure(2, weight=1)

        self.f_item.set_supplier(lambda: self.db.get_distinct_filtered("item_type", {}))
        self.f_size.set_supplier(
            lambda: self.db.get_distinct_filtered(
                "size",
                {"item_type": self.f_item.get() or None},
            )
        )
        self.f_item.cb.bind("<<ComboboxSelected>>", lambda _e: self.f_size.refresh_values(), add="+")
        self.f_item.cb.bind("<KeyRelease>", lambda _e: self.f_size.refresh_values(), add="+")

        btns = ttk.Frame(top)
        btns.grid(row=1, column=0, columnspan=3, sticky="e", padx=6, pady=(0, 6))
        ttk.Button(btns, text="Ø¬Ø¯ÙŠØ¯", command=self._clear_form).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø­ÙØ¸", command=self._save).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø­Ø°Ù", command=self._delete_selected).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT, padx=4)

        wrap = ttk.LabelFrame(self, text="Ø§Ù„Ø£ÙˆØ²Ø§Ù† Ø§Ù„Ù…Ø¹Ø±ÙØ©")
        wrap.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        cols = ("item_type", "size", "weight_grams")
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", height=16)
        for col, txt, w in [
            ("item_type", "Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù", 280),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 140),
            ("weight_grams", "Ø§Ù„ÙˆØ²Ù† Ø¨Ø§Ù„Ø¬Ø±Ø§Ù…", 140),
        ]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        ysb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        add_context_menu(self.tree)
        _bind_mousewheel(self.tree)
        self.tree.bind("<<TreeviewSelect>>", self._load_selected, add="+")

    def _selected_specs(self) -> Tuple[str, str]:
        item_type = self.f_item.get().strip()
        size = self.f_size.get().strip()
        return item_type, size

    def _clear_form(self):
        self.f_item.set("")
        self.f_size.set("")
        self.weight_var.set("")
        try:
            self.tree.selection_remove(self.tree.selection())
        except Exception:
            pass

    def _load_selected(self, *_):
        sel = self.tree.selection()
        if not sel:
            return
        vals = self.tree.item(sel[0], "values")
        if not vals:
            return
        self.f_item.set(str(vals[0] or ""))
        self.f_size.set(str(vals[1] or ""))
        self.weight_var.set(str(vals[2] or ""))

    def _refresh(self):
        rows = self.db.list_fabric_weights()
        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row.get("item_type") or "",
                    row.get("size") or "",
                    f"{float(row.get('weight_grams') or 0.0):.3f}".rstrip("0").rstrip("."),
                ),
            )
        apply_zebra_tags(self.tree)
        self.f_item.refresh_values()
        self.f_size.refresh_values()

    def _save(self):
        item_type, size = self._selected_specs()
        try:
            weight = parse_float_text(self.weight_var.get())
            if weight is None:
                raise ValueError("Invalid fabric weight")
            self.db.upsert_fabric_weight(item_type, size, weight)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­ÙØ¸", str(ex), parent=self)
            return
        self._refresh()
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass
        show_toast(self, "ØªÙ… Ø­ÙØ¸ ÙˆØ²Ù† Ø§Ù„Ù‚Ù…Ø§Ø´ Ø¨Ù†Ø¬Ø§Ø­")

    def _delete_selected(self):
        item_type, size = self._selected_specs()
        if not item_type or not size:
            messagebox.showwarning("Ù„Ù… ÙŠØªÙ… Ø§Ù„ØªØ­Ø¯ÙŠØ¯", "Ø§Ø®ØªØ± ÙˆØ²Ù†Ù‹Ø§ Ù…Ù† Ø§Ù„Ù‚Ø§Ø¦Ù…Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not messagebox.askyesno(
            "ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø­Ø°Ù",
            f"Ø­Ø°Ù ÙˆØ²Ù† Ø§Ù„Ù‚Ù…Ø§Ø´ Ù„Ù„ØµÙ†Ù '{item_type}' ÙˆØ§Ù„Ù…Ù‚Ø§Ø³ '{size}'ØŸ",
            parent=self,
        ):
            return
        deleted = self.db.delete_fabric_weight(item_type, size)
        if deleted <= 0:
            show_toast(self, "Ù„Ù… ÙŠØªÙ… Ø§Ù„Ø¹Ø«ÙˆØ± Ø¹Ù„Ù‰ Ø§Ù„ÙˆØ²Ù† Ø§Ù„Ù…Ø­Ø¯Ø¯", bg="#f59e0b")
            return
        self._clear_form()
        self._refresh()
        if self.on_change:
            try:
                self.on_change()
            except Exception:
                pass
        show_toast(self, "ØªÙ… Ø­Ø°Ù ÙˆØ²Ù† Ø§Ù„Ù‚Ù…Ø§Ø´")


class FabricCalculationFrame(ttk.Frame):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        self._last_result: Optional[Dict[str, Any]] = None
        self._build()

    def _build(self):
        ttk.Label(self, text="Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´", font=_FONTS["h1"]).pack(anchor="w", pady=(0, 12))

        top = ttk.LabelFrame(self, text="Ø§Ù„ÙÙ„Ø§ØªØ±")
        top.pack(fill=tk.X, pady=(0, 8))
        self.ftype = LabeledCombobox(top, "Ø§Ù„Ù†ÙˆØ¹", self.db, "item_type")
        self.ftype.grid(row=0, column=0, padx=6, pady=4, sticky="ew")
        self.fsch = LabeledCombobox(top, "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", self.db, "school")
        self.fsch.grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        self.fclr = LabeledCombobox(top, "Ø§Ù„Ù„ÙˆÙ†", self.db, "color")
        self.fclr.grid(row=0, column=2, padx=6, pady=4, sticky="ew")
        self.fsiz = LabeledCombobox(top, "Ø§Ù„Ù…Ù‚Ø§Ø³", self.db, "size")
        self.fsiz.grid(row=0, column=3, padx=6, pady=4, sticky="ew")
        branch_vals = [""] + [branch_display_name(n) for n in DEFAULT_BRANCH_POS_NAMES]
        branch_map = {n: branch_display_name(n) for n in DEFAULT_BRANCH_POS_NAMES}
        self.fbranch = LabeledStaticCombo(top, "Ø§Ù„ÙØ±Ø¹", branch_vals, value_map=branch_map, width=18)
        self.fbranch.grid(row=0, column=4, padx=6, pady=4, sticky="ew")

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

        self.df = DateField(top, "Ù…Ù†")
        self.df.grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.dt = DateField(top, "Ø¥Ù„Ù‰")
        self.dt.grid(row=1, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(top, text="Ø¨Ø­Ø«").grid(row=1, column=2, sticky="e", padx=4, pady=4)
        self.txt = ttk.Entry(top)
        self.txt.grid(row=1, column=3, sticky="ew", padx=6, pady=4)
        top.columnconfigure(3, weight=1)

        btns = ttk.Frame(top)
        btns.grid(row=0, column=5, rowspan=2, sticky="e", padx=6, pady=4)
        ttk.Button(btns, text="Ø§Ø­Ø³Ø¨", command=self._calculate).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„Ø£ÙˆØ²Ø§Ù†", command=self._open_weights_dialog).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="ØªØµØ¯ÙŠØ±", command=self._export).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ø·Ø¨Ø§Ø¹Ø©", command=self._print_report).pack(side=tk.LEFT, padx=4)
        ttk.Button(btns, text="Ù…Ø³Ø­", command=self._clear).pack(side=tk.LEFT, padx=4)

        summary = ttk.LabelFrame(self, text="Ø§Ù„Ù…Ù„Ø®Øµ")
        summary.pack(fill=tk.X, pady=(0, 8))
        self._sum_rows = tk.StringVar(value="0")
        self._sum_requested = tk.StringVar(value="0")
        self._sum_weight = tk.StringVar(value="0 ÙƒØ¬Ù…")
        self._sum_missing = tk.StringVar(value="0")
        summary_items = [
            ("Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ:", self._sum_rows),
            ("Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨:", self._sum_requested),
            ("Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚Ù…Ø§Ø´:", self._sum_weight),
            ("Ø£ÙˆØ²Ø§Ù† ØºÙŠØ± Ù…Ø¹Ø±ÙØ©:", self._sum_missing),
        ]
        for i, (label, var) in enumerate(summary_items):
            ttk.Label(summary, text=label).grid(row=0, column=i * 2, padx=6, pady=8, sticky="e")
            ttk.Label(summary, textvariable=var, font=("Segoe UI", 12, "bold")).grid(
                row=0, column=i * 2 + 1, padx=(0, 16), pady=8, sticky="w"
            )

        wrap = ttk.LabelFrame(self, text="Ø§Ù„Ù†ØªØ§Ø¦Ø¬")
        wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        cols = ("school", "item_type", "color", "requested_qty", "total_weight")
        self.tree = ttk.Treeview(wrap, columns=cols, show="headings", height=14)
        for col, txt, w in [
            ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 170),
            ("item_type", "Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù", 170),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 110),
            ("requested_qty", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨", 110),
            ("total_weight", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚Ù…Ø§Ø´", 140),
        ]:
            self.tree.heading(col, text=txt)
            self.tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        ysb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        add_context_menu(self.tree)
        _bind_mousewheel(self.tree)

        missing_wrap = ttk.LabelFrame(self, text="Ø£ÙˆØ²Ø§Ù† ØºÙŠØ± Ù…Ø¹Ø±ÙØ©")
        missing_wrap.pack(fill=tk.BOTH, expand=False)
        miss_cols = ("school", "item_type", "color", "size", "requested_qty")
        self.missing_tree = ttk.Treeview(missing_wrap, columns=miss_cols, show="headings", height=7)
        for col, txt, w in [
            ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 170),
            ("item_type", "Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù", 170),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 110),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 90),
            ("requested_qty", "Ø§Ù„Ù…Ø·Ù„ÙˆØ¨", 90),
        ]:
            self.missing_tree.heading(col, text=txt)
            self.missing_tree.column(col, width=w, anchor="center")
        mysb = ttk.Scrollbar(missing_wrap, orient="vertical", command=self.missing_tree.yview)
        self.missing_tree.configure(yscrollcommand=mysb.set)
        self.missing_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        mysb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        add_context_menu(self.missing_tree)
        _bind_mousewheel(self.missing_tree)

    def _filters(self) -> Dict[str, Any]:
        return {
            "item_type": self.ftype.get(),
            "school": self.fsch.get(),
            "color": self.fclr.get(),
            "size": self.fsiz.get(),
            "branch_device": self.fbranch.get() or None,
            "date_from": self.df.get() or None,
            "date_to": self.dt.get() or None,
            "text": (self.txt.get().strip() or None),
        }

    def _open_weights_dialog(self):
        FabricWeightsDialog(self, self.db, on_change=self._on_weights_changed)

    def _on_weights_changed(self):
        if self._last_result:
            self._calculate()

    def _clear(self):
        for w in (self.ftype, self.fsch, self.fclr, self.fsiz, self.fbranch):
            w.set("")
        self.df.set("")
        self.dt.set("")
        self.txt.delete(0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.missing_tree.delete(*self.missing_tree.get_children())
        self._sum_rows.set("0")
        self._sum_requested.set("0")
        self._sum_weight.set("0 ÙƒØ¬Ù…")
        self._sum_missing.set("0")
        self._last_result = None

    def _calculate(self):
        try:
            result = self.db.calculate_fabric_requirements(self._filters())
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­Ø³Ø§Ø¨", str(ex), parent=self)
            return
        self._last_result = result
        rows = list(result.get("rows") or [])
        missing = list(result.get("missing") or [])
        summary = dict(result.get("summary") or {})

        self.tree.delete(*self.tree.get_children())
        for row in rows:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    row.get("school") or "",
                    row.get("item_type") or "",
                    row.get("color") or "",
                    int(row.get("requested_qty") or 0),
                    format_weight_kg(row.get("total_weight_grams") or 0.0),
                ),
            )
        apply_zebra_tags(self.tree)

        self.missing_tree.delete(*self.missing_tree.get_children())
        for row in missing:
            self.missing_tree.insert(
                "",
                tk.END,
                values=(
                    row.get("school") or "",
                    row.get("item_type") or "",
                    row.get("color") or "",
                    row.get("size") or "",
                    int(row.get("requested_qty") or 0),
                ),
            )
        apply_zebra_tags(self.missing_tree)

        self._sum_rows.set(str(int(summary.get("row_count") or 0)))
        self._sum_requested.set(str(int(summary.get("requested_qty") or 0)))
        self._sum_weight.set(format_weight_kg(summary.get("total_weight_grams") or 0.0))
        self._sum_missing.set(str(int(summary.get("missing_count") or 0)))

    def _export(self):
        if not self._last_result:
            messagebox.showwarning("Ù„Ø§ ØªÙˆØ¬Ø¯ Ù†ØªØ§Ø¦Ø¬", "Ø§Ø­Ø³Ø¨ Ø§Ù„Ù†ØªØ§Ø¦Ø¬ Ø£ÙˆÙ„Ø§Ù‹ Ù‚Ø¨Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="ØªØµØ¯ÙŠØ± Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"fabric_requirements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        headers = [
            "school",
            "item_type",
            "color",
            "requested_qty",
            "total_weight_grams",
            "total_weight_kg",
        ]
        table: List[List[Any]] = []
        for row in self._last_result.get("rows") or []:
            grams = float(row.get("total_weight_grams") or 0.0)
            table.append([
                row.get("school") or "",
                row.get("item_type") or "",
                row.get("color") or "",
                int(row.get("requested_qty") or 0),
                grams,
                format_weight_kg(grams),
            ])
        missing = list(self._last_result.get("missing") or [])
        if missing:
            table.append(["", "", "", "", "", ""])
            table.append(["Ø£ÙˆØ²Ø§Ù† ØºÙŠØ± Ù…Ø¹Ø±ÙØ©", "", "", "", "", ""])
            for row in missing:
                table.append([
                    row.get("school") or "",
                    row.get("item_type") or "",
                    row.get("color") or "",
                    int(row.get("requested_qty") or 0),
                    row.get("size") or "",
                    "ÙˆØ²Ù† ØºÙŠØ± Ù…Ø¹Ø±Ù",
                ])
        try:
            export_to_excel(path, headers, table)
            show_toast(self, "ØªÙ… ØªØµØ¯ÙŠØ± Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´ Ø¨Ù†Ø¬Ø§Ø­")
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„ØªØµØ¯ÙŠØ±", str(ex), parent=self)

    def _print_report(self):
        if not self._last_result:
            messagebox.showwarning("Ù„Ø§ ØªÙˆØ¬Ø¯ Ù†ØªØ§Ø¦Ø¬", "Ø§Ø­Ø³Ø¨ Ø§Ù„Ù†ØªØ§Ø¦Ø¬ Ø£ÙˆÙ„Ø§Ù‹ Ù‚Ø¨Ù„ Ø§Ù„Ø·Ø¨Ø§Ø¹Ø©.", parent=self)
            return

        def _h(v: Any) -> str:
            s = str(v if v is not None else "")
            return (
                s.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace('"', "&quot;")
            )

        filters = self._filters()
        filter_bits = []
        for label, key in (
            ("Ø§Ù„ÙØ±Ø¹", "branch_device"),
            ("Ø§Ù„Ù†ÙˆØ¹", "item_type"),
            ("Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", "school"),
            ("Ø§Ù„Ù„ÙˆÙ†", "color"),
            ("Ø§Ù„Ù…Ù‚Ø§Ø³", "size"),
            ("Ù…Ù†", "date_from"),
            ("Ø¥Ù„Ù‰", "date_to"),
            ("Ø¨Ø­Ø«", "text"),
        ):
            val = str(filters.get(key) or "").strip()
            if val:
                filter_bits.append(f"{label}: {_h(val)}")
        filters_html = " | ".join(filter_bits) if filter_bits else "Ø¨Ø¯ÙˆÙ† ÙÙ„Ø§ØªØ±"

        rows_html = []
        for row in self._last_result.get("rows") or []:
            rows_html.append(
                "<tr>"
                f"<td>{_h(row.get('school') or '')}</td>"
                f"<td>{_h(row.get('item_type') or '')}</td>"
                f"<td>{_h(row.get('color') or '')}</td>"
                f"<td class='num'>{int(row.get('requested_qty') or 0)}</td>"
                f"<td class='num'>{_h(format_weight_kg(row.get('total_weight_grams') or 0.0))}</td>"
                "</tr>"
            )

        missing_html = []
        for row in self._last_result.get("missing") or []:
            missing_html.append(
                "<tr>"
                f"<td>{_h(row.get('school') or '')}</td>"
                f"<td>{_h(row.get('item_type') or '')}</td>"
                f"<td>{_h(row.get('color') or '')}</td>"
                f"<td>{_h(row.get('size') or '')}</td>"
                f"<td class='num'>{int(row.get('requested_qty') or 0)}</td>"
                "</tr>"
            )

        summary = dict(self._last_result.get("summary") or {})
        missing_section = ""
        if missing_html:
            missing_section = f"""
            <div class="subhead">Ø£ÙˆØ²Ø§Ù† ØºÙŠØ± Ù…Ø¹Ø±ÙØ©</div>
            <table>
              <thead>
                <tr>
                  <th>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©</th>
                  <th>Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù</th>
                  <th>Ø§Ù„Ù„ÙˆÙ†</th>
                  <th>Ø§Ù„Ù…Ù‚Ø§Ø³</th>
                  <th>Ø§Ù„Ù…Ø·Ù„ÙˆØ¨</th>
                </tr>
              </thead>
              <tbody>
                {''.join(missing_html)}
              </tbody>
            </table>
            """

        html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8" />
<title>Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´</title>
<style>
@page {{ size: A4 landscape; margin: 10mm; }}
body {{
  margin: 0;
  font-family: "Segoe UI", Tahoma, Arial, "Noto Sans Arabic", sans-serif;
  color: #0f172a;
  direction: rtl;
}}
.wrap {{ padding: 8px; }}
.title {{ font-size: 20px; font-weight: 700; margin-bottom: 4px; }}
.meta {{ font-size: 12px; color: #334155; margin-bottom: 4px; }}
.summary {{
  background: #f8fafc;
  border: 1px solid #cbd5e1;
  padding: 8px;
  margin: 8px 0 10px;
  font-weight: 600;
  font-size: 13px;
}}
.subhead {{ font-size: 16px; font-weight: 700; margin: 14px 0 6px; }}
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
    <div class="title">ØªÙ‚Ø±ÙŠØ± Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´</div>
    <div class="meta">ÙˆÙ‚Øª Ø§Ù„ØªÙ‚Ø±ÙŠØ±: {_h(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))}</div>
    <div class="meta">Ø§Ù„ÙÙ„Ø§ØªØ±: {filters_html}</div>
    <div class="summary">
      Ø¹Ø¯Ø¯ Ø§Ù„ØµÙÙˆÙ: {int(summary.get('row_count') or 0)} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨: {int(summary.get('requested_qty') or 0)} |
      Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚Ù…Ø§Ø´: {_h(format_weight_kg(summary.get('total_weight_grams') or 0.0))} |
      Ø£ÙˆØ²Ø§Ù† ØºÙŠØ± Ù…Ø¹Ø±ÙØ©: {int(summary.get('missing_count') or 0)}
    </div>
    <table>
      <thead>
        <tr>
          <th>Ø§Ù„Ù…Ø¯Ø±Ø³Ø©</th>
          <th>Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù</th>
          <th>Ø§Ù„Ù„ÙˆÙ†</th>
          <th>Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø·Ù„ÙˆØ¨</th>
          <th>Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù‚Ù…Ø§Ø´</th>
        </tr>
      </thead>
      <tbody>
        {''.join(rows_html)}
      </tbody>
    </table>
    {missing_section}
  </div>
</body>
</html>
"""
        path = os.path.join(
            tempfile.gettempdir(),
            f"fabric_requirements_print_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
        )
        with open(path, "w", encoding="utf-8") as fobj:
            fobj.write(html)
        _print_html_auto(path, copies=1, parent=self)


# ------------------- Branch / POS reporting (warehouse) -------------------

class BranchBillsSyncLogWindow(tk.Toplevel):
    """Shipments to branches (warehouse bills) + synced inbound queue."""

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master)
        self.db = db
        self.title("Ø³Ø¬Ù„ ÙÙˆØ§ØªÙŠØ± Ø§Ù„ÙØ±ÙˆØ¹ ÙˆØ§Ù„ÙˆØ§Ø±Ø¯ Ø§Ù„Ù…ØªØ²Ø§Ù…Ù†")
        self.geometry("1080x620")
        self._build()

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        tab_b = ttk.Frame(nb)
        nb.add(tab_b, text="Ø´Ø­Ù†Ø§Øª ØµØ§Ø¯Ø±Ø© (ÙÙˆØ§ØªÙŠØ± ÙØ±Ø¹)")
        top_b = ttk.Frame(tab_b)
        top_b.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(top_b, text="ØªØ­Ø¯ÙŠØ«", command=lambda: self._fill_branch_bills()).pack(side=tk.RIGHT)
        wrap_b = ttk.Frame(tab_b)
        wrap_b.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        cols_b = ("id", "created_at", "customer", "total", "status")
        self._tree_b = ttk.Treeview(wrap_b, columns=cols_b, show="headings", height=18)
        for col, txt, w in [
            ("id", "#", 70),
            ("created_at", "Ø§Ù„ØªØ§Ø±ÙŠØ®", 170),
            ("customer", "Ø§Ù„ÙØ±Ø¹ / Ø§Ù„Ø¹Ù…ÙŠÙ„", 260),
            ("total", "Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ", 100),
            ("status", "Ø§Ù„Ø­Ø§Ù„Ø©", 90),
        ]:
            self._tree_b.heading(col, text=txt)
            self._tree_b.column(col, width=w, anchor="center")
        sb_b = ttk.Scrollbar(wrap_b, orient="vertical", command=self._tree_b.yview)
        self._tree_b.configure(yscrollcommand=sb_b.set)
        self._tree_b.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb_b.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree_b)

        tab_q = ttk.Frame(nb)
        nb.add(tab_q, text="ÙˆØ§Ø±Ø¯ Ù…ØªØ²Ø§Ù…Ù† (Ø·Ø§Ø¨ÙˆØ± Ø§Ù„ÙØ±ÙˆØ¹)")
        top_q = ttk.Frame(tab_q)
        top_q.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(top_q, text="ØªØ­Ø¯ÙŠØ«", command=lambda: self._fill_queue()).pack(side=tk.RIGHT)
        wrap_q = ttk.Frame(tab_q)
        wrap_q.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        cols_q = ("id", "kind", "source", "target", "type", "school", "color", "size", "qty", "status", "created_at")
        self._tree_q = ttk.Treeview(wrap_q, columns=cols_q, show="headings", height=18)
        for col, txt, w in [
            ("id", "#", 55),
            ("kind", "Ø§Ù„Ù†ÙˆØ¹", 80),
            ("source", "Ø§Ù„Ù…Ø±Ø³Ù„", 110),
            ("target", "Ø§Ù„Ù‡Ø¯Ù", 110),
            ("type", "Ø§Ù„ØµÙ†Ù", 140),
            ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 120),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 80),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 70),
            ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 60),
            ("status", "Ø§Ù„Ø­Ø§Ù„Ø©", 100),
            ("created_at", "Ø§Ù„ØªØ§Ø±ÙŠØ®", 150),
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
        status_map = {"DRAFT": "Ù…Ø³ÙˆØ¯Ø©", "CONFIRMED": "Ù…Ø¤ÙƒØ¯Ø©", "VOID": "Ù…Ù„ØºØ§Ø©"}
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
                    f"{format_money(float(b.get('total') or 0))}",
                    st,
                ),
            )
        apply_zebra_tags(self._tree_b)

    def _fill_queue(self):
        self._tree_q.delete(*self._tree_q.get_children())
        for row in self.db.list_branch_inventory_queue(None):
            kind = "Ù…Ø±ØªØ¬Ø¹" if row.get("queue_kind") == "RETURN" else "ØªØ­ÙˆÙŠÙ„"
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
        self.title("Ø­Ø¬ÙˆØ²Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ (Ù…Ø±Ø¢Ø© Ù…Ù† Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©)")
        self.geometry("1220x600")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="Ù†Ù‚Ø·Ø© Ø§Ù„Ø¨ÙŠØ¹:").pack(side=tk.RIGHT, padx=4)
        pick = self.db.list_pos_reservations_mirror_device_picklist()
        self._dev = tk.StringVar(value="")
        self._dev_cb = ttk.Combobox(
            top, textvariable=self._dev, values=pick, width=34, state="readonly",
        )
        self._dev_cb.pack(side=tk.RIGHT)
        self._dev_cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh())
        self._active_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            top, text="Ø§Ù„Ù…Ø¹Ù„Ù‚Ø© ÙÙ‚Ø·", variable=self._active_only,
            command=self._refresh,
        ).pack(side=tk.RIGHT, padx=12)
        ttk.Button(top, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT)

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        tab_d = ttk.Frame(nb)
        nb.add(tab_d, text="ØªÙØµÙŠÙ„ÙŠ (ÙƒÙ„ Ø­Ø¬Ø²)")
        ttk.Label(
            tab_d,
            text=(
                "ÙƒÙ„ Ø³Ø·Ø± Ù‡Ù†Ø§ = Ø³Ø·Ø± Ø­Ø¬Ø² ÙˆØ§Ø­Ø¯ ÙÙŠ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© (Ù…Ø¹Ø±Ù‘ÙØ§Øª Ù…Ø®ØªÙ„ÙØ© ØªØ¸Ù‡Ø± Ù…Ù†ÙØµÙ„Ø©). "
                "Ù„Ø±Ø¤ÙŠØ© Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙƒÙ…ÙŠØ© Ù„Ù†ÙØ³ Ø§Ù„Ù…Ù†ØªØ¬ (Ù†ÙˆØ¹ + Ù…Ø¯Ø±Ø³Ø© + Ù„ÙˆÙ† + Ù…Ù‚Ø§Ø³) Ø¨Ø¹Ø¯ Ø¬Ù…Ø¹ ÙƒÙ„ Ø§Ù„Ø£Ø³Ø·Ø±ØŒ "
                "Ø§Ø³ØªØ®Ø¯Ù… ØªØ¨ÙˆÙŠØ¨ Â«Ù…Ø¬Ù…Ù‘Ø¹ Ø­Ø³Ø¨ Ø§Ù„Ù…Ù†ØªØ¬Â» (ÙŠÙØªØ­ Ø§ÙØªØ±Ø§Ø¶ÙŠØ§Ù‹)."
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
            ("device", "Ø§Ù„Ø¬Ù‡Ø§Ø²", 110),
            ("key", "Ø§Ù„Ù…Ø¹Ø±Ù‘Ù", 180),
            ("customer", "Ø§Ù„Ø¹Ù…ÙŠÙ„", 120),
            ("type", "Ø§Ù„Ù†ÙˆØ¹", 110),
            ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 110),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 70),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 55),
            ("qty", "Ø§Ù„ÙƒÙ…ÙŠØ©", 50),
            ("price", "Ø§Ù„Ø³Ø¹Ø±", 65),
            ("total", "Ø§Ù„Ø¥Ø¬Ù…Ø§Ù„ÙŠ", 75),
            ("paid", "Ø§Ù„Ù…Ø¯ÙÙˆØ¹", 70),
            ("status", "Ø§Ù„Ø­Ø§Ù„Ø©", 85),
            ("updated", "Ø¢Ø®Ø± ØªØ­Ø¯ÙŠØ«", 140),
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
        nb.add(tab_a, text="Ù…Ø¬Ù…Ù‘Ø¹ Ø­Ø³Ø¨ Ø§Ù„Ù…Ù†ØªØ¬")
        hint = (
            "ØªÙØ¬Ù…Ù‘Ø¹ ÙƒÙ„ Ø§Ù„Ø£Ø³Ø·Ø± Ø°Ø§Øª Ù†ÙØ³ (Ø§Ù„Ù†ÙˆØ¹ØŒ Ø§Ù„Ù…Ø¯Ø±Ø³Ø©ØŒ Ø§Ù„Ù„ÙˆÙ†ØŒ Ø§Ù„Ù…Ù‚Ø§Ø³). "
            "Ø¹Ù†Ø¯ Ø§Ø®ØªÙŠØ§Ø± Ù†Ù‚Ø·Ø© Ø¨ÙŠØ¹ ÙŠÙØ­Ø³Ø¨ Ø§Ù„Ù…Ø¬Ù…ÙˆØ¹ Ù„Ù‡Ø°Ø§ Ø§Ù„ÙØ±Ø¹ ÙÙ‚Ø· (Ø¨Ù…Ø§ ÙŠØªÙˆØ§ÙÙ‚ Ù…Ø¹ Ø§Ù„Ø§Ø³Ù… Ø£Ùˆ Ø§Ù„Ù…Ø¹Ø±Ù‘Ù Ø§Ù„Ù…Ø®Ø²Ù‘Ù†). "
            "Ø¹Ù†Ø¯ ØªØ±Ùƒ Â«Ù†Ù‚Ø·Ø© Ø§Ù„Ø¨ÙŠØ¹Â» ÙØ§Ø±ØºØ§Ù‹ ÙŠÙØ¬Ù…Ø¹ Ø¹Ø¨Ø± ÙƒÙ„ Ø§Ù„ÙØ±ÙˆØ¹."
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
            ("type", "Ø§Ù„Ù†ÙˆØ¹", 120),
            ("school", "Ø§Ù„Ù…Ø¯Ø±Ø³Ø©", 120),
            ("color", "Ø§Ù„Ù„ÙˆÙ†", 80),
            ("size", "Ø§Ù„Ù…Ù‚Ø§Ø³", 60),
            ("qty", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„ÙƒÙ…ÙŠØ©", 90),
            ("lines", "Ø¹Ø¯Ø¯ Ø§Ù„Ø­Ø¬ÙˆØ²Ø§Øª", 95),
            ("branches", "Ø¹Ø¯Ø¯ Ø§Ù„ÙØ±ÙˆØ¹", 85),
            ("avg_price", "Ø³Ø¹Ø± ÙˆØ³ÙŠØ·", 80),
            ("sum_total", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¨Ø§Ù„Øº", 95),
            ("sum_paid", "Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø¯ÙÙˆØ¹", 95),
            ("updated", "Ø¢Ø®Ø± ØªØ­Ø¯ÙŠØ«", 130),
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
                    f"{format_money(float(r.get('unit_price') or 0))}",
                    f"{format_money(float(r.get('total_amount') or 0))}",
                    f"{format_money(float(r.get('paid_amount') or 0))}",
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
                    f"{format_money(float(r.get('avg_unit_price') or 0))}",
                    f"{format_money(float(r.get('sum_total_amount') or 0))}",
                    f"{format_money(float(r.get('sum_paid_amount') or 0))}",
                    r.get("last_updated") or "",
                ),
            )
        apply_zebra_tags(self._tree_agg)


class PosBranchFinancialWindow(tk.Toplevel):
    """Day-by-day POS cashflow totals from synced events."""

    _CAT_AR = {
        "sale": "Ù…Ø¨ÙŠØ¹Ø§Øª",
        "return_bill": "Ù…Ø±ØªØ¬Ø¹Ø§Øª",
        "void_bill": "Ø¥Ù„ØºØ§Ø¡ ÙÙˆØ§ØªÙŠØ±",
        "exchange_net": "Ø§Ø³ØªØ¨Ø¯Ø§Ù„ (ØµØ§ÙÙŠ)",
        "reservation_downpayment": "Ø¹Ø±Ø¨ÙˆÙ† Ø­Ø¬Ø²",
        "reservation_payment": "Ø¯ÙØ¹Ø§Øª Ø­Ø¬Ø²",
        "reservation_collect": "ØªØ­ØµÙŠÙ„ Ø¹Ù†Ø¯ Ø§Ù„ØªØ³Ù„ÙŠÙ…",
    }

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master)
        self.db = db
        self.title("Ø§Ù„ØªØ¯ÙÙ‚Ø§Øª Ø§Ù„Ù…Ø§Ù„ÙŠØ© Ù„Ù„ÙØ±ÙˆØ¹ (ÙŠÙˆÙ…ÙŠ)")
        self.geometry("1120x580")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="Ù…Ù† ØªØ§Ø±ÙŠØ®:").pack(side=tk.RIGHT, padx=4)
        self._df = DateField(top, "")
        self._df.pack(side=tk.RIGHT)
        ttk.Label(top, text="Ø¥Ù„Ù‰:").pack(side=tk.RIGHT, padx=4)
        self._dt = DateField(top, "")
        self._dt.pack(side=tk.RIGHT)
        ttk.Label(top, text="Ø§Ù„Ø¬Ù‡Ø§Ø²:").pack(side=tk.RIGHT, padx=4)
        names = self.db.list_pos_reservations_mirror_device_picklist()
        self._dev = tk.StringVar(value="")
        ttk.Combobox(top, textvariable=self._dev, values=names, width=30, state="readonly").pack(side=tk.RIGHT, padx=4)
        ttk.Button(top, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh_summary).pack(side=tk.LEFT)

        mid = ttk.Panedwindow(self, orient=tk.VERTICAL)
        mid.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

        sum_frame = ttk.Frame(mid)
        mid.add(sum_frame, weight=2)
        cols = (
            "day", "sales", "returns", "voids", "exch", "rdep", "rpay", "rcoll", "net",
        )
        self._sum_tree = ttk.Treeview(sum_frame, columns=cols, show="headings", height=10)
        for col, txt, w in [
            ("day", "Ø§Ù„ÙŠÙˆÙ…", 100),
            ("sales", "Ù…Ø¨ÙŠØ¹Ø§Øª", 90),
            ("returns", "Ù…Ø±ØªØ¬Ø¹Ø§Øª", 90),
            ("voids", "Ø¥Ù„ØºØ§Ø¡", 80),
            ("exch", "Ø§Ø³ØªØ¨Ø¯Ø§Ù„", 90),
            ("rdep", "Ø¹Ø±Ø¨ÙˆÙ†", 80),
            ("rpay", "Ø¯ÙØ¹Ø§Øª Ø­Ø¬Ø²", 90),
            ("rcoll", "ØªØ­ØµÙŠÙ„ ØªØ³Ù„ÙŠÙ…", 95),
            ("net", "Ø§Ù„ØµØ§ÙÙŠ", 100),
        ]:
            self._sum_tree.heading(col, text=txt)
            self._sum_tree.column(col, width=w, anchor="center")
        ssb = ttk.Scrollbar(sum_frame, orient="vertical", command=self._sum_tree.yview)
        self._sum_tree.configure(yscrollcommand=ssb.set)
        self._sum_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ssb.pack(side=tk.RIGHT, fill=tk.Y)
        self._sum_tree.bind("<<TreeviewSelect>>", lambda _e: self._load_detail())
        _bind_mousewheel(self._sum_tree)

        det_fr = ttk.LabelFrame(mid, text="ØªÙØ§ØµÙŠÙ„ Ø§Ù„ÙŠÙˆÙ… Ø§Ù„Ù…Ø­Ø¯Ø¯")
        mid.add(det_fr, weight=3)
        dcols = ("cat", "amount", "etype", "eid", "meta")
        self._det_tree = ttk.Treeview(det_fr, columns=dcols, show="headings", height=14)
        for col, txt, w in [
            ("cat", "Ø§Ù„Ø¨Ù†Ø¯", 140),
            ("amount", "Ø§Ù„Ù…Ø¨Ù„Øº", 90),
            ("etype", "Ù†ÙˆØ¹ Ø§Ù„Ø­Ø¯Ø«", 160),
            ("eid", "Ù…Ø±Ø¬Ø¹", 220),
            ("meta", "Ù…Ù„Ø§Ø­Ø¸Ø©", 320),
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
                    f"{format_money(float(r.get('sales_amt') or 0))}",
                    f"{format_money(float(r.get('returns_amt') or 0))}",
                    f"{format_money(float(r.get('voids_amt') or 0))}",
                    f"{format_money(float(r.get('exchange_amt') or 0))}",
                    f"{format_money(float(r.get('res_dep_amt') or 0))}",
                    f"{format_money(float(r.get('res_pay_amt') or 0))}",
                    f"{format_money(float(r.get('res_coll_amt') or 0))}",
                    f"{format_money(float(r.get('net_amt') or 0))}",
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
                    f"{format_money(float(r.get('amount') or 0))}",
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
        self._row_branch_by_iid: Dict[str, str] = {}
        self.title("Ù…Ù„Ø®Øµ Ø¯ÙˆØ±Ø© Ø§Ù„ÙØ±ÙˆØ¹")
        self.geometry("1400x760")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="Ù…Ù† ØªØ§Ø±ÙŠØ®:").pack(side=tk.RIGHT, padx=4)
        self._df = DateField(top, "")
        self._df.pack(side=tk.RIGHT)
        ttk.Label(top, text="Ø¥Ù„Ù‰:").pack(side=tk.RIGHT, padx=4)
        self._dt = DateField(top, "")
        self._dt.pack(side=tk.RIGHT)
        ttk.Button(top, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh).pack(side=tk.LEFT)
        ttk.Button(top, text="ØªØ³Ø¬ÙŠÙ„ Ù…ØªØ­ØµÙ„ ÙØ¹Ù„ÙŠâ€¦", command=self._add_actual_receipt).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(top, text="Ø­Ø°Ù Ø§Ù„Ù…ØªØ­ØµÙ„ Ø§Ù„Ù…Ø­Ø¯Ø¯", command=self._delete_selected_receipt).pack(side=tk.LEFT, padx=(8, 0))

        cols = (
            "branch", "bills", "ship_qty", "ship_value", "cash_net",
            "actual_received", "stock_qty", "stock_value", "gap", "actual_gap",
        )
        self._tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        for col, txt, w in [
            ("branch", "Ø§Ù„ÙØ±Ø¹", 160),
            ("bills", "Ø¹Ø¯Ø¯ ÙÙˆØ§ØªÙŠØ± Ø§Ù„Ø´Ø­Ù†", 110),
            ("ship_qty", "ÙƒÙ…ÙŠØ© Ø§Ù„Ø´Ø­Ù†Ø§Øª", 95),
            ("ship_value", "Ù‚ÙŠÙ…Ø© Ø§Ù„Ø´Ø­Ù†Ø§Øª", 120),
            ("cash_net", "ØµØ§ÙÙŠ Ø§Ù„Ù…ØªØ­ØµÙ„", 110),
            ("actual_received", "Ø§Ù„Ù…Ø³ØªÙ„Ù… ÙØ¹Ù„ÙŠØ§Ù‹", 110),
            ("stock_qty", "ÙƒÙ…ÙŠØ© Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø§Ù„Ø­Ø§Ù„ÙŠ", 120),
            ("stock_value", "Ù‚ÙŠÙ…Ø© Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø§Ù„Ø­Ø§Ù„ÙŠ", 130),
            ("gap", "ÙØ¬ÙˆØ© Ø­Ø³Ø¨ POS", 120),
            ("actual_gap", "Ø§Ù„ÙØ¬ÙˆØ© Ø§Ù„ÙØ¹Ù„ÙŠØ©", 120),
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
        self._tree.bind("<<TreeviewSelect>>", lambda _e: self._refresh_receipts())

        receipts = ttk.LabelFrame(self, text="Ø§Ù„Ù…ØªØ­ØµÙ„ Ø§Ù„ÙØ¹Ù„ÙŠ Ø§Ù„Ù…Ø³Ø¬Ù„ Ù„Ù„ÙØ±Ø¹ Ø§Ù„Ù…Ø­Ø¯Ø¯")
        receipts.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 8))
        rwrap = ttk.Frame(receipts)
        rwrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        rcols = ("id", "date", "branch", "amount", "note")
        self._receipt_tree = ttk.Treeview(rwrap, columns=rcols, show="headings", height=7)
        for col, txt, w in [
            ("id", "#", 60),
            ("date", "Ø§Ù„ØªØ§Ø±ÙŠØ®", 110),
            ("branch", "Ø§Ù„ÙØ±Ø¹", 150),
            ("amount", "Ø§Ù„Ù…Ø¨Ù„Øº", 100),
            ("note", "Ù…Ù„Ø§Ø­Ø¸Ø©", 520),
        ]:
            self._receipt_tree.heading(col, text=txt)
            self._receipt_tree.column(col, width=w, anchor="center" if col != "note" else "w")
        rysb = ttk.Scrollbar(rwrap, orient="vertical", command=self._receipt_tree.yview)
        self._receipt_tree.configure(yscrollcommand=rysb.set)
        self._receipt_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._receipt_tree)

        self._sum = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._sum).pack(fill=tk.X, padx=10, pady=(0, 8))

        self._refresh()

    def _selected_branch_device(self) -> Optional[str]:
        sel = self._tree.selection()
        if not sel:
            return None
        return self._row_branch_by_iid.get(sel[0])

    def _add_actual_receipt(self) -> None:
        branch_device = self._selected_branch_device()
        if not branch_device:
            messagebox.showwarning("Ø§Ø®ØªØ± Ø§Ù„ÙØ±Ø¹", "Ø§Ø®ØªØ± ÙØ±Ø¹Ø§Ù‹ Ù…Ù† Ø§Ù„Ø¬Ø¯ÙˆÙ„ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return

        dlg = tk.Toplevel(self)
        dlg.title("ØªØ³Ø¬ÙŠÙ„ Ù…ØªØ­ØµÙ„ ÙØ¹Ù„ÙŠ")
        dlg.transient(self)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=f"Ø§Ù„ÙØ±Ø¹: {branch_display_name(branch_device)}", font=("Segoe UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 8)
        )
        ttk.Label(frm, text="Ø§Ù„ØªØ§Ø±ÙŠØ®:").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        date_field = DateField(frm, "")
        current_df = self._df.get()
        if current_df:
            date_field.set(current_df)
        date_field.grid(row=1, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frm, text="Ø§Ù„Ù…Ø¨Ù„Øº:").grid(row=2, column=0, sticky="e", padx=4, pady=4)
        amount_var = tk.StringVar()
        ttk.Entry(frm, textvariable=amount_var, width=18).grid(row=2, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(frm, text="Ù…Ù„Ø§Ø­Ø¸Ø©:").grid(row=3, column=0, sticky="ne", padx=4, pady=4)
        note_txt = tk.Text(frm, width=40, height=4)
        note_txt.grid(row=3, column=1, sticky="ew", padx=4, pady=4)
        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(8, 0))

        def _save() -> None:
            try:
                amount = parse_float_text(amount_var.get())
                if amount is None:
                    raise ValueError
            except Exception:
                messagebox.showerror("Ù…Ø¨Ù„Øº ØºÙŠØ± ØµØ§Ù„Ø­", "Ø£Ø¯Ø®Ù„ Ù…Ø¨Ù„ØºØ§Ù‹ Ø±Ù‚Ù…ÙŠØ§Ù‹ ØµØ­ÙŠØ­Ø§Ù‹.", parent=dlg)
                return
            try:
                self.db.add_branch_cash_receipt(
                    branch_device,
                    amount,
                    received_at=date_field.get() or None,
                    note=note_txt.get("1.0", tk.END).strip(),
                )
            except Exception as ex:
                messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­ÙØ¸", str(ex), parent=dlg)
                return
            dlg.destroy()
            self._refresh()

        ttk.Button(btns, text="Ø¥Ù„ØºØ§Ø¡", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="Ø­ÙØ¸", command=_save).pack(side=tk.RIGHT, padx=6)

    def _delete_selected_receipt(self) -> None:
        sel = self._receipt_tree.selection()
        if not sel:
            messagebox.showwarning("Ø§Ø®ØªØ± Ù…ØªØ­ØµÙ„Ø§Ù‹", "Ø§Ø®ØªØ± Ø³Ø¬Ù„ Ù…ØªØ­ØµÙ„ ÙØ¹Ù„ÙŠ Ù…Ù† Ø§Ù„Ù‚Ø§Ø¦Ù…Ø© Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        vals = self._receipt_tree.item(sel[0], "values")
        if not vals:
            return
        rid = parse_int_text(vals[0])
        if rid is None:
            messagebox.showerror("Ø¨ÙŠØ§Ù†Ø§Øª ØºÙŠØ± ØµØ§Ù„Ø­Ø©", "Ù…Ø¹Ø±Ù Ø§Ù„Ù…ØªØ­ØµÙ„ ØºÙŠØ± ØµØ§Ù„Ø­.", parent=self)
            return
        if not messagebox.askyesno("ØªØ£ÙƒÙŠØ¯ Ø§Ù„Ø­Ø°Ù", "Ø­Ø°Ù Ø§Ù„Ù…ØªØ­ØµÙ„ Ø§Ù„Ù…Ø­Ø¯Ø¯ØŸ", parent=self):
            return
        try:
            self.db.delete_branch_cash_receipt(rid)
        except Exception as ex:
            messagebox.showerror("ÙØ´Ù„ Ø§Ù„Ø­Ø°Ù", str(ex), parent=self)
            return
        self._refresh()

    def _refresh_receipts(self) -> None:
        self._receipt_tree.delete(*self._receipt_tree.get_children())
        branch_device = self._selected_branch_device()
        rows = self.db.list_branch_cash_receipts(
            branch_device=branch_device,
            date_from=(self._df.get() or "").strip() or None,
            date_to=(self._dt.get() or "").strip() or None,
        )
        for r in rows:
            self._receipt_tree.insert(
                "",
                tk.END,
                values=(
                    int(r.get("id") or 0),
                    r.get("received_at") or "",
                    branch_display_name(r.get("branch_device") or ""),
                    f"{format_money(float(r.get('amount') or 0.0))}",
                    r.get("note") or "",
                ),
            )
        apply_zebra_tags(self._receipt_tree)

    def _refresh(self):
        self._tree.delete(*self._tree.get_children())
        df = (self._df.get() or "").strip() or None
        dt = (self._dt.get() or "").strip() or None
        rows = self.db.list_branch_cycle_reconciliation(df, dt)
        t_ship = 0.0
        t_cash = 0.0
        t_actual = 0.0
        t_stock = 0.0
        t_gap = 0.0
        t_actual_gap = 0.0
        self._row_branch_by_iid.clear()
        for r in rows:
            t_ship += float(r.get("shipment_value") or 0.0)
            t_cash += float(r.get("cash_net") or 0.0)
            t_actual += float(r.get("actual_received") or 0.0)
            t_stock += float(r.get("stock_value") or 0.0)
            t_gap += float(r.get("cycle_gap") or 0.0)
            t_actual_gap += float(r.get("actual_gap") or 0.0)
            iid = self._tree.insert(
                "",
                tk.END,
                values=(
                    r.get("branch_name") or "",
                    int(r.get("shipment_bills_count") or 0),
                    int(r.get("shipment_qty") or 0),
                    f"{format_money(float(r.get('shipment_value') or 0.0))}",
                    f"{format_money(float(r.get('cash_net') or 0.0))}",
                    f"{format_money(float(r.get('actual_received') or 0.0))}",
                    int(r.get("stock_qty") or 0),
                    f"{format_money(float(r.get('stock_value') or 0.0))}",
                    f"{format_money(float(r.get('cycle_gap') or 0.0))}",
                    f"{format_money(float(r.get('actual_gap') or 0.0))}",
                ),
            )
            self._row_branch_by_iid[iid] = str(r.get("branch_device") or "")
        apply_zebra_tags(self._tree)
        self._sum.set(
            f"Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ù‚ÙŠÙ…Ø© Ø§Ù„Ø´Ø­Ù†Ø§Øª: {format_money(t_ship)}  |  Ø¥Ø¬Ù…Ø§Ù„ÙŠ ØµØ§ÙÙŠ Ø§Ù„Ù…ØªØ­ØµÙ„ (POS): {format_money(t_cash)}  |  "
            f"Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ø§Ù„Ù…Ø³ØªÙ„Ù… ÙØ¹Ù„ÙŠØ§Ù‹: {format_money(t_actual)}  |  Ø¥Ø¬Ù…Ø§Ù„ÙŠ Ù‚ÙŠÙ…Ø© Ø§Ù„Ù…Ø®Ø²ÙˆÙ† Ø§Ù„Ø­Ø§Ù„ÙŠ: {format_money(t_stock)}  |  "
            f"ÙØ¬ÙˆØ© POS: {format_money(t_gap)}  |  Ø§Ù„ÙØ¬ÙˆØ© Ø§Ù„ÙØ¹Ù„ÙŠØ©: {format_money(t_actual_gap)}"
        )
        self._refresh_receipts()


# ------------------- App -------------------

class PriceProfileManagerWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±")
        self.geometry("980x620")
        self.configure(bg=_UI["BG"])
        self._selected_profile_id: Optional[int] = None
        self._profile_rows: List[Dict[str, Any]] = []
        self._row_vars: List[Tuple[tk.StringVar, tk.StringVar]] = []
        self._item_type_var = tk.StringVar(value="")
        self._import_item_var = tk.StringVar(value="")
        self._profile_title_var = tk.StringVar(value="Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø³Ø¹Ø±")
        self._build()
        self._refresh_profiles()

    def _build(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill=tk.BOTH, expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(root, text="Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª")
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 8))
        right = ttk.LabelFrame(root, text="Ø§Ù„ØªÙØ§ØµÙŠÙ„")
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        self._profiles_list = tk.Listbox(left, height=24)
        self._profiles_list.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._profiles_list.bind("<<ListboxSelect>>", lambda _e: self._on_profile_selected())

        left_btns = ttk.Frame(left)
        left_btns.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Button(left_btns, text="Ø¬Ø¯ÙŠØ¯", command=self._create_profile).pack(side=tk.LEFT)
        ttk.Button(left_btns, text="Ø¥Ø¹Ø§Ø¯Ø© ØªØ³Ù…ÙŠØ©", command=self._rename_profile).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(left_btns, text="Ø­Ø°Ù", command=self._delete_profile).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Label(right, textvariable=self._profile_title_var, font=_FONTS["h2"]).grid(row=0, column=0, sticky="w", padx=8, pady=(8, 4))

        item_bar = ttk.Frame(right)
        item_bar.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        item_bar.columnconfigure(1, weight=1)
        ttk.Label(item_bar, text="Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù").grid(row=0, column=0, sticky="w")
        self._item_type_combo = LabeledCombobox(item_bar, "", self.db, "item_type")
        self._item_type_combo.var = self._item_type_var
        self._item_type_combo.cb.configure(textvariable=self._item_type_var)
        self._item_type_combo.grid(row=0, column=1, sticky="ew", padx=6)
        self._item_type_combo.cb.bind("<<ComboboxSelected>>", lambda _e: self._load_item_prices(), add="+")
        self._item_type_combo.cb.bind("<Return>", lambda _e: self._load_item_prices(), add="+")
        self._item_type_combo.cb.bind("<FocusOut>", lambda _e: self._load_item_prices(), add="+")
        ttk.Button(item_bar, text="ØªØ­Ù…ÙŠÙ„", command=self._load_item_prices).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(item_bar, text="Ø¥Ø¶Ø§ÙØ© Ù…Ù‚Ø§Ø³", command=self._add_custom_size).grid(row=0, column=3)

        ttk.Label(item_bar, text="Ø§Ø³ØªÙŠØ±Ø§Ø¯ Ù…Ø«Ù„ ØµÙ†Ù").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self._import_item_combo = LabeledCombobox(item_bar, "", self.db, "item_type")
        self._import_item_combo.var = self._import_item_var
        self._import_item_combo.cb.configure(textvariable=self._import_item_var)
        self._import_item_combo.grid(row=1, column=1, sticky="ew", padx=6, pady=(6, 0))
        self._import_item_combo.cb.bind("<<ComboboxSelected>>", lambda _e: self._import_prices_from_item(), add="+")
        ttk.Button(item_bar, text="Ø§Ø³ØªÙŠØ±Ø§Ø¯", command=self._import_prices_from_item).grid(row=1, column=2, padx=(0, 6), pady=(6, 0))

        canvas_wrap = ttk.Frame(right)
        canvas_wrap.grid(row=2, column=0, sticky="nsew", padx=8, pady=(0, 8))
        canvas_wrap.columnconfigure(0, weight=1)
        canvas_wrap.rowconfigure(0, weight=1)
        self._rows_canvas = tk.Canvas(canvas_wrap, highlightthickness=0)
        rows_scroll = ttk.Scrollbar(canvas_wrap, orient="vertical", command=self._rows_canvas.yview)
        self._rows_canvas.configure(yscrollcommand=rows_scroll.set)
        self._rows_canvas.grid(row=0, column=0, sticky="nsew")
        rows_scroll.grid(row=0, column=1, sticky="ns")
        self._rows_host = ttk.Frame(self._rows_canvas)
        self._rows_window = self._rows_canvas.create_window((0, 0), window=self._rows_host, anchor="nw")
        self._rows_host.bind("<Configure>", lambda _e=None: self._sync_rows_canvas())
        self._rows_canvas.bind("<Configure>", lambda _e=None: self._sync_rows_canvas())
        _bind_mousewheel(canvas_wrap, self._rows_canvas)

        action_bar = ttk.Frame(right)
        action_bar.grid(row=3, column=0, sticky="ew", padx=8, pady=(0, 8))
        ttk.Button(action_bar, text="Ø­ÙØ¸ Ø£Ø³Ø¹Ø§Ø± Ù‡Ø°Ø§ Ø§Ù„ØµÙ†Ù", command=self._save_current_item_prices).pack(side=tk.LEFT)
        ttk.Button(action_bar, text="ØªØ­Ø¯ÙŠØ«", command=self._refresh_current_profile).pack(side=tk.LEFT, padx=(6, 0))

    def _sync_rows_canvas(self):
        try:
            self._rows_canvas.configure(scrollregion=self._rows_canvas.bbox("all"))
            self._rows_canvas.itemconfigure(self._rows_window, width=self._rows_canvas.winfo_width())
        except Exception:
            pass

    def _refresh_profiles(self, select_profile_id: Optional[int] = None):
        self._profile_rows = self.db.list_price_profiles()
        self._profiles_list.delete(0, tk.END)
        selected_index = None
        for idx, row in enumerate(self._profile_rows):
            label = str(row.get("name") or "").strip()
            count = int(row.get("assignment_count") or 0)
            suffix = f" ({count})" if count else ""
            self._profiles_list.insert(tk.END, f"{label}{suffix}")
            if select_profile_id and int(row["id"]) == int(select_profile_id):
                selected_index = idx
        if selected_index is None and self._profile_rows:
            selected_index = 0
        if selected_index is not None:
            self._profiles_list.selection_clear(0, tk.END)
            self._profiles_list.selection_set(selected_index)
            self._profiles_list.activate(selected_index)
            self._on_profile_selected()
        else:
            self._selected_profile_id = None
            self._profile_title_var.set("Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø³Ø¹Ø±")
            self._item_type_combo.set_supplier(lambda: [])
            self._import_item_combo.set_supplier(lambda: [])
            self._item_type_var.set("")
            self._import_item_var.set("")
            self._render_price_rows([])

    def _selected_profile_row(self) -> Optional[Dict[str, Any]]:
        if self._selected_profile_id is None:
            return None
        for row in self._profile_rows:
            if int(row["id"]) == int(self._selected_profile_id):
                return row
        return None

    def _on_profile_selected(self):
        sel = self._profiles_list.curselection()
        if not sel:
            return
        idx = parse_int_text(sel[0], -1)
        if idx < 0 or idx >= len(self._profile_rows):
            return
        self._selected_profile_id = int(self._profile_rows[idx]["id"])
        self._refresh_current_profile()

    def _refresh_current_profile(self):
        row = self._selected_profile_row()
        if row is None:
            return
        self._profile_title_var.set(str(row.get("name") or "").strip())
        item_types = sorted(
            {
                *[str(v or "").strip() for v in self.db.get_distinct("item_type")],
                *self.db.list_price_profile_item_types(self._selected_profile_id),
            },
            key=lambda s: s.casefold(),
        )
        self._item_type_combo.set_supplier(lambda vals=item_types: vals[:])
        self._item_type_combo.refresh_values()
        self._refresh_import_item_choices()
        if (self._item_type_var.get() or "").strip() not in item_types:
            self._item_type_var.set(item_types[0] if item_types else "")
        self._load_item_prices()

    def _refresh_import_item_choices(self):
        if not self._selected_profile_id:
            choices: List[str] = []
        else:
            current = (self._item_type_var.get() or "").strip().casefold()
            choices = [
                item for item in self.db.list_price_profile_item_types(self._selected_profile_id)
                if item.strip().casefold() != current
            ]
        self._import_item_combo.set_supplier(lambda vals=choices: vals[:])
        self._import_item_combo.refresh_values()
        if (self._import_item_var.get() or "").strip() not in choices:
            self._import_item_var.set("")

    def _render_price_rows(self, rows: Sequence[Tuple[str, Optional[float]]]):
        for child in self._rows_host.winfo_children():
            child.destroy()
        self._row_vars = []
        if not rows:
            ttk.Label(self._rows_host, text="Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø«Ù… Ù†ÙˆØ¹ ØµÙ†Ù Ù„Ø¹Ø±Ø¶ Ø§Ù„Ù…Ù‚Ø§Ø³Ø§Øª ÙˆØ§Ù„Ø£Ø³Ø¹Ø§Ø±.").pack(anchor="w", padx=8, pady=8)
            return
        header = ttk.Frame(self._rows_host)
        header.pack(fill=tk.X, padx=8, pady=(6, 2))
        ttk.Label(header, text="Ø§Ù„Ù…Ù‚Ø§Ø³", width=20).pack(side=tk.LEFT)
        ttk.Label(header, text="Ø§Ù„Ø³Ø¹Ø±", width=20).pack(side=tk.LEFT, padx=(8, 0))
        for size, price in rows:
            row = ttk.Frame(self._rows_host)
            row.pack(fill=tk.X, padx=8, pady=2)
            size_var = tk.StringVar(value=str(size or "").strip())
            price_var = tk.StringVar(value="" if price is None else format_money(float(price)))
            ttk.Entry(row, textvariable=size_var, width=18).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=price_var, width=18).pack(side=tk.LEFT, padx=(8, 0))
            self._row_vars.append((size_var, price_var))

    def _load_item_prices(self):
        if not self._selected_profile_id:
            self._render_price_rows([])
            return
        item_type = (self._item_type_var.get() or "").strip()
        if not item_type:
            self._render_price_rows([])
            return
        self._refresh_import_item_choices()
        sizes = self.db.list_sizes_for_price_profile_item(item_type)
        existing = {
            str(row.get("size") or "").strip(): row.get("price")
            for row in self.db.list_price_profile_lines(self._selected_profile_id, item_type=item_type)
        }
        merged_sizes = list(sizes)
        for size in existing:
            if size and size not in merged_sizes:
                merged_sizes.append(size)
        merged_sizes.sort(key=warehouse_size_sort_key)
        self._render_price_rows([(size, existing.get(size)) for size in merged_sizes])

    def _import_prices_from_item(self):
        if not self._selected_profile_id:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        target_item = (self._item_type_var.get() or "").strip()
        source_item = (self._import_item_var.get() or "").strip()
        if not target_item:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ù†ÙˆØ¹ Ø§Ù„ØµÙ†Ù Ø§Ù„Ø­Ø§Ù„ÙŠ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not source_item:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ø§Ù„ØµÙ†Ù Ø§Ù„Ø°ÙŠ ØªØ±ÙŠØ¯ Ø§Ù„Ø§Ø³ØªÙŠØ±Ø§Ø¯ Ù…Ù†Ù‡.", parent=self)
            return
        if source_item.casefold() == target_item.casefold():
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± ØµÙ†ÙØ§Ù‹ Ù…Ø®ØªÙ„ÙØ§Ù‹ Ù„Ù„Ø§Ø³ØªÙŠØ±Ø§Ø¯ Ù…Ù†Ù‡.", parent=self)
            return

        source_rows = self.db.list_price_profile_lines(self._selected_profile_id, item_type=source_item)
        source_prices = {
            str(row.get("size") or "").strip(): row.get("price")
            for row in source_rows
            if str(row.get("size") or "").strip() and row.get("price") is not None
        }
        if not source_prices:
            messagebox.showinfo("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ù„ØµÙ†Ù Ø§Ù„Ù…Ø®ØªØ§Ø± Ù„Ø§ ÙŠØ­ØªÙˆÙŠ Ø¹Ù„Ù‰ Ø£Ø³Ø¹Ø§Ø± Ù…Ø­ÙÙˆØ¸Ø© ÙÙŠ Ù‡Ø°Ø§ Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„.", parent=self)
            return

        current_sizes = [sv.get().strip() for sv, _ in self._row_vars if sv.get().strip()]
        merged_sizes = list(current_sizes)
        for size in source_prices:
            if size and size not in merged_sizes:
                merged_sizes.append(size)
        merged_sizes.sort(key=warehouse_size_sort_key)

        existing_current: Dict[str, Optional[float]] = {}
        for size_var, price_var in self._row_vars:
            size = size_var.get().strip()
            raw_price = price_var.get().strip()
            if size:
                existing_current[size] = parse_float_text(raw_price) if raw_price else None

        rendered_rows: List[Tuple[str, Optional[float]]] = []
        copied = 0
        for size in merged_sizes:
            if size in source_prices:
                rendered_rows.append((size, parse_float_text(source_prices[size])))
                copied += 1
            else:
                rendered_rows.append((size, existing_current.get(size)))
        self._render_price_rows(rendered_rows)
        show_toast(self, f"ØªÙ… Ø§Ø³ØªÙŠØ±Ø§Ø¯ Ø£Ø³Ø¹Ø§Ø± {copied} Ù…Ù‚Ø§Ø³ Ù…Ù† {source_item}")

    def _add_custom_size(self):
        if not self._selected_profile_id:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        size = simpledialog.askstring("Ù…Ù‚Ø§Ø³ Ø¬Ø¯ÙŠØ¯", "Ø£Ø¯Ø®Ù„ Ø§Ø³Ù… Ø§Ù„Ù…Ù‚Ø§Ø³:", parent=self)
        size = str(size or "").strip()
        if not size:
            return
        rows = [(sv.get().strip(), pv.get().strip()) for sv, pv in self._row_vars]
        sizes = [s for s, _ in rows if s]
        if size not in sizes:
            sizes.append(size)
        sizes.sort(key=warehouse_size_sort_key)
        existing_prices = {s: p for s, p in rows}
        rendered_rows: List[Tuple[str, Optional[float]]] = []
        for label in sizes:
            raw_price = existing_prices.get(label)
            if raw_price in ("", None):
                rendered_rows.append((label, None))
                continue
            try:
                rendered_rows.append((label, parse_float_text(raw_price)))
            except (TypeError, ValueError):
                rendered_rows.append((label, None))
        self._render_price_rows(rendered_rows)

    def _save_current_item_prices(self):
        if not self._selected_profile_id:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        item_type = (self._item_type_var.get() or "").strip()
        if not item_type:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ù†ÙˆØ¹ ØµÙ†Ù Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        rows: List[Dict[str, Any]] = []
        for size_var, price_var in self._row_vars:
            size = size_var.get().strip()
            price_txt = price_var.get().strip()
            if not size:
                continue
            if not price_txt:
                continue
            price_val = parse_float_text(price_txt)
            if price_val is None:
                messagebox.showwarning("Ø³Ø¹Ø± ØºÙŠØ± ØµØ§Ù„Ø­", f"Ø§Ù„Ø³Ø¹Ø± ØºÙŠØ± ØµØ§Ù„Ø­ Ù„Ù„Ù…Ù‚Ø§Ø³ {size}.", parent=self)
                return
            rows.append({"size": size, "price": price_val})
        try:
            self.db.replace_price_profile_item_prices(self._selected_profile_id, item_type, rows)
            result = self.db.apply_price_profile_item_to_assigned_stock(
                self._selected_profile_id,
                item_type,
                note="Price profile prices updated",
            )
            self._refresh_profiles(select_profile_id=self._selected_profile_id)
            show_toast(
                self,
                f"ØªÙ… ØªØ­Ø¯ÙŠØ« {result.get('updated', 0)} ØµÙ Ù…Ø®Ø²ÙˆÙ† Ù…Ø±ØªØ¨Ø· Ø¨Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„",
                bg="#16a34a" if int(result.get("updated", 0) or 0) else "#f59e0b",
            )
            show_toast(self, "ØªÙ… Ø­ÙØ¸ Ø£Ø³Ø¹Ø§Ø± Ø§Ù„Ø¨Ø±ÙˆÙØ§ÙŠÙ„")
        except Exception as ex:
            messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=self)

    def _create_profile(self):
        name = simpledialog.askstring("Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø¬Ø¯ÙŠØ¯", "Ø§Ø³Ù… Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±:", parent=self)
        name = str(name or "").strip()
        if not name:
            return
        try:
            profile_id = self.db.create_price_profile(name)
            self._refresh_profiles(select_profile_id=profile_id)
        except Exception as ex:
            messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=self)

    def _rename_profile(self):
        row = self._selected_profile_row()
        if row is None:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        name = simpledialog.askstring("Ø¥Ø¹Ø§Ø¯Ø© ØªØ³Ù…ÙŠØ©", "Ø§Ø³Ù… Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±:", initialvalue=str(row.get("name") or "").strip(), parent=self)
        name = str(name or "").strip()
        if not name:
            return
        try:
            self.db.rename_price_profile(int(row["id"]), name)
            self._refresh_profiles(select_profile_id=int(row["id"]))
        except Exception as ex:
            messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=self)

    def _delete_profile(self):
        row = self._selected_profile_row()
        if row is None:
            messagebox.showwarning("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", "Ø§Ø®ØªØ± Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø£ÙˆÙ„Ø§Ù‹.", parent=self)
            return
        if not messagebox.askyesno("Ø­Ø°Ù", f"Ø­Ø°Ù Ø¨Ø±ÙˆÙØ§ÙŠÙ„ Ø§Ù„Ø³Ø¹Ø±: {row.get('name')}", parent=self):
            return
        try:
            self.db.delete_price_profile(int(row["id"]))
            self._refresh_profiles()
        except Exception as ex:
            messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=self)


class WarehouseApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„Ù…Ø®Ø§Ø²Ù†")
        self.geometry("1320x760")
        self.db = SqliteDatabase(DB_PATH)
        self._dark_mode = False
        # POS-like defaults for classic Tk widgets across all windows/dialogs.
        self.option_add("*Font", "{Segoe UI} 9")
        self.option_add("*Listbox.Font", "{Segoe UI} 9")
        self.option_add("*Menu.Font", "{Segoe UI} 9")
        self.option_add("*Message.font", "{Segoe UI} 9")

        # Apply palette BEFORE creating widgets so ttk picks it up
        self._apply_colorful_theme(True)

        self._build()
        self._bind_shortcuts()
        self._tick_clock()
        self.after(2600, self._poll_shipment_receipt_reviews)
        try:
            import sync_periodic

            sync_periodic.attach_periodic_sync(self, self.db.path)
        except Exception:
            pass

    def _poll_shipment_receipt_reviews(self):
        try:
            review = self.db.get_next_shipment_receipt_review()
            if review:
                self._open_shipment_receipt_review_popup(review)
                self.db.mark_shipment_receipt_review_shown(int(review["id"]))
        except Exception:
            pass
        self.after(5000, self._poll_shipment_receipt_reviews)

    def _open_shipment_receipt_review_popup(self, review: Dict[str, Any]) -> None:
        source = str(review.get("source_device") or "").strip() or "POS"
        shipment = str(review.get("shipment_uuid") or "").strip()
        wrong_count = self.db.get_branch_wrong_bill_count(source)
        payload = []
        try:
            raw = json.loads(review.get("payload_json") or "[]")
            if isinstance(raw, list):
                payload = [x for x in raw if isinstance(x, dict)]
        except Exception:
            payload = []
        diffs = [x for x in payload if int(x.get("diff_qty") or 0) != 0]
        if not diffs:
            diffs = payload
        lines_preview = []
        for d in diffs[:8]:
            lines_preview.append(
                f"- {d.get('item_type','')} | {d.get('school','')} | {d.get('color','')} | {d.get('size','')}: "
                f"Ù…Ø±Ø³Ù„ {int(d.get('expected_qty') or 0)} / Ù…Ø³ØªÙ„Ù… {int(d.get('received_qty') or 0)}"
            )
        preview_txt = "\n".join(lines_preview) if lines_preview else "(Ù„Ø§ ØªÙˆØ¬Ø¯ ÙØ±ÙˆÙ‚Ø§Øª)"
        msg = (
            f"Ù…Ø±Ø§Ø¬Ø¹Ø© ÙØ±Ù‚ Ø§Ø³ØªÙ„Ø§Ù… Ø´Ø­Ù†Ø© Ù…Ù†: {source}\n"
            f"Ù…Ø±Ø¬Ø¹ Ø§Ù„Ø´Ø­Ù†Ø©: {shipment[:8] if shipment else '-'}\n"
            f"Ø¹Ø¯Ø¯ Ù…Ø±Ø§Øª Ø¥Ø¨Ù„Ø§Øº Ù‡Ø°Ø§ Ø§Ù„ÙØ±Ø¹ Ø¨ÙˆØ¬ÙˆØ¯ Ø®Ø·Ø£: {wrong_count}\n\n"
            f"{preview_txt}\n\n"
            f"Ù‡Ù„ ØªØ±ÙŠØ¯ Ø§Ø¹ØªÙ…Ø§Ø¯ Ø§Ù„ÙƒÙ…ÙŠØ§Øª Ø§Ù„ØªÙŠ Ø£Ø¨Ù„Øº Ø¨Ù‡Ø§ Ø§Ù„ÙØ±Ø¹ØŸ"
        )
        accept = messagebox.askyesno("Ù…Ø±Ø§Ø¬Ø¹Ø© ÙØ±ÙˆÙ‚ Ø´Ø­Ù†Ø© ÙØ±Ø¹", msg, parent=self)
        note = ""
        if not accept:
            note = simpledialog.askstring("Ø³Ø¨Ø¨ Ø§Ù„Ø±ÙØ¶ (Ø§Ø®ØªÙŠØ§Ø±ÙŠ)", "Ø§ÙƒØªØ¨ Ù…Ù„Ø§Ø­Ø¸Ø© Ù„Ù„Ø±ÙØ¶:", parent=self) or ""
        try:
            out = self.db.decide_shipment_receipt_review(int(review["id"]), bool(accept), note)
            if accept:
                show_toast(
                    self,
                    f"ØªÙ… Ù‚Ø¨ÙˆÙ„ Ù…Ø±Ø§Ø¬Ø¹Ø© Ø§Ù„Ø´Ø­Ù†Ø© ({out.get('adjustments_applied', 0)} ØªØ¹Ø¯ÙŠÙ„).",
                    bg="#166534",
                    fg="#dcfce7",
                )
            else:
                show_toast(
                    self,
                    "ØªÙ… Ø±ÙØ¶ Ù…Ø±Ø§Ø¬Ø¹Ø© Ø§Ù„Ø´Ø­Ù†Ø©.",
                    bg="#854d0e",
                    fg="#fef9c3",
                )
        except Exception as ex:
            messagebox.showerror("Ø®Ø·Ø£", str(ex), parent=self)

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
            "fabric_calculation": 3,
            "sync_diagnostics": 4,
            "statistics": 5,
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
            elif idx == 4 and hasattr(self, "_sync_diagnostics"):
                self._sync_diagnostics._refresh()
            elif idx == 5 and hasattr(self, "_statistics"):
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

        # -------- POS light palette --------
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
        s.configure("TLabelframe", background=SURFACE, bordercolor=BORDER, relief="groove")
        s.configure("TLabelframe.Label", background=SURFACE, foreground=TEXT_SEC,
                    font=("Segoe UI", 10, "bold"))

        s.configure("TLabel",  background=SURFACE, foreground=TEXT)
        s.configure("TButton", background=ACCENT, foreground="white",
                    bordercolor=ACCENT, focusthickness=1, focuscolor=ACCENT_H, padding=[12, 6],
                    font=("Segoe UI", 9, "bold"))
        s.map("TButton",
            background=[("active", ACCENT_H), ("pressed", ACCENT_H)],
            bordercolor=[("focus", BRAND)])

        # entries / combos â€” white field, blue focus border
        for sty in ("TEntry", "TCombobox", "TSpinbox"):
            s.configure(sty, fieldbackground=SURFACE, background=SURFACE,
                        foreground=TEXT, bordercolor=BORDER,
                        lightcolor=BORDER, darkcolor=BORDER,
                        arrowcolor=TEXT_SEC, padding=4)
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
        s.configure("TNotebook.Tab", background=SURFACE, foreground=TEXT_DIM,
                    padding=[14, 6], font=("Segoe UI", 9, "bold"))
        s.map("TNotebook.Tab",
            background=[("selected", BG)],
            foreground=[("selected", ACCENT)])

        # -------- Treeview (tables) --------
        s.configure("Treeview",
                    background=SURFACE, fieldbackground=SURFACE, foreground=TEXT,
                    bordercolor=BORDER, rowheight=26,
                    font=("Segoe UI", 9))
        s.configure("Treeview.Heading",
                    background=SURFACE2, foreground=TEXT_SEC, bordercolor=BORDER,
                    font=("Segoe UI", 9, "bold"), padding=4)
        s.map("Treeview",
            background=[("selected", SEL_BG)],
            foreground=[("selected", SEL_FG)])

        s.configure("TScrollbar", background=SURFACE, troughcolor=BG,
                    bordercolor=BORDER, arrowcolor=TEXT)

        # -------- PanedWindow --------
        s.configure("TPanedwindow", background=BG)
        s.configure("Sash", sashthickness=4, gripcount=0)

    def _build(self):
        # ======== MENU BAR ========
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        self._colorful_var = tk.BooleanVar(value=True)
        self._dark_var = tk.BooleanVar(value=False)

        def _toggle_theme():
            self._apply_colorful_theme(self._colorful_var.get())

        def _toggle_dark():
            self._dark_mode = self._dark_var.get()
            self._apply_dark_mode(self._dark_mode)

        warehouse_menu = tk.Menu(menubar, tearoff=False)
        warehouse_menu.add_command(label="Ø¹Ø±Ø¶ Ø§Ù„Ù…Ø®Ø²ÙˆÙ† (Ø¬Ø¯ÙˆÙ„)      Ctrl+I", command=self._open_inventory)
        warehouse_menu.add_command(label="Ø³Ø¬Ù„ Ø§Ù„Ø­Ø±ÙƒØ§Øª", command=self._open_movements)
        warehouse_menu.add_command(label="Ø³Ø¬Ù„ Ø§Ù„ÙÙˆØ§ØªÙŠØ±      Ctrl+P", command=self._open_bills_history)
        warehouse_menu.add_separator()
        warehouse_menu.add_command(label="ØªØ­ÙˆÙŠÙ„ Ø¨ÙŠÙ† Ø§Ù„Ù…Ø®Ø§Ø²Ù†", command=self._open_transfer)
        warehouse_menu.add_command(label="Ø¬Ø±Ø¯ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", command=self._open_audit)
        menubar.add_cascade(label="Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„Ù…Ø®Ø²Ù†", menu=warehouse_menu)

        branches_menu = tk.Menu(menubar, tearoff=False)
        branches_menu.add_command(label="Ù…Ø®Ø²ÙˆÙ† Ø§Ù„ÙØ±ÙˆØ¹", command=self._open_branch_stock)
        branches_menu.add_command(label="Ù…Ù†ØªØ¬Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ ØºÙŠØ± Ø§Ù„Ù…Ø¹Ø§Ù„Ø¬Ø©", command=self._open_branch_inventory_queue)
        branches_menu.add_separator()
        branches_menu.add_command(label="Ø³Ø¬Ù„ Ø´Ø­Ù†Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ ÙˆØ§Ù„ÙˆØ§Ø±Ø¯ Ø§Ù„Ù…ØªØ²Ø§Ù…Ù†â€¦", command=self._open_branch_bills_sync_log)
        branches_menu.add_command(label="Ø­Ø¬ÙˆØ²Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ (Ù…Ø±Ø¢Ø©)â€¦", command=self._open_pos_reservations_mirror)
        branches_menu.add_command(label="Ø§Ù„ØªØ¯ÙÙ‚Ø§Øª Ø§Ù„Ù…Ø§Ù„ÙŠØ© Ù„Ù„ÙØ±ÙˆØ¹ (ÙŠÙˆÙ…ÙŠ)â€¦", command=self._open_pos_financial_by_day)
        branches_menu.add_command(label="Ù…Ù„Ø®Øµ Ø¯ÙˆØ±Ø© Ø§Ù„ÙØ±ÙˆØ¹â€¦", command=self._open_branch_cycle_summary)
        menubar.add_cascade(label="Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„ÙØ±ÙˆØ¹", menu=branches_menu)

        pricing_menu = tk.Menu(menubar, tearoff=False)
        pricing_menu.add_command(label="Ø¥Ù†Ø´Ø§Ø¡ Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±â€¦", command=self._open_price_profiles)
        pricing_menu.add_command(label="Ø³Ø¬Ù„ ØªØ¹Ø¯ÙŠÙ„Ø§Øª Ø£Ø³Ø¹Ø§Ø± Ø§Ù„ÙØ±ÙˆØ¹â€¦", command=self._open_price_sync_audit)
        menubar.add_cascade(label="Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", menu=pricing_menu)

        sync_menu = tk.Menu(menubar, tearoff=False)
        sync_menu.add_command(label="Ù…Ø²Ø§Ù…Ù†Ø© Ø§Ù„Ø¢Ù†â€¦", command=self._open_sync_dialog)
        sync_menu.add_command(label="Ø¥Ø¹Ø¯Ø§Ø¯Ø§Øª Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©â€¦", command=self._open_sync_setup)
        sync_menu.add_separator()
        sync_menu.add_command(label="Ø³Ø¬Ù„ ØªØ¹Ø¯ÙŠÙ„Ø§Øª Ø£Ø³Ø¹Ø§Ø± Ø§Ù„ÙØ±ÙˆØ¹â€¦", command=self._open_price_sync_audit)
        sync_menu.add_command(label="ØªØ´Ø®ÙŠØµ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø© (F11)", command=lambda: self._switch_tab("sync_diagnostics"))
        menubar.add_cascade(label="Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", menu=sync_menu)

        system_menu = tk.Menu(menubar, tearoff=False)
        system_menu.add_command(label="Ø¥Ø¹Ø¯Ø§Ø¯Ø§Øª Ø§Ù„Ù…Ø¯ÙŠØ±â€¦", command=self._open_admin)
        system_menu.add_separator()

        appearance = tk.Menu(system_menu, tearoff=False)
        appearance.add_checkbutton(label="Ù…Ø¸Ù‡Ø± Ù…Ù„ÙˆÙ‘Ù†", variable=self._colorful_var, command=_toggle_theme)
        appearance.add_checkbutton(label="Ø§Ù„ÙˆØ¶Ø¹ Ø§Ù„Ù„ÙŠÙ„ÙŠ", variable=self._dark_var, command=_toggle_dark)
        system_menu.add_cascade(label="Ø§Ù„Ù…Ø¸Ù‡Ø±", menu=appearance)

        shortcuts_menu = tk.Menu(system_menu, tearoff=False)
        shortcuts_menu.add_command(label="F1  - Ø§Ù„ÙˆØ§Ø±Ø¯", state="disabled")
        shortcuts_menu.add_command(label="F2  - Ø§Ù„Ù…Ù†ØµØ±Ù", state="disabled")
        shortcuts_menu.add_command(label="F3  - Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", state="disabled")
        shortcuts_menu.add_command(label="F4  - Ø§Ù„ÙÙˆØ§ØªÙŠØ±", state="disabled")
        shortcuts_menu.add_command(label="F5  - ØªØ­Ø¯ÙŠØ«", state="disabled")
        shortcuts_menu.add_command(label="F9  - Ù„ÙˆØ­Ø© Ø§Ù„ØªØ­ÙƒÙ…", state="disabled")
        shortcuts_menu.add_command(label="F10 - Ø§Ù„Ø¥Ø­ØµØ§Ø¦ÙŠØ§Øª", state="disabled")
        shortcuts_menu.add_command(label="F11 - ØªØ´Ø®ÙŠØµ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", state="disabled")
        system_menu.add_cascade(label="Ø§Ù„Ø§Ø®ØªØµØ§Ø±Ø§Øª", menu=shortcuts_menu)
        menubar.add_cascade(label="Ø§Ù„Ù†Ø¸Ø§Ù…", menu=system_menu)

        # ======== HEADER BAR (POS-like light + blue accents) ========
        _HBG = _UI["SURFACE"]
        _HBG2 = _UI["SURFACE2"]

        header = tk.Frame(self, bg=_HBG, height=56)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        # Brand area (right for RTL)
        brand = tk.Frame(header, bg=_HBG)
        brand.pack(side=tk.RIGHT, padx=16)
        tk.Label(brand, text="Ø¥Ø¯Ø§Ø±Ø© Ø§Ù„Ù…Ø®Ø§Ø²Ù†", bg=_HBG, fg=_UI["TEXT"],
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
            ("Ù„ÙˆØ­Ø© Ø§Ù„ØªØ­ÙƒÙ…", 0),
            ("Ø§Ù„Ù…Ù†ØµØ±Ù", 1),
            ("Ø§Ù„ÙˆØ§Ø±Ø¯", 2),
            ("Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´", 3),
            ("ØªØ´Ø®ÙŠØµ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", 4),
            ("Ø§Ù„Ø¥Ø­ØµØ§Ø¦ÙŠØ§Øª", 5),
        ]
        for text, idx in nav_items:
            b = tk.Button(nav, text=f"  {text}  ", bg=_HBG2, fg=_UI["TEXT_SEC"],
                          font=("Segoe UI", 9, "bold"), bd=0, padx=12, pady=6,
                          cursor="hand2", activebackground=_UI["BRAND_H"],
                          activeforeground="#FFFFFF",
                          command=lambda i=idx: self._select_nav(i))
            b.pack(side=tk.RIGHT, padx=3)
            _add_hover(b, _UI["BRAND_H"], _HBG2, "#FFFFFF", _UI["TEXT_SEC"])
            self._nav_btns.append(b)

        # Utility buttons (left side for RTL)
        util = tk.Frame(header, bg=_HBG)
        util.pack(side=tk.LEFT, padx=12)

        self._header_clock_var = tk.StringVar(value="")
        tk.Label(util, textvariable=self._header_clock_var, bg=_HBG, fg=_UI["TEXT_DIM"],
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=8)

        _ghost_hdr = {"bg": _HBG, "fg": _UI["TEXT_SEC"], "bd": 0, "font": ("Segoe UI", 9, "bold"),
                      "padx": 8, "pady": 3, "cursor": "hand2",
                      "activebackground": _UI["BRAND_L"], "activeforeground": _UI["SEL_FG"]}
        for txt, cmd, tip in [
            ("Ø§Ù„Ù…Ø®Ø²ÙˆÙ†", self._open_inventory, "F3 / Ctrl+I"),
            ("Ø§Ù„ÙÙˆØ§ØªÙŠØ±", self._open_bills_history, "F4 / Ctrl+P"),
            ("Ø§Ù„Ø­Ø±ÙƒØ§Øª", self._open_movements, "Ø³Ø¬Ù„ Ø§Ù„Ø­Ø±ÙƒØ§Øª"),
            ("Ø§Ù„Ù…Ø¯ÙŠØ±", self._open_admin, "Ø¥Ø¹Ø¯Ø§Ø¯Ø§Øª Ø§Ù„Ù…Ø¯ÙŠØ±"),
            ("ØªØ­ÙˆÙŠÙ„", self._open_transfer, "ØªØ­ÙˆÙŠÙ„ Ø¨ÙŠÙ† Ø§Ù„Ù…Ø®Ø§Ø²Ù†"),
            ("Ø¬Ø±Ø¯", self._open_audit, "Ø¬Ø±Ø¯ Ø§Ù„Ù…Ø®Ø²ÙˆÙ†"),
        ]:
            b = tk.Button(util, text=txt, command=cmd, **_ghost_hdr)
            b.pack(side=tk.LEFT, padx=2)
            _add_hover(b, _UI["BRAND_L"], _HBG, _UI["SEL_FG"], _UI["TEXT_SEC"])
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
        self.notebook.add(self._dashboard, text="  Ù„ÙˆØ­Ø© Ø§Ù„ØªØ­ÙƒÙ…  (F9)  ")

        self._outcome = OutcomeFrame(self.notebook, self.db)
        self.notebook.add(self._outcome, text="  Ø§Ù„Ù…Ù†ØµØ±Ù  (F2)  ")

        self._income = IncomeFrame(self.notebook, self.db)
        self.notebook.add(self._income, text="  Ø§Ù„ÙˆØ§Ø±Ø¯  (F1)  ")

        self._fabric_calculation = FabricCalculationFrame(self.notebook, self.db)
        self.notebook.add(self._fabric_calculation, text="  Ø­Ø³Ø§Ø¨ Ø§Ù„Ù‚Ù…Ø§Ø´  ")

        self._sync_diagnostics = SyncDiagnosticsFrame(self.notebook, self.db, app=self)
        self.notebook.add(self._sync_diagnostics, text="  ØªØ´Ø®ÙŠØµ Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©  (F11)  ")

        self._statistics = StatisticsFrame(self.notebook, self.db)
        self.notebook.add(self._statistics, text="  Ø§Ù„Ø¥Ø­ØµØ§Ø¦ÙŠØ§Øª  (F10)  ")

        self.notebook.select(1)
        self._highlight_nav(1)

        def _on_tab_change(event):
            try:
                idx = self.notebook.index(self.notebook.select())
                self._highlight_nav(idx)
                if idx == 0:
                    self._dashboard._refresh()
                elif idx == 4:
                    self._sync_diagnostics._refresh()
                elif idx == 5:
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

        self._status_msg = tk.StringVar(value="Ø¬Ø§Ù‡Ø²")
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
                btn.config(bg=_UI["SURFACE2"], fg=_UI["TEXT_SEC"])
                _add_hover(btn, _UI["BRAND_H"], _UI["SURFACE2"], "#FFFFFF", _UI["TEXT_SEC"])

    def _open_admin(self):
        AdminWindow(self, self.db)

    def _open_sync_dialog(self):
        try:
            import sync_ui
            sync_ui.open_sync_dialog(self, self.db.conn)
        except Exception as e:
            messagebox.showerror("Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ù†Ø§ÙØ°Ø© Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©:\n{e}", parent=self)

    def _open_branch_stock(self):
        try:
            BranchStockWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ù…Ø®Ø²ÙˆÙ† Ø§Ù„ÙØ±ÙˆØ¹", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ù†Ø§ÙØ°Ø© Ù…Ø®Ø²ÙˆÙ† Ø§Ù„ÙØ±ÙˆØ¹:\n{e}", parent=self)

    def _open_branch_inventory_queue(self):
        try:
            BranchInventoryQueueWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ù…Ù†ØªØ¬Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹ ØºÙŠØ± Ø§Ù„Ù…Ø¹Ø§Ù„Ø¬Ø©", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ù†Ø§ÙØ°Ø© Ø§Ù„Ù…Ù†ØªØ¬Ø§Øª ØºÙŠØ± Ø§Ù„Ù…Ø¹Ø§Ù„Ø¬Ø©:\n{e}", parent=self)

    def _open_sync_setup(self):
        try:
            import sync_ui
            sync_ui.open_sync_setup(self, self.db.conn)
        except Exception as e:
            messagebox.showerror("Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ø¥Ø¹Ø¯Ø§Ø¯Ø§Øª Ø§Ù„Ù…Ø²Ø§Ù…Ù†Ø©:\n{e}", parent=self)

    def _open_price_profiles(self):
        try:
            PriceProfileManagerWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ù†Ø§ÙØ°Ø© Ø¨Ø±ÙˆÙØ§ÙŠÙ„Ø§Øª Ø§Ù„Ø£Ø³Ø¹Ø§Ø±:\n{e}", parent=self)

    def _open_price_sync_audit(self):
        try:
            open_price_sync_audit_dialog(self, self.db)
        except Exception as e:
            messagebox.showerror("Ø³Ø¬Ù„ Ø£Ø³Ø¹Ø§Ø± Ø§Ù„ÙØ±ÙˆØ¹", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ø§Ù„Ø³Ø¬Ù„:\n{e}", parent=self)

    def _open_branch_bills_sync_log(self):
        try:
            BranchBillsSyncLogWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ø³Ø¬Ù„ Ø§Ù„ÙØ±ÙˆØ¹", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ø§Ù„Ù†Ø§ÙØ°Ø©:\n{e}", parent=self)

    def _open_pos_reservations_mirror(self):
        try:
            PosReservationsMirrorWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ø­Ø¬ÙˆØ²Ø§Øª Ø§Ù„ÙØ±ÙˆØ¹", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ø§Ù„Ù†Ø§ÙØ°Ø©:\n{e}", parent=self)

    def _open_pos_financial_by_day(self):
        try:
            PosBranchFinancialWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ø§Ù„ØªØ¯ÙÙ‚Ø§Øª Ø§Ù„Ù…Ø§Ù„ÙŠØ©", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ø§Ù„Ù†Ø§ÙØ°Ø©:\n{e}", parent=self)

    def _open_branch_cycle_summary(self):
        try:
            BranchCycleSummaryWindow(self, self.db)
        except Exception as e:
            messagebox.showerror("Ù…Ù„Ø®Øµ Ø¯ÙˆØ±Ø© Ø§Ù„ÙØ±ÙˆØ¹", f"ØªØ¹Ø°Ù‘Ø± ÙØªØ­ Ø§Ù„Ù†Ø§ÙØ°Ø©:\n{e}", parent=self)

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
