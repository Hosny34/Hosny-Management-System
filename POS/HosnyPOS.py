# filepath: app/warehouse_manager_nosqlite_excel_billing.py
# Python 3.7.3+

import json
import os
import hashlib
import sys, subprocess
import sqlite3
import time
import tempfile
import webbrowser
import re
import uuid
import unicodedata
import threading
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from dataclasses import dataclass
from datetime import datetime, date, timedelta   # +date for calendar
import calendar                       # ADD
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

try:
    from pos_version import APP_VERSION
except Exception:
    APP_VERSION = "1.6"


def _safe_log_component(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"[^A-Za-z0-9_.-]+", "-", text).strip("-_.")
    return text[:40]


def _pos_log_name_from_device(device_name: Any) -> str:
    component = _safe_log_component(device_name)
    if not component:
        return "HosnyPOS"
    short_name = re.sub(r"^POS[-_]+", "", component, flags=re.IGNORECASE) or component
    return f"HosnyPOS-{short_name}"


def _runtime_base_dir() -> str:
    if getattr(sys, "frozen", False):
        exe_path = os.path.abspath(getattr(sys, "executable", ""))
        if exe_path:
            return os.path.dirname(exe_path)
    return os.path.dirname(os.path.abspath(__file__))


def _initial_log_name() -> str:
    for env_name in ("HOSNY_POS_LOG_NAME", "HOSNY_DEVICE_NAME", "HOSNY_POS_DEVICE"):
        env_value = os.environ.get(env_name)
        if _safe_log_component(env_value):
            return _pos_log_name_from_device(env_value)

    seen: Set[str] = set()
    candidates = [
        os.path.join(_runtime_base_dir(), "warehouse_data.sqlite3"),
    ]
    for db_path in candidates:
        db_path = os.path.abspath(db_path)
        if db_path in seen or not os.path.exists(db_path):
            continue
        seen.add(db_path)
        try:
            conn = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True, timeout=1.0)
            try:
                row = conn.execute(
                    "SELECT device_name FROM device_identity WHERE id = 1"
                ).fetchone()
            finally:
                conn.close()
            if row and _safe_log_component(row[0]):
                return _pos_log_name_from_device(row[0])
        except Exception:
            pass
    return "HosnyPOS"


try:
    import logging_setup
    logging_setup.install_crash_logging(_initial_log_name())
except Exception:
    logging_setup = None  # type: ignore


APP_TITLE = "إدارة المخازن والمبيعات"

if "--smoke-import" in sys.argv:
    print("HosnyPOS smoke import ok")
    raise SystemExit(0)

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


def _summarize_sync_payload_for_ui(event_type: str, payload: Dict[str, Any]) -> str:
    try:
        parts = []
        for key in ("bill_id", "reservation_id", "shipment_uuid", "target_device", "source_device", "customer", "total", "payment_method"):
            value = (payload or {}).get(key)
            if value not in (None, ""):
                parts.append(f"{key}={value}")
        for list_key in ("items", "lines", "reservations", "return_lines", "take_lines"):
            value = (payload or {}).get(list_key)
            if isinstance(value, list):
                parts.append(f"{list_key}={len(value)}")
        return ", ".join(parts) or str(event_type or "")
    except Exception:
        return str(event_type or "")


def parse_int_text(value: Any, default: int = 0) -> int:
    text = _strip_digit_marks(value).strip()
    try:
        return int(text)
    except Exception:
        nums = re.findall(r"-?\d+", text)
        return int(nums[0]) if nums else default


PREFERRED_ITEM_TYPE_ORDER = {
    "تيشيرت صيفي": 0,
    "تيشيرت خريفي": 1,
    "تيشيرت شتوي": 2,
    "شروال": 3,
    "جاكيت": 4,
    "تيشيرت رياضي": 5,
    "تيشيرت رياضي كم": 6,
    "تيشيرت رياضي شتوي": 7,
    "شروال رياضي": 8,
    "جاكيت رياضي": 9,
}


def _normalize_item_order_label(value: Any) -> str:
    text = unicodedata.normalize("NFKC", str(value or "").strip())
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return " ".join(text.split())


def item_type_sort_key(item_type: Any) -> Tuple[Any, ...]:
    item_clean = _normalize_item_order_label(item_type)
    return (
        PREFERRED_ITEM_TYPE_ORDER.get(item_clean, len(PREFERRED_ITEM_TYPE_ORDER)),
        item_clean.casefold(),
    )


def sort_item_type_values(values: Sequence[Any]) -> List[str]:
    return sorted(
        [str(v or "").strip() for v in values if str(v or "").strip()],
        key=item_type_sort_key,
    )


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

APP_BASE_DIR = _runtime_base_dir()
DB_PATH = os.path.join(APP_BASE_DIR, "warehouse_data.sqlite3")
LEGACY_JSON_PATH = os.path.join(APP_BASE_DIR, "warehouse_data.json")

try:
    if logging_setup is not None:  # type: ignore[name-defined]
        logging_setup.configure_context(app="POS", version=APP_VERSION, db_path=DB_PATH)  # type: ignore[union-attr]
except Exception:
    pass

ADMIN_PASSWORD_PLAIN = "112233"
ADMIN_PASSWORD_HASH_PREFIX = "sha256$"
WAREHOUSE_DEVICE_NAME = "WAREHOUSE"
WAREHOUSE_RETURN_LABEL = "الى المصنع"
# Bill rows for stock sent to warehouse review — not retail sales (no POS income / dashboard sales).
WAREHOUSE_RETURN_BILL_TYPE = "WAREHOUSE_RETURN"
BRANCH_TRANSFER_BILL_TYPE = "BRANCH_TRANSFER_REQUEST"
PAYMENT_METHOD_CASH = "CASH"
PAYMENT_METHOD_VISA = "VISA"
BRANCH_TARGET_PREFIX = "فرع: "
POS_CASHIER_LOCKDOWN = True
POS_DISCOUNT_PRESETS = (5, 10, 15)
POS_MANAGER_FEATURE_DEFAULTS = {
    "allow_inventory_window": not POS_CASHIER_LOCKDOWN,
    "allow_manual_incoming": not POS_CASHIER_LOCKDOWN,
    "allow_bulk_price": not POS_CASHIER_LOCKDOWN,
    "allow_excel_import": not POS_CASHIER_LOCKDOWN,
    "allow_manual_adjustment": not POS_CASHIER_LOCKDOWN,
    "allow_reset_counts": not POS_CASHIER_LOCKDOWN,
    "allow_inventory_delete": not POS_CASHIER_LOCKDOWN,
    "allow_inventory_price_edit": not POS_CASHIER_LOCKDOWN,
    "allow_inventory_specs_edit": not POS_CASHIER_LOCKDOWN,
    "allow_size_profile_edit": not POS_CASHIER_LOCKDOWN,
    "allow_stock_audit": True,
}
DEFAULT_BRANCH_POS_NAMES = [
    "POS-ZAY",
    "POS-OCT",
    "POS-OBO",
    "POS-GESR",
    "POS-BAH",
    "POS-CEN",
    "POS-TEST",
]
BRANCH_UI_NAME_BY_DEVICE = {
    "POS-ZAY": "فرع زايد",
    "POS-OCT": "فرع اكتوبر",
    "POS-BAH": "فرع بهتيم",
    "POS-CEN": "فرع السنتر",
    "POS-OBO": "فرع العبور",
    "POS-GESR": "فرع جسر السويس",
    "POS-TEST": "فرع الاختبار",
}
BRANCH_DEVICE_BY_UI_NAME = {v: k for k, v in BRANCH_UI_NAME_BY_DEVICE.items()}
BRANCH_GENERAL_SCHOOL_TARGET_BY_DEVICE = {
    "POS-ZAY": "عام اكتوبر",
    "POS-OCT": "عام اكتوبر",
    "POS-BAH": "عام شبرا",
    "POS-CEN": "عام شبرا",
}
GENERAL_SHARED_SCHOOL_NAME = "عام"
RECEIPT_SUPPORT_CALL_SETTING = "receipt_support_call"
RECEIPT_SUPPORT_WHATSAPP_SETTING = "receipt_support_whatsapp"
RECEIPT_FONT_FILE = os.path.join("Fonts", "V100009_.TTF")
RECEIPT_FONT_NAME = "HosnyReceiptFont"
RECEIPT_FONT_STACK = '"%s", Tahoma, Arial, "Segoe UI", sans-serif' % RECEIPT_FONT_NAME

ALLOWED_NUMERIC_RANGES = {
    (0, 24): [str(i) for i in range(0, 26, 2)],
    (0, 16): [str(i) for i in range(0, 18, 2)],
    (6, 22): [str(i) for i in range(6, 24, 2)],
    (14, 28): [str(i) for i in range(14, 30, 2)],
    (18, 30): [str(i) for i in range(18, 32, 2)],
    (0, 9):  [str(i) for i in range(0, 10, 1)],
    (30, 62): [str(i) for i in range(30, 63, 2)],

}
NUMERIC_RANGE_LABELS = [
    f"{a} → {b}" for (a, b) in ALLOWED_NUMERIC_RANGES.keys()
]

ALPHA_SIZES = ["S", "M", "L", "XL", "2XL", "3XL", "4XL", "5XL"]


def parse_numeric_range_label(label: Any) -> Tuple[Optional[int], Optional[int]]:
    text = western_digits(label).strip()
    if not text:
        return None, None
    nums = re.findall(r"\d+", text)
    if len(nums) < 2:
        raise ValueError(f"Invalid size range: {text}")
    return int(nums[0]), int(nums[1])


def merged_numeric_size_labels_from_profile(
    r1s: Optional[int],
    r1e: Optional[int],
    r2s: Optional[int],
    r2e: Optional[int],
) -> List[str]:
    """Union of predefined numeric columns for range 1 and range 2 (no duplicate sizes)."""
    chunks: List[List[str]] = []
    if r1s is not None and r1e is not None:
        labs = ALLOWED_NUMERIC_RANGES.get((r1s, r1e))
        if labs is None:
            labs = [str(i) for i in range(int(r1s), int(r1e) + 1)]
        if labs:
            chunks.append(list(labs))
    if r2s is not None and r2e is not None:
        labs = ALLOWED_NUMERIC_RANGES.get((r2s, r2e))
        if labs is None:
            labs = [str(i) for i in range(int(r2s), int(r2e) + 1)]
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


def fmt_local_ts(value: Any, empty: str = "—") -> str:
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


def _cashier_lockdown_message(action: str = "") -> str:
    base = "هذا الإجراء متوقف في وضع الكاشير. المخزن فقط هو المخوّل بتغيير المخزون أو الأسعار."
    if action:
        return f"{action}\n\n{base}"
    return base


def _payment_method_label(method: Optional[str]) -> str:
    code = str(method or PAYMENT_METHOD_CASH).strip().upper()
    if code == PAYMENT_METHOD_VISA:
        return "فيزا"
    return "كاش"


def _normalize_payment_method(method: Optional[str]) -> str:
    code = str(method or PAYMENT_METHOD_CASH).strip().upper()
    if code not in (PAYMENT_METHOD_CASH, PAYMENT_METHOD_VISA):
        raise ValueError("طريقة الدفع غير مدعومة.")
    return code

def _round_up_to_step(value: float, step: int = 5) -> float:
    import math
    step_f = float(step or 1)
    if value <= 0:
        return 0.0
    return float(math.ceil(float(value) / step_f) * step_f)

def _allocate_reservation_down_payments(
    totals: Sequence[float],
    paid_amount: float,
    *,
    round_step: int = 5,
) -> List[float]:
    paid = max(0.0, float(paid_amount or 0.0))
    clean_totals = [max(0.0, float(t or 0.0)) for t in totals]
    if not clean_totals:
        return []

    total_due = sum(clean_totals)
    if paid - total_due > 1e-6:
        raise ValueError("العربون لا يمكن أن يزيد عن إجمالي الحجز.")
    if abs(paid - total_due) <= 0.01:
        return [round(t, 2) for t in clean_totals]

    allocations: List[float] = [0.0 for _ in clean_totals]
    remaining = paid
    open_indexes = {idx for idx, total in enumerate(clean_totals) if total > 1e-9}
    while remaining > 0.01 and open_indexes:
        share = _round_up_to_step(remaining / len(open_indexes), round_step)
        progressed = False
        for idx in list(sorted(open_indexes)):
            capacity = max(0.0, clean_totals[idx] - allocations[idx])
            if capacity <= 0.01:
                open_indexes.discard(idx)
                continue
            alloc = min(share, capacity, remaining)
            allocations[idx] = round(allocations[idx] + alloc, 2)
            remaining = round(remaining - alloc, 2)
            progressed = True
            if clean_totals[idx] - allocations[idx] <= 0.01:
                open_indexes.discard(idx)
            if remaining <= 0.01:
                break
        if not progressed:
            break

    if remaining > 0.01:
        for idx in sorted(open_indexes):
            capacity = max(0.0, clean_totals[idx] - allocations[idx])
            alloc = min(capacity, remaining)
            allocations[idx] = round(allocations[idx] + alloc, 2)
            remaining = round(remaining - alloc, 2)
            if remaining <= 0.01:
                break
    return [min(round(a, 2), round(clean_totals[idx], 2)) for idx, a in enumerate(allocations)]


def _reservation_bill_id_from_rows(rows: Sequence[Dict[str, Any]]) -> str:
    ids: List[int] = []
    for row in rows or []:
        try:
            rid = int(row.get("id") or 0)
        except Exception:
            rid = 0
        if rid > 0:
            ids.append(rid)
    return str(min(ids)) if ids else ""


RESERVATION_STATUS_PENDING = "معلق"
RESERVATION_STATUS_DELIVERED = "تم التسليم"
RESERVATION_STATUS_CANCELLED = "ملغي"


def _is_reservation_delivered(status: Any) -> bool:
    return str(status or "").strip() == RESERVATION_STATUS_DELIVERED


def _is_reservation_cancelled(status: Any) -> bool:
    return str(status or "").strip() == RESERVATION_STATUS_CANCELLED


def _is_reservation_active(status: Any) -> bool:
    return not (_is_reservation_delivered(status) or _is_reservation_cancelled(status))


def _normalize_customer_phone(value: Any) -> str:
    return _strip_digit_marks(value).strip()


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


def _feature_restricted_message(action: str = "") -> str:
    base = "هذه الميزة مقيدة حالياً من تبويب صلاحيات المدير."
    if action:
        return f"{action}\n\n{base}"
    return base


def _extract_warehouse_target(value: Any) -> Optional[str]:
    raw = str(value or "").strip()
    if not raw:
        return None
    if raw == WAREHOUSE_RETURN_LABEL:
        return WAREHOUSE_DEVICE_NAME
    if raw.upper() == WAREHOUSE_DEVICE_NAME:
        return WAREHOUSE_DEVICE_NAME
    if raw.startswith(f"{WAREHOUSE_RETURN_LABEL}:"):
        name = raw.split(":", 1)[1].strip()
        if name.upper() == WAREHOUSE_DEVICE_NAME:
            return WAREHOUSE_DEVICE_NAME
    return None


def _extract_branch_target(value: Any) -> Optional[str]:
    raw = str(value or "").strip()
    if not raw:
        return None
    if raw.startswith(BRANCH_TARGET_PREFIX):
        raw = raw[len(BRANCH_TARGET_PREFIX):].strip()
    if raw in DEFAULT_BRANCH_POS_NAMES:
        return raw
    if raw in BRANCH_DEVICE_BY_UI_NAME:
        return BRANCH_DEVICE_BY_UI_NAME[raw]
    return None


def _branch_display_name(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    return BRANCH_UI_NAME_BY_DEVICE.get(raw, raw)


def _receipt_pos_name(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return "POS"
    return _branch_display_name(raw)


def _lookup_receipt_pos_name(conn: sqlite3.Connection) -> str:
    try:
        row = conn.execute(
            "SELECT device_name FROM device_identity WHERE id = 1"
        ).fetchone()
        return _receipt_pos_name(row[0] if row else "")
    except Exception:
        return "POS"


def _lookup_receipt_support(conn: sqlite3.Connection, key: str) -> str:
    try:
        row = conn.execute(
            "SELECT value FROM app_settings WHERE key = ?",
            (str(key),),
        ).fetchone()
        return western_digits((row[0] if row else "") or "").strip()
    except Exception:
        return ""


def _lookup_receipt_branding(conn: sqlite3.Connection) -> Tuple[str, str, str]:
    return (
        _lookup_receipt_pos_name(conn),
        _lookup_receipt_support(conn, RECEIPT_SUPPORT_CALL_SETTING),
        _lookup_receipt_support(conn, RECEIPT_SUPPORT_WHATSAPP_SETTING),
    )


def _receipt_bill_type_label(value: Any) -> str:
    code = str(value or "SALE").strip().upper()
    return {
        "SALE": "بيع",
        "RETURN": "مرتجع",
        "EXCHANGE": "استبدال",
        WAREHOUSE_RETURN_BILL_TYPE: WAREHOUSE_RETURN_LABEL,
        BRANCH_TRANSFER_BILL_TYPE: "تحويل فرع",
        "RESERVATION": "حجز",
    }.get(code, code or "بيع")


def _app_resource_path(relative_path: str) -> str:
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_dir, *relative_path.replace("\\", "/").split("/"))


def _receipt_font_face_css() -> str:
    font_path = _app_resource_path(RECEIPT_FONT_FILE)
    if not os.path.exists(font_path):
        return ""
    return (
        '@font-face { font-family: "%s"; src: url("%s") format("truetype"); '
        "font-weight: 400 900; font-style: normal; }\n"
    ) % (RECEIPT_FONT_NAME, _file_url(font_path))


@dataclass
class StockRow:
    id: int
    item_type: str
    school: str
    color: str
    size: str
    unit_price: float
    count: int


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

def _ensure_html_document_encoding(html: str) -> str:
    if "<meta charset" in html.lower():
        return html
    meta = '<meta charset="utf-8">'
    lower = html.lower()
    head_idx = lower.find("<head>")
    if head_idx >= 0:
        insert_at = head_idx + len("<head>")
        return html[:insert_at] + meta + html[insert_at:]
    return meta + html

def _write_html_file(path: str, html: str) -> None:
    html = _ensure_html_document_encoding(str(html or ""))
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(html)

def _print_html_auto_legacy_ie(path: str, copies: int = 1, parent: Optional[tk.Widget] = None) -> None:
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
                while True:
                    time.sleep(0.1)
                    try:
                        if not ie.Busy and int(ie.ReadyState) == 4:
                            doc = getattr(ie, "Document", None)
                            if doc is not None and str(getattr(doc, "readyState", "")).lower() == "complete":
                                break
                    except Exception:
                        break
                time.sleep(0.2)
                ie.ExecWB(6, 2)
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
            with open(path, "r", encoding="utf-8-sig") as f:
                html = f.read()
            edge_meta = '<meta http-equiv="X-UA-Compatible" content="IE=edge" />'
            if "</head>" in html:
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

            _write_html_file(auto_path, html)
        except Exception:
            auto_path = path

        for _ in range(copies):
            webbrowser.open_new_tab(_file_url(auto_path))
            time.sleep(0.05)

        if parent is not None:
            messagebox.showinfo("الطباعة", f"تم فتح الفاتورة للطباعة تلقائيًا.\nعدد النسخ: {copies}.", parent=parent)
    except Exception as ex:
        if parent is not None:
            messagebox.showerror("فشل الطباعة", f"{ex}", parent=parent)

def _html_receipt_to_text(path: str) -> str:
    """Best-effort plain-text receipt for old Windows print engines."""
    import re
    from html import unescape

    with open(path, "r", encoding="utf-8-sig") as f:
        html = f.read()

    text = re.sub(r"(?is)<script.*?</script>", "", html)
    text = re.sub(r"(?is)<style.*?</style>", "", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"(?i)</(div|p|h1|h2|h3|tr|table)>", "\n", text)
    text = re.sub(r"(?i)</t[dh]>", "    ", text)
    text = re.sub(r"(?is)<[^>]+>", "", text)
    text = unescape(text)

    lines = []
    for line in text.splitlines():
        clean = " ".join(line.split())
        if clean:
            lines.append(clean)
    return western_digits("\n".join(lines).strip()) + "\n"


def _rtf_escape(value: Any) -> str:
    text = western_digits(str(value or ""))
    out = []
    for ch in text:
        if ch == "\\":
            out.append(r"\\")
        elif ch == "{":
            out.append(r"\{")
        elif ch == "}":
            out.append(r"\}")
        elif ch == "\n":
            out.append(r"\line ")
        elif ch == "\r":
            continue
        else:
            code = ord(ch)
            if 32 <= code <= 126:
                out.append(ch)
            else:
                if code > 32767:
                    code -= 65536
                out.append(r"\u%d?" % code)
    return "".join(out)


def _rtf_table_row(cells: List[str], widths: List[int], bold: bool = False, font_size: int = 22) -> str:
    parts = [r"\trowd\trgaph60\trleft0"]
    pos = 0
    for width in widths:
        pos += int(width)
        parts.append(
            r"\clbrdrt\brdrs\brdrw10\clbrdrl\brdrs\brdrw10\clbrdrb\brdrs\brdrw10\clbrdrr\brdrs\brdrw10\cellx%d" % pos
        )
    cell_prefix = r"\pard\intbl\rtlpar\qr "
    if bold:
        cell_prefix += r"\b "
    cell_prefix += r"\fs%d " % int(font_size)
    for cell in cells:
        parts.append(cell_prefix + _rtf_escape(cell) + r"\cell")
    parts.append(r"\row")
    return "\n".join(parts)


def save_bill_as_rtf(
    path: str,
    bill: Dict[str, Any],
    items: List[Dict[str, Any]],
    pos_name: str = "",
    support_call: str = "",
    support_whatsapp: str = "",
    extra_summary_rows: Optional[Sequence[Tuple[str, Any]]] = None,
) -> None:
    customer = western_digits(bill.get("customer") or "").strip()
    customer_phone = western_digits(bill.get("customer_phone") or "").strip()
    pos_name = _receipt_pos_name(pos_name)
    created_at = western_digits(bill.get("created_at") or "")
    extra_rows = list(extra_summary_rows or [])

    rows = []
    widths = [2650, 500, 650, 650]
    rows.append(_rtf_table_row(["الصنف", "العدد", "السعر", "الإجمالي"], widths, bold=True, font_size=22))
    for line in items:
        line_total = line.get("line_total")
        if line_total is None:
            try:
                line_total = float(line.get("unit_price", 0)) * int(line.get("qty", 0))
            except Exception:
                line_total = 0
        item_name = "%s - %s\n%s / %s" % (
            western_digits(line.get("item_type") or ""),
            western_digits(line.get("school") or ""),
            western_digits(line.get("color") or ""),
            western_digits(line.get("size") or ""),
        )
        rows.append(
            _rtf_table_row(
                [
                    item_name,
                    str(int(line.get("qty") or 0)),
                    format_money(float(line.get("unit_price") or 0)),
                    format_money(float(line_total or 0)),
                ],
                widths,
                bold=False,
                font_size=21,
            )
        )

    support_lines = []
    if support_call:
        support_lines.append(r"\pard\rtlpar\qc\fs22 " + _rtf_escape("للاتصال: %s" % support_call) + r"\par")
    if support_whatsapp:
        support_lines.append(r"\pard\rtlpar\qc\fs22 " + _rtf_escape("واتساب: %s" % support_whatsapp) + r"\par")

    # WordPad prints RTF as paged media. A fixed large height makes thermal
    # printers feed a near-A4 blank tail after every receipt, so size the page
    # to the receipt content instead.
    paper_height = 3300
    paper_height += max(1, len(items)) * 520
    paper_height += len(extra_rows) * 330
    if customer:
        paper_height += 260
    if customer_phone:
        paper_height += 260
    paper_height += len(support_lines) * 260
    paper_height = max(4200, min(32000, int(paper_height)))

    content = [
        r"{\rtf1\ansi\deff0\uc1",
        r"{\fonttbl{\f0 Tahoma;}}",
        r"\paperw4535\paperh%d\margl120\margr120\margt80\margb80" % paper_height,
        r"\pard\rtlpar\qc\b\fs30 " + _rtf_escape(pos_name) + r"\par",
        r"\pard\rtlpar\qc\b\fs28 " + _rtf_escape("فاتورة #%s" % bill["id"]) + r"\par",
        r"\pard\rtlpar\qc\fs22 " + _rtf_escape(created_at) + r"\par",
    ]
    if customer:
        content.append(r"\pard\rtlpar\qr\fs22 " + _rtf_escape("العميل: %s" % customer) + r"\par")
    if customer_phone:
        content.append(r"\pard\rtlpar\qr\fs22 " + _rtf_escape("رقم العميل: %s" % customer_phone) + r"\par")
    content.extend(
        [
            r"\pard\rtlpar\qr\fs16 ------------------------------------------------\par",
            "\n".join(rows),
            r"\pard\rtlpar\qr\fs16 ------------------------------------------------\par",
            r"\pard\rtlpar\qc\b\fs28 " + _rtf_escape("الإجمالي: %s" % format_money(float(bill["total"]))) + r"\par",
        ]
    )
    for label, value in extra_rows:
        content.append(
            r"\pard\rtlpar\qc\b\fs24 "
            + _rtf_escape("%s: %s" % (label, format_money(float(value or 0.0))))
            + r"\par"
        )
    content.append(r"\pard\rtlpar\qc\fs22 " + _rtf_escape("شكرا لتعاملكم معنا") + r"\par")
    content.extend(support_lines)
    content.append("}")
    with open(path, "w", encoding="ascii", errors="ignore") as f:
        f.write("\n".join(content))

def _print_text_with_notepad(path: str, copies: int) -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        for _ in range(max(1, int(copies))):
            proc = subprocess.Popen(["notepad.exe", "/p", path])
            try:
                proc.wait(timeout=20)
            except Exception:
                pass
        return True
    except Exception:
        return False


def _print_rtf_with_wordpad(path: str, copies: int) -> bool:
    if not sys.platform.startswith("win"):
        return False
    win_dir = os.environ.get("SystemRoot", r"C:\Windows")
    candidates = [
        os.path.join(win_dir, "write.exe"),
        os.path.join(win_dir, "System32", "write.exe"),
        os.path.join(win_dir, "System32", "wordpad.exe"),
        "write.exe",
        "wordpad.exe",
    ]
    for exe_path in candidates:
        try:
            for _ in range(max(1, int(copies))):
                proc = subprocess.Popen([exe_path, "/p", path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                try:
                    proc.wait(timeout=25)
                except Exception:
                    pass
            return True
        except Exception:
            continue
    return False

def _make_autoprint_html(path: str) -> str:
    auto_path = path[:-5] + "_autoprint.html" if path.lower().endswith(".html") else path + "_autoprint.html"
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            html = f.read()
        script = """
<script>
  window.onload=function(){
    setTimeout(function(){
      try{ window.focus(); window.print(); }catch(e){}
    }, 250);
  };
</script>
"""
        if "</head>" in html:
            if "X-UA-Compatible" not in html:
                html = html.replace("</head>", '<meta http-equiv="X-UA-Compatible" content="IE=edge" /></head>')
            if "window.print()" not in html:
                html = html.replace("</head>", script + "</head>")
        elif "window.print()" not in html:
            html += script
        _write_html_file(auto_path, html)
        return auto_path
    except Exception:
        return path

def _print_html_with_ie_silent(path: str, copies: int) -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32com.client  # type: ignore
    except Exception:
        return False
    ie = None
    try:
        ie = win32com.client.Dispatch("InternetExplorer.Application")
        ie.Visible = False
        url = _file_url(path)
        for _ in range(max(1, int(copies))):
            ie.Navigate(url)
            deadline = time.time() + 20
            while time.time() < deadline:
                time.sleep(0.1)
                try:
                    doc = getattr(ie, "Document", None)
                    ready = str(getattr(doc, "readyState", "")).lower() if doc is not None else ""
                    if not ie.Busy and int(ie.ReadyState) == 4 and ready in ("", "complete"):
                        break
                except Exception:
                    break
            time.sleep(0.4)
            # OLECMDID_PRINT=6, OLECMDEXECOPT_DONTPROMPTUSER=2.
            ie.ExecWB(6, 2)
            time.sleep(0.7)
        return True
    except Exception:
        return False
    finally:
        try:
            if ie is not None:
                ie.Quit()
        except Exception:
            pass


def _print_html_with_shell(path: str, copies: int) -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import win32api  # type: ignore
        for _ in range(max(1, int(copies))):
            rc = win32api.ShellExecute(0, "print", path, None, os.path.dirname(path) or ".", 0)
            if int(rc) <= 32:
                return False
            time.sleep(0.25)
        return True
    except Exception:
        try:
            for _ in range(max(1, int(copies))):
                os.startfile(path, "print")  # type: ignore[attr-defined]
                time.sleep(0.25)
            return True
        except Exception:
            return False


def _find_windows_browser_exe() -> Optional[str]:
    """Prefer a modern browser for local receipt previews on older Windows.

    Windows 7 can still route file:// HTML through Internet Explorer even when
    Chrome is selected in Default Programs. Launching the browser executable
    directly avoids IE active-content warnings and preserves CSS receipt sizing
    much better than the legacy IE print dialog.
    """
    if not sys.platform.startswith("win"):
        return None

    candidates: List[str] = []
    env_names = (
        "HOSNY_RECEIPT_BROWSER",
        "CHROME_PATH",
        "EDGE_PATH",
        "FIREFOX_PATH",
    )
    for env_name in env_names:
        value = os.environ.get(env_name)
        if value:
            candidates.append(value)

    for base_env in ("ProgramFiles", "ProgramFiles(x86)", "LocalAppData"):
        base = os.environ.get(base_env)
        if not base:
            continue
        candidates.extend([
            os.path.join(base, "Google", "Chrome", "Application", "chrome.exe"),
            os.path.join(base, "Microsoft", "Edge", "Application", "msedge.exe"),
            os.path.join(base, "Mozilla Firefox", "firefox.exe"),
        ])

    candidates.extend([
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Users\%USERNAME%\AppData\Local\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Mozilla Firefox\firefox.exe",
        r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe",
    ])

    seen: Set[str] = set()
    for raw in candidates:
        path = os.path.expandvars(str(raw or "").strip().strip('"'))
        if not path:
            continue
        key = os.path.normcase(os.path.abspath(path))
        if key in seen:
            continue
        seen.add(key)
        if os.path.exists(path):
            return path
    return None


def _open_receipt_preview(path: str) -> bool:
    url = _file_url(path)
    browser = _find_windows_browser_exe()
    if browser:
        try:
            subprocess.Popen([browser, "--new-window", url], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            pass

    try:
        if webbrowser.open_new_tab(url):
            return True
    except Exception:
        pass

    if sys.platform.startswith("win"):
        try:
            os.startfile(path)  # type: ignore[attr-defined]
            return True
        except Exception:
            return False
    return False


def _print_html_auto(path: str, copies: int = 1, parent: Optional[tk.Widget] = None) -> None:
    copies = max(1, int(copies))

    try:
        auto_path = _make_autoprint_html(path)
        for _ in range(copies):
            opened = _open_receipt_preview(auto_path)
            if not opened:
                raise RuntimeError("Could not open the auto-print receipt file.")
    except Exception as ex:
        if parent is not None:
            messagebox.showerror("Print failed", f"{ex}", parent=parent)


def _print_receipt_direct_or_fallback(
    html_path: str,
    bill: Dict[str, Any],
    items: List[Dict[str, Any]],
    *,
    copies: int = 1,
    parent: Optional[tk.Widget] = None,
    pos_name: str = "",
    support_call: str = "",
    support_whatsapp: str = "",
    extra_summary_rows: Optional[Sequence[Tuple[str, Any]]] = None,
) -> None:
    copies = max(1, int(copies))
    receipt_id = (bill or {}).get("id")
    try:
        if logging_setup is not None:  # type: ignore[name-defined]
            logging_setup.log_event(  # type: ignore[union-attr]
                "print.receipt.direct.start",
                receipt_id=receipt_id,
                bill_type=(bill or {}).get("bill_type"),
                copies=copies,
                html_path=html_path,
                mode="html",
            )
    except Exception:
        pass

    try:
        _print_html_auto(html_path, copies=copies, parent=None)
        try:
            if logging_setup is not None:  # type: ignore[name-defined]
                logging_setup.log_event(  # type: ignore[union-attr]
                    "print.receipt.direct.done",
                    receipt_id=receipt_id,
                    copies=copies,
                    html_path=html_path,
                    mode="browser_autoprint",
                )
        except Exception:
            pass
        return
    except Exception:
        try:
            logging_setup.log_exception("print.receipt.direct.failed", receipt_id=receipt_id, copies=copies)  # type: ignore[union-attr]
        except Exception:
            pass
        pass

def save_bill_as_html(
    path: str,
    bill: Dict[str, Any],
    items: List[Dict[str, Any]],
    pos_name: str = "",
    support_call: str = "",
    support_whatsapp: str = "",
    shift_id: Any = "",
    user_name: str = "مدير",
    extra_summary_rows: Optional[List[Tuple[str, Any]]] = None,
) -> None:
    """Generate receipt-style HTML for thermal printers (80mm width)."""

    def _fmtf(x: Any) -> str:
        try:
            return f"{format_money(float(x))}"
        except Exception:
            return "0"

    def _receipt_date_time(value: Any) -> Tuple[str, str]:
        raw = western_digits(value or "")
        try:
            dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
            return dt.strftime("%d/%m/%Y"), dt.strftime("%I:%M%p")
        except Exception:
            parts = raw.replace("T", " ").split()
            return (parts[0] if parts else ""), (parts[1][:5] if len(parts) > 1 else "")

    rows_html = ""
    item_count = 0
    for ln in items:
        line_total = _fmtf(ln.get('line_total') or (float(ln.get('unit_price', 0)) * int(ln.get('qty', 0))))
        qty = int(ln.get('qty') or 0)
        item_count += qty
        rows_html += f"""<tr>
<td class="num">{line_total}</td>
<td class="num">{_fmtf(ln['unit_price'])}</td>
<td class="item">{_html(ln.get('item_type') or '')} - {_html(ln.get('school') or '')}<br>{_html(ln.get('color') or '')} - مقاس <span class="digits">{_html(ln.get('size') or '')}</span></td>
<td class="qty">{qty}</td>
</tr>
"""

    customer = _html(bill.get('customer') or '')
    customer_phone = _receipt_phone_html(bill.get('customer_phone') or '')
    pos_name = _html(_receipt_pos_name(pos_name))
    support_call = _receipt_phone_html(support_call)
    support_whatsapp = _receipt_phone_html(support_whatsapp)
    customer_line = f"<div>العميل: {customer}</div>" if customer else ""
    customer_phone_line = f"<div>رقم العميل: {customer_phone}</div>" if customer_phone else ""
    pos_line = f"<div class=\"branch\">{pos_name}</div>" if pos_name else ""
    support_call_line = f"<div>للاتصال: {support_call}</div>" if support_call else ""
    support_whatsapp_line = f"<div>واتساب: {support_whatsapp}</div>" if support_whatsapp else ""
    receipt_date, receipt_time = _receipt_date_time(bill.get("created_at"))
    bill_type_label = _html(_receipt_bill_type_label(bill.get("bill_type")))
    total_value = format_money(float(bill['total']))
    extra_summary_html = ""
    for label, value in (extra_summary_rows or []):
        extra_summary_html += (
            '<tr class="summary"><td class="num">%s</td><td class="label" colspan="3">%s</td></tr>\n'
            % (format_money(value), _html(str(label)))
        )

    html = f"""<!DOCTYPE html>
<html lang="ar" dir="ltr">
<head>
<meta charset="utf-8">
<title>فاتورة #{bill['id']}</title>
<style>
  {_receipt_font_face_css()}
  @page {{ size: 80mm auto; margin: 1mm; }}
  * {{ box-sizing: border-box; }}
  html, body {{
    width: 80mm;
    min-height: 0;
    margin: 0;
    padding: 0;
  }}
  body {{
    color: #000;
    font-family: {RECEIPT_FONT_STACK};
    font-size: 11px;
    font-weight: 800;
    margin: 0;
    padding: 0;
    width: 80mm;
    direction: ltr;
    -webkit-font-smoothing: none;
    print-color-adjust: exact;
  }}
  .receipt {{
    display: block;
    min-height: auto;
    height: auto;
    max-height: none;
    overflow: visible;
    padding: 1.5mm;
    position: relative;
    width: 78mm;
  }}
  table, tbody, tr {{
    break-inside: avoid;
    page-break-inside: avoid;
  }}
  @media screen {{
    body {{
      background: #f8fafc;
      padding: 6mm 0;
    }}
    .receipt {{
      background: #fff;
      box-shadow: 0 0 0 1px #d1d5db;
      margin: 0 auto;
    }}
  }}
  .receipt-stamp {{
    color: #000;
    direction: ltr;
    font-family: Tahoma, Arial, "Segoe UI", sans-serif;
    font-size: 10px;
    font-weight: 900;
    left: 1.5mm;
    line-height: 1.25;
    position: absolute;
    text-align: left;
    top: 1.5mm;
  }}
  .brand-row {{
    align-items: flex-start;
    direction: ltr;
    display: grid;
    gap: 5mm;
    grid-template-columns: 1fr 1fr;
    margin: 2mm 0 1mm;
  }}
  .ossni-logo {{
    align-items: center;
    background: #333;
    color: #fff;
    display: flex;
    font-size: 18px;
    font-weight: 900;
    height: 18mm;
    justify-content: center;
    letter-spacing: 1px;
    margin-top: 8mm;
  }}
  .brand-mark {{
    text-align: center;
  }}
  .brand-box {{
    border: 1.5px solid #111;
    color: #000;
    display: inline-block;
    font-family: {RECEIPT_FONT_STACK};
    font-size: 28px;
    font-weight: 900;
    line-height: 1;
    min-width: 31mm;
    padding: 3mm 2mm 2mm;
  }}
  .tagline {{
    font-size: 16px;
    color: #000;
    font-weight: 900;
    margin-top: 1mm;
  }}
  .meta {{
    border-collapse: collapse;
    direction: rtl;
    font-size: 11px;
    margin-top: 1.5mm;
    table-layout: fixed;
    width: 100%;
  }}
  .meta td {{
    border: 1.4px dashed #333;
    color: #000;
    font-weight: 900;
    line-height: 1.25;
    padding: 0.8mm 1mm;
    text-align: right;
    white-space: nowrap;
  }}
  .meta .label {{
    width: 30%;
  }}
  .meta .value {{
    text-align: center;
    width: 20%;
  }}
  .digits, .meta .value, .items .num, .items .qty {{
    font-family: Tahoma, Arial, "Segoe UI", sans-serif;
  }}
  .receipt-phone {{
    direction: ltr;
    display: inline-block;
    font-family: Tahoma, Arial, "Segoe UI", sans-serif;
    font-weight: 900;
    unicode-bidi: embed;
  }}
  .items {{
    border-collapse: collapse;
    direction: ltr;
    font-size: 11px;
    margin-top: 4mm;
    table-layout: fixed;
    width: 100%;
  }}
  .items th, .items td {{
    border: 1.3px solid #555;
    color: #000;
    padding: 0.9mm 0.8mm;
    overflow: hidden;
    vertical-align: middle;
  }}
  .items th {{
    font-size: 11px;
    font-weight: 900;
    text-align: center;
  }}
  .items .num {{
    font-size: 12px;
    font-weight: 900;
    text-align: center;
    width: 18%;
  }}
  .items .item {{
    direction: rtl;
    font-family: {RECEIPT_FONT_STACK};
    font-size: 9px;
    font-weight: 900;
    line-height: 1.25;
    max-width: 0;
    overflow-wrap: break-word;
    text-align: right;
    white-space: normal;
    width: 56%;
    word-break: break-word;
  }}
  .items .qty {{
    font-size: 12px;
    font-weight: 900;
    text-align: center;
    width: 12%;
  }}
  .summary td {{
    border-left: 0;
    border-right: 0;
  }}
  .summary .label {{
    direction: rtl;
    font-size: 14px;
    font-weight: 900;
    text-align: right;
  }}
  .notes {{
    border-bottom: 1.3px solid #555;
    border-top: 1.3px solid #555;
    color: #000;
    direction: rtl;
    font-family: {RECEIPT_FONT_STACK};
    font-size: 12px;
    font-weight: 900;
    margin-top: 1mm;
    padding: 1mm 0;
    text-align: center;
  }}
  .footer {{
    border-top: 1.3px solid #555;
    color: #000;
    direction: rtl;
    font-family: {RECEIPT_FONT_STACK};
    font-size: 11px;
    font-weight: 900;
    line-height: 1.35;
    margin-top: auto;
    padding-top: 2mm;
    text-align: center;
  }}
</style>
</head>
<body>
<div class="receipt">
  <div class="receipt-stamp">
    <div>{receipt_date}</div>
    <div>{receipt_time}</div>
  </div>
  <div class="brand-row">
    <div class="ossni-logo">حسنى</div>
    <div class="brand-mark">
      <div class="brand-box">حسنى</div>
      <div class="tagline">للزي المدرسي</div>
    </div>
  </div>
  <table class="meta">
    <colgroup>
      <col style="width:30%">
      <col style="width:20%">
      <col style="width:30%">
      <col style="width:20%">
    </colgroup>
    <tbody>
      <tr><td class="label">رقم الفاتورة</td><td class="value">{western_digits(bill['id'])}</td><td class="label">نوع الفاتورة</td><td class="value">{bill_type_label}</td></tr>
    </tbody>
  </table>
  <table class="items">
    <colgroup>
      <col style="width:18%">
      <col style="width:14%">
      <col style="width:56%">
      <col style="width:12%">
    </colgroup>
    <thead>
      <tr>
        <th>الإجمالي</th>
        <th>السعر</th>
        <th>الصنف</th>
        <th>العدد</th>
      </tr>
    </thead>
    <tbody>
      {rows_html}
      <tr class="summary"><td class="num">{total_value}</td><td class="label" colspan="2">الإجمالي</td><td class="qty">{item_count}</td></tr>
      {extra_summary_html}
    </tbody>
  </table>
  <div class="notes">{customer_line}{customer_phone_line}</div>
  <div class="footer">
    <div>رقم التليفون: {support_call or support_whatsapp}</div>
    {support_call_line}
    {support_whatsapp_line}
  </div>
</div>
</body>
</html>"""

    _write_html_file(path, html)

def _html(s: str) -> str:
    return western_digits(s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _num_html(value: Any) -> str:
    return f'<span class="digits">{_html(str(value))}</span>'


def _receipt_phone_html(value: Any) -> str:
    text = western_digits(str(value or "").strip())
    if not text:
        return ""
    return f'<span class="receipt-phone" dir="ltr">{_html(text)}</span>'


def _money_html(value: Any) -> str:
    return _num_html(format_money(value))

# Normalize size text for matching (handles Arabic/Persian digits & case)
_AR_DIGITS = _DIGIT_TRANSLATION
def _normalize_size_label(s: str) -> str:
    return (s or "").strip().translate(_AR_DIGITS).upper()


def _parse_money_amount(raw: Any) -> float:
    """Parse user-entered money (Arabic/Latin digits, common decimal separators)."""
    if raw is None:
        return 0.0
    s = str(raw).strip().translate(_AR_DIGITS)
    for ch in ("\u066b", "\u060c"):  # Arabic decimal / thousands comma
        s = s.replace(ch, ".")
    s = s.replace(" ", "").replace("\u00a0", "").replace("\u2009", "")
    if not s:
        return 0.0
    # If only comma as separator, treat as decimal (e.g. 40,50)
    if "," in s and "." not in s:
        s = s.replace(",", ".")
    else:
        s = s.replace(",", "")
    return float(s)


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


class PosAuditLogger:
    """Append-only, hash-chained daily audit log for POS business operations."""

    SCHEMA_VERSION = 1
    MAX_DETAIL_TEXT = 500
    MAX_LIST_ITEMS = 80
    SENSITIVE_KEYS = {"password", "plain", "secret", "token", "hash"}

    def __init__(self, db_path: str, device_supplier=None) -> None:
        self.disabled = str(os.environ.get("HOSNY_POS_AUDIT_DISABLED", "")).strip().lower() in ("1", "true", "yes")
        self.device_supplier = device_supplier
        self.session_id = uuid.uuid4().hex
        self._seq = 0
        self._last_hash = ""
        self.log_dir = self._resolve_log_dir(db_path)
        self.state_path = os.path.join(self.log_dir, "pos-audit-state.json")
        if not self.disabled:
            os.makedirs(self.log_dir, exist_ok=True)
            self._load_state()

    def _resolve_log_dir(self, db_path: str) -> str:
        db_dir = os.path.dirname(os.path.abspath(db_path or DB_PATH))
        runtime_dir = os.path.abspath(_runtime_base_dir())
        try:
            common = os.path.commonpath([db_dir, runtime_dir])
        except Exception:
            common = ""
        if common == runtime_dir:
            return os.path.join(runtime_dir, "audit_logs")
        db_stem = _safe_log_component(os.path.splitext(os.path.basename(db_path or DB_PATH))[0]) or "pos-db"
        return os.path.join(db_dir, f"audit_logs-{db_stem}")

    def _load_state(self) -> None:
        try:
            with open(self.state_path, "r", encoding="utf-8") as f:
                state = json.load(f)
            self._seq = int(state.get("seq") or 0)
            self._last_hash = str(state.get("last_hash") or "")
        except Exception:
            self._seq = 0
            self._last_hash = ""

    def _write_state(self, entry_date: str, entry_hash: str, seq: int) -> None:
        state = {
            "schema_version": self.SCHEMA_VERSION,
            "date": entry_date,
            "seq": int(seq),
            "last_hash": str(entry_hash or ""),
            "updated_at": datetime.now().isoformat(timespec="seconds"),
        }
        tmp_path = f"{self.state_path}.tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
            f.write("\n")
            try:
                f.flush()
                os.fsync(f.fileno())
            except Exception:
                pass
        os.replace(tmp_path, self.state_path)

    def _device_name(self) -> str:
        try:
            value = self.device_supplier() if self.device_supplier else None
        except Exception:
            value = None
        return str(value or os.environ.get("HOSNY_DEVICE_NAME") or os.environ.get("COMPUTERNAME") or "pos").strip()

    def _sanitize(self, value: Any, depth: int = 0) -> Any:
        if depth > 4:
            return "<max-depth>"
        if value is None or isinstance(value, (bool, int, float)):
            return value
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, sqlite3.Row):
            return self._sanitize(dict(value), depth + 1)
        if isinstance(value, dict):
            out: Dict[str, Any] = {}
            for key, item in value.items():
                key_text = str(key)
                if key_text.lower() in self.SENSITIVE_KEYS:
                    out[key_text] = "<redacted>"
                else:
                    out[key_text] = self._sanitize(item, depth + 1)
            return out
        if isinstance(value, (list, tuple, set)):
            seq = list(value)
            items = [self._sanitize(item, depth + 1) for item in seq[: self.MAX_LIST_ITEMS]]
            if len(seq) > self.MAX_LIST_ITEMS:
                items.append({"truncated_count": len(seq) - self.MAX_LIST_ITEMS})
            return items
        text = str(value)
        if len(text) > self.MAX_DETAIL_TEXT:
            return text[: self.MAX_DETAIL_TEXT] + "...<truncated>"
        return text

    def write(self, event: str, details: Optional[Dict[str, Any]] = None, *, actor: str = "system", shift_id: Optional[int] = None) -> None:
        if self.disabled:
            return
        ts = datetime.now().isoformat(timespec="seconds")
        entry_date = ts[:10]
        self._seq += 1
        base = {
            "schema_version": self.SCHEMA_VERSION,
            "ts": ts,
            "date": entry_date,
            "device": self._device_name(),
            "app_version": APP_VERSION,
            "session_id": self.session_id,
            "seq": self._seq,
            "event": str(event or "unknown"),
            "shift_id": shift_id,
            "actor": str(actor or "system"),
            "details": self._sanitize(details or {}),
            "prev_hash": self._last_hash,
        }
        canonical = json.dumps(base, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        entry_hash = hashlib.sha256(canonical.encode("utf-8")).hexdigest()
        entry = dict(base)
        entry["entry_hash"] = entry_hash
        line = json.dumps(entry, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n"
        path = os.path.join(self.log_dir, f"pos-audit-{entry_date}.jsonl")
        flags = os.O_APPEND | os.O_CREAT | os.O_WRONLY
        fd = os.open(path, flags, 0o600)
        try:
            os.write(fd, line.encode("utf-8"))
            try:
                os.fsync(fd)
            except Exception:
                pass
        finally:
            os.close(fd)
        self._last_hash = entry_hash
        self._write_state(entry_date, entry_hash, self._seq)


class SqliteDatabase:
    """SQLite persistence layer."""

    def __init__(self, path: str = DB_PATH, legacy_json: str = LEGACY_JSON_PATH) -> None:
        self.path = path
        try:
            logging_setup.log_event(  # type: ignore[union-attr]
                "db.open.start",
                db_path=os.path.abspath(self.path),
                exists=os.path.exists(self.path),
                size_bytes=(os.path.getsize(self.path) if os.path.exists(self.path) else 0),
            )
        except Exception:
            pass
        self.conn = sqlite3.connect(self.path, timeout=30.0, isolation_level=None, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.active_shift_id: Optional[int] = None
        self.audit = PosAuditLogger(self.path, device_supplier=self._current_device_name)

        started = time.time()
        self._apply_pragmas()
        self._init_schema()
        self._migrate_from_json_if_empty(legacy_json)
        normalized_ranges = self._normalize_legacy_size_profile_ranges()
        self._log_db_open_step("repair_stock_integrity.start")
        self.repair_stock_integrity()
        self._log_db_open_step("repair_stock_integrity.done")
        self._log_db_open_step("branch_catalog_visibility.start")
        visibility_repair = self.ensure_branch_catalog_stock_rows_lightweight()
        self._log_db_open_step("branch_catalog_visibility.done", result=visibility_repair)
        rename_repair = self._maybe_run_heavy_startup_repair(
            "repair_unsafe_spec_value_rename_damage",
            self.repair_unsafe_spec_value_rename_damage,
        )
        self._log_db_open_step("repair_missing_branch_reclassification_counts.start")
        try:
            reclass_repair = self.repair_missing_branch_reclassification_counts()
        except Exception:
            import traceback
            traceback.print_exc()
            reclass_repair = {"skipped": True, "error": "exception", "repair_name": "repair_missing_branch_reclassification_counts"}
        self._log_db_open_step("repair_missing_branch_reclassification_counts.done", result=reclass_repair)
        allow_snapshot_count_repairs = str(
            os.environ.get("HOSNY_POS_ENABLE_SNAPSHOT_COUNT_REPAIRS") or ""
        ).strip().lower() in {"1", "true", "yes", "on"}
        if allow_snapshot_count_repairs:
            self._log_db_open_step("repair_branch_group_general_school_counts.start")
            try:
                branch_group_repair = self.repair_branch_group_general_school_counts_from_snapshots()
            except Exception:
                import traceback
                traceback.print_exc()
                branch_group_repair = {"skipped": True, "error": "exception", "repair_name": "repair_branch_group_general_school_counts_from_snapshots"}
            self._log_db_open_step("repair_branch_group_general_school_counts.done", result=branch_group_repair)
            self._log_db_open_step("repair_stock_counts_from_snapshots.start")
            try:
                snapshot_count_repair = self.repair_stock_counts_from_snapshots()
            except Exception:
                import traceback
                traceback.print_exc()
                snapshot_count_repair = {"skipped": True, "error": "exception", "repair_name": "repair_stock_counts_from_snapshots"}
            self._log_db_open_step("repair_stock_counts_from_snapshots.done", result=snapshot_count_repair)
        else:
            branch_group_repair = {
                "skipped": True,
                "reason": "snapshot_count_repairs_disabled",
                "repair_name": "repair_branch_group_general_school_counts_from_snapshots",
            }
            snapshot_count_repair = {
                "skipped": True,
                "reason": "snapshot_count_repairs_disabled",
                "repair_name": "repair_stock_counts_from_snapshots",
            }
            self._log_db_open_step("repair_branch_group_general_school_counts.skipped", result=branch_group_repair)
            self._log_db_open_step("repair_stock_counts_from_snapshots.skipped", result=snapshot_count_repair)
        cleanup = self._maybe_run_heavy_startup_repair(
            "cleanup_unowned_branch_catalog_rows",
            self.cleanup_unowned_branch_catalog_rows,
        )
        self._log_db_open_step("source_truth_reset.start")
        source_truth_reset = self.reset_source_truth_sync_events_for_current_version()
        self._log_db_open_step("source_truth_reset.done")

        self._audit("app_started", {
            "db_path": os.path.abspath(self.path),
        })
        try:
            logging_setup.configure_context(device_name=self._current_device_name())  # type: ignore[union-attr]
            logging_setup.log_event(  # type: ignore[union-attr]
                "db.open.done",
                elapsed_ms=int((time.time() - started) * 1000),
                normalized_size_ranges=normalized_ranges,
                spec_rename_repair=rename_repair,
                branch_reclassification_repair=reclass_repair,
                branch_group_school_repair=branch_group_repair,
                stock_snapshot_count_repair=snapshot_count_repair,
                branch_catalog_cleanup=cleanup,
                branch_catalog_visibility=visibility_repair,
                source_truth_reset=source_truth_reset,
            )
        except Exception:
            pass

    def _log_db_open_step(self, step_name: str, **details: Any) -> None:
        try:
            logging_setup.log_event(  # type: ignore[union-attr]
                "db.open.step",
                step=step_name,
                **details,
            )
        except Exception:
            pass

    def _maybe_run_heavy_startup_repair(self, name: str, func) -> Dict[str, Any]:
        """Avoid blocking POS launch with historical whole-database repairs.

        The repair/cleanup functions remain available from the relevant UI
        actions and tests, but startup must open the cashier window quickly.
        New sync events are still gated strictly by the sync appliers.
        """
        enabled = str(os.environ.get("HOSNY_POS_HEAVY_STARTUP_REPAIR") or "").strip().lower()
        if enabled not in {"1", "true", "yes", "on"}:
            result: Dict[str, Any] = {"skipped": True, "reason": "deferred_after_startup", "repair_name": name}
            self._log_db_open_step(name + ".skipped", **result)
            return result
        self._log_db_open_step(name + ".start")
        try:
            result = func()
        except Exception:
            import traceback
            traceback.print_exc()
            result = {"skipped": True, "error": "exception", "repair_name": name}
        self._log_db_open_step(name + ".done", result=result)
        return result

    def _normalize_legacy_size_profile_ranges(self) -> int:
        try:
            with self.conn:
                c1 = self.conn.execute(
                    """
                    UPDATE size_profiles
                       SET num_start_1 = 30,
                           updated_at = datetime('now')
                     WHERE num_start_1 = 32
                       AND num_end_1 = 62
                    """
                ).rowcount
                c2 = self.conn.execute(
                    """
                    UPDATE size_profiles
                       SET num_start_2 = 30,
                           updated_at = datetime('now')
                     WHERE num_start_2 = 32
                       AND num_end_2 = 62
                    """
                ).rowcount
            return int(c1 or 0) + int(c2 or 0)
        except Exception:
            return 0

    def reset_source_truth_sync_events_for_current_version(self) -> Dict[str, int]:
        """Let an updated POS apply source-of-truth events skipped by older builds."""
        key = "pos_source_truth_reset_version"
        try:
            row = self.conn.execute("SELECT value FROM app_settings WHERE key = ?", (key,)).fetchone()
            if row is not None and str(row["value"] or "") == APP_VERSION:
                return {"reset": 0, "already_current": 1}
        except sqlite3.OperationalError:
            return {"reset": 0, "already_current": 0}
        try:
            with self.conn:
                legacy_spec_renames = self.conn.execute(
                    """
                    UPDATE sync_inbox
                       SET apply_status = 'skipped',
                           apply_error = 'legacy SPEC_RENAMED replay suppressed; use branch reclassification/source truth',
                           apply_at = COALESCE(apply_at, ?)
                     WHERE event_type = 'SPEC_RENAMED'
                       AND apply_status IS NULL
                       AND COALESCE(apply_attempts, 0) > 0
                    """,
                    (now_iso(),),
                )
                cur = self.conn.execute(
                    """
                    UPDATE sync_inbox
                       SET apply_status = NULL,
                           apply_error = NULL
                     WHERE event_type IN ('POS_OWNERSHIP_SNAPSHOT', 'PRICE_UPDATE', 'BRANCH_STOCK_RECLASSIFIED')
                       AND apply_status = 'skipped'
                    """
                )
                reset = int(cur.rowcount or 0)
                self.conn.execute(
                    "INSERT OR REPLACE INTO app_settings(key, value) VALUES(?, ?)",
                    (key, APP_VERSION),
                )
            return {
                "reset": reset,
                "already_current": 0,
                "legacy_spec_renames_suppressed": int(legacy_spec_renames.rowcount or 0),
            }
        except sqlite3.OperationalError:
            return {"reset": 0, "already_current": 0}


    def repair_stock_integrity(self) -> Dict[str, int]:
        fixed_negative = 0
        merged_rows = 0
        restored_reservation_deductions = 0
        with self.conn:
            negative_rows = self.conn.execute(
                "SELECT * FROM stocks WHERE COALESCE(count,0) < 0"
            ).fetchall()
            for s in negative_rows:
                qty = abs(int(s["count"] or 0))
                self.conn.execute("UPDATE stocks SET count=0 WHERE id=?", (int(s["id"]),))
                self.conn.execute(
                    """INSERT INTO movements
                    (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        now_iso(), "ADJUST_IN", int(s["id"]), qty,
                        "Automatic repair: negative stock clamped to zero",
                        None, s["item_type"], s["school"], s["color"], s["size"],
                        float(s["unit_price"]),
                    ),
                )
                fixed_negative += 1

            legacy_reservation_deductions = self.conn.execute(
                """
                SELECT
                    m.*,
                    COALESCE(b.bill_type, '') AS bill_type
                FROM movements m
                LEFT JOIN bills b ON b.id = m.bill_id
                WHERE COALESCE(m.qty,0) > 0
                  AND (
                        (
                            UPPER(COALESCE(b.bill_type,'')) = 'RESERVATION'
                            AND UPPER(COALESCE(m.direction,'')) = 'OUT'
                        )
                     OR (
                            UPPER(COALESCE(m.direction,'')) = 'RESERVE'
                            AND m.stock_id IS NOT NULL
                        )
                  )
                ORDER BY m.id ASC
                """
            ).fetchall()
            for m in legacy_reservation_deductions:
                marker = f"Automatic repair: restored stock from legacy reservation deduction movement #{int(m['id'])}"
                already = self.conn.execute(
                    "SELECT 1 FROM movements WHERE direction='ADJUST_IN' AND note=? LIMIT 1",
                    (marker,),
                ).fetchone()
                if already:
                    continue

                qty = int(m["qty"] or 0)
                if qty <= 0:
                    continue
                stock_id = int(m["stock_id"] or 0)
                stock_row = None
                if stock_id > 0:
                    stock_row = self.conn.execute(
                        "SELECT * FROM stocks WHERE id=?",
                        (stock_id,),
                    ).fetchone()

                if stock_row is not None:
                    self.conn.execute(
                        "UPDATE stocks SET count=count+? WHERE id=?",
                        (qty, stock_id),
                    )
                    restored_stock_id = stock_id
                    item_type = stock_row["item_type"]
                    school = stock_row["school"]
                    color = stock_row["color"]
                    size = stock_row["size"]
                    unit_price = float(stock_row["unit_price"] or m["unit_price"] or 0)
                else:
                    item_type = str(m["item_type"] or "").strip()
                    school = str(m["school"] or "").strip()
                    color = str(m["color"] or "").strip()
                    size = str(m["size"] or "").strip()
                    unit_price = float(m["unit_price"] or 0)
                    if not all((item_type, school, color, size)):
                        continue
                    restored_stock_id = self.add_or_update_stock_row(
                        item_type, school, color, size, unit_price, qty,
                    )

                self.conn.execute(
                    """INSERT INTO movements
                    (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        now_iso(), "ADJUST_IN", int(restored_stock_id), qty,
                        marker, m["bill_id"], item_type, school, color, size,
                        unit_price,
                    ),
                )
                restored_reservation_deductions += 1

            duplicate_groups = self.conn.execute(
                """
                SELECT item_type,school,color,size,unit_price
                FROM stocks
                GROUP BY item_type,school,color,size,unit_price
                HAVING COUNT(*) > 1
                """
            ).fetchall()
            for g in duplicate_groups:
                rows = self.conn.execute(
                    """
                    SELECT id,COALESCE(count,0) AS count
                    FROM stocks
                    WHERE item_type=? AND school=? AND color=? AND size=? AND unit_price=?
                    ORDER BY id ASC
                    """,
                    (g["item_type"], g["school"], g["color"], g["size"], float(g["unit_price"])),
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
        return {
            "negative_fixed": fixed_negative,
            "duplicates_merged": merged_rows,
            "reservation_deductions_restored": restored_reservation_deductions,
        }


    def add_or_update_stock_row(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
        unit_price: float,
        qty: int,
    ) -> int:
        item_type = str(item_type or "").strip()
        school = str(school or "").strip()
        color = str(color or "").strip()
        size = str(size or "").strip()
        qty = int(qty or 0)
        if qty < 0:
            raise ValueError("Stock quantity cannot be negative")
        row = self.conn.execute(
            """
            SELECT id FROM stocks
            WHERE item_type=? AND school=? AND color=? AND size=? AND unit_price=?
            ORDER BY id ASC
            LIMIT 1
            """,
            (item_type, school, color, size, float(unit_price)),
        ).fetchone()
        if row:
            self.conn.execute(
                "UPDATE stocks SET count=count+? WHERE id=?",
                (qty, int(row["id"])),
            )
            return int(row["id"])
        cur = self.conn.execute(
            """INSERT INTO stocks(item_type,school,color,size,unit_price,count)
            VALUES(?,?,?,?,?,?)""",
            (item_type, school, color, size, float(unit_price), qty),
        )
        return int(cur.lastrowid)

    def ensure_branch_catalog_stock_rows_lightweight(self) -> Dict[str, int]:
        """Create missing zero-count stock rows from branch ownership anchors.

        This is intentionally cheap enough for Inventory/Stock Audit refreshes.
        The heavier source-truth cleanup remains manual/deferred and must not
        run while the cashier is opening interactive windows.
        """
        created = 0
        price_updates = 0
        inbox_specs = 0
        seen: Set[Tuple[str, str, str, str]] = set()

        def _clean_local(value: Any) -> str:
            return _strip_digit_marks(value).strip()

        def _complete_spec_from(raw: Dict[str, Any]) -> Optional[Dict[str, str]]:
            spec = {
                "item_type": _clean_local(raw.get("item_type")),
                "school": _clean_local(raw.get("school")),
                "color": _clean_local(raw.get("color")),
                "size": _clean_local(raw.get("size")),
            }
            return spec if all(spec.values()) else None

        def _load_rename_history() -> List[Tuple[Dict[str, str], Dict[str, str]]]:
            try:
                rows = self.conn.execute(
                    """
                    SELECT payload_json
                      FROM sync_inbox
                     WHERE event_type = 'SPEC_RENAMED'
                     ORDER BY server_seq ASC
                    """
                ).fetchall()
            except sqlite3.OperationalError:
                return []
            history: List[Tuple[Dict[str, str], Dict[str, str]]] = []
            for row in rows:
                try:
                    payload = json.loads(row["payload_json"] or "{}")
                except Exception:
                    continue
                old_spec = _complete_spec_from(payload.get("old_spec") or {})
                new_spec_raw = payload.get("new_spec") or {}
                if old_spec is None or not isinstance(new_spec_raw, dict):
                    continue
                new_spec = {
                    k: _clean_local(new_spec_raw.get(k) or old_spec.get(k) or "")
                    for k in ("item_type", "school", "color", "size")
                }
                if all(new_spec.values()) and old_spec != new_spec:
                    history.append((old_spec, new_spec))
            return history

        rename_history = _load_rename_history()

        def _apply_local_renames(spec: Dict[str, str]) -> Dict[str, str]:
            current = {k: _clean_local(spec.get(k)) for k in ("item_type", "school", "color", "size")}
            for old_spec, new_spec in rename_history:
                if all(
                    current.get(k, "").casefold() == old_spec.get(k, "").casefold()
                    for k in ("item_type", "school", "color", "size")
                ):
                    current = dict(new_spec)
            return current

        def _load_delete_filters() -> List[Dict[str, str]]:
            try:
                self.conn.execute(
                    """
                    CREATE TABLE IF NOT EXISTS branch_catalog_delete_tombstones(
                        item_type TEXT,
                        school TEXT,
                        color TEXT,
                        size TEXT,
                        delete_server_seq INTEGER,
                        source_event_uuid TEXT,
                        note TEXT,
                        created_at TEXT NOT NULL
                    )
                    """
                )
                try:
                    cols = {
                        str(row["name"])
                        for row in self.conn.execute("PRAGMA table_info(branch_catalog_delete_tombstones)").fetchall()
                    }
                    if "delete_server_seq" not in cols:
                        self.conn.execute("ALTER TABLE branch_catalog_delete_tombstones ADD COLUMN delete_server_seq INTEGER")
                except sqlite3.OperationalError:
                    pass
                rows = self.conn.execute(
                    "SELECT item_type, school, color, size, delete_server_seq FROM branch_catalog_delete_tombstones"
                ).fetchall()
            except sqlite3.OperationalError:
                return []
            out = []
            for row in rows:
                filt = {field: _clean_local(row[field]) for field in ("item_type", "school", "color", "size")}
                try:
                    filt["delete_server_seq"] = int(row["delete_server_seq"]) if row["delete_server_seq"] is not None else None  # type: ignore[assignment]
                except (TypeError, ValueError):
                    filt["delete_server_seq"] = None  # type: ignore[assignment]
                out.append(filt)
            return out

        delete_filters = _load_delete_filters()

        def _deleted_by_filter(spec: Dict[str, str], event_seq: Optional[int] = None) -> bool:
            for filt in delete_filters:
                delete_seq = filt.get("delete_server_seq")
                if isinstance(delete_seq, int) and event_seq is not None and event_seq > delete_seq:
                    continue
                if isinstance(delete_seq, int) and event_seq is None:
                    continue
                matched = True
                for field in ("item_type", "school", "color", "size"):
                    value = filt.get(field) or ""
                    if value and value.casefold() != _clean_local(spec.get(field)).casefold():
                        matched = False
                        break
                if matched:
                    return True
            return False

        def _remember(
            specs: Dict[str, Any],
            unit_price: Any,
            *,
            from_inbox: bool = False,
            event_seq: Optional[int] = None,
        ) -> None:
            nonlocal created, price_updates, inbox_specs
            spec = _complete_spec_from(specs)
            if spec is None:
                return
            spec = _apply_local_renames(spec)
            if _deleted_by_filter(spec, event_seq):
                return
            item_type = spec["item_type"]
            school = spec["school"]
            color = spec["color"]
            size = spec["size"]
            key = (item_type.casefold(), school.casefold(), color.casefold(), size.casefold())
            if key in seen:
                return
            seen.add(key)
            if from_inbox:
                inbox_specs += 1
            try:
                price = float(unit_price or 0)
            except (TypeError, ValueError):
                price = 0.0
            existing = self.conn.execute(
                """
                SELECT id, unit_price
                  FROM stocks
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                 ORDER BY id ASC
                 LIMIT 1
                """,
                (item_type, school, color, size),
            ).fetchone()
            if existing is None:
                self.conn.execute(
                    """
                    INSERT INTO stocks(item_type, school, color, size, unit_price, count)
                    VALUES (?, ?, ?, ?, ?, 0)
                    """,
                    (item_type, school, color, size, price),
                )
                created += 1
                self._upsert_history(
                    {"item_type": item_type, "school": school, "color": color, "size": size}
                )
            elif price > 0 and abs(float(existing["unit_price"] or 0) - price) >= 0.001:
                cur = self.conn.execute(
                    "UPDATE stocks SET unit_price = ? WHERE id = ? AND COALESCE(count, 0) = 0",
                    (price, int(existing["id"])),
                )
                price_updates += int(cur.rowcount or 0)

        try:
            with self.conn:
                try:
                    rows = self.conn.execute(
                        """
                        SELECT item_type, school, color, size, unit_price
                          FROM incoming_shipment_items_pending
                         WHERE (
                                COALESCE(expected_qty, 0) > 0
                                OR COALESCE(received_qty, 0) > 0
                           )
                           AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    rows = []
                for row in rows:
                    _remember(dict(row), row["unit_price"])

                try:
                    rows = self.conn.execute(
                        """
                        SELECT item_type, school, color, size, unit_price, source_event_uuid
                          FROM branch_catalog_definitions
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    rows = []
                for row in rows:
                    source_uuid = _clean_local(row["source_event_uuid"])
                    seq_row = self.conn.execute(
                        "SELECT server_seq FROM sync_inbox WHERE event_uuid = ? LIMIT 1",
                        (source_uuid,),
                    ).fetchone() if source_uuid else None
                    try:
                        event_seq = int(seq_row["server_seq"]) if seq_row else None
                    except (TypeError, ValueError):
                        event_seq = None
                    _remember(dict(row), row["unit_price"], event_seq=event_seq)

                cancelled_shipments: Set[str] = set()
                try:
                    rows = self.conn.execute(
                        """
                        SELECT payload_json
                          FROM sync_inbox
                         WHERE event_type = 'STOCK_TRANSFER_CANCELLED'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    rows = []
                for row in rows:
                    try:
                        payload = json.loads(row["payload_json"] or "{}")
                    except Exception:
                        continue
                    shipment_uuid = _clean_local(payload.get("shipment_uuid") or payload.get("bill_uuid"))
                    if shipment_uuid:
                        cancelled_shipments.add(shipment_uuid.casefold())

                try:
                    rows = self.conn.execute(
                        """
                        SELECT payload_json, server_seq
                          FROM sync_inbox
                         WHERE event_type = 'STOCK_TRANSFER_OUT'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    rows = []
                for row in rows:
                    try:
                        payload = json.loads(row["payload_json"] or "{}")
                    except Exception:
                        continue
                    shipment_uuid = _clean_local(payload.get("shipment_uuid"))
                    if shipment_uuid and shipment_uuid.casefold() in cancelled_shipments:
                        continue
                    note_text = _clean_local(payload.get("note")).casefold()
                    is_reservation_definition = "reservation" in note_text or "حجز" in note_text
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        try:
                            qty = int(float(item.get("qty") or 0))
                        except (TypeError, ValueError):
                            qty = 0
                        catalog_only = bool(item.get("catalog_only"))
                        if qty <= 0 and not (catalog_only and is_reservation_definition):
                            continue
                        try:
                            event_seq = int(row["server_seq"])
                        except (TypeError, ValueError):
                            event_seq = None
                        _remember(item, item.get("unit_price"), from_inbox=True, event_seq=event_seq)

                try:
                    rows = self.conn.execute(
                        """
                        SELECT payload_json, server_seq
                          FROM sync_inbox
                         WHERE event_type = 'BRANCH_STOCK_RECLASSIFIED'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    rows = []
                for row in rows:
                    try:
                        payload = json.loads(row["payload_json"] or "{}")
                    except Exception:
                        continue
                    to_spec = payload.get("to_spec") or {}
                    if not isinstance(to_spec, dict):
                        continue
                    try:
                        event_seq = int(row["server_seq"])
                    except (TypeError, ValueError):
                        event_seq = None
                    _remember(to_spec, to_spec.get("unit_price"), from_inbox=True, event_seq=event_seq)

                stale_renamed_zero_rows = 0
            return {
                "created": created,
                "price_updates": price_updates,
                "owned_specs": len(seen),
                "inbox_specs": inbox_specs,
                "stale_renamed_zero_rows": stale_renamed_zero_rows,
                "stale_renamed_zero_cleanup": "manual_delete_definition_only",
            }
        except Exception:
            import traceback
            traceback.print_exc()
            return {"created": 0, "price_updates": 0, "owned_specs": 0, "inbox_specs": 0, "error": 1}


    def deduct_stock_for_specs(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
        qty: int,
        *,
        note: str,
        bill_id: Optional[int] = None,
        preferred_stock_id: Optional[int] = None,
    ) -> List[Tuple[sqlite3.Row, int]]:
        qty_needed = int(qty or 0)
        if qty_needed <= 0:
            raise ValueError("Qty must be > 0")
        rows = self.conn.execute(
            """
            SELECT * FROM stocks
            WHERE count>0
              AND LOWER(TRIM(item_type))=LOWER(TRIM(?))
              AND LOWER(TRIM(school))=LOWER(TRIM(?))
              AND LOWER(TRIM(color))=LOWER(TRIM(?))
              AND LOWER(TRIM(size))=LOWER(TRIM(?))
            ORDER BY CASE WHEN id=? THEN 0 ELSE 1 END, id ASC
            """,
            (
                str(item_type or "").strip(),
                str(school or "").strip(),
                str(color or "").strip(),
                str(size or "").strip(),
                int(preferred_stock_id or 0),
            ),
        ).fetchall()
        remaining = qty_needed
        chunks: List[Tuple[sqlite3.Row, int]] = []
        for s in rows:
            if remaining <= 0:
                break
            take = min(int(s["count"] or 0), remaining)
            if take <= 0:
                continue
            self.conn.execute(
                "UPDATE stocks SET count=count-? WHERE id=?",
                (int(take), int(s["id"])),
            )
            self.conn.execute(
                """INSERT INTO movements
                (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    now_iso(), "OUT", int(s["id"]), int(take), note, bill_id,
                    s["item_type"], s["school"], s["color"], s["size"],
                    float(s["unit_price"]),
                ),
            )
            chunks.append((s, take))
            remaining -= take
        if remaining > 0:
            raise ValueError("Insufficient stock for reservation delivery")
        return chunks


    def _require_shift(self):
        self.assert_clock_sane()
        if self.active_shift_id is None:
            raise ValueError("لا يمكن إجراء عمليات بدون وردية مفتوحة.")

    @staticmethod
    def _parse_db_datetime(value: Any) -> Optional[datetime]:
        raw = _strip_digit_marks(value).strip()
        if not raw:
            return None
        try:
            dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
            if dt.tzinfo is not None:
                dt = dt.astimezone().replace(tzinfo=None)
            return dt
        except Exception:
            try:
                return datetime.strptime(raw[:19].replace("T", " "), "%Y-%m-%d %H:%M:%S")
            except Exception:
                return None

    def _table_has_column(self, table: str, column: str) -> bool:
        try:
            return column in {r["name"] for r in self.conn.execute(f"PRAGMA table_info({table})")}
        except Exception:
            return False

    def _latest_recorded_datetime(self) -> Optional[datetime]:
        latest: Optional[datetime] = None
        sources = [
            "SELECT MAX(started_at) FROM shifts",
            "SELECT MAX(ended_at) FROM shifts",
            "SELECT MAX(created_at) FROM bills",
            "SELECT MAX(ts) FROM movements",
            "SELECT MAX(created_at) FROM reservations",
        ]
        for sql in sources:
            try:
                row = self.conn.execute(sql).fetchone()
            except Exception:
                continue
            dt = self._parse_db_datetime(row[0] if row else None)
            if dt and (latest is None or dt > latest):
                latest = dt
        return latest

    def assert_clock_sane(self) -> None:
        latest = self._latest_recorded_datetime()
        if not latest:
            return
        if datetime.now() + timedelta(minutes=10) < latest:
            raise ValueError(
                "تاريخ أو وقت الكمبيوتر أقدم من آخر عملية مسجلة في النظام. "
                f"آخر عملية محلية: {latest.strftime('%Y-%m-%d %H:%M:%S')}. "
                "اضبط التاريخ والوقت أولا ثم افتح البرنامج مرة أخرى."
            )

    def _backfill_legacy_shift_ids(self) -> None:
        shifts = self.conn.execute(
            """
            SELECT id, started_at, ended_at
            FROM shifts
            WHERE ended_at IS NOT NULL
            ORDER BY started_at, id
            """
        ).fetchall()
        if not shifts:
            return
        specs = [("bills", "created_at"), ("reservations", "created_at"), ("movements", "ts")]
        with self.conn:
            for s in shifts:
                started = str(s["started_at"] or "")
                ended = str(s["ended_at"] or "")
                if not started or not ended:
                    continue
                for table, ts_col in specs:
                    if not self._table_has_column(table, "shift_id"):
                        continue
                    self.conn.execute(
                        f"""
                        UPDATE {table}
                           SET shift_id=?
                         WHERE shift_id IS NULL
                           AND {ts_col} >= ?
                           AND {ts_col} <= ?
                        """,
                        (int(s["id"]), started, ended),
                    )

    def _exact_shift_activity_count(self, shift_id: int) -> int:
        total = 0
        for table in ("bills", "reservations", "movements"):
            if not self._table_has_column(table, "shift_id"):
                continue
            try:
                row = self.conn.execute(
                    f"SELECT COUNT(*) FROM {table} WHERE shift_id=?",
                    (int(shift_id),),
                ).fetchone()
                total += int(row[0] if row else 0)
            except Exception:
                pass
        return total

    def _close_empty_impossible_open_shifts(self) -> None:
        row = self.conn.execute(
            "SELECT MAX(ended_at) FROM shifts WHERE status='CLOSED' AND ended_at IS NOT NULL"
        ).fetchone()
        max_closed = str(row[0] or "") if row else ""
        max_closed_dt = self._parse_db_datetime(max_closed)
        open_rows = self.conn.execute(
            "SELECT id, started_at FROM shifts WHERE status='OPEN' ORDER BY started_at"
        ).fetchall()
        with self.conn:
            for r in open_rows:
                sid = int(r["id"])
                started_dt = self._parse_db_datetime(r["started_at"])
                if not started_dt:
                    continue
                overlaps_closed_shift = bool(max_closed_dt and started_dt < max_closed_dt)
                is_empty_stale_shift = datetime.now() - started_dt > timedelta(hours=18)
                if not overlaps_closed_shift and not is_empty_stale_shift:
                    continue
                if self._exact_shift_activity_count(sid) > 0:
                    continue
                closed_at = max_closed if overlaps_closed_shift else now_iso()
                reason = (
                    "Auto-closed empty overlapping shift after clock rollback detection."
                    if overlaps_closed_shift
                    else "Auto-closed empty stale shift."
                )
                self.conn.execute(
                    """
                    UPDATE shifts
                       SET status='CLOSED',
                           ended_at=?,
                           note=COALESCE(note || '\n', '') || ?
                     WHERE id=? AND status='OPEN'
                    """,
                    (closed_at, reason, sid),
                )

    def _allow_legacy_shift_time_fallback(self, shift: Dict[str, Any]) -> bool:
        started_dt = self._parse_db_datetime(shift.get("started_at"))
        ended_dt = self._parse_db_datetime(shift.get("ended_at"))
        if ended_dt is not None:
            return True
        if not started_dt:
            return False
        row = self.conn.execute(
            """
            SELECT MAX(ended_at)
            FROM shifts
            WHERE status='CLOSED'
              AND ended_at IS NOT NULL
              AND ended_at > ?
            """,
            (shift.get("started_at"),),
        ).fetchone()
        if row and row[0]:
            return False
        if datetime.now() - started_dt > timedelta(hours=18):
            return False
        return True

    def _shift_scope_sql(self, shift: Dict[str, Any], ts_col: str) -> Tuple[str, Tuple[Any, ...]]:
        sid = int(shift["id"])
        started = shift["started_at"]
        ended = shift["ended_at"] or now_iso()
        if self._allow_legacy_shift_time_fallback(shift):
            return (
                f"(shift_id=? OR (shift_id IS NULL AND {ts_col} >= ? AND {ts_col} <= ?))",
                (sid, started, ended),
            )
        return "(shift_id=?)", (sid,)

    def _stored_closed_shift_summary(self, shift: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        if str(shift.get("status") or "").upper() != "CLOSED":
            return None
        raw = str(shift.get("summary_json") or "").strip()
        if not raw:
            return None
        try:
            stored = json.loads(raw)
        except Exception:
            return None
        if not isinstance(stored, dict):
            return None
        row = dict(shift)
        row.update(stored)
        row["id"] = int(shift["id"])
        row["shift_id"] = int(shift["id"])
        row["started_at"] = shift.get("started_at") or stored.get("started_at") or ""
        row["ended_at"] = shift.get("ended_at") or stored.get("ended_at") or ""
        row["status"] = shift.get("status") or "CLOSED"
        row.setdefault("expense_count", 0)
        row.setdefault("expense_total", 0.0)
        row.setdefault("expenses", [])
        return row

    def _missing_reservation_payment_totals(self, shift: Dict[str, Any]) -> Dict[str, float]:
        """Recover reservation money if payment movement rows are missing.

        Reservation summaries normally use RESERVE_PAY/DELIVER_PAY movements.
        Some older/corrupted shifts have reservation paid_amount values without
        matching money movements, which made cash/visa summaries drop money that
        was already saved and synced.
        """
        res_scope, res_params = self._shift_scope_sql(shift, "created_at")
        try:
            rows = self.conn.execute(
                f"""
                SELECT
                    UPPER(COALESCE(NULLIF(TRIM(r.payment_method), ''), '{PAYMENT_METHOD_CASH}')) AS method,
                    COALESCE(SUM(
                        CASE
                            WHEN COALESCE(r.paid_amount, 0) > COALESCE((
                                SELECT SUM(m.unit_price)
                                  FROM movements m
                                 WHERE m.direction IN ('RESERVE_PAY', 'DELIVER_PAY')
                                   AND (m.note = 'عربون حجز #' || r.id OR m.note LIKE '%' || '#' || r.id)
                            ), 0)
                            THEN COALESCE(r.paid_amount, 0) - COALESCE((
                                SELECT SUM(m.unit_price)
                                  FROM movements m
                                 WHERE m.direction IN ('RESERVE_PAY', 'DELIVER_PAY')
                                   AND (m.note = 'عربون حجز #' || r.id OR m.note LIKE '%' || '#' || r.id)
                            ), 0)
                            ELSE 0
                        END
                    ), 0) AS missing_total
                  FROM reservations r
                 WHERE {res_scope}
                 GROUP BY UPPER(COALESCE(NULLIF(TRIM(r.payment_method), ''), '{PAYMENT_METHOD_CASH}'))
                """,
                res_params,
            ).fetchall()
        except sqlite3.OperationalError:
            return {"paid": 0.0, "cash_total": 0.0, "visa_total": 0.0}
        out = {"paid": 0.0, "cash_total": 0.0, "visa_total": 0.0}
        for row in rows:
            total = float(row["missing_total"] or 0.0)
            if total <= 1e-9:
                continue
            out["paid"] += total
            if str(row["method"] or "").upper() == PAYMENT_METHOD_VISA:
                out["visa_total"] += total
            else:
                out["cash_total"] += total
        return out

    def _apply_pragmas(self) -> None:
        cur = self.conn.cursor()
        cur.execute("PRAGMA journal_mode=WAL;")
        cur.execute("PRAGMA synchronous=NORMAL;")
        cur.execute("PRAGMA foreign_keys=ON;")
        cur.execute("PRAGMA busy_timeout=30000;")
        cur.execute("PRAGMA temp_store=MEMORY;")
        cur.execute("PRAGMA cache_size=-20000;")
        cur.close()

    def _record_sync_event(self, event_type: str, payload: Dict[str, Any]) -> Optional[str]:
        """Append a business event to sync_outbox. Fail-safe.

        Phase 1: events are written but never pushed. If the sync layer
        is not present or the outbox insert fails, the caller continues
        without error — we must never break a business operation.
        """
        try:
            from sync_core import record_event
            return record_event(self.conn, event_type, payload)
        except Exception:
            import traceback
            traceback.print_exc()
            return None

    def _record_sync_event_or_raise(self, event_type: str, payload: Dict[str, Any]) -> str:
        if event_type in {
            "SALE_CREATED",
            "SALE_RETURNED",
            "SALE_EXCHANGED",
            "SALE_VOIDED",
            "SALE_BILL_TYPE_CORRECTED",
            "RESERVATION_CREATED",
            "RESERVATION_PAYMENT_UPDATED",
            "RESERVATION_DELIVERED",
            "SHIFT_OPENED",
            "SHIFT_CLOSED",
        }:
            local_ts = str(payload.get("created_at") or payload.get("ended_at") or payload.get("started_at") or now_iso())
            payload.setdefault("created_at", local_ts)
            payload.setdefault("business_day", local_ts[:10])
        event_uuid = self._record_sync_event(event_type, payload)
        if not event_uuid:
            raise RuntimeError(f"Failed to record required sync event: {event_type}")
        try:
            logging_setup.log_event(  # type: ignore[union-attr]
                "sync.outbox.recorded",
                event_type=event_type,
                event_uuid=event_uuid,
                payload_summary=_summarize_sync_payload_for_ui(event_type, payload),
            )
        except Exception:
            pass
        return str(event_uuid)

    def _audit(self, event: str, details: Optional[Dict[str, Any]] = None, *, actor: str = "system") -> None:
        try:
            logging_setup.log_event(  # type: ignore[union-attr]
                f"business.{event}",
                actor=actor,
                shift_id=self.active_shift_id,
                details=details or {},
            )
        except Exception:
            pass
        try:
            self.audit.write(event, details or {}, actor=actor, shift_id=self.active_shift_id)
        except Exception:
            import traceback
            traceback.print_exc()

    def _record_warehouse_return_event(
        self,
        return_uuid: str,
        note: str,
        lines: List[Dict[str, Any]],
    ) -> str:
        """Append a warehouse-targeted stock return event to sync_outbox."""
        return self._record_targeted_inventory_event(
            event_type="STOCK_RETURN_TO_WAREHOUSE",
            target_scope="warehouse",
            payload={
                "return_uuid": return_uuid,
                "from_device": self._current_device_name() or "pos",
                "to_device": WAREHOUSE_DEVICE_NAME,
                "note": note or "",
                "items": lines,
            },
        )

    def _record_transfer_via_warehouse_event(
        self,
        request_uuid: str,
        target_device: str,
        note: str,
        lines: List[Dict[str, Any]],
    ) -> str:
        return self._record_targeted_inventory_event(
            event_type="POS_TRANSFER_VIA_WAREHOUSE",
            target_scope="warehouse",
            payload={
                "request_uuid": request_uuid,
                "from_device": self._current_device_name() or "pos",
                "target_device": target_device,
                "note": note or "",
                "items": lines,
            },
        )

    def _current_device_name(self) -> Optional[str]:
        try:
            ident = self.conn.execute(
                "SELECT device_name FROM device_identity WHERE id = 1"
            ).fetchone()
            if ident is not None:
                return ident[0]
        except Exception:
            pass
        return None

    def _record_targeted_inventory_event(
        self,
        event_type: str,
        target_scope: str,
        payload: Dict[str, Any],
    ) -> str:
        import json as _json
        import uuid as _uuid
        try:
            from sync_client import _utc_now_iso as _sync_now  # type: ignore
        except Exception:
            _sync_now = lambda: datetime.utcnow().isoformat(timespec="seconds") + "Z"
        event_uuid = str(_uuid.uuid4())
        now = _sync_now()

        try:
            has_target = False
            try:
                cur = self.conn.execute("PRAGMA table_info(sync_outbox)")
                has_target = any(r[1] == "target_scope" for r in cur.fetchall())
            except Exception:
                has_target = False

            if has_target:
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
                        now,
                        target_scope,
                    ),
                )
            else:
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
                        event_type,
                        _json.dumps(payload, ensure_ascii=False, default=str),
                        now,
                    ),
                )
            return str(event_uuid)
        except Exception:
            import traceback
            traceback.print_exc()
            raise RuntimeError(f"Failed to record required targeted sync event: {event_type}")

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
                unit_price REAL NOT NULL,
                count INTEGER NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_stocks_specs
            ON stocks(item_type, school, color, size);
            CREATE INDEX IF NOT EXISTS idx_stocks_count
            ON stocks(count);

            CREATE TABLE IF NOT EXISTS movements(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                ts TEXT NOT NULL,
                direction TEXT NOT NULL, -- IN | OUT | ADJUST_OUT | OUT_FACTORY | PRICE_UPDATE | RESERVE | RESERVE_PAY | DELIVER_PAY
                stock_id INTEGER,
                qty INTEGER NOT NULL,
                note TEXT,
                bill_id INTEGER,
                item_type TEXT,
                school TEXT,
                color TEXT,
                size TEXT,
                unit_price REAL,
                shift_id INTEGER,
                payment_method TEXT NOT NULL DEFAULT 'CASH'
            );
            CREATE INDEX IF NOT EXISTS idx_movements_ts ON movements(ts DESC);
            CREATE INDEX IF NOT EXISTS idx_movements_dir ON movements(direction);
            CREATE INDEX IF NOT EXISTS idx_movements_specs ON movements(item_type,school,color,size);

            CREATE TABLE IF NOT EXISTS stock_audit_reports(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                reason TEXT,
                diff_count INTEGER NOT NULL DEFAULT 0,
                total_diff INTEGER NOT NULL DEFAULT 0,
                total_value REAL NOT NULL DEFAULT 0,
                bucket_key TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_stock_audit_reports_created
                ON stock_audit_reports(created_at DESC);

            CREATE TABLE IF NOT EXISTS stock_audit_report_lines(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                report_id INTEGER NOT NULL,
                stock_id INTEGER,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                expected INTEGER NOT NULL,
                actual INTEGER NOT NULL,
                diff INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                diff_value REAL NOT NULL,
                FOREIGN KEY(report_id) REFERENCES stock_audit_reports(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_stock_audit_lines_report
                ON stock_audit_report_lines(report_id);
            CREATE INDEX IF NOT EXISTS idx_stock_audit_lines_specs
                ON stock_audit_report_lines(item_type, school, color, size);

            CREATE TABLE IF NOT EXISTS bills(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                customer TEXT,
                customer_phone TEXT,
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

            CREATE TABLE IF NOT EXISTS branch_catalog_definitions(
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                unit_price REAL NOT NULL DEFAULT 0,
                source_event_uuid TEXT,
                note TEXT,
                created_at TEXT NOT NULL,
                PRIMARY KEY(item_type, school, color, size)
            );

            CREATE TABLE IF NOT EXISTS hidden_definitions(
                field TEXT NOT NULL,
                value TEXT NOT NULL,
                hidden_at TEXT NOT NULL,
                PRIMARY KEY(field, value)
            );

            CREATE TABLE IF NOT EXISTS reservations(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                customer TEXT,
                customer_phone TEXT,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                qty INTEGER NOT NULL,
                unit_price REAL NOT NULL,
                total_amount REAL NOT NULL,
                paid_amount REAL NOT NULL DEFAULT 0,
                payment_method TEXT NOT NULL DEFAULT 'CASH',
                status TEXT NOT NULL DEFAULT 'معلق',
                note TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_reservations_status ON reservations(status);
            CREATE INDEX IF NOT EXISTS idx_reservations_created ON reservations(created_at DESC);

            CREATE TABLE IF NOT EXISTS reservation_alerts(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sync_event_uuid TEXT NOT NULL UNIQUE,
                request_uuid TEXT,
                customer TEXT NOT NULL,
                branch_device TEXT,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                qty INTEGER NOT NULL,
                hold_until TEXT,
                note TEXT,
                created_at TEXT NOT NULL,
                shown_at TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_reservation_alerts_shown
            ON reservation_alerts(shown_at, created_at DESC);

            CREATE TABLE IF NOT EXISTS incoming_shipment_items_pending(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                shipment_uuid TEXT NOT NULL,
                line_index INTEGER NOT NULL,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                unit_price REAL NOT NULL,
                expected_qty INTEGER NOT NULL,
                received_qty INTEGER,
                status TEXT NOT NULL DEFAULT 'PENDING',
                UNIQUE(shipment_uuid, line_index)
            );
            CREATE INDEX IF NOT EXISTS idx_incoming_ship_items_shipment
            ON incoming_shipment_items_pending(shipment_uuid, status);

            CREATE TABLE IF NOT EXISTS incoming_shipment_alerts(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sync_event_uuid TEXT NOT NULL UNIQUE,
                shipment_uuid TEXT NOT NULL,
                from_device TEXT,
                note TEXT,
                total_qty INTEGER NOT NULL DEFAULT 0,
                created_at TEXT NOT NULL,
                shown_at TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_incoming_ship_alerts_shown
            ON incoming_shipment_alerts(shown_at, created_at DESC);

            CREATE TABLE IF NOT EXISTS shifts(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                started_at TEXT NOT NULL,
                ended_at TEXT,
                status TEXT NOT NULL DEFAULT 'OPEN',
                summary_json TEXT,
                note TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_shifts_status ON shifts(status);
            CREATE INDEX IF NOT EXISTS idx_shifts_started ON shifts(started_at DESC);

            CREATE TABLE IF NOT EXISTS expenses(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                amount REAL NOT NULL,
                note TEXT,
                shift_id INTEGER NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_expenses_shift_id ON expenses(shift_id);
            CREATE INDEX IF NOT EXISTS idx_expenses_created ON expenses(created_at DESC);

            CREATE TABLE IF NOT EXISTS income_bills(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                supplier TEXT,
                total_qty INTEGER NOT NULL DEFAULT 0,
                total_value REAL NOT NULL DEFAULT 0,
                note TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_income_bills_created ON income_bills(created_at DESC);

            CREATE TABLE IF NOT EXISTS income_bill_items(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                income_bill_id INTEGER NOT NULL,
                item_type TEXT NOT NULL,
                school TEXT NOT NULL,
                color TEXT NOT NULL,
                size TEXT NOT NULL,
                unit_price REAL NOT NULL,
                qty INTEGER NOT NULL,
                line_total REAL NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_income_bill_items_bill ON income_bill_items(income_bill_id);

            CREATE TABLE IF NOT EXISTS app_settings(
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
            """
        )

        try:
            new_admin = f"{ADMIN_PASSWORD_HASH_PREFIX}{hashlib.sha256(ADMIN_PASSWORD_PLAIN.encode('utf-8')).hexdigest()}"
            old_admin = f"{ADMIN_PASSWORD_HASH_PREFIX}{hashlib.sha256('1234'.encode('utf-8')).hexdigest()}"
            row = self.conn.execute(
                "SELECT value FROM app_settings WHERE key = 'admin_password'"
            ).fetchone()
            stored = str(row["value"]) if row else ""
            if not stored or stored in ("1234", old_admin):
                self.conn.execute(
                    "INSERT OR REPLACE INTO app_settings(key, value) VALUES('admin_password', ?)",
                    (new_admin,),
                )
        except Exception:
            pass

        # ensure 'origin' exists
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bill_items)")}
            if "origin" not in cols:
                self.conn.execute(
                    "ALTER TABLE bill_items ADD COLUMN origin TEXT NOT NULL DEFAULT 'STOCK'"
                )
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(stock_audit_reports)")}
            if cols and "reason" not in cols:
                self.conn.execute("ALTER TABLE stock_audit_reports ADD COLUMN reason TEXT")
            if cols and "diff_count" not in cols:
                self.conn.execute("ALTER TABLE stock_audit_reports ADD COLUMN diff_count INTEGER NOT NULL DEFAULT 0")
            if cols and "total_diff" not in cols:
                self.conn.execute("ALTER TABLE stock_audit_reports ADD COLUMN total_diff INTEGER NOT NULL DEFAULT 0")
            if cols and "total_value" not in cols:
                self.conn.execute("ALTER TABLE stock_audit_reports ADD COLUMN total_value REAL NOT NULL DEFAULT 0")
            if cols and "bucket_key" not in cols:
                self.conn.execute("ALTER TABLE stock_audit_reports ADD COLUMN bucket_key TEXT")
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_stock_audit_reports_bucket "
                "ON stock_audit_reports(reason, bucket_key)"
            )
        except Exception:
            pass

        # Migration: bill_type column on bills table
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bills)")}
            if "bill_type" not in cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN bill_type TEXT NOT NULL DEFAULT 'SALE'")
            if "payment_method" not in cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass

        # Migration: payment_method on POS money movements/reservations.
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if cols and "payment_method" not in cols:
                self.conn.execute("ALTER TABLE movements ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(reservations)")}
            if cols and "payment_method" not in cols:
                self.conn.execute("ALTER TABLE reservations ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass

        # Migration: immutable/voidable finalized bills
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bills)")}
            if "status" not in cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN status TEXT NOT NULL DEFAULT 'CONFIRMED'")
            if "void_reason" not in cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN void_reason TEXT")
            if "voided_at" not in cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN voided_at TEXT")
        except Exception:
            pass

        # Migration: remove warehouse_no/package_no from existing tables
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(stocks)")}
            if "warehouse_no" in cols:
                self.conn.executescript("""
                    CREATE TABLE IF NOT EXISTS stocks_new(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        item_type TEXT NOT NULL,
                        school TEXT NOT NULL,
                        color TEXT NOT NULL,
                        size TEXT NOT NULL,
                        unit_price REAL NOT NULL,
                        count INTEGER NOT NULL
                    );
                    INSERT INTO stocks_new(item_type, school, color, size, unit_price, count)
                        SELECT item_type, school, color, size, unit_price, SUM(count)
                        FROM stocks GROUP BY item_type, school, color, size, unit_price;
                    DROP TABLE stocks;
                    ALTER TABLE stocks_new RENAME TO stocks;
                    CREATE INDEX IF NOT EXISTS idx_stocks_specs ON stocks(item_type, school, color, size);
                    CREATE INDEX IF NOT EXISTS idx_stocks_count ON stocks(count);
                """)
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if "warehouse_no" in cols:
                self.conn.executescript("""
                    CREATE TABLE IF NOT EXISTS movements_new(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        ts TEXT NOT NULL,
                        direction TEXT NOT NULL,
                        stock_id INTEGER,
                        qty INTEGER NOT NULL,
                        note TEXT,
                        bill_id INTEGER,
                        item_type TEXT,
                        school TEXT,
                        color TEXT,
                        size TEXT,
                        unit_price REAL
                    );
                    INSERT INTO movements_new(ts, direction, stock_id, qty, note, bill_id, item_type, school, color, size, unit_price)
                        SELECT ts, direction, stock_id, qty, note, bill_id, item_type, school, color, size, unit_price
                        FROM movements;
                    DROP TABLE movements;
                    ALTER TABLE movements_new RENAME TO movements;
                    CREATE INDEX IF NOT EXISTS idx_movements_ts ON movements(ts DESC);
                    CREATE INDEX IF NOT EXISTS idx_movements_dir ON movements(direction);
                    CREATE INDEX IF NOT EXISTS idx_movements_specs ON movements(item_type,school,color,size);
                """)
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bill_items)")}
            if "warehouse_no" in cols:
                self.conn.executescript("""
                    CREATE TABLE IF NOT EXISTS bill_items_new(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        bill_id INTEGER NOT NULL,
                        item_type TEXT NOT NULL,
                        school TEXT NOT NULL,
                        color TEXT NOT NULL,
                        size TEXT NOT NULL,
                        unit_price REAL NOT NULL,
                        qty INTEGER NOT NULL,
                        line_total REAL NOT NULL,
                        origin TEXT NOT NULL DEFAULT 'STOCK'
                    );
                    INSERT INTO bill_items_new(bill_id, item_type, school, color, size, unit_price, qty, line_total, origin)
                        SELECT bill_id, item_type, school, color, size, unit_price, qty, line_total, COALESCE(origin,'STOCK')
                        FROM bill_items;
                    DROP TABLE bill_items;
                    ALTER TABLE bill_items_new RENAME TO bill_items;
                    CREATE INDEX IF NOT EXISTS idx_bill_items_bill ON bill_items(bill_id);
                """)
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if cols and "payment_method" not in cols:
                self.conn.execute("ALTER TABLE movements ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(reservations)")}
            if cols and "payment_method" not in cols:
                self.conn.execute("ALTER TABLE reservations ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass

        # Drop packages table (no longer needed)
        try:
            self.conn.execute("DROP TABLE IF EXISTS packages")
        except Exception:
            pass

        # Migration: remove has_badge
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(stocks)")}
            if "has_badge" in cols:
                self.conn.executescript("""
                    CREATE TABLE IF NOT EXISTS stocks_new2(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        item_type TEXT NOT NULL, school TEXT NOT NULL, color TEXT NOT NULL,
                        size TEXT NOT NULL, unit_price REAL NOT NULL, count INTEGER NOT NULL
                    );
                    INSERT INTO stocks_new2(item_type, school, color, size, unit_price, count)
                        SELECT item_type, school, color, size, unit_price, SUM(count)
                        FROM stocks GROUP BY item_type, school, color, size, unit_price;
                    DROP TABLE stocks;
                    ALTER TABLE stocks_new2 RENAME TO stocks;
                    CREATE INDEX IF NOT EXISTS idx_stocks_specs ON stocks(item_type, school, color, size);
                    CREATE INDEX IF NOT EXISTS idx_stocks_count ON stocks(count);
                """)
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if "has_badge" in cols:
                self.conn.executescript("""
                    CREATE TABLE IF NOT EXISTS movements_new2(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        ts TEXT NOT NULL, direction TEXT NOT NULL, stock_id INTEGER,
                        qty INTEGER NOT NULL, note TEXT, bill_id INTEGER,
                        item_type TEXT, school TEXT, color TEXT, size TEXT, unit_price REAL
                    );
                    INSERT INTO movements_new2(ts, direction, stock_id, qty, note, bill_id, item_type, school, color, size, unit_price)
                        SELECT ts, direction, stock_id, qty, note, bill_id, item_type, school, color, size, unit_price
                        FROM movements;
                    DROP TABLE movements;
                    ALTER TABLE movements_new2 RENAME TO movements;
                    CREATE INDEX IF NOT EXISTS idx_movements_ts ON movements(ts DESC);
                    CREATE INDEX IF NOT EXISTS idx_movements_dir ON movements(direction);
                    CREATE INDEX IF NOT EXISTS idx_movements_specs ON movements(item_type,school,color,size);
                """)
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if cols and "payment_method" not in cols:
                self.conn.execute("ALTER TABLE movements ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(reservations)")}
            if cols and "payment_method" not in cols:
                self.conn.execute("ALTER TABLE reservations ADD COLUMN payment_method TEXT NOT NULL DEFAULT 'CASH'")
        except Exception:
            pass

        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bill_items)")}
            if "has_badge" in cols:
                self.conn.executescript("""
                    CREATE TABLE IF NOT EXISTS bill_items_new2(
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        bill_id INTEGER NOT NULL, item_type TEXT NOT NULL, school TEXT NOT NULL,
                        color TEXT NOT NULL, size TEXT NOT NULL, unit_price REAL NOT NULL,
                        qty INTEGER NOT NULL, line_total REAL NOT NULL,
                        origin TEXT NOT NULL DEFAULT 'STOCK'
                    );
                    INSERT INTO bill_items_new2(bill_id, item_type, school, color, size, unit_price, qty, line_total, origin)
                        SELECT bill_id, item_type, school, color, size, unit_price, qty, line_total, COALESCE(origin,'STOCK')
                        FROM bill_items;
                    DROP TABLE bill_items;
                    ALTER TABLE bill_items_new2 RENAME TO bill_items;
                    CREATE INDEX IF NOT EXISTS idx_bill_items_bill ON bill_items(bill_id);
                """)
        except Exception:
            pass

        # Optional customer phone/number for bills and reservations.
        try:
            bill_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(bills)")}
            if "customer_phone" not in bill_cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN customer_phone TEXT")
            if "shift_id" not in bill_cols:
                self.conn.execute("ALTER TABLE bills ADD COLUMN shift_id INTEGER")
            res_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(reservations)")}
            if "customer_phone" not in res_cols:
                self.conn.execute("ALTER TABLE reservations ADD COLUMN customer_phone TEXT")
            if "shift_id" not in res_cols:
                self.conn.execute("ALTER TABLE reservations ADD COLUMN shift_id INTEGER")
            mov_cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(movements)")}
            if "shift_id" not in mov_cols:
                self.conn.execute("ALTER TABLE movements ADD COLUMN shift_id INTEGER")
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_bills_customer_phone_norm "
                "ON bills(LOWER(TRIM(customer_phone)))"
            )
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_bills_shift_id ON bills(shift_id)"
            )
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_movements_shift_id ON movements(shift_id)"
            )
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_reservations_shift_id ON reservations(shift_id)"
            )
            self.conn.commit()
        except Exception:
            pass

        # Migrate reservation statuses from English to Arabic
        try:
            self.conn.execute("UPDATE reservations SET status='معلق' WHERE status='PENDING'")
            self.conn.execute("UPDATE reservations SET status='تم التسليم' WHERE status='COMPLETED'")
            self.conn.commit()
        except Exception:
            pass

        # Reservation bill grouping for partial delivery flow.
        try:
            cols = {r["name"] for r in self.conn.execute("PRAGMA table_info(reservations)")}
            if "reservation_group_uuid" not in cols:
                self.conn.execute("ALTER TABLE reservations ADD COLUMN reservation_group_uuid TEXT")
            self.conn.execute(
                "UPDATE reservations "
                "SET reservation_group_uuid = COALESCE(reservation_group_uuid, 'legacy-' || id) "
                "WHERE COALESCE(TRIM(reservation_group_uuid), '') = ''"
            )
            self.conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_reservations_group ON reservations(reservation_group_uuid, status)"
            )
            self.conn.commit()
        except Exception:
            pass

        try:
            self._repair_legacy_reservation_down_payment_allocations()
        except Exception:
            import traceback
            traceback.print_exc()

        # ------------------- Sync layer (Phase 1) -------------------
        # Additive: creates sync_outbox/inbox/state/device_identity tables
        # and backfills a `uuid` column on every syncable domain table.
        # Failures here must never break the main app.
        try:
            from sync_core import apply_sync_migration, ensure_device_identity
            apply_sync_migration(self.conn)
            ensure_device_identity(
                self.conn,
                default_name="POS-UNCONFIGURED",
                default_role="pos",
            )
            # Performance indexes for common filtered reads (safe/additive).
            self.conn.executescript("""
                CREATE INDEX IF NOT EXISTS idx_stocks_specs_norm
                    ON stocks(LOWER(TRIM(item_type)), LOWER(TRIM(school)), LOWER(TRIM(color)), LOWER(TRIM(size)));
                CREATE INDEX IF NOT EXISTS idx_stocks_specs_price_count
                    ON stocks(item_type, school, color, size, unit_price, count);
                CREATE INDEX IF NOT EXISTS idx_stocks_low_count_specs
                    ON stocks(count, item_type, school, color, size);
                CREATE INDEX IF NOT EXISTS idx_bills_customer_norm
                    ON bills(LOWER(TRIM(customer)));
                CREATE INDEX IF NOT EXISTS idx_bills_status_created
                    ON bills(status, created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_bills_type_created
                    ON bills(bill_type, created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_movements_bill_id
                    ON movements(bill_id);
                CREATE INDEX IF NOT EXISTS idx_movements_dir_ts_specs
                    ON movements(direction, ts, school, item_type, color);
                CREATE INDEX IF NOT EXISTS idx_bill_items_specs
                    ON bill_items(item_type, school, color, size);
                CREATE INDEX IF NOT EXISTS idx_reservations_status_created
                    ON reservations(status, created_at DESC);
                CREATE INDEX IF NOT EXISTS idx_reservations_specs_norm
                    ON reservations(LOWER(TRIM(item_type)), LOWER(TRIM(school)), LOWER(TRIM(color)), LOWER(TRIM(size)), status);
                CREATE INDEX IF NOT EXISTS idx_sync_outbox_status_seq
                    ON sync_outbox(status, local_seq);
                CREATE INDEX IF NOT EXISTS idx_sync_inbox_status_seq
                    ON sync_inbox(apply_status, server_seq);
            """)
        except Exception:
            import traceback
            traceback.print_exc()

        try:
            self._backfill_legacy_shift_ids()
            self._close_empty_impossible_open_shifts()
        except Exception:
            import traceback
            traceback.print_exc()

        cur.close()

    def _repair_legacy_reservation_down_payment_allocations(self) -> Dict[str, int]:
        repaired_groups = 0
        updated_rows = 0
        rewritten_movements = 0

        groups = self.conn.execute(
            """
            SELECT reservation_group_uuid
            FROM reservations
            WHERE COALESCE(TRIM(reservation_group_uuid), '') <> ''
            GROUP BY reservation_group_uuid
            HAVING COUNT(*) > 1
               AND SUM(CASE WHEN status='تم التسليم' THEN 1 ELSE 0 END) = 0
               AND SUM(COALESCE(paid_amount, 0)) > 0
               AND (
                    SUM(CASE WHEN COALESCE(paid_amount, 0) <= 0.000001 THEN 1 ELSE 0 END) > 0
                    OR SUM(CASE WHEN COALESCE(paid_amount, 0) - COALESCE(total_amount, 0) > 0.001 THEN 1 ELSE 0 END) > 0
                    OR ABS(SUM(COALESCE(paid_amount, 0)) - SUM(COALESCE(total_amount, 0))) <= 0.01
               )
            """
        ).fetchall()

        with self.conn:
            for g in groups:
                group_uuid = str(g["reservation_group_uuid"] or "").strip()
                rows = self.conn.execute(
                    """
                    SELECT *
                    FROM reservations
                    WHERE reservation_group_uuid=?
                    ORDER BY id ASC
                    """,
                    (group_uuid,),
                ).fetchall()
                if len(rows) <= 1:
                    continue

                totals = [float(r["total_amount"] or 0.0) for r in rows]
                total_paid = round(sum(float(r["paid_amount"] or 0.0) for r in rows), 2)
                if total_paid <= 1e-9 or total_paid - sum(totals) > 1e-6:
                    continue

                current = [round(float(r["paid_amount"] or 0.0), 2) for r in rows]
                allocations = _allocate_reservation_down_payments(totals, total_paid)
                if len(allocations) != len(rows):
                    continue
                if all(abs(current[idx] - allocations[idx]) <= 0.001 for idx in range(len(rows))):
                    continue

                has_zero_sibling = any(v <= 0.001 for v in current)
                has_overpaid_row = any(
                    current[idx] - float(totals[idx] or 0.0) > 0.001
                    for idx in range(len(rows))
                )
                is_fully_paid = abs(total_paid - sum(totals)) <= 0.01
                legacy_single_carrier = max(current) >= total_paid - 0.001
                if not (has_overpaid_row or is_fully_paid or (has_zero_sibling and legacy_single_carrier)):
                    continue

                repaired_groups += 1
                for idx, row in enumerate(rows):
                    rid = int(row["id"])
                    alloc = round(float(allocations[idx]), 2)
                    if abs(float(row["paid_amount"] or 0.0) - alloc) > 0.001:
                        self.conn.execute(
                            "UPDATE reservations SET paid_amount=? WHERE id=?",
                            (alloc, rid),
                        )
                        updated_rows += 1

                    moves = self.conn.execute(
                        """
                        SELECT *
                        FROM movements
                        WHERE direction='RESERVE_PAY'
                          AND note LIKE ?
                        ORDER BY id ASC
                        """,
                        (f"%#{rid}%",),
                    ).fetchall()
                    if alloc > 1e-9:
                        if moves:
                            first = moves[0]
                            self.conn.execute(
                                """
                                UPDATE movements
                                   SET unit_price=?,
                                       item_type=?,
                                       school=?,
                                       color=?,
                                       size=?,
                                       payment_method=COALESCE(payment_method, ?),
                                       shift_id=COALESCE(shift_id, ?)
                                 WHERE id=?
                                """,
                                (
                                    alloc,
                                    row["item_type"],
                                    row["school"],
                                    row["color"],
                                    row["size"],
                                    row["payment_method"] if "payment_method" in row.keys() else PAYMENT_METHOD_CASH,
                                    row["shift_id"] if "shift_id" in row.keys() else None,
                                    int(first["id"]),
                                ),
                            )
                            for extra in moves[1:]:
                                self.conn.execute("DELETE FROM movements WHERE id=?", (int(extra["id"]),))
                            rewritten_movements += 1
                        else:
                            self.conn.execute(
                                """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                                (
                                    row["created_at"] or now_iso(),
                                    "RESERVE_PAY",
                                    None,
                                    0,
                                    f"عربون حجز #{rid}",
                                    None,
                                    row["item_type"],
                                    row["school"],
                                    row["color"],
                                    row["size"],
                                    alloc,
                                    row["shift_id"] if "shift_id" in row.keys() else None,
                                    row["payment_method"] if "payment_method" in row.keys() else PAYMENT_METHOD_CASH,
                                ),
                            )
                            rewritten_movements += 1
                    else:
                        for move in moves:
                            self.conn.execute("DELETE FROM movements WHERE id=?", (int(move["id"]),))
                            rewritten_movements += 1

        return {
            "groups": repaired_groups,
            "rows": updated_rows,
            "movements": rewritten_movements,
        }

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
        values = (
            r1_start, r1_end,
            r2_start, r2_end,
            int(has_alpha),
        )
        cur = self.conn.execute(
            """
            UPDATE size_profiles
               SET num_start_1 = ?,
                   num_end_1   = ?,
                   num_start_2 = ?,
                   num_end_2   = ?,
                   has_alpha   = ?,
                   updated_at  = datetime('now')
             WHERE item_type = ?
               AND school = ?
               AND color = ?
            """,
            (*values, item_type, school, color),
        )
        if cur.rowcount == 0:
            self.conn.execute(
                """
                INSERT OR IGNORE INTO size_profiles (
                    item_type, school, color,
                    num_start_1, num_end_1,
                    num_start_2, num_end_2,
                    has_alpha, updated_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
                """,
                (item_type, school, color, *values),
            )

        self.conn.commit()

    def get_effective_price(self, item_type, school, color, size):
        """
        Priority:
        1) exact last price (history)
        2) default price for item
        3) None if unknown
        """
        try:
            p = self.last_price_for_specs(item_type, school, color, size)
            if p is not None:
                return float(p)
        except Exception:
            pass

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
            "INSERT OR IGNORE INTO item_defaults (item_type, default_price) VALUES (?, ?)",
            (item_type, float(price))
        )
        self.conn.commit()

    def purge_definition(self, field: str, value: str) -> int:
        if field not in ("item_type", "school", "color", "size"):
            raise ValueError("حقل غير مدعوم")

        value = (value or "").strip()
        if not value:
            return 0

        with self.conn:
            cur = self.conn.execute(
                f"SELECT COUNT(*) FROM stocks WHERE {field} = ?",
                (value,),
            )
            count = int(cur.fetchone()[0] or 0)

            self.conn.execute(
                f"DELETE FROM stocks WHERE {field} = ?",
                (value,),
            )

            self.conn.execute(
                "DELETE FROM spec_history WHERE field = ? AND value = ?",
                (field, value),
            )

        return count

    def list_schools_all(self) -> List[str]:
        cur = self.conn.cursor()
        try:
            cur.execute("""
                SELECT DISTINCT TRIM(school) AS s FROM stocks
                WHERE """ + self._hidden_definition_sql("school", "school") + """
                UNION
                SELECT DISTINCT TRIM(school) AS s FROM bill_items
                WHERE """ + self._hidden_definition_sql("school", "school") + """
                UNION
                SELECT DISTINCT TRIM(school) AS s FROM reservations
                WHERE """ + self._hidden_definition_sql("school", "school") + """
                UNION
                SELECT DISTINCT TRIM(school) AS s FROM movements
                WHERE """ + self._hidden_definition_sql("school", "school") + """
                UNION
                SELECT DISTINCT TRIM(value) AS s FROM spec_history
                WHERE field='school'
                  AND """ + self._hidden_definition_sql("school", "value") + """
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
        """Cash-flow by school/spec.

        This report intentionally tracks money movement, not only retail SALE
        bills. It includes sale cash/visa in, reservation deposits and delivery
        collections, returns, sale void refunds, and the paid/refunded
        difference on exchange bills.
        """
        cleaned = [str(s or "").strip() for s in schools if str(s or "").strip()]
        if not cleaned:
            return []
        schools_lower = [s.lower() for s in cleaned]
        placeholders = ",".join("?" for _ in schools_lower)
        buckets: Dict[Tuple[str, str, str, str], Dict[str, Any]] = {}

        def bucket(school: Any, item_type: Any, color: Any, size: Any) -> Dict[str, Any]:
            key = (
                str(school or "").strip(),
                str(item_type or "").strip(),
                str(color or "").strip(),
                str(size or "").strip(),
            )
            row = buckets.get(key)
            if row is None:
                row = {
                    "school": key[0],
                    "item_type": key[1],
                    "color": key[2],
                    "size": key[3],
                    "total_qty": 0,
                    "sales_total": 0.0,
                    "reservation_paid": 0.0,
                    "reservation_delivered": 0.0,
                    "return_total": 0.0,
                    "void_total": 0.0,
                    "exchange_in": 0.0,
                    "exchange_out": 0.0,
                    "cash_in": 0.0,
                    "cash_out": 0.0,
                    "visa_in": 0.0,
                    "visa_out": 0.0,
                    "net_total": 0.0,
                }
                buckets[key] = row
            return row

        def add_money(
            row: Dict[str, Any],
            amount: float,
            *,
            method: str = PAYMENT_METHOD_CASH,
            category: str = "",
            qty_delta: int = 0,
        ) -> None:
            amt = float(amount or 0.0)
            method_u = str(method or PAYMENT_METHOD_CASH).strip().upper()
            row["total_qty"] = int(row.get("total_qty") or 0) + int(qty_delta or 0)
            if category:
                row[category] = float(row.get(category) or 0.0) + abs(amt)
            if amt >= 0:
                if method_u == PAYMENT_METHOD_VISA:
                    row["visa_in"] += amt
                else:
                    row["cash_in"] += amt
            else:
                if method_u == PAYMENT_METHOD_VISA:
                    row["visa_out"] += abs(amt)
                else:
                    row["cash_out"] += abs(amt)
            row["net_total"] += amt

        def date_filter_sql(alias: str, column: str) -> Tuple[List[str], List[Any]]:
            parts: List[str] = []
            params: List[Any] = []
            col = f"{alias}.{column}" if alias else column
            if date_from:
                parts.append(f"date({col}) >= date(?)")
                params.append(date_from)
            if date_to:
                parts.append(f"date({col}) <= date(?)")
                params.append(date_to)
            return parts, params

        cur = self.conn.cursor()
        try:
            # Retail sales: canceled sales are not counted as normal sales.
            # The void itself is a separate cash-out event below.
            sale_where = [
                "(COALESCE(b.bill_type,'SALE')='SALE' OR b.bill_type IS NULL)",
                "UPPER(COALESCE(b.status,'CONFIRMED')) != 'VOID'",
                f"LOWER(TRIM(bi.school)) IN ({placeholders})",
            ]
            sale_args: List[Any] = list(schools_lower)
            extra, extra_args = date_filter_sql("b", "created_at")
            sale_where.extend(extra)
            sale_args.extend(extra_args)
            cur.execute(
                f"""
                SELECT
                    TRIM(bi.school) AS school,
                    TRIM(bi.item_type) AS item_type,
                    TRIM(bi.color) AS color,
                    TRIM(bi.size) AS size,
                    COALESCE(SUM(bi.qty), 0) AS total_qty,
                    COALESCE(SUM(bi.line_total), 0) AS total_amount,
                    COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method
                FROM bill_items bi
                JOIN bills b ON b.id = bi.bill_id
                WHERE {' AND '.join(sale_where)}
                GROUP BY TRIM(bi.school), TRIM(bi.item_type), TRIM(bi.color), TRIM(bi.size), COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}')
                """,
                sale_args,
            )
            for r in cur.fetchall():
                row = bucket(r["school"], r["item_type"], r["color"], r["size"])
                amount = float(r["total_amount"] or 0.0)
                add_money(
                    row,
                    amount,
                    method=r["payment_method"],
                    category="sales_total",
                    qty_delta=int(r["total_qty"] or 0),
                )

            # Reservation deposits, edits, delivery collections, and cancelled
            # reservation refunds are direct POS cash movements.
            mov_where = [
                "m.direction IN ('RESERVE_PAY','DELIVER_PAY','RESERVE_REFUND')",
                f"LOWER(TRIM(m.school)) IN ({placeholders})",
            ]
            mov_args: List[Any] = list(schools_lower)
            extra, extra_args = date_filter_sql("m", "ts")
            mov_where.extend(extra)
            mov_args.extend(extra_args)
            cur.execute(
                f"""
                SELECT
                    TRIM(m.school) AS school,
                    TRIM(m.item_type) AS item_type,
                    TRIM(m.color) AS color,
                    TRIM(m.size) AS size,
                    m.direction AS direction,
                    COALESCE(SUM(m.unit_price), 0) AS total_amount,
                    COALESCE(m.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method
                FROM movements m
                WHERE {' AND '.join(mov_where)}
                GROUP BY TRIM(m.school), TRIM(m.item_type), TRIM(m.color), TRIM(m.size), m.direction, COALESCE(m.payment_method, '{PAYMENT_METHOD_CASH}')
                """,
                mov_args,
            )
            for r in cur.fetchall():
                row = bucket(r["school"], r["item_type"], r["color"], r["size"])
                amount = float(r["total_amount"] or 0.0)
                direction = str(r["direction"]).upper()
                if direction == "DELIVER_PAY":
                    category = "reservation_delivered"
                    signed_amount = amount
                elif direction == "RESERVE_REFUND":
                    category = "reservation_paid"
                    signed_amount = -abs(amount)
                else:
                    category = "reservation_paid"
                    signed_amount = amount
                add_money(row, signed_amount, method=r["payment_method"], category=category)

            # Return bills are money out on the return date.
            ret_where = [
                "COALESCE(b.bill_type,'SALE')='RETURN'",
                "UPPER(COALESCE(b.status,'CONFIRMED'))='CONFIRMED'",
                f"LOWER(TRIM(bi.school)) IN ({placeholders})",
            ]
            ret_args: List[Any] = list(schools_lower)
            extra, extra_args = date_filter_sql("b", "created_at")
            ret_where.extend(extra)
            ret_args.extend(extra_args)
            cur.execute(
                f"""
                SELECT
                    TRIM(bi.school) AS school,
                    TRIM(bi.item_type) AS item_type,
                    TRIM(bi.color) AS color,
                    TRIM(bi.size) AS size,
                    COALESCE(SUM(bi.qty), 0) AS total_qty,
                    COALESCE(SUM(bi.line_total), 0) AS total_amount,
                    COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method
                FROM bill_items bi
                JOIN bills b ON b.id = bi.bill_id
                WHERE {' AND '.join(ret_where)}
                GROUP BY TRIM(bi.school), TRIM(bi.item_type), TRIM(bi.color), TRIM(bi.size), COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}')
                """,
                ret_args,
            )
            for r in cur.fetchall():
                row = bucket(r["school"], r["item_type"], r["color"], r["size"])
                add_money(
                    row,
                    -float(r["total_amount"] or 0.0),
                    method=r["payment_method"],
                    category="return_total",
                    qty_delta=int(r["total_qty"] or 0),
                )

            # Sale voids are refunds on the void date, attributed to original items.
            void_where = [
                "COALESCE(b.bill_type,'SALE')='SALE'",
                "UPPER(COALESCE(b.status,'CONFIRMED'))='VOID'",
                "COALESCE(TRIM(b.voided_at),'') <> ''",
                f"LOWER(TRIM(bi.school)) IN ({placeholders})",
            ]
            void_args: List[Any] = list(schools_lower)
            extra, extra_args = date_filter_sql("b", "voided_at")
            void_where.extend(extra)
            void_args.extend(extra_args)
            cur.execute(
                f"""
                SELECT
                    TRIM(bi.school) AS school,
                    TRIM(bi.item_type) AS item_type,
                    TRIM(bi.color) AS color,
                    TRIM(bi.size) AS size,
                    COALESCE(SUM(bi.line_total), 0) AS total_amount,
                    COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method
                FROM bill_items bi
                JOIN bills b ON b.id = bi.bill_id
                WHERE {' AND '.join(void_where)}
                GROUP BY TRIM(bi.school), TRIM(bi.item_type), TRIM(bi.color), TRIM(bi.size), COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}')
                """,
                void_args,
            )
            for r in cur.fetchall():
                row = bucket(r["school"], r["item_type"], r["color"], r["size"])
                add_money(
                    row,
                    -float(r["total_amount"] or 0.0),
                    method=r["payment_method"],
                    category="void_total",
                )

            # Exchange bills: only the net paid/refunded difference is real cash.
            exch_where = [
                "COALESCE(b.bill_type,'SALE')='EXCHANGE'",
                "UPPER(COALESCE(b.status,'CONFIRMED'))='CONFIRMED'",
            ]
            exch_args: List[Any] = []
            extra, extra_args = date_filter_sql("b", "created_at")
            exch_where.extend(extra)
            exch_args.extend(extra_args)
            cur.execute(
                f"""
                SELECT b.id, COALESCE(b.total, 0) AS total,
                       COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method
                FROM bills b
                WHERE {' AND '.join(exch_where)}
                  AND EXISTS (
                      SELECT 1 FROM bill_items bi
                      WHERE bi.bill_id = b.id
                        AND LOWER(TRIM(bi.school)) IN ({placeholders})
                  )
                """,
                exch_args + list(schools_lower),
            )
            exchange_bills = cur.fetchall()
            for b in exchange_bills:
                diff = float(b["total"] or 0.0)
                if abs(diff) <= 1e-9:
                    continue
                side = "STOCK" if diff > 0 else "RETURN"
                cur.execute(
                    f"""
                    SELECT
                        TRIM(school) AS school,
                        TRIM(item_type) AS item_type,
                        TRIM(color) AS color,
                        TRIM(size) AS size,
                        COALESCE(qty, 0) AS qty,
                        COALESCE(line_total, 0) AS line_total
                    FROM bill_items
                    WHERE bill_id=?
                      AND UPPER(COALESCE(origin,''))=?
                      AND LOWER(TRIM(school)) IN ({placeholders})
                    """,
                    (int(b["id"]), side, *schools_lower),
                )
                lines = [dict(x) for x in cur.fetchall()]
                denom = sum(abs(float(x.get("line_total") or 0.0)) for x in lines)
                if not lines or denom <= 1e-9:
                    continue
                for ln in lines:
                    line_total = abs(float(ln.get("line_total") or 0.0))
                    share = line_total / denom
                    amount = abs(diff) * share
                    row = bucket(ln["school"], ln["item_type"], ln["color"], ln["size"])
                    add_money(
                        row,
                        amount if diff > 0 else -amount,
                        method=b["payment_method"],
                        category="exchange_in" if diff > 0 else "exchange_out",
                        qty_delta=int(ln.get("qty") or 0),
                    )

            rows = list(buckets.values())
            rows = [
                r for r in rows
                if abs(float(r.get("cash_in") or 0.0))
                or abs(float(r.get("cash_out") or 0.0))
                or abs(float(r.get("visa_in") or 0.0))
                or abs(float(r.get("visa_out") or 0.0))
            ]
            rows.sort(key=lambda r: (
                str(r.get("school") or "").lower(),
                item_type_sort_key(r.get("item_type")),
                str(r.get("color") or "").lower(),
                str(r.get("size") or "").lower(),
            ))
            return rows
        finally:
            cur.close()

    def get_school_accounts_day_report(
        self,
        schools: Sequence[str],
        *,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Bill-level daily money report for selected schools.

        Rows are money events, not item/spec lines. Summing row cash and visa
        exactly equals the returned day totals.
        """
        cleaned = [str(s or "").strip() for s in schools if str(s or "").strip()]
        if not cleaned:
            return {"rows": [], "total_day": 0.0, "total_cash": 0.0, "total_visa": 0.0}
        schools_lower = [s.lower() for s in cleaned]
        placeholders = ",".join("?" for _ in schools_lower)
        rows: List[Dict[str, Any]] = []

        def date_parts(alias: str, column: str) -> Tuple[List[str], List[Any]]:
            col = f"{alias}.{column}" if alias else column
            parts: List[str] = []
            args: List[Any] = []
            if date_from:
                parts.append(f"date({col}) >= date(?)")
                args.append(date_from)
            if date_to:
                parts.append(f"date({col}) <= date(?)")
                args.append(date_to)
            return parts, args

        def split_amount(amount: float, method: Any) -> Tuple[float, float]:
            value = float(amount or 0.0)
            method_u = str(method or PAYMENT_METHOD_CASH).strip().upper()
            if method_u == PAYMENT_METHOD_VISA:
                return 0.0, value
            return value, 0.0

        def add_row(
            *,
            key: str,
            event_at: Any,
            bill_no: Any,
            bill_type: str,
            customer: Any,
            phone: Any = "",
            schools_text: Any,
            payment_method: Any,
            amount: float,
            status: Any = "",
        ) -> None:
            cash, visa = split_amount(amount, payment_method)
            rows.append({
                "key": key,
                "event_at": str(event_at or ""),
                "bill_no": str(bill_no or ""),
                "bill_type": bill_type,
                "customer": str(customer or ""),
                "customer_phone": str(phone or ""),
                "schools": str(schools_text or ""),
                "payment_method": str(payment_method or PAYMENT_METHOD_CASH).strip().upper() or PAYMENT_METHOD_CASH,
                "cash_total": cash,
                "visa_total": visa,
                "total_paid": cash + visa,
                "status": str(status or ""),
            })

        cur = self.conn.cursor()
        try:
            bill_base = [
                f"""
                EXISTS (
                    SELECT 1 FROM bill_items bi
                     WHERE bi.bill_id = b.id
                       AND LOWER(TRIM(bi.school)) IN ({placeholders})
                )
                """
            ]
            bill_base_args: List[Any] = list(schools_lower)

            bill_date, bill_date_args = date_parts("b", "created_at")
            bill_where = bill_base + bill_date
            cur.execute(
                f"""
                SELECT
                    b.id,
                    b.created_at,
                    COALESCE(b.customer, '') AS customer,
                    COALESCE(b.customer_phone, '') AS customer_phone,
                    COALESCE(b.total, 0) AS total,
                    COALESCE(b.bill_type, 'SALE') AS bill_type,
                    COALESCE(b.status, 'CONFIRMED') AS status,
                    COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method,
                    (
                        SELECT GROUP_CONCAT(school_name)
                          FROM (
                              SELECT DISTINCT TRIM(bi2.school) AS school_name
                                FROM bill_items bi2
                               WHERE bi2.bill_id = b.id
                                 AND LOWER(TRIM(bi2.school)) IN ({placeholders})
                               ORDER BY school_name
                          )
                    ) AS schools_text
                FROM bills b
                WHERE {' AND '.join(bill_where)}
                  AND COALESCE(b.bill_type, 'SALE') IN ('SALE', 'RETURN', 'EXCHANGE')
                  AND UPPER(COALESCE(b.status, 'CONFIRMED')) != 'VOID'
                ORDER BY b.created_at DESC, b.id DESC
                """,
                list(schools_lower) + bill_base_args + bill_date_args,
            )
            for b in cur.fetchall():
                bt = str(b["bill_type"] or "SALE").upper()
                total = float(b["total"] or 0.0)
                if bt == "RETURN":
                    amount = -abs(total)
                elif bt == "EXCHANGE":
                    amount = total
                else:
                    amount = total
                add_row(
                    key=f"bill:{bt}:{b['id']}",
                    event_at=b["created_at"],
                    bill_no=b["id"],
                    bill_type=bt,
                    customer=b["customer"],
                    phone=b["customer_phone"],
                    schools_text=b["schools_text"],
                    payment_method=b["payment_method"],
                    amount=amount,
                    status=b["status"],
                )

            void_date, void_date_args = date_parts("b", "voided_at")
            void_where = bill_base + [
                "COALESCE(b.bill_type, 'SALE') = 'SALE'",
                "UPPER(COALESCE(b.status, 'CONFIRMED')) = 'VOID'",
                "COALESCE(TRIM(b.voided_at), '') <> ''",
            ] + void_date
            cur.execute(
                f"""
                SELECT
                    b.id,
                    b.voided_at AS event_at,
                    COALESCE(b.customer, '') AS customer,
                    COALESCE(b.customer_phone, '') AS customer_phone,
                    COALESCE(b.total, 0) AS total,
                    COALESCE(b.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method,
                    (
                        SELECT GROUP_CONCAT(school_name)
                          FROM (
                              SELECT DISTINCT TRIM(bi2.school) AS school_name
                                FROM bill_items bi2
                               WHERE bi2.bill_id = b.id
                                 AND LOWER(TRIM(bi2.school)) IN ({placeholders})
                               ORDER BY school_name
                          )
                    ) AS schools_text
                FROM bills b
                WHERE {' AND '.join(void_where)}
                ORDER BY b.voided_at DESC, b.id DESC
                """,
                list(schools_lower) + bill_base_args + void_date_args,
            )
            for b in cur.fetchall():
                add_row(
                    key=f"bill:VOID:{b['id']}",
                    event_at=b["event_at"],
                    bill_no=b["id"],
                    bill_type="VOID",
                    customer=b["customer"],
                    phone=b["customer_phone"],
                    schools_text=b["schools_text"],
                    payment_method=b["payment_method"],
                    amount=-abs(float(b["total"] or 0.0)),
                    status="VOID",
                )

            movement_date, movement_date_args = date_parts("m", "ts")
            movement_where = [
                "m.direction IN ('RESERVE_PAY', 'DELIVER_PAY', 'RESERVE_REFUND')",
                f"LOWER(TRIM(m.school)) IN ({placeholders})",
            ] + movement_date
            cur.execute(
                f"""
                SELECT
                    m.id,
                    m.ts,
                    m.direction,
                    COALESCE(m.note, '') AS note,
                    TRIM(m.school) AS school,
                    COALESCE(m.unit_price, 0) AS amount,
                    COALESCE(m.payment_method, '{PAYMENT_METHOD_CASH}') AS payment_method,
                    COALESCE((
                        SELECT MAX(NULLIF(TRIM(r.customer), ''))
                          FROM reservations r
                         WHERE LOWER(TRIM(r.school)) = LOWER(TRIM(m.school))
                           AND LOWER(TRIM(r.item_type)) = LOWER(TRIM(m.item_type))
                           AND LOWER(TRIM(r.color)) = LOWER(TRIM(m.color))
                           AND LOWER(TRIM(r.size)) = LOWER(TRIM(m.size))
                           AND (m.note LIKE '%' || '#' || r.id || '%' OR date(r.created_at) <= date(m.ts))
                    ), '') AS customer,
                    COALESCE((
                        SELECT MAX(NULLIF(TRIM(r.customer_phone), ''))
                          FROM reservations r
                         WHERE LOWER(TRIM(r.school)) = LOWER(TRIM(m.school))
                           AND LOWER(TRIM(r.item_type)) = LOWER(TRIM(m.item_type))
                           AND LOWER(TRIM(r.color)) = LOWER(TRIM(m.color))
                           AND LOWER(TRIM(r.size)) = LOWER(TRIM(m.size))
                           AND (m.note LIKE '%' || '#' || r.id || '%' OR date(r.created_at) <= date(m.ts))
                    ), '') AS customer_phone
                FROM movements m
                WHERE {' AND '.join(movement_where)}
                ORDER BY m.ts DESC, m.id DESC
                """,
                list(schools_lower) + movement_date_args,
            )
            reservation_labels = {
                "RESERVE_PAY": "RESERVATION",
                "DELIVER_PAY": "RESERVATION_DELIVERY",
                "RESERVE_REFUND": "RESERVATION_REFUND",
            }
            for m in cur.fetchall():
                direction = str(m["direction"] or "").upper()
                amount = float(m["amount"] or 0.0)
                if direction == "RESERVE_REFUND":
                    amount = -abs(amount)
                bill_no = str(m["note"] or "").strip() or str(m["id"])
                add_row(
                    key=f"movement:{m['id']}",
                    event_at=m["ts"],
                    bill_no=bill_no,
                    bill_type=reservation_labels.get(direction, direction),
                    customer=m["customer"],
                    phone=m["customer_phone"],
                    schools_text=m["school"],
                    payment_method=m["payment_method"],
                    amount=amount,
                    status="",
                )

            rows.sort(key=lambda r: (str(r.get("event_at") or ""), str(r.get("key") or "")), reverse=True)
            total_cash = sum(float(r.get("cash_total") or 0.0) for r in rows)
            total_visa = sum(float(r.get("visa_total") or 0.0) for r in rows)
            return {
                "rows": rows,
                "total_day": total_cash + total_visa,
                "total_cash": total_cash,
                "total_visa": total_visa,
            }
        finally:
            cur.close()

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
            rows.sort(key=lambda t: (*item_type_sort_key(t[0]), t[1].casefold()))
            return rows
        finally:
            cur.close()

    def _size_row(self, school, item_type, color, size):
        """Return dict with size, current count, and last price for a specific size."""
        size = _normalize_size_label(size)
        cur = self.conn.execute(
            """
            SELECT COALESCE(SUM(count),0) AS c
              FROM stocks
             WHERE LOWER(TRIM(item_type))=LOWER(?)
               AND LOWER(TRIM(school))=LOWER(?)
               AND LOWER(TRIM(color))=LOWER(?)
               AND LOWER(TRIM(REPLACE(REPLACE(size, char(8206), ''), char(8207), '')))=LOWER(?)
            """,
            (item_type, school, color, size),
        )
        count = int(cur.fetchone()["c"] or 0)
        last_price = self.get_effective_price(item_type, school, color, size)
        return {"size": size, "count": count, "last_price": last_price}

    def list_sizes_for_item(self, school, item_type, color):
        ranges = self._get_size_ranges(item_type, school, color)
        out = []
        seen: Set[str] = set()

        def _add_size(sz: Any) -> None:
            label = _normalize_size_label(sz)
            if not label:
                return
            key = label.casefold()
            if key in seen:
                return
            seen.add(key)
            out.append(self._size_row(school, item_type, color, label))

        for r in ranges:
            if r["range_type"] == "NUMERIC":
                labels = ALLOWED_NUMERIC_RANGES.get((r["start"], r["end"]))
                if labels is None:
                    labels = [str(sz) for sz in range(r["start"], r["end"] + 1)]
                for sz in labels:
                    _add_size(str(sz))

            elif r["range_type"] == "ALPHA":
                for sz in r["alpha_set"]:
                    _add_size(sz)

        try:
            rows = self.current_inventory({"school": school, "item_type": item_type, "color": color})
        except Exception:
            rows = []
        for row in rows:
            _add_size(row.get("size"))

        def _sort_size_row(row: Dict[str, Any]):
            sz = _normalize_size_label(row.get("size"))
            if sz.isdigit():
                return (0, int(sz), "")
            up = sz.upper()
            if up in ALPHA_SIZES:
                return (1, ALPHA_SIZES.index(up), "")
            return (2, 9999, sz.casefold())

        out.sort(key=_sort_size_row)
        return out

    def last_price_for_specs(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
    ) -> Optional[float]:
        it = (item_type or "").strip()
        sc = (school or "").strip()
        cl = (color or "").strip()
        sz_raw = (size or "").strip()

        if not (it and sc and cl and sz_raw):
            return None

        sz_norm = _normalize_size_label(sz_raw)
        cur = self.conn.cursor()

        try:
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
        `reservations` so audit/history views stay consistent.
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

            for tbl in ("movements", "bill_items", "reservations"):
                try:
                    self.conn.execute(
                        f"UPDATE {tbl} SET {set_sql} WHERE {where_sql}",
                        (*new_vals, *where_args),
                    )
                except sqlite3.OperationalError:
                    pass

    def get_distinct_filtered(self, target: str, constraints: Dict[str, Any], *, available_only: bool = False) -> List[str]:
        valid = {"item_type", "school", "color", "size"}
        if target not in valid:
            return []

        where: List[str] = ["count > 0"] if available_only else ["1=1"]
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

        cur = self.conn.cursor()
        try:
            cur.execute(
                f"SELECT DISTINCT TRIM({target}) AS v FROM stocks WHERE {' AND '.join(where)}"
                + f" AND {self._hidden_definition_sql(target, target)}"
                f" ORDER BY LOWER(TRIM({target})) ASC",
                args,
            )

            values = [r["v"] for r in cur.fetchall() if r["v"] not in (None, "")]
            return sort_item_type_values(values) if target == "item_type" else values
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

    def update_prices(self, filters: Dict[str, Any], new_price: float, note: str = "Price update") -> int:
        if not self.is_manager_feature_enabled("allow_inventory_price_edit"):
            raise PermissionError(_feature_restricted_message("تعديل الأسعار من نقطة البيع غير مسموح به حالياً."))
        if new_price is None or float(new_price) < 0:
            raise ValueError("New price must be a non-negative number.")
        where: List[str] = []
        args: List[Any] = []

        if "id" in filters and filters["id"]:
            where.append("id = ?")
            args.append(int(filters["id"]))
        else:
            for k in ("item_type", "school", "color", "size"):
                v_raw = (filters.get(k) or "").strip()
                if not v_raw:
                    continue
                if k == "size":
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
                    (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        ts, "PRICE_UPDATE", int(r["id"]), 0, note, None,
                        r["item_type"], r["school"], r["color"], r["size"],
                        float(new_price),
                    ),
                )
            self._record_sync_event("PRICE_UPDATE", {
                "mode":    "fixed",
                "new_price": float(new_price),
                "filters": {k: v for k, v in (filters or {}).items() if v not in (None, "", [])},
                "updated_count": len(rows),
                "scope":   "filter",
                "note":    note,
            })
        cur.close()
        return len(rows)

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
                        """INSERT INTO stocks(id,item_type,school,color,size,unit_price,count)
                           VALUES(?,?,?,?,?,?,?)""",
                        (
                            int(s["id"]), s["item_type"], s["school"], s["color"], s["size"],
                            float(s["unit_price"]), int(s["count"]),
                        ),
                    )
                for m in data.get("movements", []):
                    self.conn.execute(
                        """INSERT INTO movements
                           (id,ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            int(m["id"]), m["ts"], m["direction"], m.get("stock_id"), int(m["qty"]),
                            m.get("note"), m.get("bill_id"), m.get("item_type"), m.get("school"),
                            m.get("color"), m.get("size"),
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
                           (id,bill_id,item_type,school,color,size,unit_price,qty,line_total)
                           VALUES(?,?,?,?,?,?,?,?,?)""",
                        (
                            int(bi["id"]), int(bi["bill_id"]), bi["item_type"], bi["school"], bi["color"], bi["size"],
                            float(bi["unit_price"]),
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
                  AND {self._hidden_definition_sql(field, field)}
                ORDER BY LOWER(v)
            """)
            values = [r["v"] for r in cur.fetchall() if r["v"]]
            return sort_item_type_values(values) if field == "item_type" else values
        finally:
            cur.close()

    # -------- Stocks --------
    def add_stock(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
        unit_price: Optional[float],
        count: int,
    ) -> int:
        self._require_shift()

        if count <= 0:
            raise ValueError("Count must be > 0")

        price: Optional[float] = None
        user_provided_price = unit_price is not None and str(unit_price) != ""

        if user_provided_price:
            price = float(unit_price)
            if price < 0:
                raise ValueError("Price must be >= 0")
        else:
            price = self.get_effective_price(item_type, school, color, size)

        if price is None:
            raise ValueError("يجب إدخال السعر مرة واحدة على الأقل لهذا الصنف.")

        if user_provided_price:
            self.ensure_default_price(item_type, price)

        with self.conn:
            stock_id = self.add_or_update_stock_row(
                item_type.strip(),
                school.strip(),
                color.strip(),
                size.strip(),
                float(price),
                int(count),
            )

            self.conn.execute(
                """INSERT INTO movements
                (ts,direction,stock_id,qty,note,bill_id,
                    item_type,school,color,size,unit_price)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
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
                    float(price),
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

        return int(stock_id)

    def cleanup_unused_specs(self):
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

    def repair_unsafe_spec_value_rename_damage(self) -> Dict[str, int]:
        """Undo old broad SPEC_RENAMED value updates using shipment evidence."""
        fields = ("item_type", "school", "color", "size")

        def _clean(value: Any) -> str:
            return str(value or "").strip()

        def _key(spec: Dict[str, str]) -> Tuple[str, str, str, str]:
            return tuple(_clean(spec.get(k)).casefold() for k in fields)

        try:
            rename_rows = self.conn.execute(
                """
                SELECT payload_json
                  FROM sync_inbox
                 WHERE event_type='SPEC_RENAMED'
                   AND payload_json LIKE '%value_renames%'
                """
            ).fetchall()
        except sqlite3.OperationalError:
            return {"rename_events": 0, "candidate_specs": 0, "updated_rows": 0, "deleted_zero_duplicates": 0}

        unsafe_renames: List[Dict[str, str]] = []
        protected_specs: Set[Tuple[str, str, str, str]] = set()
        for row in rename_rows:
            try:
                payload = json.loads(row["payload_json"] or "{}")
            except Exception:
                continue
            if bool(payload.get("allow_global_value_renames")):
                continue
            old_raw = payload.get("old_spec") or {}
            new_raw = payload.get("new_spec") or {}
            old_spec = {k: _clean(old_raw.get(k)) for k in fields}
            new_spec = {k: _clean(new_raw.get(k)) for k in fields}
            if all(old_spec.values()) and all(new_spec.values()):
                protected_specs.add(_key(old_spec))
            renames = payload.get("value_renames") or []
            if not isinstance(renames, list):
                continue
            for raw in renames:
                if not isinstance(raw, dict):
                    continue
                field = _clean(raw.get("field"))
                old_value = _clean(raw.get("old_value"))
                new_value = _clean(raw.get("new_value"))
                if field in fields and old_value and new_value and old_value != new_value:
                    unsafe_renames.append({
                        "field": field,
                        "old_value": old_value,
                        "new_value": new_value,
                    })

        if not unsafe_renames:
            return {"rename_events": 0, "candidate_specs": 0, "updated_rows": 0, "deleted_zero_duplicates": 0}

        try:
            shipment_rows = self.conn.execute(
                "SELECT payload_json FROM sync_inbox WHERE event_type='STOCK_TRANSFER_OUT'"
            ).fetchall()
        except sqlite3.OperationalError:
            shipment_rows = []

        candidates: Dict[Tuple[str, str, str, str, str, str, str], Dict[str, str]] = {}
        for row in shipment_rows:
            try:
                payload = json.loads(row["payload_json"] or "{}")
            except Exception:
                continue
            for item in payload.get("items") or []:
                if not isinstance(item, dict):
                    continue
                try:
                    qty = int(float(item.get("qty") or 0))
                except (TypeError, ValueError):
                    qty = 0
                if qty <= 0:
                    continue
                shipped = {k: _clean(item.get(k)) for k in fields}
                if not all(shipped.values()) or _key(shipped) in protected_specs:
                    continue
                for rename in unsafe_renames:
                    field = rename["field"]
                    if shipped[field].casefold() != rename["old_value"].casefold():
                        continue
                    corrupt = dict(shipped)
                    corrupt[field] = rename["new_value"]
                    key = (field, rename["old_value"].casefold(), rename["new_value"].casefold(), *_key(shipped))
                    candidates[key] = {
                        **shipped,
                        "corrupt_item_type": corrupt["item_type"],
                        "corrupt_school": corrupt["school"],
                        "corrupt_color": corrupt["color"],
                        "corrupt_size": corrupt["size"],
                    }

        updated_total = 0
        deleted_zero = 0
        if candidates:
            with self.conn:
                for candidate in candidates.values():
                    src_args = (
                        candidate["corrupt_item_type"],
                        candidate["corrupt_school"],
                        candidate["corrupt_color"],
                        candidate["corrupt_size"],
                    )
                    exists = self.conn.execute(
                        """
                        SELECT 1
                          FROM stocks
                         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                         LIMIT 1
                        """,
                        src_args,
                    ).fetchone()
                    if exists is None:
                        continue
                    dst_values = (
                        candidate["item_type"],
                        candidate["school"],
                        candidate["color"],
                        candidate["size"],
                    )
                    for table in (
                        "stocks",
                        "movements",
                        "bill_items",
                        "reservations",
                        "reservation_alerts",
                        "incoming_shipment_items_pending",
                        "branch_catalog_definitions",
                        "stock_audit_report_lines",
                        "size_profiles",
                    ):
                        try:
                            cur = self.conn.execute(
                                f"""
                                UPDATE {table}
                                   SET item_type = ?, school = ?, color = ?, size = ?
                                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                                """,
                                (*dst_values, *src_args),
                            )
                            updated_total += int(cur.rowcount or 0)
                        except sqlite3.OperationalError:
                            continue
                try:
                    cur = self.conn.execute(
                        """
                        DELETE FROM stocks
                         WHERE COALESCE(count, 0) = 0
                           AND EXISTS (
                                SELECT 1
                                  FROM stocks AS s2
                                 WHERE s2.id <> stocks.id
                                   AND LOWER(TRIM(s2.item_type)) = LOWER(TRIM(stocks.item_type))
                                   AND LOWER(TRIM(s2.school)) = LOWER(TRIM(stocks.school))
                                   AND LOWER(TRIM(s2.color)) = LOWER(TRIM(stocks.color))
                                   AND LOWER(TRIM(s2.size)) = LOWER(TRIM(stocks.size))
                                   AND COALESCE(s2.count, 0) > 0
                           )
                        """
                    )
                    deleted_zero = int(cur.rowcount or 0)
                except sqlite3.OperationalError:
                    deleted_zero = 0
                for candidate in candidates.values():
                    for field in fields:
                        self.conn.execute(
                            "INSERT OR IGNORE INTO spec_history(field, value) VALUES (?, ?)",
                            (field, candidate[field]),
                        )

        result = {
            "rename_events": len(unsafe_renames),
            "candidate_specs": len(candidates),
            "updated_rows": int(updated_total),
            "deleted_zero_duplicates": int(deleted_zero),
        }
        if updated_total or deleted_zero:
            self._audit("unsafe_spec_value_rename_repair", result)
        return result

    def _upsert_branch_catalog_definition_compat(
        self,
        item_type: str,
        school: str,
        color: str,
        size: str,
        unit_price: float,
        source_event_uuid: str,
        note: str,
    ) -> None:
        now = now_iso()
        cur = self.conn.execute(
            """
            UPDATE branch_catalog_definitions
               SET unit_price = ?,
                   source_event_uuid = ?,
                   note = ?,
                   created_at = ?
             WHERE item_type = ?
               AND school = ?
               AND color = ?
               AND size = ?
            """,
            (
                float(unit_price or 0),
                str(source_event_uuid or ""),
                str(note or ""),
                now,
                item_type,
                school,
                color,
                size,
            ),
        )
        if int(cur.rowcount or 0) == 0:
            self.conn.execute(
                """
                INSERT OR IGNORE INTO branch_catalog_definitions
                    (item_type, school, color, size, unit_price, source_event_uuid, note, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    item_type,
                    school,
                    color,
                    size,
                    float(unit_price or 0),
                    str(source_event_uuid or ""),
                    str(note or ""),
                    now,
                ),
            )

    def repair_missing_branch_reclassification_counts(self) -> Dict[str, int]:
        """Restore quantities for branch reclassification events that never landed.

        Older runs could mark a branch-owned target spec while the quantity move
        itself failed or was later removed. The event UUID in the movement note is
        the idempotency anchor used by the sync applier, so this repair only adds
        count when no RECLASS_IN movement exists for that exact event.
        """
        fields = ("item_type", "school", "color", "size")

        def _clean(value: Any) -> str:
            return str(value or "").strip()

        def _complete_spec(raw: Any) -> Optional[Dict[str, str]]:
            if not isinstance(raw, dict):
                return None
            spec = {field: _clean(raw.get(field)) for field in fields}
            return spec if all(spec.values()) else None

        def _stock_sum(spec: Dict[str, str]) -> int:
            row = self.conn.execute(
                """
                SELECT COALESCE(SUM(count), 0) AS qty
                  FROM stocks
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                """,
                (spec["item_type"], spec["school"], spec["color"], spec["size"]),
            ).fetchone()
            return int((row["qty"] if row else 0) or 0)

        def _add_stock(spec: Dict[str, str], qty: int, unit_price: float, note: str) -> int:
            row = self.conn.execute(
                """
                SELECT id
                  FROM stocks
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                   AND ABS(COALESCE(unit_price, 0) - ?) < 0.001
                 ORDER BY id ASC
                 LIMIT 1
                """,
                (spec["item_type"], spec["school"], spec["color"], spec["size"], float(unit_price)),
            ).fetchone()
            if row:
                stock_id = int(row["id"])
                self.conn.execute(
                    "UPDATE stocks SET count = COALESCE(count, 0) + ?, unit_price = ? WHERE id = ?",
                    (int(qty), float(unit_price), stock_id),
                )
            else:
                cur = self.conn.execute(
                    """
                    INSERT INTO stocks(item_type, school, color, size, unit_price, count)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        spec["item_type"],
                        spec["school"],
                        spec["color"],
                        spec["size"],
                        float(unit_price),
                        int(qty),
                    ),
                )
                stock_id = int(cur.lastrowid)
            self.conn.execute(
                """
                INSERT INTO movements
                    (ts, direction, stock_id, qty, note, bill_id,
                     item_type, school, color, size, unit_price)
                VALUES (?, 'RECLASS_IN', ?, ?, ?, NULL, ?, ?, ?, ?, ?)
                """,
                (
                    now_iso(),
                    stock_id,
                    int(qty),
                    note,
                    spec["item_type"],
                    spec["school"],
                    spec["color"],
                    spec["size"],
                    float(unit_price),
                ),
            )
            return stock_id

        try:
            rows = self.conn.execute(
                """
                SELECT event_uuid, payload_json
                  FROM sync_inbox
                 WHERE event_type = 'BRANCH_STOCK_RECLASSIFIED'
                """
            ).fetchall()
        except sqlite3.OperationalError:
            return {"events_checked": 0, "events_repaired": 0, "qty_restored": 0, "already_applied": 0}

        events_checked = 0
        events_repaired = 0
        qty_restored = 0
        already_applied = 0
        skipped_source_present = 0
        with self.conn:
            for row in rows:
                event_uuid = _clean(row["event_uuid"])
                if not event_uuid:
                    continue
                events_checked += 1
                if self.conn.execute(
                    """
                    SELECT 1
                      FROM movements
                     WHERE direction = 'RECLASS_IN'
                       AND note LIKE ?
                     LIMIT 1
                    """,
                    ("%" + event_uuid + "%",),
                ).fetchone():
                    already_applied += 1
                    continue
                try:
                    payload = json.loads(row["payload_json"] or "{}")
                except Exception:
                    continue
                from_spec = _complete_spec(payload.get("from_spec"))
                to_spec = _complete_spec(payload.get("to_spec"))
                if from_spec is None or to_spec is None:
                    continue
                try:
                    qty = int(payload.get("qty") or 0)
                except (TypeError, ValueError):
                    qty = 0
                if qty <= 0:
                    continue
                # If the source stock is still present, the normal sync applier
                # should handle the real move so we do not create duplicates.
                if _stock_sum(from_spec) >= qty:
                    skipped_source_present += 1
                    continue
                try:
                    unit_price = float((payload.get("to_spec") or {}).get("unit_price") or (payload.get("from_spec") or {}).get("unit_price") or 0)
                except (TypeError, ValueError):
                    unit_price = 0.0
                note = f"{_clean(payload.get('note')) or 'Branch stock reclassification repair'} #{event_uuid}"
                _add_stock(to_spec, qty, unit_price, note)
                self._upsert_branch_catalog_definition_compat(
                    to_spec["item_type"],
                    to_spec["school"],
                    to_spec["color"],
                    to_spec["size"],
                    float(unit_price),
                    event_uuid,
                    "Branch stock reclassification",
                )
                for field in fields:
                    self.conn.execute(
                        "INSERT OR IGNORE INTO spec_history(field, value) VALUES (?, ?)",
                        (field, to_spec[field]),
                    )
                events_repaired += 1
                qty_restored += int(qty)

        result = {
            "events_checked": int(events_checked),
            "events_repaired": int(events_repaired),
            "qty_restored": int(qty_restored),
            "already_applied": int(already_applied),
            "skipped_source_present": int(skipped_source_present),
        }
        if events_repaired:
            self._audit("branch_reclassification_count_repair", result)
        return result

    def repair_branch_group_general_school_counts_from_snapshots(self) -> Dict[str, int]:
        """Restore branch-group counts lost during عام school splitting.

        Some older warehouse corrections renamed branch rows with SPEC_RENAMED
        events instead of quantity reclassification events. That could leave the
        POS with catalog/price rows under the new branch-group school but with
        part of the previous quantity silently gone. Local POS stock snapshots
        are the safest source for this one-time recovery.
        """
        fields = ("item_type", "school", "color", "size")
        source_school = GENERAL_SHARED_SCHOOL_NAME
        device = (self._current_device_name() or "").strip()
        target_school = BRANCH_GENERAL_SCHOOL_TARGET_BY_DEVICE.get(device)
        if not target_school:
            return {
                "skipped": 1,
                "reason": "branch_not_mapped",
                "device": device,
                "qty_restored": 0,
            }

        def _clean(value: Any) -> str:
            return str(value or "").strip()

        def _qty_from(raw: Dict[str, Any]) -> int:
            for key in ("count", "qty", "quantity"):
                try:
                    return int(raw.get(key) or 0)
                except Exception:
                    continue
            return 0

        def _price_from(raw: Dict[str, Any]) -> float:
            for key in ("unit_price", "price"):
                try:
                    return float(raw.get(key) or 0)
                except Exception:
                    continue
            return 0.0

        def _mapped_key(raw: Dict[str, Any]) -> Optional[Tuple[str, str, str, str]]:
            item_type = _clean(raw.get("item_type"))
            school = _clean(raw.get("school"))
            color = _clean(raw.get("color"))
            size = _normalize_size_label(_strip_digit_marks(raw.get("size")))
            if not (item_type and school and color and size):
                return None
            if school not in (source_school, target_school):
                return None
            return (item_type, target_school, color, size)

        def _add_stock(key: Tuple[str, str, str, str], qty: int, unit_price: float, note: str) -> int:
            item_type, school, color, size = key
            row = self.conn.execute(
                """
                SELECT id
                  FROM stocks
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                   AND ABS(COALESCE(unit_price, 0) - ?) < 0.001
                 ORDER BY id ASC
                 LIMIT 1
                """,
                (item_type, school, color, size, float(unit_price)),
            ).fetchone()
            if row:
                stock_id = int(row["id"])
                self.conn.execute(
                    "UPDATE stocks SET count = COALESCE(count, 0) + ?, unit_price = ? WHERE id = ?",
                    (int(qty), float(unit_price), stock_id),
                )
            else:
                cur = self.conn.execute(
                    """
                    INSERT INTO stocks(item_type, school, color, size, unit_price, count)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (item_type, school, color, size, float(unit_price), int(qty)),
                )
                stock_id = int(cur.lastrowid)
            self.conn.execute(
                """
                INSERT INTO movements
                    (ts, direction, stock_id, qty, note, bill_id,
                     item_type, school, color, size, unit_price)
                VALUES (?, 'ADJUST_IN', ?, ?, ?, NULL, ?, ?, ?, ?, ?)
                """,
                (now_iso(), stock_id, int(qty), note, item_type, school, color, size, float(unit_price)),
            )
            return stock_id

        def _current_group_counts() -> Dict[Tuple[str, str, str, str], int]:
            counts: Dict[Tuple[str, str, str, str], int] = {}
            rows = self.conn.execute(
                """
                SELECT item_type, school, color, size, unit_price, COALESCE(SUM(count), 0) AS qty
                  FROM stocks
                 WHERE school IN (?, ?)
                 GROUP BY item_type, school, color, size, unit_price
                """,
                (source_school, target_school),
            ).fetchall()
            for row in rows:
                key = _mapped_key(dict(row))
                if key is not None:
                    counts[key] = counts.get(key, 0) + int(row["qty"] or 0)
            return counts

        try:
            snapshot_rows = self.conn.execute(
                """
                SELECT local_seq, created_at, payload_json
                 FROM sync_outbox
                 WHERE event_type = 'POS_STOCK_SNAPSHOT'
                 ORDER BY local_seq DESC
                 LIMIT 150
                """
            ).fetchall()
        except (sqlite3.OperationalError, sqlite3.DatabaseError):
            return {"skipped": 1, "reason": "no_sync_outbox", "device": device, "qty_restored": 0}

        current_counts = _current_group_counts()
        current_total = sum(max(0, qty) for qty in current_counts.values())
        best_snapshot: Optional[Dict[str, Any]] = None
        for row in snapshot_rows:
            try:
                payload = json.loads(row["payload_json"] or "{}")
            except Exception:
                continue
            rows = payload.get("rows") or payload.get("items") or payload.get("stocks") or []
            if not isinstance(rows, list):
                continue
            counts: Dict[Tuple[str, str, str, str], int] = {}
            prices: Dict[Tuple[str, str, str, str], float] = {}
            total = 0
            for raw in rows:
                if not isinstance(raw, dict):
                    continue
                key = _mapped_key(raw)
                if key is None:
                    continue
                qty = _qty_from(raw)
                if qty <= 0:
                    continue
                counts[key] = counts.get(key, 0) + qty
                if key not in prices or prices[key] <= 0:
                    prices[key] = _price_from(raw)
                total += qty
            if total > current_total:
                best_snapshot = {
                    "local_seq": int(row["local_seq"] or 0),
                    "created_at": _clean(row["created_at"]) or _clean(payload.get("snapshot_at")),
                    "counts": counts,
                    "prices": prices,
                    "total": total,
                }
                break

        if not best_snapshot:
            cleanup = {
                "wrong_zero_stock_rows": 0,
                "wrong_zero_catalog_rows": 0,
                "cleanup": "manual_delete_definition_only",
            }
            return {
                "device": device,
                "source_school": source_school,
                "target_school": target_school,
                "snapshot_seq": 0,
                "qty_restored": 0,
                "rows_restored": 0,
                "current_total": int(current_total),
                "snapshot_total": 0,
                **cleanup,
            }

        expected_counts = dict(best_snapshot["counts"])
        expected_prices = dict(best_snapshot.get("prices") or {})
        snapshot_at = _clean(best_snapshot.get("created_at"))
        plus_dirs = {"IN", "RETURN_IN", "ADJUST_IN", "RECLASS_IN", "TRANSFER_IN"}
        minus_dirs = {"OUT", "ADJUST_OUT", "RECLASS_OUT", "TRANSFER_OUT", "OUT_FACTORY"}
        if snapshot_at:
            movement_rows = self.conn.execute(
                """
                SELECT direction, item_type, school, color, size, unit_price, qty, note
                  FROM movements
                  WHERE school IN (?, ?)
                    AND REPLACE(ts, 'T', ' ') > ?
                 ORDER BY ts ASC, id ASC
                """,
                (source_school, target_school, snapshot_at),
            ).fetchall()
            for row in movement_rows:
                direction = _clean(row["direction"]).upper()
                note_text = _clean(row["note"]).casefold()
                if "branch group school snapshot repair" in note_text:
                    continue
                if direction not in plus_dirs and direction not in minus_dirs:
                    continue
                key = _mapped_key(dict(row))
                if key is None:
                    continue
                qty = abs(int(row["qty"] or 0))
                if key not in expected_prices or expected_prices[key] <= 0:
                    expected_prices[key] = _price_from(dict(row))
                if direction in plus_dirs:
                    expected_counts[key] = expected_counts.get(key, 0) + qty
                else:
                    expected_counts[key] = expected_counts.get(key, 0) - qty

        restored_qty = 0
        restored_rows = 0
        cleanup = {"wrong_zero_stock_rows": 0, "wrong_zero_catalog_rows": 0}
        with self.conn:
            current_counts = _current_group_counts()
            note = (
                f"Branch group school snapshot repair {source_school} -> {target_school} "
                f"snapshot_seq={best_snapshot['local_seq']}"
            )
            for key, expected_qty in expected_counts.items():
                expected_qty = int(expected_qty or 0)
                if expected_qty <= 0:
                    continue
                current_qty = int(current_counts.get(key, 0) or 0)
                missing = expected_qty - current_qty
                if missing <= 0:
                    continue
                unit_price = float(expected_prices.get(key) or 0)
                _add_stock(key, missing, unit_price, note)
                item_type, school, color, size = key
                self._upsert_branch_catalog_definition_compat(
                    item_type,
                    school,
                    color,
                    size,
                    float(unit_price),
                    f"snapshot-repair-{best_snapshot['local_seq']}",
                    "Branch group school snapshot repair",
                )
                for field, value in zip(fields, (item_type, school, color, size)):
                    self.conn.execute(
                        "INSERT OR IGNORE INTO spec_history(field, value) VALUES (?, ?)",
                        (field, value),
                    )
                restored_qty += missing
                restored_rows += 1
            cleanup = {
                "wrong_zero_stock_rows": 0,
                "wrong_zero_catalog_rows": 0,
                "cleanup": "manual_delete_definition_only",
            }

        result = {
            "device": device,
            "source_school": source_school,
            "target_school": target_school,
            "snapshot_seq": int(best_snapshot["local_seq"]),
            "snapshot_total": int(best_snapshot["total"]),
            "current_total": int(current_total),
            "qty_restored": int(restored_qty),
            "rows_restored": int(restored_rows),
            **cleanup,
        }
        if restored_qty or cleanup.get("wrong_zero_stock_rows") or cleanup.get("wrong_zero_catalog_rows"):
            self._audit("branch_group_general_school_snapshot_repair", result)
        return result

    def repair_stock_counts_from_snapshots(self) -> Dict[str, int]:
        """Restore broad count loss from the POS's own stock snapshots.

        This covers failures that are not school-mapping-specific, like a sync
        repair zeroing whole item groups. It uses the latest materially larger
        local POS_STOCK_SNAPSHOT, then replays later stock movements so normal
        sales/returns/reclassification after that snapshot are preserved.
        """
        fields = ("item_type", "school", "color", "size")

        def _clean(value: Any) -> str:
            return _strip_digit_marks(value).strip()

        def _qty_from(raw: Dict[str, Any]) -> int:
            for key in ("count", "qty", "quantity"):
                try:
                    return int(raw.get(key) or 0)
                except Exception:
                    continue
            return 0

        def _price_from(raw: Dict[str, Any]) -> float:
            for key in ("unit_price", "price"):
                try:
                    return float(raw.get(key) or 0)
                except Exception:
                    continue
            return 0.0

        def _key_from(raw: Dict[str, Any]) -> Optional[Tuple[str, str, str, str]]:
            item_type = _clean(raw.get("item_type"))
            school = _clean(raw.get("school"))
            color = _clean(raw.get("color"))
            size = _normalize_size_label(_strip_digit_marks(raw.get("size")))
            if not (item_type and school and color and size):
                return None
            return (item_type, school, color, size)

        def _local_cutoff(value: Any) -> str:
            text = str(value or "").strip()
            if not text:
                return ""
            try:
                raw = text.replace("Z", "")
                if "T" in raw:
                    raw = raw.split(".")[0]
                    dt = datetime.fromisoformat(raw)
                    # Sync timestamps are UTC; POS movement timestamps are local Cairo time.
                    return (dt + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                pass
            return text.replace("T", " ").split(".")[0].replace("Z", "")

        def _current_counts() -> Tuple[Dict[Tuple[str, str, str, str], int], Dict[Tuple[str, str, str, str], float]]:
            counts: Dict[Tuple[str, str, str, str], int] = {}
            prices: Dict[Tuple[str, str, str, str], float] = {}
            rows = self.conn.execute(
                """
                SELECT item_type, school, color, size, unit_price, COALESCE(SUM(count), 0) AS qty
                  FROM stocks
                 GROUP BY item_type, school, color, size, unit_price
                """
            ).fetchall()
            for row in rows:
                key = _key_from(dict(row))
                if key is None:
                    continue
                counts[key] = counts.get(key, 0) + int(row["qty"] or 0)
                if key not in prices or prices[key] <= 0:
                    prices[key] = float(row["unit_price"] or 0)
            return counts, prices

        def _add_stock(key: Tuple[str, str, str, str], qty: int, unit_price: float, note: str) -> int:
            item_type, school, color, size = key
            row = self.conn.execute(
                """
                SELECT id
                  FROM stocks
                 WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                   AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                   AND ABS(COALESCE(unit_price, 0) - ?) < 0.001
                 ORDER BY id ASC
                 LIMIT 1
                """,
                (item_type, school, color, size, float(unit_price)),
            ).fetchone()
            if row:
                stock_id = int(row["id"])
                self.conn.execute(
                    "UPDATE stocks SET count = COALESCE(count, 0) + ?, unit_price = ? WHERE id = ?",
                    (int(qty), float(unit_price), stock_id),
                )
            else:
                cur = self.conn.execute(
                    """
                    INSERT INTO stocks(item_type, school, color, size, unit_price, count)
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (item_type, school, color, size, float(unit_price), int(qty)),
                )
                stock_id = int(cur.lastrowid)
            self.conn.execute(
                """
                INSERT INTO movements
                    (ts, direction, stock_id, qty, note, bill_id,
                     item_type, school, color, size, unit_price)
                VALUES (?, 'ADJUST_IN', ?, ?, ?, NULL, ?, ?, ?, ?, ?)
                """,
                (now_iso(), stock_id, int(qty), note, item_type, school, color, size, float(unit_price)),
            )
            return stock_id

        try:
            snapshot_rows = self.conn.execute(
                """
                SELECT local_seq, created_at, payload_json
                  FROM sync_outbox
                 WHERE event_type = 'POS_STOCK_SNAPSHOT'
                 ORDER BY local_seq DESC
                 LIMIT 300
                """
            ).fetchall()
        except (sqlite3.OperationalError, sqlite3.DatabaseError):
            return {"skipped": 1, "reason": "no_local_stock_snapshots", "qty_restored": 0}

        current_counts, current_prices = _current_counts()
        current_total = sum(max(0, qty) for qty in current_counts.values())
        best_snapshot: Optional[Dict[str, Any]] = None
        for row in snapshot_rows:
            try:
                payload = json.loads(row["payload_json"] or "{}")
            except Exception:
                continue
            rows = payload.get("rows") or payload.get("items") or payload.get("stocks") or []
            if not isinstance(rows, list):
                continue
            counts: Dict[Tuple[str, str, str, str], int] = {}
            prices: Dict[Tuple[str, str, str, str], float] = {}
            for raw in rows:
                if not isinstance(raw, dict):
                    continue
                key = _key_from(raw)
                if key is None:
                    continue
                qty = _qty_from(raw)
                if qty <= 0:
                    continue
                counts[key] = counts.get(key, 0) + qty
                if key not in prices or prices[key] <= 0:
                    prices[key] = _price_from(raw)
            total = sum(counts.values())
            if total <= current_total + max(50, int(current_total * 0.05)):
                continue
            best_snapshot = {
                "local_seq": int(row["local_seq"] or 0),
                "created_at": _clean(payload.get("snapshot_at")) or _clean(row["created_at"]),
                "counts": counts,
                "prices": prices,
                "total": int(total),
            }
            break

        if not best_snapshot:
            return {
                "skipped": 0,
                "reason": "no_materially_higher_snapshot",
                "current_total": int(current_total),
                "qty_restored": 0,
            }

        expected_counts = dict(best_snapshot["counts"])
        expected_prices = dict(best_snapshot["prices"])
        plus_dirs = {"IN", "RETURN", "RETURN_IN", "ADJUST_IN", "RECLASS_IN", "TRANSFER_IN", "SHIPMENT_RECEIVED"}
        minus_dirs = {"OUT", "SALE", "RESERVE", "SHIPMENT_CANCEL", "ADJUST_OUT", "RECLASS_OUT", "TRANSFER_OUT", "OUT_FACTORY"}
        invalid_shipment_cancel_notes: Set[str] = set()
        try:
            cancel_rows = self.conn.execute(
                """
                SELECT payload_json
                  FROM sync_inbox
                 WHERE event_type = 'STOCK_TRANSFER_CANCELLED'
                """
            ).fetchall()
        except (sqlite3.OperationalError, sqlite3.DatabaseError):
            cancel_rows = []
        for cancel_row in cancel_rows:
            try:
                cancel_payload = json.loads(cancel_row["payload_json"] or "{}")
            except Exception:
                continue
            shipment_uuid = _clean(cancel_payload.get("shipment_uuid") or cancel_payload.get("bill_uuid"))
            if not shipment_uuid:
                continue
            known = self.conn.execute(
                "SELECT 1 FROM incoming_shipment_items_pending WHERE shipment_uuid = ? LIMIT 1",
                (shipment_uuid,),
            ).fetchone()
            if known is None:
                cancel_note = _clean(cancel_payload.get("note"))
                if cancel_note:
                    invalid_shipment_cancel_notes.add(cancel_note.casefold())
        cutoff = _local_cutoff(best_snapshot.get("created_at"))
        movement_rows = []
        if cutoff:
            try:
                movement_rows = self.conn.execute(
                    """
                    SELECT direction, item_type, school, color, size, unit_price, qty, note
                      FROM movements
                     WHERE REPLACE(ts, 'T', ' ') > ?
                       AND direction <> 'PRICE_UPDATE'
                     ORDER BY ts ASC, id ASC
                    """,
                    (cutoff,),
                ).fetchall()
            except (sqlite3.OperationalError, sqlite3.DatabaseError):
                movement_rows = []
        for row in movement_rows:
            direction = _clean(row["direction"]).upper()
            note_text = _clean(row["note"]).casefold()
            if "stock snapshot count repair" in note_text:
                continue
            if direction == "SHIPMENT_CANCEL" and note_text in invalid_shipment_cancel_notes:
                continue
            if direction not in plus_dirs and direction not in minus_dirs:
                continue
            key = _key_from(dict(row))
            if key is None:
                continue
            qty = abs(int(row["qty"] or 0))
            if key not in expected_prices or expected_prices[key] <= 0:
                expected_prices[key] = _price_from(dict(row))
            if direction in plus_dirs:
                expected_counts[key] = expected_counts.get(key, 0) + qty
            else:
                expected_counts[key] = expected_counts.get(key, 0) - qty

        restored_qty = 0
        restored_rows = 0
        with self.conn:
            current_counts, current_prices = _current_counts()
            note = f"Stock snapshot count repair snapshot_seq={best_snapshot['local_seq']}"
            for key, expected_qty in expected_counts.items():
                expected_qty = max(0, int(expected_qty or 0))
                current_qty = int(current_counts.get(key, 0) or 0)
                missing = expected_qty - current_qty
                if missing <= 0:
                    continue
                unit_price = float(expected_prices.get(key) or current_prices.get(key) or 0)
                _add_stock(key, missing, unit_price, note)
                item_type, school, color, size = key
                self._upsert_branch_catalog_definition_compat(
                    item_type,
                    school,
                    color,
                    size,
                    unit_price,
                    f"stock-snapshot-repair-{best_snapshot['local_seq']}",
                    "Stock snapshot count repair",
                )
                for field, value in zip(fields, key):
                    self.conn.execute(
                        "INSERT OR IGNORE INTO spec_history(field, value) VALUES (?, ?)",
                        (field, value),
                    )
                restored_qty += int(missing)
                restored_rows += 1

        result = {
            "snapshot_seq": int(best_snapshot["local_seq"]),
            "snapshot_total": int(best_snapshot["total"]),
            "current_total_before": int(current_total),
            "movement_rows_replayed": int(len(movement_rows)),
            "qty_restored": int(restored_qty),
            "rows_restored": int(restored_rows),
        }
        if restored_qty:
            self._audit("stock_snapshot_count_repair", result)
        return result

    def _cleanup_wrong_branch_group_school_rows(self, device: str, target_school: str) -> Dict[str, int]:
        return {
            "wrong_zero_stock_rows": 0,
            "wrong_zero_catalog_rows": 0,
            "cleanup": "manual_delete_definition_only",
        }
        wrong_schools = sorted(
            {
                school for school in BRANCH_GENERAL_SCHOOL_TARGET_BY_DEVICE.values()
                if school and school != target_school
            }
        )
        cleanup_schools = list(dict.fromkeys([GENERAL_SHARED_SCHOOL_NAME] + wrong_schools))
        if not cleanup_schools:
            return {"wrong_zero_stock_rows": 0, "wrong_zero_catalog_rows": 0}
        placeholders = ",".join("?" for _ in cleanup_schools)
        deleted_stocks = 0
        deleted_catalog = 0
        try:
            cur = self.conn.execute(
                f"DELETE FROM stocks WHERE COALESCE(count, 0) = 0 AND school IN ({placeholders})",
                cleanup_schools,
            )
            deleted_stocks = int(cur.rowcount or 0)
        except sqlite3.OperationalError:
            deleted_stocks = 0
        try:
            cur = self.conn.execute(
                f"""
                DELETE FROM branch_catalog_definitions
                 WHERE school IN ({placeholders})
                   AND NOT EXISTS (
                        SELECT 1 FROM stocks
                         WHERE stocks.item_type = branch_catalog_definitions.item_type
                           AND stocks.school = branch_catalog_definitions.school
                           AND stocks.color = branch_catalog_definitions.color
                           AND stocks.size = branch_catalog_definitions.size
                           AND COALESCE(stocks.count, 0) > 0
                   )
                """,
                cleanup_schools,
            )
            deleted_catalog = int(cur.rowcount or 0)
        except sqlite3.OperationalError:
            deleted_catalog = 0
        return {
            "wrong_zero_stock_rows": int(deleted_stocks),
            "wrong_zero_catalog_rows": int(deleted_catalog),
        }

    def cleanup_unowned_branch_catalog_rows(self) -> Dict[str, int]:
        """Remove catalog-only rows that this POS branch never received or owned.

        A branch owns a spec only when it has a warehouse bill shipment
        trail, a pending/confirmed shipment row, or the warehouse explicitly
        sent it as a reservation definition recorded in
        branch_catalog_definitions. Current stock quantity alone is not proof
        of branch ownership, because a bad catalog/branch sync can create
        positive rows for the wrong branch.

        Counts are deliberately not rebuilt from shipment quantities here.
        Shipments and reservation definitions decide visibility/ownership;
        POS sales, returns, audits, and stock movements decide current count.
        """
        return {
            "skipped": 1,
            "reason": "automatic_cleanup_disabled",
            "manual_cleanup": "use warehouse delete-definition",
            "stock_rows": 0,
            "size_profiles": 0,
            "spec_history": 0,
            "catalog_definitions": 0,
            "restored_stock_rows": 0,
            "hidden_definitions": 0,
            "price_updates": 0,
        }
        def _norm_expr(table_alias: str, column: str) -> str:
            return f"LOWER(TRIM(COALESCE({table_alias}.{column}, '')))"

        try:
            with self.conn:
                def _complete_spec_from(raw: Dict[str, Any]) -> Optional[Dict[str, str]]:
                    spec = {
                        "item_type": str(raw.get("item_type") or "").strip(),
                        "school": str(raw.get("school") or "").strip(),
                        "color": str(raw.get("color") or "").strip(),
                        "size": str(raw.get("size") or "").strip(),
                    }
                    return spec if all(spec.values()) else None

                def _load_rename_history() -> List[Tuple[Dict[str, str], Dict[str, str], List[Dict[str, str]]]]:
                    try:
                        rows = self.conn.execute(
                            """
                            SELECT payload_json
                              FROM sync_inbox
                             WHERE event_type = 'SPEC_RENAMED'
                             ORDER BY server_seq ASC
                            """
                        ).fetchall()
                    except sqlite3.OperationalError:
                        return []
                    history: List[Tuple[Dict[str, str], Dict[str, str], List[Dict[str, str]]]] = []
                    for row in rows:
                        try:
                            payload = json.loads(row["payload_json"] or "{}")
                        except Exception:
                            continue
                        old_spec_raw = payload.get("old_spec") or {}
                        new_spec_raw = payload.get("new_spec") or {}
                        if not isinstance(old_spec_raw, dict) or not isinstance(new_spec_raw, dict):
                            continue
                        old_spec = _complete_spec_from(old_spec_raw)
                        if old_spec is None:
                            continue
                        new_spec = {
                            k: str(new_spec_raw.get(k) or old_spec.get(k) or "").strip()
                            for k in ("item_type", "school", "color", "size")
                        }
                        renames = [
                            {
                                "field": str(v.get("field") or "").strip(),
                                "old_value": str(v.get("old_value") or "").strip(),
                                "new_value": str(v.get("new_value") or "").strip(),
                            }
                            for v in (payload.get("value_renames") or [])
                            if isinstance(v, dict)
                        ] if bool(payload.get("allow_global_value_renames")) else []
                        if old_spec == new_spec and not renames:
                            continue
                        history.append((old_spec, new_spec, renames))
                    return history

                rename_history = _load_rename_history()

                def _apply_local_renames(spec: Dict[str, str]) -> Dict[str, str]:
                    current = {k: str(spec.get(k) or "").strip() for k in ("item_type", "school", "color", "size")}
                    for old_spec, new_spec, value_renames in rename_history:
                        if all(
                            current.get(k, "").casefold() == str(old_spec.get(k) or "").strip().casefold()
                            for k in ("item_type", "school", "color", "size")
                        ):
                            current = {k: str(new_spec.get(k) or current.get(k) or "").strip() for k in current}
                        for rename in value_renames:
                            field = rename.get("field")
                            if field not in current:
                                continue
                            old_value = str(rename.get("old_value") or "").strip()
                            new_value = str(rename.get("new_value") or "").strip()
                            if old_value and new_value and current[field].casefold() == old_value.casefold():
                                current[field] = new_value
                    return current

                def _spec_key(spec: Dict[str, str]) -> Tuple[str, str, str, str]:
                    return tuple(str(spec[k] or "").strip().casefold() for k in ("item_type", "school", "color", "size"))  # type: ignore[return-value]

                def _load_delete_filters() -> List[Dict[str, str]]:
                    try:
                        self.conn.execute(
                            """
                            CREATE TABLE IF NOT EXISTS branch_catalog_delete_tombstones(
                                item_type TEXT,
                                school TEXT,
                                color TEXT,
                                size TEXT,
                                delete_server_seq INTEGER,
                                source_event_uuid TEXT,
                                note TEXT,
                                created_at TEXT NOT NULL
                            )
                            """
                        )
                        try:
                            cols = {
                                str(row["name"])
                                for row in self.conn.execute("PRAGMA table_info(branch_catalog_delete_tombstones)").fetchall()
                            }
                            if "delete_server_seq" not in cols:
                                self.conn.execute("ALTER TABLE branch_catalog_delete_tombstones ADD COLUMN delete_server_seq INTEGER")
                        except sqlite3.OperationalError:
                            pass
                        rows = self.conn.execute(
                            "SELECT item_type, school, color, size, delete_server_seq FROM branch_catalog_delete_tombstones"
                        ).fetchall()
                    except sqlite3.OperationalError:
                        return []
                    out = []
                    for row in rows:
                        filt = {
                            "item_type": str(row["item_type"] or "").strip(),
                            "school": str(row["school"] or "").strip(),
                            "color": str(row["color"] or "").strip(),
                            "size": str(row["size"] or "").strip(),
                        }
                        try:
                            filt["delete_server_seq"] = int(row["delete_server_seq"]) if row["delete_server_seq"] is not None else None  # type: ignore[assignment]
                        except (TypeError, ValueError):
                            filt["delete_server_seq"] = None  # type: ignore[assignment]
                        out.append(filt)
                    return out

                delete_filters = _load_delete_filters()

                def _deleted_by_filter(spec: Dict[str, str], event_seq: Optional[int] = None) -> bool:
                    for filt in delete_filters:
                        delete_seq = filt.get("delete_server_seq")
                        if isinstance(delete_seq, int) and event_seq is not None and event_seq > delete_seq:
                            continue
                        if isinstance(delete_seq, int) and event_seq is None:
                            continue
                        matched = True
                        for field in ("item_type", "school", "color", "size"):
                            value = str(filt.get(field) or "").strip()
                            if value and value.casefold() != str(spec.get(field) or "").strip().casefold():
                                matched = False
                                break
                        if matched:
                            return True
                    return False

                price_overrides: Dict[Tuple[str, str, str, str], float] = {}
                try:
                    price_rows = self.conn.execute(
                        """
                        SELECT payload_json
                          FROM sync_inbox
                         WHERE event_type = 'PRICE_UPDATE'
                         ORDER BY server_seq ASC
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    price_rows = []
                for price_row in price_rows:
                    try:
                        payload = json.loads(price_row["payload_json"] or "{}")
                    except Exception:
                        continue
                    raw_filters = payload.get("filters") or {}
                    if not isinstance(raw_filters, dict):
                        continue
                    spec = _complete_spec_from(raw_filters)
                    if spec is None:
                        continue
                    try:
                        new_price = float(payload.get("new_price"))
                    except (TypeError, ValueError):
                        continue
                    if new_price < 0:
                        continue
                    price_overrides[_spec_key(spec)] = new_price
                    price_overrides[_spec_key(_apply_local_renames(spec))] = new_price

                self.conn.executescript(
                    """
                    DROP TABLE IF EXISTS temp._owned_branch_specs;
                    CREATE TEMP TABLE _owned_branch_specs(
                        item_type TEXT NOT NULL,
                        school TEXT NOT NULL,
                        color TEXT NOT NULL,
                        size TEXT NOT NULL,
                        PRIMARY KEY(item_type, school, color, size)
                    );
                    DROP TABLE IF EXISTS temp._owned_branch_values;
                    CREATE TEMP TABLE _owned_branch_values(
                        field TEXT NOT NULL,
                        value TEXT NOT NULL,
                        PRIMARY KEY(field, value)
                    );
                    """
                )

                restored_stock_rows = 0
                price_updates = 0
                positive_stock_preserved = self.conn.execute(
                    """
                    INSERT OR IGNORE INTO _owned_branch_specs(item_type, school, color, size)
                    SELECT LOWER(TRIM(COALESCE(item_type, ''))),
                           LOWER(TRIM(COALESCE(school, ''))),
                           LOWER(TRIM(COALESCE(color, ''))),
                           LOWER(TRIM(COALESCE(size, '')))
                      FROM stocks
                     WHERE COALESCE(count, 0) > 0
                       AND COALESCE(TRIM(item_type), '') <> ''
                       AND COALESCE(TRIM(school), '') <> ''
                       AND COALESCE(TRIM(color), '') <> ''
                       AND COALESCE(TRIM(size), '') <> ''
                    """
                ).rowcount
                audit_specs_preserved = 0
                try:
                    audit_specs_preserved = int(self.conn.execute(
                        """
                        INSERT OR IGNORE INTO _owned_branch_specs(item_type, school, color, size)
                        SELECT LOWER(TRIM(COALESCE(item_type, ''))),
                               LOWER(TRIM(COALESCE(school, ''))),
                               LOWER(TRIM(COALESCE(color, ''))),
                               LOWER(TRIM(COALESCE(size, '')))
                          FROM stock_audit_report_lines
                         WHERE COALESCE(TRIM(item_type), '') <> ''
                           AND COALESCE(TRIM(school), '') <> ''
                           AND COALESCE(TRIM(color), '') <> ''
                           AND COALESCE(TRIM(size), '') <> ''
                        """
                    ).rowcount or 0)
                except sqlite3.OperationalError:
                    audit_specs_preserved = 0

                deleted_definitions = self.conn.execute(
                    """
                    DELETE FROM branch_catalog_definitions
                     WHERE LOWER(COALESCE(note, '')) NOT LIKE '%reservation%'
                       AND COALESCE(note, '') NOT LIKE '%حجز%'
                       AND LOWER(COALESCE(note, '')) NOT LIKE '%reclassification%'
                       AND NOT EXISTS (
                            SELECT 1 FROM stocks
                             WHERE LOWER(TRIM(COALESCE(stocks.item_type, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                               AND LOWER(TRIM(COALESCE(stocks.school, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                               AND LOWER(TRIM(COALESCE(stocks.color, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                               AND LOWER(TRIM(COALESCE(stocks.size, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
                               AND COALESCE(stocks.count, 0) > 0
                       )
                       AND NOT EXISTS (
                            SELECT 1 FROM incoming_shipment_items_pending
                             WHERE LOWER(TRIM(COALESCE(incoming_shipment_items_pending.item_type, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                               AND LOWER(TRIM(COALESCE(incoming_shipment_items_pending.school, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                               AND LOWER(TRIM(COALESCE(incoming_shipment_items_pending.color, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                               AND LOWER(TRIM(COALESCE(incoming_shipment_items_pending.size, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
                               AND (COALESCE(expected_qty, 0) > 0 OR COALESCE(received_qty, 0) > 0)
                               AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
                       )
                       AND NOT EXISTS (
                            SELECT 1 FROM stock_audit_report_lines
                             WHERE LOWER(TRIM(COALESCE(stock_audit_report_lines.item_type, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.item_type, '')))
                               AND LOWER(TRIM(COALESCE(stock_audit_report_lines.school, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.school, '')))
                               AND LOWER(TRIM(COALESCE(stock_audit_report_lines.color, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.color, '')))
                               AND LOWER(TRIM(COALESCE(stock_audit_report_lines.size, ''))) =
                                   LOWER(TRIM(COALESCE(branch_catalog_definitions.size, '')))
                       )
                    """
                ).rowcount

                def _remember_owned_spec(
                    spec: Dict[str, str],
                    unit_price: float,
                    *,
                    create_stock_row: bool,
                    definition_event_uuid: str = "",
                    definition_note: str = "",
                    event_seq: Optional[int] = None,
                ) -> None:
                    truth_spec = _apply_local_renames(spec)
                    if not all(truth_spec.values()):
                        return
                    if _deleted_by_filter(truth_spec, event_seq):
                        return
                    key = _spec_key(truth_spec)
                    effective_price = float(price_overrides.get(key, unit_price or 0.0))
                    self.conn.execute(
                        """
                        INSERT OR IGNORE INTO _owned_branch_specs(item_type, school, color, size)
                        VALUES (LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)), LOWER(TRIM(?)))
                        """,
                        (
                            truth_spec["item_type"],
                            truth_spec["school"],
                            truth_spec["color"],
                            truth_spec["size"],
                        ),
                    )
                    for field, value in truth_spec.items():
                        self.conn.execute(
                            "INSERT OR IGNORE INTO spec_history(field, value) VALUES (?, ?)",
                            (field, value),
                        )
                    if definition_note:
                        self._upsert_branch_catalog_definition_compat(
                            truth_spec["item_type"],
                            truth_spec["school"],
                            truth_spec["color"],
                            truth_spec["size"],
                            effective_price,
                            definition_event_uuid,
                            definition_note,
                        )
                    if not create_stock_row:
                        return
                    existing = self.conn.execute(
                        """
                        SELECT id, unit_price
                          FROM stocks
                         WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                           AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                           AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                           AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                         ORDER BY id ASC
                         LIMIT 1
                        """,
                        (
                            truth_spec["item_type"],
                            truth_spec["school"],
                            truth_spec["color"],
                            truth_spec["size"],
                        ),
                    ).fetchone()
                    nonlocal restored_stock_rows, price_updates
                    if existing is None:
                        self.conn.execute(
                            """
                            INSERT INTO stocks(item_type, school, color, size, unit_price, count)
                            VALUES (?, ?, ?, ?, ?, 0)
                            """,
                            (
                                truth_spec["item_type"],
                                truth_spec["school"],
                                truth_spec["color"],
                                truth_spec["size"],
                                effective_price,
                            ),
                        )
                        restored_stock_rows += 1
                    elif abs(float(existing["unit_price"] or 0) - effective_price) >= 0.001:
                        self.conn.execute(
                            "UPDATE stocks SET unit_price = ? WHERE id = ?",
                            (effective_price, int(existing["id"])),
                        )
                        price_updates += 1

                try:
                    cancel_rows = self.conn.execute(
                        """
                        SELECT payload_json
                          FROM sync_inbox
                         WHERE event_type = 'STOCK_TRANSFER_CANCELLED'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    cancel_rows = []
                cancelled_shipments: Set[str] = set()
                for cancel_row in cancel_rows:
                    try:
                        cancel_payload = json.loads(cancel_row["payload_json"] or "{}")
                    except Exception:
                        continue
                    shipment_uuid = str(cancel_payload.get("shipment_uuid") or cancel_payload.get("bill_uuid") or "").strip()
                    if shipment_uuid:
                        cancelled_shipments.add(shipment_uuid.casefold())

                try:
                    inbox_rows = self.conn.execute(
                        """
                        SELECT event_uuid, payload_json, server_seq
                          FROM sync_inbox
                         WHERE event_type = 'STOCK_TRANSFER_OUT'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    inbox_rows = []
                for inbox_row in inbox_rows:
                    try:
                        payload = json.loads(inbox_row["payload_json"] or "{}")
                    except Exception:
                        continue
                    shipment_uuid = str(payload.get("shipment_uuid") or "").strip()
                    if shipment_uuid and shipment_uuid.casefold() in cancelled_shipments:
                        continue
                    note_text = str(payload.get("note") or "")
                    is_reservation_definition = (
                        "reservation" in note_text.casefold()
                        or "\u062d\u062c\u0632" in note_text
                    )
                    event_uuid = str(inbox_row["event_uuid"] or "")
                    for item in payload.get("items") or []:
                        if not isinstance(item, dict):
                            continue
                        item_type = str(item.get("item_type") or "").strip()
                        school = str(item.get("school") or "").strip()
                        color = str(item.get("color") or "").strip()
                        size = str(item.get("size") or "").strip()
                        if not (item_type and school and color and size):
                            continue
                        try:
                            qty = int(float(item.get("qty") or 0))
                        except (TypeError, ValueError):
                            qty = 0
                        catalog_only = bool(item.get("catalog_only"))
                        owns_spec = qty > 0 or (catalog_only and is_reservation_definition)
                        if not owns_spec:
                            continue
                        try:
                            unit_price = float(item.get("unit_price") or 0)
                        except (TypeError, ValueError):
                            unit_price = 0.0
                        try:
                            event_seq = int(inbox_row["server_seq"])
                        except (TypeError, ValueError):
                            event_seq = None
                        _remember_owned_spec(
                            {
                                "item_type": item_type,
                                "school": school,
                                "color": color,
                                "size": size,
                            },
                            unit_price,
                            create_stock_row=True,
                            definition_event_uuid=(event_uuid if catalog_only and is_reservation_definition else ""),
                            definition_note=(note_text if catalog_only and is_reservation_definition else ""),
                            event_seq=event_seq,
                        )

                try:
                    reclass_rows = self.conn.execute(
                        """
                        SELECT event_uuid, payload_json, server_seq
                          FROM sync_inbox
                         WHERE event_type = 'BRANCH_STOCK_RECLASSIFIED'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    reclass_rows = []
                for reclass_row in reclass_rows:
                    try:
                        payload = json.loads(reclass_row["payload_json"] or "{}")
                    except Exception:
                        continue
                    try:
                        event_seq = int(reclass_row["server_seq"])
                    except (TypeError, ValueError):
                        event_seq = None
                    event_uuid = str(reclass_row["event_uuid"] or "")
                    reclass_specs: List[Dict[str, Any]] = []
                    to_spec = payload.get("to_spec") or {}
                    if isinstance(to_spec, dict):
                        reclass_specs.append(to_spec)
                    catalog_rows = payload.get("catalog_rows") or []
                    if isinstance(catalog_rows, list):
                        reclass_specs.extend([r for r in catalog_rows if isinstance(r, dict)])
                    for raw_spec in reclass_specs:
                        spec = _complete_spec_from(raw_spec)
                        if spec is None:
                            continue
                        try:
                            unit_price = float(raw_spec.get("unit_price") or 0)
                        except (TypeError, ValueError):
                            unit_price = 0.0
                        _remember_owned_spec(
                            spec,
                            unit_price,
                            create_stock_row=True,
                            definition_event_uuid=event_uuid,
                            definition_note="Branch stock reclassification",
                            event_seq=event_seq,
                        )

                try:
                    pending_rows = self.conn.execute(
                        """
                        SELECT item_type, school, color, size, unit_price
                         FROM incoming_shipment_items_pending
                         WHERE (
                                COALESCE(expected_qty, 0) > 0
                                OR COALESCE(received_qty, 0) > 0
                           )
                           AND UPPER(COALESCE(status, '')) <> 'CANCELLED'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    pending_rows = []
                for pending in pending_rows:
                    spec = _complete_spec_from(dict(pending))
                    if spec is None:
                        continue
                    try:
                        unit_price = float(pending["unit_price"] or 0)
                    except (TypeError, ValueError):
                        unit_price = 0.0
                    _remember_owned_spec(spec, unit_price, create_stock_row=True)

                try:
                    definition_rows = self.conn.execute(
                        """
                        SELECT item_type, school, color, size, unit_price,
                               source_event_uuid, note
                          FROM branch_catalog_definitions
                         WHERE LOWER(COALESCE(note, '')) LIKE '%reservation%'
                            OR COALESCE(note, '') LIKE '%حجز%'
                            OR LOWER(COALESCE(note, '')) LIKE '%reclassification%'
                        """
                    ).fetchall()
                except sqlite3.OperationalError:
                    definition_rows = []
                for definition in definition_rows:
                    spec = _complete_spec_from(dict(definition))
                    if spec is None:
                        continue
                    try:
                        unit_price = float(definition["unit_price"] or 0)
                    except (TypeError, ValueError):
                        unit_price = 0.0
                    source_uuid = str(definition["source_event_uuid"] or "").strip()
                    event_seq = None
                    if source_uuid:
                        seq_row = self.conn.execute(
                            "SELECT server_seq FROM sync_inbox WHERE event_uuid = ? LIMIT 1",
                            (source_uuid,),
                        ).fetchone()
                        try:
                            event_seq = int(seq_row["server_seq"]) if seq_row else None
                        except (TypeError, ValueError):
                            event_seq = None
                    _remember_owned_spec(
                        spec,
                        unit_price,
                        create_stock_row=True,
                        definition_event_uuid=source_uuid,
                        definition_note=str(definition["note"] or ""),
                        event_seq=event_seq,
                    )

                spec_sources: List[Tuple[str, str]] = []
                for table_name, where_sql in spec_sources:
                    try:
                        self.conn.execute(
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
                        pass

                deleted_stock = self.conn.execute(
                    """
                    DELETE FROM stocks
                     WHERE NOT EXISTS (
                            SELECT 1 FROM _owned_branch_specs o
                             WHERE o.item_type = LOWER(TRIM(COALESCE(stocks.item_type, '')))
                               AND o.school = LOWER(TRIM(COALESCE(stocks.school, '')))
                               AND o.color = LOWER(TRIM(COALESCE(stocks.color, '')))
                               AND o.size = LOWER(TRIM(COALESCE(stocks.size, '')))
                       )
                       AND COALESCE(count, 0) <= 0
                    """
                ).rowcount

                deleted_profiles = self.conn.execute(
                    """
                    DELETE FROM size_profiles
                     WHERE NOT EXISTS (
                            SELECT 1 FROM _owned_branch_specs o
                             WHERE o.item_type = LOWER(TRIM(COALESCE(size_profiles.item_type, '')))
                               AND o.school = LOWER(TRIM(COALESCE(size_profiles.school, '')))
                               AND o.color = LOWER(TRIM(COALESCE(size_profiles.color, '')))
                       )
                       AND NOT EXISTS (
                            SELECT 1 FROM stocks
                             WHERE LOWER(TRIM(COALESCE(stocks.item_type, ''))) =
                                   LOWER(TRIM(COALESCE(size_profiles.item_type, '')))
                               AND LOWER(TRIM(COALESCE(stocks.school, ''))) =
                                   LOWER(TRIM(COALESCE(size_profiles.school, '')))
                               AND LOWER(TRIM(COALESCE(stocks.color, ''))) =
                                   LOWER(TRIM(COALESCE(size_profiles.color, '')))
                               AND COALESCE(stocks.count, 0) > 0
                       )
                    """
                ).rowcount

                for field, column in (
                    ("item_type", "item_type"),
                    ("school", "school"),
                    ("color", "color"),
                    ("size", "size"),
                ):
                    self.conn.execute(
                        f"""
                        INSERT OR IGNORE INTO _owned_branch_values(field, value)
                        SELECT ?, {column}
                          FROM _owned_branch_specs
                         WHERE COALESCE(TRIM({column}), '') <> ''
                        """,
                        (field,),
                    )
                    for table_name, where_sql in spec_sources:
                        try:
                            self.conn.execute(
                                f"""
                                INSERT OR IGNORE INTO _owned_branch_values(field, value)
                                SELECT ?, LOWER(TRIM(COALESCE({column}, '')))
                                  FROM {table_name}
                                 WHERE {where_sql}
                                   AND COALESCE(TRIM({column}), '') <> ''
                                """,
                                (field,),
                            )
                        except sqlite3.OperationalError:
                            pass

                deleted_hidden = self.conn.execute(
                    """
                    DELETE FROM hidden_definitions
                     WHERE EXISTS (
                            SELECT 1 FROM _owned_branch_values v
                             WHERE v.field = hidden_definitions.field
                               AND v.value = LOWER(TRIM(COALESCE(hidden_definitions.value, '')))
                       )
                    """
                ).rowcount

                deleted_history = self.conn.execute(
                    """
                    DELETE FROM spec_history
                     WHERE NOT EXISTS (
                            SELECT 1 FROM _owned_branch_values v
                             WHERE v.field = spec_history.field
                               AND v.value = LOWER(TRIM(COALESCE(spec_history.value, '')))
                       )
                    """
                ).rowcount

                self.conn.executescript(
                    """
                    DROP TABLE IF EXISTS temp._owned_branch_specs;
                    DROP TABLE IF EXISTS temp._owned_branch_values;
                    """
                )
            return {
                "stock_rows": int(deleted_stock or 0),
                "size_profiles": int(deleted_profiles or 0),
                "spec_history": int(deleted_history or 0),
                "catalog_definitions": int(deleted_definitions or 0),
                "restored_stock_rows": int(restored_stock_rows or 0),
                "positive_stock_preserved": int(positive_stock_preserved or 0),
                "audit_specs_preserved": int(audit_specs_preserved or 0),
                "hidden_definitions": int(deleted_hidden or 0),
                "price_updates": int(price_updates or 0),
            }
        except Exception:
            import traceback
            traceback.print_exc()
            return {
                "stock_rows": 0,
                "size_profiles": 0,
                "spec_history": 0,
                "catalog_definitions": 0,
                "restored_stock_rows": 0,
                "hidden_definitions": 0,
                "price_updates": 0,
            }

    def _filters_where(self, filters: Dict[str, Any], prefix: str = "") -> Tuple[str, List[Any]]:
        """Include count = 0 stock rows so schools/items/colors stay visible when out of stock."""
        where: List[str] = ["1=1"]
        args: List[Any] = []

        candidate_keys_txt = ("item_type", "school", "color", "size")
        all_keys = candidate_keys_txt

        multi_keys = [k for k in all_keys if isinstance(filters.get(k), list) and len(filters.get(k) or []) > 0]
        if len(multi_keys) > 1:
            raise ValueError("يمكن اختيار أكثر من قيمة لحقل واحد فقط، وليس لأكثر من حقل.")

        for k in candidate_keys_txt:
            v = filters.get(k)
            if v and not isinstance(v, list):
                v = str(v).strip()
                if v:
                    col = f"{prefix}{k}"
                    if k == "size":
                        col = f"TRIM(REPLACE(REPLACE({col}, char(8206), ''), char(8207), ''))"
                        v = _normalize_size_label(v)
                    else:
                        col = f"TRIM({col})"
                    where.append(f"LOWER({col}) = LOWER(?)")
                    args.append(v)

        if multi_keys:
            mk = multi_keys[0]
            vals = [x for x in (filters.get(mk) or []) if x not in (None, "")]
            if vals:
                placeholders = ",".join(["?"] * len(vals))
                col = f"{prefix}{mk}"
                if mk == "size":
                    col = f"TRIM(REPLACE(REPLACE({col}, char(8206), ''), char(8207), ''))"
                    args.extend([_normalize_size_label(str(x).strip()).lower() for x in vals])
                else:
                    col = f"TRIM({col})"
                    args.extend([str(x).strip().lower() for x in vals])
                where.append(f"LOWER({col}) IN ({placeholders})")

        return (" AND ".join(where)) if where else "1=1", args

    def search_stocks(self, filters: Dict[str, Any]) -> List[StockRow]:
        where, args = self._filters_where(filters)
        cur = self.conn.cursor()
        cur.execute(
            f"""SELECT id,item_type,school,color,size,unit_price,count
                FROM stocks
                WHERE {where}
                ORDER BY id ASC""",
            args,
        )
        rows = [
            StockRow(
                id=r["id"], item_type=r["item_type"], school=r["school"], color=r["color"],
                size=r["size"],
                unit_price=r["unit_price"], count=r["count"],
            )
            for r in cur.fetchall()
        ]
        cur.close()
        return rows

    def current_inventory(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        where, args = self._filters_where(filters, prefix="s.")
        having_sql = "HAVING SUM(s.count) > 0" if (filters or {}).get("hide_zero") else ""
        size_expr = "COALESCE(TRIM(REPLACE(REPLACE(s.size, char(8206), ''), char(8207), '')), '')"

        cur = self.conn.cursor()
        cur.execute(
            f"""
            SELECT
                MIN(s.id)              AS id,
                COALESCE(s.item_type, '') AS item_type,
                COALESCE(s.school, '')    AS school,
                COALESCE(s.color, '')     AS color,
                {size_expr}               AS size,
                COALESCE(s.unit_price, 0) AS unit_price,
                COALESCE(SUM(COALESCE(s.count, 0)), 0) AS count,
                COALESCE(SUM(COALESCE(s.count, 0) * COALESCE(s.unit_price, 0)), 0) AS value
            FROM stocks s
            WHERE {where}
            GROUP BY
                COALESCE(s.item_type, ''),
                COALESCE(s.school, ''),
                COALESCE(s.color, ''),
                {size_expr},
                COALESCE(s.unit_price, 0)
            {having_sql}
            ORDER BY
                s.item_type,
                s.school,
                s.color,

                CASE
                    WHEN {size_expr} NOT GLOB '*[^0-9]*' THEN 0
                    ELSE 1
                END,

                CASE
                    WHEN {size_expr} NOT GLOB '*[^0-9]*'
                    THEN CAST({size_expr} AS INTEGER)
                    ELSE NULL
                END,

                CASE UPPER({size_expr})
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
                    unit_price=r["unit_price"],
                    count=int(r["count"]),
                    value=float(r["value"]),
                )
            )

        cur.close()
        return rows

    def low_stock_summary(self, *, threshold: int = 5, limit: int = 80) -> Tuple[int, List[Dict[str, Any]]]:
        threshold = max(1, int(threshold or 5))
        limit = max(1, int(limit or 80))
        hidden_sql = " AND ".join(
            self._hidden_definition_sql(field, f"s.{field}")
            for field in ("item_type", "school", "color", "size")
        )
        base_sql = f"""
            FROM (
                SELECT
                    COALESCE(s.item_type, '') AS item_type,
                    COALESCE(s.school, '') AS school,
                    COALESCE(s.color, '') AS color,
                    COALESCE(s.size, '') AS size,
                    COALESCE(SUM(COALESCE(s.count, 0)), 0) AS count
                FROM stocks s
                WHERE {hidden_sql}
                GROUP BY
                    COALESCE(s.item_type, ''),
                    COALESCE(s.school, ''),
                    COALESCE(s.color, ''),
                    COALESCE(s.size, '')
                HAVING count > 0 AND count <= ?
            ) low
        """
        total_row = self.conn.execute(f"SELECT COUNT(*) AS c {base_sql}", (threshold,)).fetchone()
        rows = self.conn.execute(
            f"""
            SELECT item_type, school, color, size, count
            {base_sql}
            ORDER BY count ASC, item_type, school, color, size
            LIMIT ?
            """,
            (threshold, limit),
        ).fetchall()
        return int(total_row["c"] if total_row else 0), [dict(r) for r in rows]

    # -------- Billing --------
    def create_bill(
        self,
        customer: str,
        bill_lines: List[Dict[str, Any]],
        payment_method: str = PAYMENT_METHOD_CASH,
        customer_phone: str = "",
    ) -> int:
        self._require_shift()
        if not bill_lines:
            raise ValueError("Bill has no items")
        for line in bill_lines:
            if bool(line.get("allow_factory_fill")):
                raise ValueError("POS sale bills cannot contain factory/no-stock items")
        payment_method = str(payment_method or PAYMENT_METHOD_CASH).strip().upper()
        if payment_method not in (PAYMENT_METHOD_CASH, PAYMENT_METHOD_VISA):
            raise ValueError("طريقة الدفع غير مدعومة.")
        customer_phone = _normalize_customer_phone(customer_phone)

        def _candidates_for_line(line: Dict[str, Any]) -> List[sqlite3.Row]:
            cur = self.conn.cursor()
            try:
                if line.get("stock_id"):
                    cur.execute(
                        "SELECT * FROM stocks WHERE id=? AND count>0 ORDER BY id ASC",
                        (int(line["stock_id"]),),
                    )
                else:
                    where_parts = ["count > 0"]
                    args: List[Any] = []
                    for k in ("item_type", "school", "color", "size"):
                        v = (line.get(k) or "").strip()
                        if v:
                            where_parts.append(f"LOWER(TRIM({k})) = LOWER(?)")
                            args.append(v)
                    cur.execute(f"SELECT * FROM stocks WHERE {' AND '.join(where_parts)} ORDER BY id ASC", args)
                return cur.fetchall()
            finally:
                cur.close()

        with self.conn:
            bill_cur = self.conn.execute(
                "INSERT INTO bills(created_at,customer,customer_phone,total,bill_type,status,payment_method,shift_id) VALUES(?,?,?,?,?,?,?,?)",
                (now_iso(), (customer or "").strip() or None, customer_phone or None, 0.0, "SALE", "CONFIRMED", payment_method, self.active_shift_id),
            )
            bill_id = int(bill_cur.lastrowid)
            total = 0.0

            for line in bill_lines:
                try:
                    qty_needed = int(line["qty"])
                except Exception:
                    raise ValueError("Qty must be > 0")
                if qty_needed <= 0:
                    raise ValueError("Qty must be > 0")

                allow_factory = False
                cands = _candidates_for_line(line)

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
                    raise ValueError(
                        f"لا توجد كمية كافية للصنف "
                        f"{line.get('item_type','(النوع?)')} / {line.get('school','(المدرسة?)')} / {line.get('size','(المقاس?)')} "
                        f"(المطلوب {qty_needed}، المتاح {qty_needed-remaining})."
                    )

                taken = qty_needed - remaining

                price_base: Optional[float] = None
                if "unit_price" in line and line["unit_price"] is not None and str(line["unit_price"]) != "":
                    price_base = float(line["unit_price"])
                elif taken > 0:
                    value_sum = sum(float(s["unit_price"]) * take for s, take in chunks)
                    qty_sum = sum(take for _, take in chunks)
                    if qty_sum > 0:
                        price_base = round(value_sum / qty_sum, 4)
                if price_base is None:
                    price_base = 0.0

                if taken > 0:
                    for s, take in chunks:
                        stock_price = float(price_base)
                        line_total_chunk = stock_price * int(take)
                        self.conn.execute(
                            """INSERT INTO bill_items
                            (bill_id,item_type,school,color,size,unit_price,qty,line_total,origin)
                            VALUES(?,?,?,?,?,?,?,?,?)""",
                            (
                                bill_id,
                                line["item_type"], line["school"], line["color"], line["size"],
                                stock_price, int(take), line_total_chunk, "STOCK",
                            ),
                        )
                        total += line_total_chunk

                if remaining > 0 and allow_factory:
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
                        (bill_id,item_type,school,color,size,unit_price,qty,line_total,origin)
                        VALUES(?,?,?,?,?,?,?,?,?)""",
                        (
                            bill_id,
                            line["item_type"], line["school"], line["color"], line["size"],
                            factory_price, int(remaining), factory_line_total, "FACTORY",
                        ),
                    )
                    total += factory_line_total

                for s, take in chunks:
                    self.conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?", (int(take), int(s["id"])))
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            now_iso(), "OUT", int(s["id"]), int(take), "Bill deduction", bill_id,
                            s["item_type"], s["school"], s["color"], s["size"],
                            float(s["unit_price"]),
                        ),
                    )

                if remaining > 0 and allow_factory:
                    self.conn.execute(
                        """INSERT INTO movements
                        (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            now_iso(), "OUT_FACTORY", None, int(remaining), "Factory direct", bill_id,
                            line["item_type"], line["school"], line["color"], line["size"],
                            float(price_base),
                        ),
                    )

            self.cleanup_unused_specs()
            self.conn.execute("UPDATE bills SET total=? WHERE id=?", (float(total), bill_id))
            wh_tgt = _extract_warehouse_target(customer)
            if wh_tgt:
                self.conn.execute(
                    "UPDATE bills SET bill_type=?, total=? WHERE id=?",
                    (WAREHOUSE_RETURN_BILL_TYPE, 0.0, bill_id),
                )

            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id = ?", (bill_id,)
            ).fetchone()
            bill_uuid = bill_uuid_row[0] if bill_uuid_row else None
            items_payload = [
                {
                    "item_type": r[0], "school": r[1], "color": r[2], "size": r[3],
                    "unit_price": float(r[4]), "qty": int(r[5]),
                    "line_total": float(r[6]), "origin": r[7],
                }
                for r in self.conn.execute(
                    "SELECT item_type,school,color,size,unit_price,qty,line_total,origin "
                    "FROM bill_items WHERE bill_id = ?", (bill_id,)
                ).fetchall()
            ]
            warehouse_target = _extract_warehouse_target(customer)
            branch_target = _extract_branch_target(customer)
            source_device = (self._current_device_name() or "").strip()
            if branch_target and branch_target == source_device:
                branch_target = None
            if warehouse_target:
                self._record_warehouse_return_event(
                    return_uuid=bill_uuid or "",
                    note=f"bill #{bill_id}",
                    lines=[
                        {
                            "item_type": it["item_type"],
                            "school": it["school"],
                            "color": it["color"],
                            "size": it["size"],
                            "unit_price": float(it["unit_price"]),
                            "qty": int(it["qty"]),
                        }
                        for it in items_payload if int(it["qty"]) > 0
                    ],
                )
            elif branch_target:
                # Branch-to-branch via warehouse is a transfer request, not a retail cash sale.
                self.conn.execute(
                    "UPDATE bills SET bill_type=? WHERE id=?",
                    (BRANCH_TRANSFER_BILL_TYPE, bill_id),
                )
                self._record_transfer_via_warehouse_event(
                    request_uuid=bill_uuid or "",
                    target_device=branch_target,
                    note=f"bill #{bill_id}",
                    lines=[
                        {
                            "item_type": it["item_type"],
                            "school": it["school"],
                            "color": it["color"],
                            "size": it["size"],
                            "unit_price": float(it["unit_price"]),
                            "qty": int(it["qty"]),
                        }
                        for it in items_payload if int(it["qty"]) > 0
                    ],
                )
            else:
                self._record_sync_event_or_raise("SALE_CREATED", {
                    "bill_uuid": bill_uuid,
                    "bill_id": bill_id,
                    "customer": (customer or "").strip() or None,
                    "customer_phone": customer_phone or None,
                    "total": float(total),
                    "payment_method": payment_method,
                    "items": items_payload,
                    "shift_id": self.active_shift_id,
                })

            bill_type_row = self.conn.execute(
                "SELECT COALESCE(bill_type,'SALE') FROM bills WHERE id=?", (bill_id,)
            ).fetchone()
            self._audit("sale_created", {
                "bill_id": bill_id,
                "bill_uuid": bill_uuid,
                "bill_type": bill_type_row[0] if bill_type_row else "SALE",
                "customer": (customer or "").strip() or None,
                "customer_phone": customer_phone or None,
                "total": float(total),
                "payment_method": payment_method,
                "warehouse_target": warehouse_target,
                "branch_target": branch_target,
                "item_count": len(items_payload),
                "qty_total": sum(int(it.get("qty") or 0) for it in items_payload),
                "items": items_payload,
            }, actor="cashier")
            return bill_id

    # -------- Return Bill Methods --------
    def create_return_bill(
        self,
        customer: str,
        return_lines: List[Dict[str, Any]],
        customer_phone: str = "",
        payment_method: str = PAYMENT_METHOD_CASH,
    ) -> int:
        """Create a return bill – adds items back to stock."""
        self._require_shift()
        if not return_lines:
            raise ValueError("لا توجد أصناف في فاتورة المرتجع")
        customer_phone = _normalize_customer_phone(customer_phone)
        payment_method = _normalize_payment_method(payment_method)

        with self.conn:
            bill_cur = self.conn.execute(
                "INSERT INTO bills(created_at,customer,customer_phone,total,bill_type,status,payment_method,shift_id) VALUES(?,?,?,?,?,?,?,?)",
                (now_iso(), (customer or "").strip() or None, customer_phone or None, 0.0, "RETURN", "CONFIRMED", payment_method, self.active_shift_id),
            )
            bill_id = int(bill_cur.lastrowid)
            total = 0.0

            for line in return_lines:
                qty = int(line["qty"])
                if qty <= 0:
                    raise ValueError("الكمية يجب أن تكون أكبر من 0")
                price = float(line.get("unit_price") or 0)
                line_total = price * qty
                total += line_total

                # Add stock back
                stock_id = self.add_or_update_stock_row(
                    line["item_type"].strip(), line["school"].strip(),
                    line["color"].strip(), line["size"].strip(),
                    price, qty,
                )

                # Record bill item
                self.conn.execute(
                    """INSERT INTO bill_items(bill_id,item_type,school,color,size,unit_price,qty,line_total,origin)
                    VALUES(?,?,?,?,?,?,?,?,?)""",
                    (bill_id, line["item_type"], line["school"], line["color"],
                     line["size"], price, qty, line_total, "RETURN"),
                )

                # Record movement
                self.conn.execute(
                    """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), "RETURN_IN", stock_id, qty,
                     "مرتجع", bill_id,
                     line["item_type"], line["school"], line["color"],
                     line["size"], price),
                )

                self._upsert_history({
                    "item_type": line.get("item_type"),
                    "school": line.get("school"),
                    "color": line.get("color"),
                    "size": line.get("size"),
                })

            self.conn.execute("UPDATE bills SET total=? WHERE id=?", (float(total), bill_id))

            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id = ?", (bill_id,)
            ).fetchone()
            self._record_sync_event_or_raise("SALE_RETURNED", {
                "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                "bill_id": bill_id,
                "customer": (customer or "").strip() or None,
                "customer_phone": customer_phone or None,
                "total": float(total),
                "payment_method": payment_method,
                "lines": [
                    {
                        "item_type": ln.get("item_type"),
                        "school":    ln.get("school"),
                        "color":     ln.get("color"),
                        "size":      ln.get("size"),
                        "unit_price": float(ln.get("unit_price") or 0),
                        "qty":        int(ln["qty"]),
                    }
                    for ln in return_lines
                ],
                "shift_id": self.active_shift_id,
            })
            self._audit("return_bill_created", {
                "bill_id": bill_id,
                "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                "customer": (customer or "").strip() or None,
                "customer_phone": customer_phone or None,
                "total": float(total),
                "payment_method": payment_method,
                "item_count": len(return_lines),
                "qty_total": sum(int(ln.get("qty") or 0) for ln in return_lines),
                "lines": return_lines,
            }, actor="cashier")
            return bill_id

    def create_exchange_bill(self, customer: str,
                             return_lines: List[Dict[str, Any]],
                             take_lines: List[Dict[str, Any]],
                             customer_phone: str = "",
                             payment_method: str = PAYMENT_METHOD_CASH) -> int:
        """Create an exchange bill – returns items to stock and takes new items."""
        self._require_shift()
        if not return_lines and not take_lines:
            raise ValueError("لا توجد أصناف في فاتورة الاستبدال")
        customer_phone = _normalize_customer_phone(customer_phone)
        payment_method = _normalize_payment_method(payment_method)

        def _candidates_for_line(line: Dict[str, Any]) -> List[Any]:
            cur = self.conn.cursor()
            try:
                where_parts = ["count > 0"]
                args: List[Any] = []
                for k in ("item_type", "school", "color", "size"):
                    v = (line.get(k) or "").strip()
                    if v:
                        where_parts.append(f"LOWER(TRIM({k})) = LOWER(?)")
                        args.append(v)
                cur.execute(f"SELECT * FROM stocks WHERE {' AND '.join(where_parts)} ORDER BY id ASC", args)
                return cur.fetchall()
            finally:
                cur.close()

        with self.conn:
            bill_cur = self.conn.execute(
                "INSERT INTO bills(created_at,customer,customer_phone,total,bill_type,status,payment_method,shift_id) VALUES(?,?,?,?,?,?,?,?)",
                (now_iso(), (customer or "").strip() or None, customer_phone or None, 0.0, "EXCHANGE", "CONFIRMED", payment_method, self.active_shift_id),
            )
            bill_id = int(bill_cur.lastrowid)
            return_total = 0.0
            take_total = 0.0

            # --- Process returned items (add to stock) ---
            for line in return_lines:
                qty = int(line["qty"])
                if qty <= 0:
                    raise ValueError("الكمية يجب أن تكون أكبر من 0")
                price = float(line.get("unit_price") or 0)
                line_total = price * qty
                return_total += line_total

                stock_id = self.add_or_update_stock_row(
                    line["item_type"].strip(), line["school"].strip(),
                    line["color"].strip(), line["size"].strip(),
                    price, qty,
                )

                self.conn.execute(
                    """INSERT INTO bill_items(bill_id,item_type,school,color,size,unit_price,qty,line_total,origin)
                    VALUES(?,?,?,?,?,?,?,?,?)""",
                    (bill_id, line["item_type"], line["school"], line["color"],
                     line["size"], price, qty, line_total, "RETURN"),
                )
                self.conn.execute(
                    """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), "RETURN_IN", stock_id, qty,
                     "استبدال - مرتجع", bill_id,
                     line["item_type"], line["school"], line["color"],
                     line["size"], price),
                )
                self._upsert_history({
                    "item_type": line.get("item_type"),
                    "school": line.get("school"),
                    "color": line.get("color"),
                    "size": line.get("size"),
                })

            # --- Process taken items (deduct from stock) ---
            for line in take_lines:
                qty_needed = int(line["qty"])
                if qty_needed <= 0:
                    raise ValueError("الكمية يجب أن تكون أكبر من 0")

                cands = _candidates_for_line(line)
                remaining = qty_needed
                chunks: List[Tuple[Any, int]] = []
                for s in cands:
                    if remaining <= 0:
                        break
                    take = min(int(s["count"]), remaining)
                    if take > 0:
                        chunks.append((s, take))
                        remaining -= take

                if remaining > 0:
                    raise ValueError(
                        f"لا توجد كمية كافية للصنف "
                        f"{line.get('item_type','?')} / {line.get('school','?')} / {line.get('size','?')} "
                        f"(المطلوب {qty_needed}، المتاح {qty_needed - remaining})."
                    )

                price = float(line.get("unit_price") or 0)

                for s, take_qty in chunks:
                    lt = price * int(take_qty)
                    take_total += lt
                    self.conn.execute(
                        """INSERT INTO bill_items(bill_id,item_type,school,color,size,unit_price,qty,line_total,origin)
                        VALUES(?,?,?,?,?,?,?,?,?)""",
                        (bill_id, line["item_type"], line["school"], line["color"],
                         line["size"], price, int(take_qty), lt, "STOCK"),
                    )
                    self.conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?",
                                     (int(take_qty), int(s["id"])))
                    self.conn.execute(
                        """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (now_iso(), "OUT", int(s["id"]), int(take_qty),
                         "استبدال - مأخوذ", bill_id,
                         s["item_type"], s["school"], s["color"],
                         s["size"], float(s["unit_price"])),
                    )

            self.cleanup_unused_specs()
            diff = take_total - return_total
            self.conn.execute("UPDATE bills SET total=? WHERE id=?", (float(diff), bill_id))

            bill_uuid_row = self.conn.execute(
                "SELECT uuid FROM bills WHERE id = ?", (bill_id,)
            ).fetchone()
            self._record_sync_event_or_raise("SALE_EXCHANGED", {
                "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                "bill_id": bill_id,
                "customer": (customer or "").strip() or None,
                "customer_phone": customer_phone or None,
                "return_total": float(return_total),
                "take_total": float(take_total),
                "diff": float(diff),
                "payment_method": payment_method,
                "return_lines": [
                    {
                        "item_type": ln.get("item_type"),
                        "school":    ln.get("school"),
                        "color":     ln.get("color"),
                        "size":      ln.get("size"),
                        "unit_price": float(ln.get("unit_price") or 0),
                        "qty":        int(ln["qty"]),
                    }
                    for ln in return_lines
                ],
                "take_lines": [
                    {
                        "item_type": ln.get("item_type"),
                        "school":    ln.get("school"),
                        "color":     ln.get("color"),
                        "size":      ln.get("size"),
                        "unit_price": float(ln.get("unit_price") or 0),
                        "qty":        int(ln["qty"]),
                    }
                    for ln in take_lines
                ],
                "shift_id": self.active_shift_id,
            })
            self._audit("exchange_bill_created", {
                "bill_id": bill_id,
                "bill_uuid": bill_uuid_row[0] if bill_uuid_row else None,
                "customer": (customer or "").strip() or None,
                "customer_phone": customer_phone or None,
                "return_total": float(return_total),
                "take_total": float(take_total),
                "diff": float(diff),
                "payment_method": payment_method,
                "return_lines": return_lines,
                "take_lines": take_lines,
            }, actor="cashier")
            return bill_id

    # -------- New Reservation Methods --------
    def create_reservation(
        self,
        customer,
        lines,
        paid_amount=0.0,
        customer_phone: str = "",
        payment_method: str = PAYMENT_METHOD_CASH,
    ):
        """Create reservation records for items."""
        self._require_shift()
        if not lines:
            raise ValueError("لا توجد أصناف للحجز")
        customer_phone = _normalize_customer_phone(customer_phone)
        payment_method = _normalize_payment_method(payment_method)
        created = []
        import uuid as _uuid
        group_uuid = f"grp-{_uuid.uuid4()}"
        line_totals = [float(line["unit_price"]) * int(line["qty"]) for line in lines]
        paid_allocations = _allocate_reservation_down_payments(line_totals, float(paid_amount or 0.0))
        with self.conn:
            for idx, line in enumerate(lines):
                total = float(line_totals[idx])
                alloc_paid = float(paid_allocations[idx] if idx < len(paid_allocations) else 0.0)
                cur = self.conn.execute(
                    """INSERT INTO reservations(created_at,customer,customer_phone,item_type,school,color,size,qty,unit_price,total_amount,paid_amount,payment_method,status,note,reservation_group_uuid,shift_id)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), (customer or "").strip(), customer_phone or None, line["item_type"], line["school"],
                     line["color"], line["size"], int(line["qty"]), float(line["unit_price"]),
                     total, alloc_paid, payment_method, RESERVATION_STATUS_PENDING, line.get("note", ""), group_uuid, self.active_shift_id),
                )
                created.append(cur.lastrowid)
                self.conn.execute(
                    """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), "RESERVE", None, int(line["qty"]), f"حجز - {customer}",
                     None, line["item_type"], line["school"], line["color"], line["size"],
                     float(line["unit_price"]), self.active_shift_id),
                )
                if alloc_paid > 1e-9:
                    self.conn.execute(
                        """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (now_iso(), "RESERVE_PAY", None, 0, f"عربون حجز #{cur.lastrowid}",
                         None, line["item_type"], line["school"], line["color"], line["size"],
                         alloc_paid, self.active_shift_id, payment_method),
                    )
            if created:
                uuid_rows = self.conn.execute(
                    "SELECT id, uuid FROM reservations WHERE id IN (%s)"
                    % ",".join("?" * len(created)),
                    tuple(int(x) for x in created),
                ).fetchall()
                id_to_uuid = {int(r[0]): r[1] for r in uuid_rows}
                self._record_sync_event_or_raise("RESERVATION_CREATED", {
                    "reservation_group_uuid": group_uuid,
                    "customer": (customer or "").strip() or None,
                    "customer_phone": customer_phone or None,
                    "paid_amount": float(paid_amount),
                    "payment_method": payment_method,
                    "reservations": [
                        {
                            "reservation_uuid": id_to_uuid.get(int(rid)),
                            "reservation_id":   int(rid),
                            "item_type": ln.get("item_type"),
                            "school":    ln.get("school"),
                            "color":     ln.get("color"),
                            "size":      ln.get("size"),
                            "qty":        int(ln["qty"]),
                            "unit_price": float(ln["unit_price"]),
                            "paid_amount": float(paid_allocations[idx] if idx < len(paid_allocations) else 0.0),
                        }
                        for idx, (rid, ln) in enumerate(zip(created, lines))
                    ],
                    "shift_id": self.active_shift_id,
                })
                self._audit("reservation_created", {
                    "reservation_group_uuid": group_uuid,
                    "reservation_ids": [int(x) for x in created],
                    "customer": (customer or "").strip() or None,
                    "customer_phone": customer_phone or None,
                    "total": float(sum(line_totals)),
                    "paid_amount": float(paid_amount or 0.0),
                    "payment_method": payment_method,
                    "item_count": len(lines),
                    "qty_total": sum(int(ln.get("qty") or 0) for ln in lines),
                    "lines": [
                        {
                            "reservation_id": int(rid),
                            "item_type": ln.get("item_type"),
                            "school": ln.get("school"),
                            "color": ln.get("color"),
                            "size": ln.get("size"),
                            "qty": int(ln["qty"]),
                            "unit_price": float(ln["unit_price"]),
                            "paid_amount": float(paid_allocations[idx] if idx < len(paid_allocations) else 0.0),
                        }
                        for idx, (rid, ln) in enumerate(zip(created, lines))
                    ],
                }, actor="cashier")
        return created

    def get_available_qty_for_reservation(self, item_type: str, school: str, color: str, size: str) -> int:
        row = self.conn.execute(
            """
            SELECT COALESCE(SUM(count),0)
            FROM stocks
            WHERE LOWER(TRIM(item_type))=LOWER(TRIM(?))
              AND LOWER(TRIM(school))=LOWER(TRIM(?))
              AND LOWER(TRIM(color))=LOWER(TRIM(?))
              AND LOWER(TRIM(size))=LOWER(TRIM(?))
            """,
            (item_type, school, color, size),
        ).fetchone()
        on_hand = int((row[0] if row else 0) or 0)
        return max(0, on_hand)

    def get_next_reservation_alert(self) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            """
            SELECT id, customer, item_type, school, color, size, qty, hold_until, note, created_at
            FROM reservation_alerts
            WHERE shown_at IS NULL
            ORDER BY id ASC
            LIMIT 1
            """
        ).fetchone()
        return dict(row) if row else None

    def mark_reservation_alert_shown(self, alert_id: int) -> None:
        self.conn.execute(
            "UPDATE reservation_alerts SET shown_at=? WHERE id=?",
            (now_iso(), int(alert_id)),
        )

    def get_next_incoming_shipment_alert(self) -> Optional[Dict[str, Any]]:
        row = self.conn.execute(
            """
            SELECT id, sync_event_uuid, shipment_uuid, from_device, note, total_qty, created_at
            FROM incoming_shipment_alerts AS a
            WHERE EXISTS (
                SELECT 1
                FROM incoming_shipment_items_pending AS p
                WHERE p.shipment_uuid = a.shipment_uuid
                  AND p.status = 'PENDING'
            )
              AND a.shown_at IS NULL
            ORDER BY a.id ASC
            LIMIT 1
            """
        ).fetchone()
        return dict(row) if row else None

    def reset_incoming_shipment_alert_shown(self, alert_id: int) -> None:
        self.conn.execute(
            "UPDATE incoming_shipment_alerts SET shown_at=NULL WHERE id=?",
            (int(alert_id),),
        )

    def mark_incoming_shipment_alert_shown(self, alert_id: int) -> None:
        self.conn.execute(
            "UPDATE incoming_shipment_alerts SET shown_at=? WHERE id=?",
            (now_iso(), int(alert_id)),
        )

    def mark_incoming_shipment_confirmed(self, shipment_uuid: str) -> None:
        self.conn.execute(
            """
            UPDATE incoming_shipment_alerts
            SET shown_at=COALESCE(shown_at, ?)
            WHERE shipment_uuid=?
            """,
            (now_iso(), str(shipment_uuid or "").strip()),
        )

    def list_pending_shipment_items(self, shipment_uuid: str) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            """
            SELECT * FROM incoming_shipment_items_pending
            WHERE shipment_uuid=? AND status='PENDING'
            ORDER BY line_index ASC
            """,
            (str(shipment_uuid or "").strip(),),
        )
        return [dict(r) for r in cur.fetchall()]

    def list_grouped_pending_shipment_items(self, shipment_uuid: str) -> List[Dict[str, Any]]:
        rows = self.list_pending_shipment_items(shipment_uuid)
        grouped: Dict[Tuple[str, str, str, str, float], Dict[str, Any]] = {}
        ordered: List[Dict[str, Any]] = []
        for row in rows:
            key = (
                str(row.get("item_type") or ""),
                str(row.get("school") or ""),
                str(row.get("color") or ""),
                str(row.get("size") or ""),
                float(row.get("unit_price") or 0.0),
            )
            bucket = grouped.get(key)
            if bucket is None:
                bucket = {
                    "group_key": len(ordered),
                    "item_type": key[0],
                    "school": key[1],
                    "color": key[2],
                    "size": key[3],
                    "unit_price": key[4],
                    "expected_qty": 0,
                    "line_indexes": [],
                    "source_rows": [],
                }
                grouped[key] = bucket
                ordered.append(bucket)
            bucket["expected_qty"] = int(bucket["expected_qty"]) + int(row.get("expected_qty") or 0)
            bucket["line_indexes"].append(int(row["line_index"]))
            bucket["source_rows"].append(dict(row))
        return ordered

    def list_pending_incoming_shipments(self) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            """
            SELECT
                a.id,
                a.sync_event_uuid,
                a.shipment_uuid,
                a.from_device,
                a.note,
                a.total_qty,
                a.created_at,
                a.shown_at,
                COUNT(p.id) AS pending_lines,
                COALESCE(SUM(p.expected_qty), 0) AS pending_qty
            FROM incoming_shipment_alerts AS a
            JOIN incoming_shipment_items_pending AS p
              ON p.shipment_uuid = a.shipment_uuid
             AND p.status = 'PENDING'
            GROUP BY
                a.id, a.sync_event_uuid, a.shipment_uuid, a.from_device,
                a.note, a.total_qty, a.created_at, a.shown_at
            ORDER BY a.created_at ASC, a.id ASC
            """
        )
        return [dict(r) for r in cur.fetchall()]

    def confirm_incoming_shipment(
        self,
        shipment_uuid: str,
        receipt_lines: List[Dict[str, Any]],
        note: str = "",
    ) -> Dict[str, Any]:
        """Apply cashier-confirmed shipment quantities and report differences."""
        self._require_shift()
        ship = str(shipment_uuid or "").strip()
        if not ship:
            raise ValueError("shipment_uuid مطلوب.")
        rows = self.list_pending_shipment_items(ship)
        if not rows:
            raise ValueError("لا توجد بنود شحنة معلقة لهذا المرجع.")
        by_idx = {int(r["line_index"]): r for r in rows}
        recv_by_idx: Dict[int, int] = {}
        for ln in (receipt_lines or []):
            idx = int(ln.get("line_index"))
            qty = int(ln.get("received_qty"))
            if idx not in by_idx:
                raise ValueError("يوجد بند غير صالح في التأكيد.")
            if qty < 0:
                raise ValueError("الكمية المستلمة لا يمكن أن تكون سالبة.")
            recv_by_idx[idx] = qty
        if set(recv_by_idx.keys()) != set(by_idx.keys()):
            raise ValueError("يجب إدخال الكمية المستلمة لكل بند.")

        from_dev_row = self.conn.execute(
            "SELECT from_device FROM incoming_shipment_alerts WHERE shipment_uuid=? ORDER BY id DESC LIMIT 1",
            (ship,),
        ).fetchone()
        from_device = str(from_dev_row[0] if from_dev_row else "").strip() or "WAREHOUSE-MAIN"
        diffs: List[Dict[str, Any]] = []

        with self.conn:
            for idx, src in by_idx.items():
                expected = int(src["expected_qty"] or 0)
                received = int(recv_by_idx.get(idx, 0))
                if received > 0:
                    self.conn.execute(
                        """INSERT INTO stocks(item_type, school, color, size, unit_price, count)
                           VALUES(?,?,?,?,?,?)""",
                        (
                            src["item_type"], src["school"], src["color"], src["size"],
                            float(src["unit_price"] or 0.0), received,
                        ),
                    )
                    self.conn.execute(
                        """INSERT INTO movements
                           (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            now_iso(), "IN", None, received,
                            f"استلام شحنة مؤكد #{ship[:8]}", None,
                            src["item_type"], src["school"], src["color"], src["size"],
                            float(src["unit_price"] or 0.0),
                        ),
                    )
                self.conn.execute(
                    "UPDATE incoming_shipment_items_pending SET received_qty=?, status='CONFIRMED' WHERE id=?",
                    (received, int(src["id"])),
                )
                if expected != received:
                    diffs.append(
                        {
                            "line_index": idx,
                            "item_type": src["item_type"],
                            "school": src["school"],
                            "color": src["color"],
                            "size": src["size"],
                            "unit_price": float(src["unit_price"] or 0.0),
                            "expected_qty": expected,
                            "received_qty": received,
                            "diff_qty": int(received - expected),
                        }
                    )

            self._record_targeted_inventory_event(
                "SHIPMENT_RECEIPT_REPORTED",
                "warehouse",
                {
                    "shipment_uuid": ship,
                    "from_device": from_device,
                    "confirmed_at": now_iso(),
                    "lines": diffs or [
                        {
                            "line_index": int(r["line_index"]),
                            "item_type": r["item_type"],
                            "school": r["school"],
                            "color": r["color"],
                            "size": r["size"],
                            "unit_price": float(r["unit_price"] or 0.0),
                            "expected_qty": int(r["expected_qty"] or 0),
                            "received_qty": int(recv_by_idx.get(int(r["line_index"]), 0)),
                            "diff_qty": int(recv_by_idx.get(int(r["line_index"]), 0) - int(r["expected_qty"] or 0)),
                        }
                        for r in rows
                    ],
                    "has_diff": bool(diffs),
                    "note": (note or "").strip() or None,
                    "shift_id": self.active_shift_id,
                },
            )
            self.mark_incoming_shipment_confirmed(ship)
        summary = {"shipment_uuid": ship, "has_diff": bool(diffs), "lines": len(rows)}
        self._audit("incoming_shipment_confirmed", {
            **summary,
            "from_device": from_device,
            "note": (note or "").strip() or None,
            "lines_detail": [
                {
                    "line_index": int(r["line_index"]),
                    "item_type": r["item_type"],
                    "school": r["school"],
                    "color": r["color"],
                    "size": r["size"],
                    "unit_price": float(r["unit_price"] or 0.0),
                    "expected_qty": int(r["expected_qty"] or 0),
                    "received_qty": int(recv_by_idx.get(int(r["line_index"]), 0)),
                    "diff_qty": int(recv_by_idx.get(int(r["line_index"]), 0) - int(r["expected_qty"] or 0)),
                }
                for r in rows
            ],
        }, actor="cashier")
        return summary

    def confirm_grouped_incoming_shipment(
        self,
        shipment_uuid: str,
        grouped_receipt_lines: List[Dict[str, Any]],
        note: str = "",
    ) -> Dict[str, Any]:
        rows = self.list_pending_shipment_items(shipment_uuid)
        by_idx = {int(r["line_index"]): dict(r) for r in rows}
        payload: List[Dict[str, Any]] = []
        seen: Set[int] = set()

        for group in (grouped_receipt_lines or []):
            line_indexes = [int(x) for x in (group.get("line_indexes") or [])]
            if not line_indexes:
                raise ValueError("يوجد بند مجمع غير صالح في التأكيد.")
            received_total = int(group.get("received_qty") or 0)
            if received_total < 0:
                raise ValueError("الكمية المستلمة لا يمكن أن تكون سالبة.")
            expected_total = 0
            for idx in line_indexes:
                if idx not in by_idx:
                    raise ValueError("يوجد بند مجمع غير صالح في التأكيد.")
                expected_total += int(by_idx[idx].get("expected_qty") or 0)

            remaining = received_total
            for idx in line_indexes:
                expected = int(by_idx[idx].get("expected_qty") or 0)
                assigned = min(expected, remaining)
                payload.append({"line_index": idx, "received_qty": assigned})
                seen.add(idx)
                remaining -= assigned
            if remaining > 0:
                payload[-1]["received_qty"] = int(payload[-1]["received_qty"]) + int(remaining)

        if set(by_idx.keys()) != seen:
            raise ValueError("يجب إدخال الكمية المستلمة لكل بند.")

        return self.confirm_incoming_shipment(shipment_uuid, payload, note)

    def list_reservations(self, status=None, date_from=None, date_to=None, school=None, item_type=None, color=None):
        where = ["1=1"]
        args = []
        if status:
            where.append("status = ?")
            args.append(status)
        if date_from:
            where.append("date(created_at) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(created_at) <= date(?)")
            args.append(date_to)
        if school:
            where.append("LOWER(TRIM(school)) = LOWER(?)")
            args.append(school.strip())
        if item_type:
            where.append("LOWER(TRIM(item_type)) = LOWER(?)")
            args.append(item_type.strip())
        if color:
            where.append("LOWER(TRIM(color)) = LOWER(?)")
            args.append(color.strip())
        cur = self.conn.execute(
            f"SELECT * FROM reservations WHERE {' AND '.join(where)} ORDER BY id DESC", args)
        return [dict(r) for r in cur.fetchall()]

    def count_reservations(self, status=None, date_from=None, date_to=None, school=None, item_type=None, color=None) -> int:
        where = ["1=1"]
        args = []
        if status:
            where.append("status = ?")
            args.append(status)
        if date_from:
            where.append("date(created_at) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(created_at) <= date(?)")
            args.append(date_to)
        if school:
            where.append("LOWER(TRIM(school)) = LOWER(?)")
            args.append(school.strip())
        if item_type:
            where.append("LOWER(TRIM(item_type)) = LOWER(?)")
            args.append(item_type.strip())
        if color:
            where.append("LOWER(TRIM(color)) = LOWER(?)")
            args.append(color.strip())
        row = self.conn.execute(
            f"SELECT COUNT(*) AS c FROM reservations WHERE {' AND '.join(where)}",
            args,
        ).fetchone()
        return int(row["c"] if row else 0)

    def list_reservation_totals(self, status=None, date_from=None, date_to=None, school=None, item_type=None, color=None) -> List[Dict[str, Any]]:
        where = ["1=1"]
        args = []
        if status:
            where.append("status = ?")
            args.append(status)
        if date_from:
            where.append("date(created_at) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(created_at) <= date(?)")
            args.append(date_to)
        if school:
            where.append("LOWER(TRIM(school)) = LOWER(?)")
            args.append(str(school).strip())
        if item_type:
            where.append("LOWER(TRIM(item_type)) = LOWER(?)")
            args.append(str(item_type).strip())
        if color:
            where.append("LOWER(TRIM(color)) = LOWER(?)")
            args.append(str(color).strip())
        cur = self.conn.execute(
            f"""
            SELECT TRIM(item_type) AS item_type,
                   TRIM(school) AS school,
                   TRIM(color) AS color,
                   TRIM(size) AS size,
                   unit_price,
                   SUM(qty) AS qty,
                   SUM(total_amount) AS total_amount,
                   SUM(paid_amount) AS paid_amount,
                   COUNT(*) AS lines_count
              FROM reservations
             WHERE {' AND '.join(where)}
             GROUP BY TRIM(item_type), TRIM(school), TRIM(color), TRIM(size), unit_price
             ORDER BY TRIM(school), TRIM(color), TRIM(item_type), TRIM(size), unit_price
            """,
            args,
        )
        return [dict(r) for r in cur.fetchall()]

    def list_reservation_group_items(self, reservation_id: int) -> List[Dict[str, Any]]:
        rid = int(reservation_id)
        row = self.conn.execute(
            "SELECT reservation_group_uuid FROM reservations WHERE id=?",
            (rid,),
        ).fetchone()
        if not row:
            return []
        group_uuid = str(row[0] or "").strip() or f"legacy-{rid}"
        if group_uuid.startswith("legacy-"):
            row2 = self.conn.execute("SELECT * FROM reservations WHERE id=? ORDER BY id ASC", (rid,)).fetchone()
            return [dict(row2)] if row2 else []
        rows = self.conn.execute(
            "SELECT * FROM reservations WHERE reservation_group_uuid=? ORDER BY id ASC",
            (group_uuid,),
        ).fetchall()
        return [dict(r) for r in rows]

    def _stock_qty_for_specs(self, item_type: str, school: str, color: str, size: str) -> int:
        row = self.conn.execute(
            """
            SELECT COALESCE(SUM(count), 0)
            FROM stocks
            WHERE LOWER(TRIM(item_type)) = LOWER(TRIM(?))
              AND LOWER(TRIM(school)) = LOWER(TRIM(?))
              AND LOWER(TRIM(color)) = LOWER(TRIM(?))
              AND LOWER(TRIM(size)) = LOWER(TRIM(?))
            """,
            (item_type, school, color, size),
        ).fetchone()
        return int((row[0] if row else 0) or 0)

    def _ensure_reservation_delivery_stock(self, rows: Sequence[sqlite3.Row]) -> None:
        required: Dict[Tuple[str, str, str, str], int] = {}
        for row in rows:
            key = (
                str(row["item_type"] or "").strip(),
                str(row["school"] or "").strip(),
                str(row["color"] or "").strip(),
                str(row["size"] or "").strip(),
            )
            required[key] = required.get(key, 0) + int(row["qty"] or 0)

        shortages: List[str] = []
        for (item_type, school, color, size), need_qty in required.items():
            have_qty = self._stock_qty_for_specs(item_type, school, color, size)
            if have_qty < need_qty:
                shortages.append(
                    f"{item_type} / {school} / {color} / {size}: المتاح {have_qty} والمطلوب {need_qty}"
                )

        if shortages:
            raise ValueError(
                "لا يمكن تسليم الحجز لأن بعض الأصناف غير موجودة في المخزون الحالي:\n"
                + "\n".join(shortages)
            )

    def update_reservation_payment(self, res_id, new_paid, payment_method: str = PAYMENT_METHOD_CASH):
        payment_method = _normalize_payment_method(payment_method)
        rid = int(res_id)
        row_prev = self.conn.execute(
            "SELECT paid_amount, total_amount, item_type, school, color, size FROM reservations WHERE id=?",
            (rid,),
        ).fetchone()
        if not row_prev:
            raise ValueError("الحجز غير موجود")
        old_paid = float(row_prev["paid_amount"] or 0.0)
        new_paid_f = float(new_paid)
        total_amount = float(row_prev["total_amount"] or 0.0)
        if new_paid_f < -1e-9:
            raise ValueError("المبلغ المدفوع لا يمكن أن يكون أقل من صفر")
        if new_paid_f - total_amount > 1e-6:
            raise ValueError("المبلغ المدفوع لا يمكن أن يزيد عن إجمالي الحجز")
        self.conn.execute(
            "UPDATE reservations SET paid_amount=? WHERE id=?",
            (new_paid_f, rid))
        delta = new_paid_f - old_paid
        if abs(delta) > 1e-9:
            self.conn.execute(
                """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (now_iso(), "RESERVE_PAY", None, 0, f"تعديل عربون حجز #{rid}",
                 None, row_prev["item_type"], row_prev["school"], row_prev["color"], row_prev["size"],
                 delta, self.active_shift_id, payment_method),
            )
        row = self.conn.execute(
            "SELECT uuid FROM reservations WHERE id=?", (rid,)
        ).fetchone()
        self._record_sync_event_or_raise("RESERVATION_PAYMENT_UPDATED", {
            "reservation_uuid": row[0] if row else None,
            "reservation_id":   rid,
            "paid_amount":      new_paid_f,
            "payment_method":   payment_method,
        })
        self._audit("reservation_payment_updated", {
            "reservation_id": rid,
            "reservation_uuid": row[0] if row else None,
            "old_paid_amount": old_paid,
            "new_paid_amount": new_paid_f,
            "delta": float(delta),
            "total_amount": total_amount,
            "payment_method": payment_method,
            "item_type": row_prev["item_type"],
            "school": row_prev["school"],
            "color": row_prev["color"],
            "size": row_prev["size"],
        }, actor="cashier")

    def complete_reservation(self, res_id):
        self.deliver_reservation(int(res_id), collected_amount=0.0)

    def deliver_reservation(self, res_id: int, collected_amount: float = 0.0, payment_method: str = PAYMENT_METHOD_CASH):
        """Mark reservation as delivered, collect remaining payment, record movement."""
        payment_method = _normalize_payment_method(payment_method)
        self._require_shift()
        rid = int(res_id)
        cur = self.conn.execute("SELECT * FROM reservations WHERE id=?", (rid,))
        row = cur.fetchone()
        if not row:
            raise ValueError("الحجز غير موجود")
        if _is_reservation_delivered(row["status"]):
            raise ValueError("تم تسليم هذا الحجز مسبقاً")
        if _is_reservation_cancelled(row["status"]):
            raise ValueError("تم إلغاء هذا الحجز.")
        self._ensure_reservation_delivery_stock([row])
        collected = float(collected_amount)
        if collected < -1e-9:
            raise ValueError("المبلغ المحصل لا يمكن أن يكون أقل من صفر")
        total_amount = float(row["total_amount"] or 0.0)
        new_paid = float(row["paid_amount"]) + collected
        if new_paid - total_amount > 1e-6:
            raise ValueError("المبلغ المحصل يزيد عن المتبقي من الحجز")
        with self.conn:
            self.deduct_stock_for_specs(
                row["item_type"], row["school"], row["color"], row["size"],
                int(row["qty"] or 0),
                note=f"Reservation delivered #{rid}",
            )
            self.conn.execute(
                "UPDATE reservations SET status=?, paid_amount=? WHERE id=?",
                (RESERVATION_STATUS_DELIVERED, new_paid, rid))
            if collected > 1e-9:
                self.conn.execute(
                    """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (now_iso(), "DELIVER_PAY", None, 0,
                     f"تحصيل باقي حجز #{rid}",
                     None, row["item_type"], row["school"], row["color"], row["size"],
                     collected, self.active_shift_id, payment_method))
            self._record_sync_event_or_raise("RESERVATION_DELIVERED", {
                "reservation_uuid": row["uuid"] if "uuid" in row.keys() else None,
                "reservation_id":   rid,
                "collected_amount": collected,
                "paid_amount_total": float(new_paid),
                "payment_method": payment_method,
                "shift_id": self.active_shift_id,
            })
            self._audit("reservation_delivered", {
                "reservation_id": rid,
                "reservation_uuid": row["uuid"] if "uuid" in row.keys() else None,
                "customer": row["customer"] if "customer" in row.keys() else None,
                "customer_phone": row["customer_phone"] if "customer_phone" in row.keys() else None,
                "collected_amount": collected,
                "paid_amount_total": float(new_paid),
                "total_amount": total_amount,
                "payment_method": payment_method,
                "item_type": row["item_type"],
                "school": row["school"],
                "color": row["color"],
                "size": row["size"],
                "qty": int(row["qty"] or 0),
            }, actor="cashier")

    def deliver_reservation_items(self, reservation_ids: Sequence[int], collected_amount: float = 0.0, payment_method: str = PAYMENT_METHOD_CASH) -> Dict[str, Any]:
        """Deliver selected reservation items with strict payment rules."""
        payment_method = _normalize_payment_method(payment_method)
        self._require_shift()
        ids = sorted({int(x) for x in reservation_ids})
        if not ids:
            raise ValueError("اختر عنصر حجز واحد على الأقل.")
        ph = ",".join("?" for _ in ids)
        rows = self.conn.execute(
            f"SELECT * FROM reservations WHERE id IN ({ph})",
            tuple(ids),
        ).fetchall()
        if len(rows) != len(ids):
            raise ValueError("بعض عناصر الحجز غير موجودة.")
        pending_rows = [r for r in rows if _is_reservation_active(r["status"])]
        if not pending_rows:
            raise ValueError("العناصر المحددة غير معلقة.")
        self._ensure_reservation_delivery_stock(pending_rows)
        group_uuid = str((rows[0]["reservation_group_uuid"] or "")).strip()
        selected_ids = {int(r["id"]) for r in pending_rows}
        selected_total = sum(float(r["total_amount"] or 0.0) for r in pending_rows)
        selected_credit = sum(float(r["paid_amount"] or 0.0) for r in pending_rows)

        selected_unpaid = max(0.0, selected_total - selected_credit)
        required_collect = selected_unpaid
        collect = max(0.0, float(collected_amount))
        if abs(collect - required_collect) > 1e-6:
            raise ValueError(
                f"المبلغ المطلوب لهذه العملية هو {format_money(required_collect)} "
                f"(المتبقي على العناصر المحددة)."
            )

        cash_alloc: Dict[int, float] = {}
        for r in pending_rows:
            rid = int(r["id"])
            total = float(r["total_amount"] or 0.0)
            paid = float(r["paid_amount"] or 0.0)
            cash_alloc[rid] = max(0.0, total - paid)

        with self.conn:
            for r in pending_rows:
                rid = int(r["id"])
                row_total = float(r["total_amount"] or 0.0)
                add_paid = float(cash_alloc.get(rid, 0.0))
                new_paid = row_total
                self.deduct_stock_for_specs(
                    r["item_type"], r["school"], r["color"], r["size"],
                    int(r["qty"] or 0),
                    note=f"Reservation delivered #{rid}",
                )
                self.conn.execute(
                    "UPDATE reservations SET status=?, paid_amount=? WHERE id=?",
                    (RESERVATION_STATUS_DELIVERED, new_paid, rid),
                )
                if add_paid > 1e-9:
                    self.conn.execute(
                        """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        (now_iso(), "DELIVER_PAY", None, 0,
                         f"تحصيل باقي حجز #{rid}",
                         None, r["item_type"], r["school"], r["color"], r["size"],
                         add_paid, self.active_shift_id, payment_method),
                    )
                self._record_sync_event_or_raise("RESERVATION_DELIVERED", {
                    "reservation_uuid": r["uuid"] if "uuid" in r.keys() else None,
                    "reservation_id":   rid,
                    "collected_amount": add_paid,
                    "paid_amount_total": float(new_paid),
                    "payment_method": payment_method,
                    "shift_id": self.active_shift_id,
                })
        group_done = False
        if group_uuid:
            left_cnt = self.conn.execute(
                "SELECT COUNT(*) FROM reservations WHERE reservation_group_uuid=? AND status NOT IN (?, ?)",
                (group_uuid, RESERVATION_STATUS_DELIVERED, RESERVATION_STATUS_CANCELLED),
            ).fetchone()
            group_done = int((left_cnt[0] if left_cnt else 0) or 0) == 0
        self._audit("reservation_items_delivered", {
            "reservation_ids": [int(r["id"]) for r in pending_rows],
            "group_uuid": group_uuid or None,
            "delivered_items": len(pending_rows),
            "collected_amount": collect,
            "required_collect": required_collect,
            "payment_method": payment_method,
            "group_completed": group_done,
            "items": [
                {
                    "reservation_id": int(r["id"]),
                    "reservation_uuid": r["uuid"] if "uuid" in r.keys() else None,
                    "customer": r["customer"] if "customer" in r.keys() else None,
                    "customer_phone": r["customer_phone"] if "customer_phone" in r.keys() else None,
                    "item_type": r["item_type"],
                    "school": r["school"],
                    "color": r["color"],
                    "size": r["size"],
                    "qty": int(r["qty"] or 0),
                    "total_amount": float(r["total_amount"] or 0.0),
                    "previous_paid_amount": float(r["paid_amount"] or 0.0),
                    "collected_amount": float(cash_alloc.get(int(r["id"]), 0.0)),
                }
                for r in pending_rows
            ],
        }, actor="cashier")
        return {
            "delivered_items": len(pending_rows),
            "collected_amount": collect,
            "group_completed": group_done,
            "group_uuid": group_uuid or None,
        }

    def cancel_reservation_items(
        self,
        reservation_ids: Sequence[int],
        *,
        reason: str = "",
        refund_payment_method: str = PAYMENT_METHOD_CASH,
    ) -> Dict[str, Any]:
        """Cancel selected pending reservation rows and reallocate their deposit.

        The selected rows remain as cancelled audit rows. Their allocated
        down-payment is moved onto the still-pending rows in the same visible
        reservation bill. If no pending rows remain, that money is recorded as a
        refund for shift totals.
        """
        self._require_shift()
        refund_payment_method = _normalize_payment_method(refund_payment_method)
        ids = sorted({int(x) for x in reservation_ids})
        if not ids:
            raise ValueError("اختر عنصر حجز واحد على الأقل.")
        ph = ",".join("?" for _ in ids)
        rows = self.conn.execute(
            f"SELECT * FROM reservations WHERE id IN ({ph}) ORDER BY id ASC",
            tuple(ids),
        ).fetchall()
        if len(rows) != len(ids):
            raise ValueError("بعض عناصر الحجز غير موجودة.")
        active_rows = [r for r in rows if _is_reservation_active(r["status"])]
        if not active_rows:
            raise ValueError("العناصر المحددة غير معلقة.")

        group_uuid = str(active_rows[0]["reservation_group_uuid"] or "").strip()
        if group_uuid:
            for r in active_rows:
                if str(r["reservation_group_uuid"] or "").strip() != group_uuid:
                    raise ValueError("اختر عناصر من نفس فاتورة الحجز فقط.")
        else:
            group_uuid = f"legacy-{int(active_rows[0]['id'])}"
            if len(active_rows) > 1:
                raise ValueError("اختر عناصر من نفس فاتورة الحجز فقط.")

        cancelled_ids = {int(r["id"]) for r in active_rows}
        cancelled_paid = round(sum(float(r["paid_amount"] or 0.0) for r in active_rows), 2)
        reason_text = (reason or "").strip() or "Reservation item cancelled"

        with self.conn:
            for r in active_rows:
                rid = int(r["id"])
                self.conn.execute(
                    "UPDATE reservations SET status=?, paid_amount=0 WHERE id=?",
                    (RESERVATION_STATUS_CANCELLED, rid),
                )
                self.conn.execute(
                    """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        now_iso(), "RESERVATION_CANCELLED", None, int(r["qty"] or 0),
                        f"إلغاء بند حجز #{rid}: {reason_text}",
                        None, r["item_type"], r["school"], r["color"], r["size"],
                        0.0, self.active_shift_id, refund_payment_method,
                    ),
                )
                self._record_sync_event_or_raise("RESERVATION_CANCELLED", {
                    "reservation_uuid": r["uuid"] if "uuid" in r.keys() else None,
                    "reservation_id": rid,
                    "reason": reason_text,
                    "refunded_amount": 0.0,
                    "shift_id": self.active_shift_id,
                })

            if group_uuid.startswith("legacy-"):
                remaining_rows = []
            else:
                remaining_rows = self.conn.execute(
                    """
                    SELECT *
                      FROM reservations
                     WHERE reservation_group_uuid=?
                       AND status NOT IN (?, ?)
                     ORDER BY id ASC
                    """,
                    (group_uuid, RESERVATION_STATUS_DELIVERED, RESERVATION_STATUS_CANCELLED),
                ).fetchall()

            refund_amount = 0.0
            if remaining_rows and cancelled_paid > 1e-9:
                existing_paid = sum(float(r["paid_amount"] or 0.0) for r in remaining_rows)
                total_paid_to_allocate = round(existing_paid + cancelled_paid, 2)
                remaining_totals = [float(r["total_amount"] or 0.0) for r in remaining_rows]
                allocations = _allocate_reservation_down_payments(remaining_totals, total_paid_to_allocate)
                for r, alloc in zip(remaining_rows, allocations):
                    self.conn.execute(
                        "UPDATE reservations SET paid_amount=? WHERE id=?",
                        (float(alloc), int(r["id"])),
                    )
            elif cancelled_paid > 1e-9:
                refund_amount = cancelled_paid
                first = active_rows[0]
                self.conn.execute(
                    """INSERT INTO movements(ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        now_iso(), "RESERVE_REFUND", None, 0,
                        f"استرداد عربون حجز #{_reservation_bill_id_from_rows([dict(r) for r in active_rows])}: {reason_text}",
                        None, first["item_type"], first["school"], first["color"], first["size"],
                        float(refund_amount), self.active_shift_id, refund_payment_method,
                    ),
                )

        summary = {
            "cancelled_items": len(active_rows),
            "redistributed_amount": 0.0 if refund_amount > 1e-9 else cancelled_paid,
            "refund_amount": refund_amount,
            "group_uuid": None if group_uuid.startswith("legacy-") else group_uuid,
        }
        self._audit("reservation_items_cancelled", {
            **summary,
            "reservation_ids": [int(r["id"]) for r in active_rows],
            "reason": reason_text,
            "refund_payment_method": refund_payment_method,
            "cancelled_paid_amount": cancelled_paid,
            "items": [
                {
                    "reservation_id": int(r["id"]),
                    "reservation_uuid": r["uuid"] if "uuid" in r.keys() else None,
                    "customer": r["customer"] if "customer" in r.keys() else None,
                    "customer_phone": r["customer_phone"] if "customer_phone" in r.keys() else None,
                    "item_type": r["item_type"],
                    "school": r["school"],
                    "color": r["color"],
                    "size": r["size"],
                    "qty": int(r["qty"] or 0),
                    "total_amount": float(r["total_amount"] or 0.0),
                    "paid_amount": float(r["paid_amount"] or 0.0),
                }
                for r in active_rows
            ],
        }, actor="manager")
        return summary

    # -------- Shift Management --------
    def start_shift(self) -> int:
        self.assert_clock_sane()
        cur = self.conn.execute("SELECT id FROM shifts WHERE status='OPEN' ORDER BY started_at DESC, id DESC LIMIT 1")
        if cur.fetchone():
            raise ValueError("يوجد وردية مفتوحة بالفعل.")
        started_at = now_iso()
        business_day = started_at[:10]
        with self.conn:
            c = self.conn.execute(
                "INSERT INTO shifts(started_at, status) VALUES(?, 'OPEN')",
                (started_at,)
            )
            shift_id = int(c.lastrowid)
            shift_uuid_row = self.conn.execute(
                "SELECT uuid FROM shifts WHERE id=?", (shift_id,)
            ).fetchone()
            self._record_sync_event_or_raise("SHIFT_OPENED", {
                "shift_uuid": shift_uuid_row[0] if shift_uuid_row else None,
                "shift_id":   shift_id,
                "started_at": started_at,
                "business_day": business_day,
                "status": "OPEN",
                "source_event": "shift_opened",
            })
            try:
                self.audit.write("shift_started", {
                    "shift_id": shift_id,
                    "shift_uuid": shift_uuid_row[0] if shift_uuid_row else None,
                }, actor="cashier", shift_id=shift_id)
            except Exception:
                import traceback
                traceback.print_exc()
            return shift_id

    def get_open_shift(self) -> Optional[Dict[str, Any]]:
        cur = self.conn.execute("SELECT * FROM shifts WHERE status='OPEN' ORDER BY started_at DESC, id DESC LIMIT 1")
        row = cur.fetchone()
        return dict(row) if row else None

    def end_shift(self, shift_id: int, summary_json: str = "") -> None:
        self.assert_clock_sane()
        ended_at = now_iso()
        with self.conn:
            shift_before = self.conn.execute(
                "SELECT started_at, uuid FROM shifts WHERE id=? AND status='OPEN'",
                (int(shift_id),)
            ).fetchone()
            cur = self.conn.execute(
                "UPDATE shifts SET ended_at=?, status='CLOSED', summary_json=? WHERE id=? AND status='OPEN'",
                (ended_at, summary_json, int(shift_id))
            )
            if cur.rowcount != 1:
                raise ValueError("Shift is not open or was already closed")
            shift_uuid_row = self.conn.execute(
                "SELECT uuid FROM shifts WHERE id=?", (int(shift_id),)
            ).fetchone()
            started_at = str(shift_before["started_at"] if shift_before and "started_at" in shift_before.keys() else "")
            self._record_sync_event_or_raise("SHIFT_CLOSED", {
                "shift_uuid":  shift_uuid_row[0] if shift_uuid_row else None,
                "shift_id":    int(shift_id),
                "started_at":  started_at,
                "ended_at":    ended_at,
                "business_day": (ended_at or started_at)[:10],
                "status": "CLOSED",
                "source_event": "shift_closed",
                "summary_json": summary_json or "",
            })
            self._audit("shift_ended", {
                "shift_id": int(shift_id),
                "shift_uuid": shift_uuid_row[0] if shift_uuid_row else None,
                "summary_json": summary_json or "",
            }, actor="cashier")

    def add_expense(self, amount: float, note: str = "") -> int:
        self._require_shift()
        amount = float(amount or 0.0)
        if amount <= 0:
            raise ValueError("أدخل مبلغ مصروف أكبر من صفر.")
        sid = int(self.active_shift_id or 0)
        created_at = now_iso()
        clean_note = (note or "").strip()
        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO expenses(created_at, amount, note, shift_id) VALUES(?,?,?,?)",
                (created_at, amount, clean_note or None, sid),
            )
            expense_id = int(cur.lastrowid)
            self._audit(
                "expense_logged",
                {"expense_id": expense_id, "amount": amount, "note": clean_note, "shift_id": sid},
                actor="cashier",
            )
        return expense_id

    def list_shift_expenses(self, shift_id: int) -> List[Dict[str, Any]]:
        rows = self.conn.execute(
            """
            SELECT id, created_at, amount, note, shift_id
            FROM expenses
            WHERE shift_id=?
            ORDER BY created_at ASC, id ASC
            """,
            (int(shift_id),),
        ).fetchall()
        return [dict(r) for r in rows]

    def get_shift_expense_summary(self, shift_id: int) -> Dict[str, Any]:
        items = self.list_shift_expenses(int(shift_id))
        total = sum(float(r.get("amount") or 0.0) for r in items)
        return {"expense_count": len(items), "expense_total": total, "expenses": items}

    def get_shift_summary(self, shift_id: int) -> Dict[str, Any]:
        cur = self.conn.execute("SELECT * FROM shifts WHERE id=?", (int(shift_id),))
        shift = cur.fetchone()
        if not shift:
            raise ValueError("الوردية غير موجودة")
        stored_summary = self._stored_closed_shift_summary(dict(shift))
        if stored_summary is not None:
            return stored_summary
        started = shift["started_at"]
        ended = shift["ended_at"] or now_iso()
        sid = int(shift["id"])
        bill_scope, bill_params = self._shift_scope_sql(dict(shift), "created_at")
        movement_scope, movement_params = self._shift_scope_sql(dict(shift), "ts")

        # Gross retail sales for this shift. VOID is subtracted where it happened.
        cur = self.conn.execute(
            f"""
            SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills
             WHERE {bill_scope}
               AND (bill_type='SALE' OR bill_type IS NULL)
               AND UPPER(COALESCE(status,'CONFIRMED')) != 'VOID'
            """,
            bill_params,
        )
        sales = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
            FROM bills
             WHERE {bill_scope}
               AND (bill_type='SALE' OR bill_type IS NULL)
               AND UPPER(COALESCE(status,'CONFIRMED')) != 'VOID'
            """,
            bill_params,
        )
        sales_by_method = dict(cur.fetchone())

        # Returns
        cur = self.conn.execute(
            f"SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills WHERE {bill_scope} AND bill_type='RETURN'",
            bill_params)
        returns = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
            FROM bills
            WHERE {bill_scope} AND bill_type='RETURN'
            """,
            bill_params,
        )
        returns_by_method = dict(cur.fetchone())

        # Exchanges (total can be positive=customer paid or negative=refund)
        cur = self.conn.execute(
            f"SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills WHERE {bill_scope} AND bill_type='EXCHANGE'",
            bill_params)
        exchanges = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
            FROM bills
            WHERE {bill_scope} AND bill_type='EXCHANGE'
            """,
            bill_params,
        )
        exchanges_by_method = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT COUNT(*) as cnt, COALESCE(SUM(total_amount),0) as total
            FROM reservations
            WHERE {bill_scope}
              AND status != ?
            """,
            (*bill_params, RESERVATION_STATUS_CANCELLED),
        )
        reservations = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT
                COALESCE(SUM(unit_price),0) as paid,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN unit_price ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN unit_price ELSE 0 END),0) as visa_total
            FROM movements
            WHERE direction='RESERVE_PAY' AND {movement_scope}
            """,
            movement_params,
        )
        reserve_pay = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT
                COALESCE(SUM(unit_price),0) as total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN unit_price ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN unit_price ELSE 0 END),0) as visa_total
            FROM movements
            WHERE direction='RESERVE_REFUND' AND {movement_scope}
            """,
            movement_params,
        )
        reserve_refund = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT
                COUNT(*) as cnt,
                COALESCE(SUM(unit_price),0) as total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN unit_price ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN unit_price ELSE 0 END),0) as visa_total
            FROM movements
            WHERE direction='DELIVER_PAY' AND {movement_scope}
            """,
            movement_params,
        )
        delivery_pay = dict(cur.fetchone())
        cur = self.conn.execute(
            f"""
            SELECT COUNT(*) as cnt, COALESCE(SUM(unit_price),0) as total
            FROM movements
            WHERE direction='VOID_PAY' AND {movement_scope}
            """,
            movement_params,
        )
        voids = dict(cur.fetchone())
        void_method_scope = movement_scope.replace("shift_id", "m.shift_id").replace("ts", "m.ts")
        cur = self.conn.execute(
            f"""
            SELECT
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(b.payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN m.unit_price ELSE 0 END),0) as cash_total,
                COALESCE(SUM(CASE WHEN UPPER(COALESCE(b.payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN m.unit_price ELSE 0 END),0) as visa_total
            FROM movements m
            LEFT JOIN bills b ON b.id = m.bill_id
            WHERE m.direction='VOID_PAY' AND {void_method_scope}
            """,
            movement_params,
        )
        voids_by_method = dict(cur.fetchone())
        if self._allow_legacy_shift_time_fallback(dict(shift)):
            cur = self.conn.execute(
                """
                SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total
                FROM bills b
                WHERE UPPER(COALESCE(b.status,'CONFIRMED'))='VOID'
                  AND b.voided_at >= ? AND b.voided_at <= ?
                  AND NOT EXISTS (
                      SELECT 1 FROM movements m
                      WHERE m.direction='VOID_PAY' AND m.bill_id=b.id
                  )
                """,
                (started, ended),
            )
            legacy_voids = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
                FROM bills b
                WHERE UPPER(COALESCE(b.status,'CONFIRMED'))='VOID'
                  AND b.voided_at >= ? AND b.voided_at <= ?
                  AND NOT EXISTS (
                      SELECT 1 FROM movements m
                      WHERE m.direction='VOID_PAY' AND m.bill_id=b.id
                  )
                """,
                (started, ended),
            )
            legacy_voids_by_method = dict(cur.fetchone())
        else:
            legacy_voids = {"cnt": 0, "total": 0}
            legacy_voids_by_method = {"cash_total": 0, "visa_total": 0}
        reservation_cash = float(reserve_pay["cash_total"])
        reservation_visa = float(reserve_pay["visa_total"])
        missing_reserve_pay = self._missing_reservation_payment_totals(dict(shift))
        reservation_cash += float(missing_reserve_pay["cash_total"])
        reservation_visa += float(missing_reserve_pay["visa_total"])
        delivery_cash = float(delivery_pay["cash_total"])
        delivery_visa = float(delivery_pay["visa_total"])
        void_count = int(voids["cnt"] or 0) + int(legacy_voids["cnt"] or 0)
        void_total = float(voids["total"]) + float(legacy_voids["total"])
        void_cash = float(voids_by_method["cash_total"] or 0.0) + float(legacy_voids_by_method["cash_total"] or 0.0)
        void_visa = float(voids_by_method["visa_total"] or 0.0) + float(legacy_voids_by_method["visa_total"] or 0.0)
        return_cash = float(returns_by_method["cash_total"])
        return_visa = float(returns_by_method["visa_total"])
        exchange_cash = float(exchanges_by_method["cash_total"])
        exchange_visa = float(exchanges_by_method["visa_total"])
        retail_sales_total = float(sales["total"])
        money_in_total = retail_sales_total + float(reserve_pay["paid"]) + float(missing_reserve_pay["paid"]) + float(delivery_pay["total"])
        expense_summary = self.get_shift_expense_summary(sid)
        return {
            "shift_id": shift["id"], "started_at": started, "ended_at": ended,
            "inflow_count": 0, "inflow_total_qty": 0,
            "inflow_items": [],
            "sales_count": sales["cnt"],
            "retail_sales_total": retail_sales_total,
            "sales_total": money_in_total,
            "money_in_total": money_in_total,
            "sales_cash_total": float(sales_by_method["cash_total"]),
            "sales_visa_total": float(sales_by_method["visa_total"]),
            "res_count": reservations["cnt"], "res_total": float(reservations["total"]),
            "res_paid": float(reserve_pay["paid"]) + float(missing_reserve_pay["paid"]),
            "res_refund": float(reserve_refund["total"]),
            "deliver_count": delivery_pay["cnt"], "deliver_total": float(delivery_pay["total"]),
            "void_count": void_count, "void_total": void_total,
            "return_count": returns["cnt"], "return_total": float(returns["total"]),
            "exchange_count": exchanges["cnt"], "exchange_total": float(exchanges["total"]),
            "expense_count": expense_summary["expense_count"],
            "expense_total": expense_summary["expense_total"],
            "expenses": expense_summary["expenses"],
            "cash_collected": float(sales_by_method["cash_total"]) + reservation_cash + delivery_cash - float(reserve_refund["cash_total"]) - void_cash - return_cash + exchange_cash,
            "visa_collected": float(sales_by_method["visa_total"]) + reservation_visa + delivery_visa - float(reserve_refund["visa_total"]) - void_visa - return_visa + exchange_visa,
        }

    def get_all_shifts(self, date_from: Optional[str] = None, date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        """Return all shifts with computed summary columns."""
        wheres: List[str] = []
        params: List[Any] = []
        if date_from:
            wheres.append("s.started_at >= ?")
            params.append(date_from)
        if date_to:
            wheres.append("s.started_at <= ?")
            params.append(date_to + "T23:59:59")
        where_sql = (" WHERE " + " AND ".join(wheres)) if wheres else ""
        rows = self.conn.execute(
            f"SELECT * FROM shifts s{where_sql} ORDER BY s.started_at DESC", params
        ).fetchall()
        results: List[Dict[str, Any]] = []
        for row in rows:
            r = dict(row)
            stored_summary = self._stored_closed_shift_summary(r)
            if stored_summary is not None:
                results.append(stored_summary)
                continue
            started = r["started_at"]
            ended = r["ended_at"] or now_iso()
            sid = int(r["id"])
            bill_scope, bill_params = self._shift_scope_sql(r, "created_at")
            movement_scope, movement_params = self._shift_scope_sql(r, "ts")
            # Gross retail sales. VOID is subtracted in the shift where it happened.
            cur = self.conn.execute(
                f"""
                SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills
                 WHERE {bill_scope}
                   AND (bill_type='SALE' OR bill_type IS NULL)
                   AND UPPER(COALESCE(status,'CONFIRMED')) != 'VOID'
                """,
                bill_params,
            )
            sales = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
                FROM bills
                 WHERE {bill_scope}
                   AND (bill_type='SALE' OR bill_type IS NULL)
                   AND UPPER(COALESCE(status,'CONFIRMED')) != 'VOID'
                """,
                bill_params,
            )
            sales_by_method = dict(cur.fetchone())
            # returns
            cur = self.conn.execute(
                f"SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills WHERE {bill_scope} AND bill_type='RETURN'",
                bill_params)
            returns = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
                FROM bills
                WHERE {bill_scope} AND bill_type='RETURN'
                """,
                bill_params,
            )
            returns_by_method = dict(cur.fetchone())
            # exchanges
            cur = self.conn.execute(
                f"SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills WHERE {bill_scope} AND bill_type='EXCHANGE'",
                bill_params)
            exchanges = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
                FROM bills
                WHERE {bill_scope} AND bill_type='EXCHANGE'
                """,
                bill_params,
            )
            exchanges_by_method = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT COUNT(*) as cnt, COALESCE(SUM(total_amount),0) as total
                FROM reservations
                WHERE {bill_scope}
                  AND status != ?
                """,
                (*bill_params, RESERVATION_STATUS_CANCELLED),
            )
            reservations = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(unit_price),0) as paid,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN unit_price ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN unit_price ELSE 0 END),0) as visa_total
                FROM movements
                WHERE direction='RESERVE_PAY' AND {movement_scope}
                """,
                movement_params,
            )
            reserve_pay = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(unit_price),0) as total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN unit_price ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN unit_price ELSE 0 END),0) as visa_total
                FROM movements
                WHERE direction='RESERVE_REFUND' AND {movement_scope}
                """,
                movement_params,
            )
            reserve_refund = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT
                    COUNT(*) as cnt,
                    COALESCE(SUM(unit_price),0) as total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN unit_price ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN unit_price ELSE 0 END),0) as visa_total
                FROM movements
                WHERE direction='DELIVER_PAY' AND {movement_scope}
                """,
                movement_params,
            )
            delivery_pay = dict(cur.fetchone())
            cur = self.conn.execute(
                f"""
                SELECT COUNT(*) as cnt, COALESCE(SUM(unit_price),0) as total
                FROM movements
                WHERE direction='VOID_PAY' AND {movement_scope}
                """,
                movement_params,
            )
            voids = dict(cur.fetchone())
            void_method_scope = movement_scope.replace("shift_id", "m.shift_id").replace("ts", "m.ts")
            cur = self.conn.execute(
                f"""
                SELECT
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(b.payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN m.unit_price ELSE 0 END),0) as cash_total,
                    COALESCE(SUM(CASE WHEN UPPER(COALESCE(b.payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN m.unit_price ELSE 0 END),0) as visa_total
                FROM movements m
                LEFT JOIN bills b ON b.id = m.bill_id
                WHERE m.direction='VOID_PAY' AND {void_method_scope}
                """,
                movement_params,
            )
            voids_by_method = dict(cur.fetchone())
            if self._allow_legacy_shift_time_fallback(r):
                cur = self.conn.execute(
                    """
                    SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total
                    FROM bills b
                    WHERE UPPER(COALESCE(b.status,'CONFIRMED'))='VOID'
                      AND b.voided_at >= ? AND b.voided_at <= ?
                      AND NOT EXISTS (
                          SELECT 1 FROM movements m
                          WHERE m.direction='VOID_PAY' AND m.bill_id=b.id
                      )
                    """,
                    (started, ended),
                )
                legacy_voids = dict(cur.fetchone())
                cur = self.conn.execute(
                    f"""
                    SELECT
                        COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_CASH}' THEN total ELSE 0 END),0) as cash_total,
                        COALESCE(SUM(CASE WHEN UPPER(COALESCE(payment_method,'{PAYMENT_METHOD_CASH}'))='{PAYMENT_METHOD_VISA}' THEN total ELSE 0 END),0) as visa_total
                    FROM bills b
                    WHERE UPPER(COALESCE(b.status,'CONFIRMED'))='VOID'
                      AND b.voided_at >= ? AND b.voided_at <= ?
                      AND NOT EXISTS (
                          SELECT 1 FROM movements m
                          WHERE m.direction='VOID_PAY' AND m.bill_id=b.id
                      )
                    """,
                    (started, ended),
                )
                legacy_voids_by_method = dict(cur.fetchone())
            else:
                legacy_voids = {"cnt": 0, "total": 0}
                legacy_voids_by_method = {"cash_total": 0, "visa_total": 0}
            reservation_cash = float(reserve_pay["cash_total"])
            reservation_visa = float(reserve_pay["visa_total"])
            missing_reserve_pay = self._missing_reservation_payment_totals(r)
            reservation_cash += float(missing_reserve_pay["cash_total"])
            reservation_visa += float(missing_reserve_pay["visa_total"])
            delivery_cash = float(delivery_pay["cash_total"])
            delivery_visa = float(delivery_pay["visa_total"])
            void_count = int(voids["cnt"] or 0) + int(legacy_voids["cnt"] or 0)
            void_total = float(voids["total"]) + float(legacy_voids["total"])
            void_cash = float(voids_by_method["cash_total"] or 0.0) + float(legacy_voids_by_method["cash_total"] or 0.0)
            void_visa = float(voids_by_method["visa_total"] or 0.0) + float(legacy_voids_by_method["visa_total"] or 0.0)
            return_cash = float(returns_by_method["cash_total"])
            return_visa = float(returns_by_method["visa_total"])
            exchange_cash = float(exchanges_by_method["cash_total"])
            exchange_visa = float(exchanges_by_method["visa_total"])
            retail_sales_total = float(sales["total"])
            money_in_total = retail_sales_total + float(reserve_pay["paid"]) + float(missing_reserve_pay["paid"]) + float(delivery_pay["total"])
            r["sales_count"] = sales["cnt"]
            r["retail_sales_total"] = retail_sales_total
            r["sales_total"] = money_in_total
            r["money_in_total"] = money_in_total
            r["sales_cash_total"] = float(sales_by_method["cash_total"])
            r["sales_visa_total"] = float(sales_by_method["visa_total"])
            r["res_count"] = reservations["cnt"]
            r["res_total"] = float(reservations["total"])
            r["res_paid"] = float(reserve_pay["paid"]) + float(missing_reserve_pay["paid"])
            r["res_refund"] = float(reserve_refund["total"])
            r["deliver_count"] = delivery_pay["cnt"]
            r["deliver_total"] = float(delivery_pay["total"])
            r["void_count"] = void_count
            r["void_total"] = void_total
            r["inflow_count"] = 0
            r["inflow_total_qty"] = 0
            r["return_count"] = returns["cnt"]
            r["return_total"] = float(returns["total"])
            r["exchange_count"] = exchanges["cnt"]
            r["exchange_total"] = float(exchanges["total"])
            expense_summary = self.get_shift_expense_summary(sid)
            r["expense_count"] = expense_summary["expense_count"]
            r["expense_total"] = expense_summary["expense_total"]
            r["expenses"] = expense_summary["expenses"]
            # Cash summary excludes stock receipts but includes customer money.
            r["cash_collected"] = r["sales_cash_total"] + reservation_cash + delivery_cash - float(reserve_refund["cash_total"]) - void_cash - return_cash + exchange_cash
            r["visa_collected"] = r["sales_visa_total"] + reservation_visa + delivery_visa - float(reserve_refund["visa_total"]) - void_visa - return_visa + exchange_visa
            results.append(r)
        return results

    # -------- Income Bill Methods --------
    def create_income_bill(self, supplier: str, lines: List[Dict[str, Any]], note: str = "") -> int:
        """Create an income bill that groups multiple stock additions."""
        if not self.is_manager_feature_enabled("allow_manual_incoming"):
            raise PermissionError(_feature_restricted_message("الوارد اليدوي غير مسموح به حالياً في نقطة البيع."))
        self._require_shift()
        if not lines:
            raise ValueError("لا توجد أصناف في فاتورة الوارد")

        total_qty = 0
        total_value = 0.0

        with self.conn:
            cur = self.conn.execute(
                "INSERT INTO income_bills(created_at, supplier, total_qty, total_value, note) VALUES(?,?,0,0,?)",
                (now_iso(), (supplier or "").strip() or None, (note or "").strip() or None)
            )
            bill_id = int(cur.lastrowid)

            for line in lines:
                item_type = str(line["item_type"]).strip()
                school = str(line["school"]).strip()
                color = str(line["color"]).strip()
                size = str(line["size"]).strip()
                qty = int(line["qty"])
                price = float(line["unit_price"])

                if qty <= 0:
                    continue

                line_total = price * qty
                total_qty += qty
                total_value += line_total

                # Add to inventory
                self.add_stock(
                    item_type=item_type, school=school, color=color,
                    size=size, unit_price=price, count=qty,
                )

                # Record in income bill items
                self.conn.execute(
                    """INSERT INTO income_bill_items
                    (income_bill_id, item_type, school, color, size, unit_price, qty, line_total)
                    VALUES(?,?,?,?,?,?,?,?)""",
                    (bill_id, item_type, school, color, size, price, qty, line_total)
                )

            self.conn.execute(
                "UPDATE income_bills SET total_qty=?, total_value=? WHERE id=?",
                (total_qty, total_value, bill_id)
            )
            ib_uuid_row = self.conn.execute(
                "SELECT uuid FROM income_bills WHERE id=?", (bill_id,)
            ).fetchone()
            self._record_sync_event("INCOME_BILL_CREATED", {
                "income_bill_uuid": ib_uuid_row[0] if ib_uuid_row else None,
                "income_bill_id":   bill_id,
                "supplier":    (supplier or "").strip() or None,
                "note":        (note or "").strip() or None,
                "total_qty":   int(total_qty),
                "total_value": float(total_value),
                "lines": [
                    {
                        "item_type": str(ln["item_type"]).strip(),
                        "school":    str(ln["school"]).strip(),
                        "color":     str(ln["color"]).strip(),
                        "size":      str(ln["size"]).strip(),
                        "unit_price": float(ln["unit_price"]),
                        "qty":        int(ln["qty"]),
                    }
                    for ln in lines if int(ln.get("qty") or 0) > 0
                ],
                "shift_id": self.active_shift_id,
            })
            return bill_id

    def list_income_bills(self, date_from: Optional[str] = None, date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        where = ["1=1"]
        args: list = []
        if date_from:
            where.append("date(created_at) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(created_at) <= date(?)")
            args.append(date_to)
        cur = self.conn.execute(
            f"SELECT * FROM income_bills WHERE {' AND '.join(where)} ORDER BY id DESC", args
        )
        return [dict(r) for r in cur.fetchall()]

    def list_income_bill_items(self, income_bill_id: int) -> List[Dict[str, Any]]:
        cur = self.conn.execute(
            "SELECT * FROM income_bill_items WHERE income_bill_id=? ORDER BY id", (int(income_bill_id),)
        )
        return [dict(r) for r in cur.fetchall()]

    def get_income_stats(self, date_from: Optional[str] = None, date_to: Optional[str] = None,
                         school: Optional[str] = None, item_type: Optional[str] = None,
                         color: Optional[str] = None) -> Dict[str, Any]:
        """Get income statistics optionally filtered."""
        where = ["1=1"]
        args: list = []
        if date_from:
            where.append("date(ib.created_at) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(ib.created_at) <= date(?)")
            args.append(date_to)

        item_where = ["1=1"]
        item_args: list = []
        if school:
            item_where.append("LOWER(TRIM(ibi.school)) = LOWER(?)")
            item_args.append(school.strip())
        if item_type:
            item_where.append("LOWER(TRIM(ibi.item_type)) = LOWER(?)")
            item_args.append(item_type.strip())
        if color:
            item_where.append("LOWER(TRIM(ibi.color)) = LOWER(?)")
            item_args.append(color.strip())

        cur = self.conn.execute(
            f"""SELECT COUNT(DISTINCT ib.id) as bill_count,
                       COALESCE(SUM(ibi.qty), 0) as total_qty,
                       COALESCE(SUM(ibi.line_total), 0) as total_value
                FROM income_bills ib
                JOIN income_bill_items ibi ON ibi.income_bill_id = ib.id
                WHERE {' AND '.join(where)} AND {' AND '.join(item_where)}""",
            args + item_args
        )
        row = cur.fetchone()
        return {
            "bill_count": int(row["bill_count"]),
            "total_qty": int(row["total_qty"]),
            "total_value": float(row["total_value"]),
        }

    # -------- Bulk Price Update --------
    def bulk_update_prices(self, mode: str, value: float, constraints: Optional[Dict[str, Any]] = None) -> int:
        """Update prices in bulk. mode='percentage' or 'fixed'. Returns count of updated rows."""
        if not self.is_manager_feature_enabled("allow_bulk_price"):
            raise PermissionError(_feature_restricted_message("تعديل الأسعار من نقطة البيع غير مسموح به حالياً."))
        import math
        where_parts = ["count > 0"]
        args: list = []
        if constraints:
            for k in ("item_type", "school", "color", "size"):
                v = constraints.get(k)
                if v and str(v).strip():
                    where_parts.append(f"LOWER(TRIM({k})) = LOWER(?)")
                    args.append(str(v).strip())

        where = " AND ".join(where_parts)
        cur = self.conn.execute(f"SELECT id, unit_price FROM stocks WHERE {where}", args)
        rows = cur.fetchall()

        if not rows:
            return 0

        def _round_up_5(x: float) -> float:
            return float(math.ceil(x / 5.0) * 5)

        updated = 0
        with self.conn:
            for r in rows:
                old_price = float(r["unit_price"])
                if mode == "percentage":
                    new_price = _round_up_5(old_price * (1 + value / 100.0))
                else:
                    new_price = old_price + value

                if new_price < 0:
                    new_price = 0.0

                if abs(new_price - old_price) > 0.001:
                    self.conn.execute("UPDATE stocks SET unit_price=? WHERE id=?", (new_price, int(r["id"])))
                    self.conn.execute(
                        """INSERT INTO movements(ts,direction,stock_id,qty,note,item_type,school,color,size,unit_price)
                        SELECT ?,'PRICE_UPDATE',id,0,?,item_type,school,color,size,? FROM stocks WHERE id=?""",
                        (now_iso(), f"تعديل سعر جماعي: {format_money(old_price)} -> {format_money(new_price)}", new_price, int(r["id"]))
                    )
                    updated += 1
            if updated > 0:
                self._record_sync_event("PRICE_UPDATE", {
                    "mode":        mode,
                    "value":       float(value),
                    "constraints": constraints or {},
                    "updated_count": int(updated),
                    "scope":       "bulk",
                })
        return updated

    def get_sales_stats(self, date_from=None, date_to=None, school=None, item_type=None, color=None):
        """Retail sale bill totals plus separate reservation cash fields."""
        # Bill stats - if spec filters exist, join to bill_items
        bill_where = ["1=1"]
        bill_args = []
        if date_from:
            bill_where.append("date(b.created_at) >= date(?)")
            bill_args.append(date_from)
        if date_to:
            bill_where.append("date(b.created_at) <= date(?)")
            bill_args.append(date_to)

        bill_where.append("(COALESCE(b.bill_type,'SALE')='SALE' OR b.bill_type IS NULL)")
        bill_where.append("UPPER(COALESCE(b.status,'CONFIRMED')) != 'VOID'")

        if school or item_type or color:
            # Need to join bill_items for spec filtering
            spec_where = []
            if school:
                spec_where.append("LOWER(TRIM(bi.school)) = LOWER(?)")
                bill_args.append(school.strip())
            if item_type:
                spec_where.append("LOWER(TRIM(bi.item_type)) = LOWER(?)")
                bill_args.append(item_type.strip())
            if color:
                spec_where.append("LOWER(TRIM(bi.color)) = LOWER(?)")
                bill_args.append(color.strip())

            cur = self.conn.execute(
                f"""SELECT COUNT(DISTINCT b.id) as cnt,
                           COALESCE(SUM(bi.line_total), 0) as total
                    FROM bills b
                    JOIN bill_items bi ON bi.bill_id = b.id
                    WHERE {' AND '.join(bill_where)} AND {' AND '.join(spec_where)}""",
                bill_args
            )
        else:
            cur = self.conn.execute(
                f"SELECT COUNT(*) as cnt, COALESCE(SUM(total),0) as total FROM bills b WHERE {' AND '.join(bill_where)}",
                bill_args
            )

        bill_row = dict(cur.fetchone())

        # Reservation stats
        res_where = ["status != ?"]
        res_args = [RESERVATION_STATUS_CANCELLED]
        if date_from:
            res_where.append("date(created_at) >= date(?)")
            res_args.append(date_from)
        if date_to:
            res_where.append("date(created_at) <= date(?)")
            res_args.append(date_to)
        if school:
            res_where.append("LOWER(TRIM(school)) = LOWER(?)")
            res_args.append(school.strip())
        if item_type:
            res_where.append("LOWER(TRIM(item_type)) = LOWER(?)")
            res_args.append(item_type.strip())
        if color:
            res_where.append("LOWER(TRIM(color)) = LOWER(?)")
            res_args.append(color.strip())

        cur = self.conn.execute(
            f"""SELECT COUNT(*) as cnt, COALESCE(SUM(total_amount),0) as total,
                       COALESCE(SUM(paid_amount),0) as paid
                FROM reservations WHERE {' AND '.join(res_where)}""",
            res_args
        )
        res_row = dict(cur.fetchone())

        pwhere = ["direction='RESERVE_PAY'"]
        pargs: List[Any] = []
        if date_from:
            pwhere.append("date(ts) >= date(?)")
            pargs.append(date_from)
        if date_to:
            pwhere.append("date(ts) <= date(?)")
            pargs.append(date_to)
        if school:
            pwhere.append("LOWER(TRIM(school)) = LOWER(?)")
            pargs.append(school.strip())
        if item_type:
            pwhere.append("LOWER(TRIM(item_type)) = LOWER(?)")
            pargs.append(item_type.strip())
        if color:
            pwhere.append("LOWER(TRIM(color)) = LOWER(?)")
            pargs.append(color.strip())
        cur = self.conn.execute(
            f"SELECT COALESCE(SUM(unit_price), 0) AS pt FROM movements WHERE {' AND '.join(pwhere)}",
            pargs,
        )
        pr = cur.fetchone()
        reserve_cash = float(pr["pt"] if pr and pr["pt"] is not None else 0)

        dwhere = ["direction='DELIVER_PAY'"]
        dargs: List[Any] = []
        if date_from:
            dwhere.append("date(ts) >= date(?)")
            dargs.append(date_from)
        if date_to:
            dwhere.append("date(ts) <= date(?)")
            dargs.append(date_to)
        if school:
            dwhere.append("LOWER(TRIM(school)) = LOWER(?)")
            dargs.append(school.strip())
        if item_type:
            dwhere.append("LOWER(TRIM(item_type)) = LOWER(?)")
            dargs.append(item_type.strip())
        if color:
            dwhere.append("LOWER(TRIM(color)) = LOWER(?)")
            dargs.append(color.strip())
        cur = self.conn.execute(
            f"SELECT COALESCE(SUM(unit_price), 0) AS dt FROM movements WHERE {' AND '.join(dwhere)}",
            dargs,
        )
        dr = cur.fetchone()
        deliver_cash = float(dr["dt"] if dr and dr["dt"] is not None else 0)

        rwhere = ["direction='RESERVE_REFUND'"]
        rargs: List[Any] = []
        if date_from:
            rwhere.append("date(ts) >= date(?)")
            rargs.append(date_from)
        if date_to:
            rwhere.append("date(ts) <= date(?)")
            rargs.append(date_to)
        if school:
            rwhere.append("LOWER(TRIM(school)) = LOWER(?)")
            rargs.append(school.strip())
        if item_type:
            rwhere.append("LOWER(TRIM(item_type)) = LOWER(?)")
            rargs.append(item_type.strip())
        if color:
            rwhere.append("LOWER(TRIM(color)) = LOWER(?)")
            rargs.append(color.strip())
        cur = self.conn.execute(
            f"SELECT COALESCE(SUM(unit_price), 0) AS rt FROM movements WHERE {' AND '.join(rwhere)}",
            rargs,
        )
        rr = cur.fetchone()
        reserve_refund = float(rr["rt"] if rr and rr["rt"] is not None else 0)

        retail_sales_total = float(bill_row["total"])
        money_in_total = retail_sales_total + reserve_cash + deliver_cash
        return {
            "sales_count": bill_row["cnt"],
            "retail_sales_total": retail_sales_total,
            "sales_total": money_in_total,
            "money_in_total": money_in_total,
            "res_count": res_row["cnt"], "res_total": float(res_row["total"]),
            "res_paid": reserve_cash,
            "deliver_cash": deliver_cash,
            "reserve_refund": reserve_refund,
        }

    def get_item_movement_stats(self, date_from=None, date_to=None, school=None, item_type=None, color=None):
        """Get received/sold/reserved counts per item."""
        where = ["direction IN ('IN','OUT','OUT_FACTORY','RESERVE','ADJUST_OUT')"]
        args: list = []
        if date_from:
            where.append("date(ts) >= date(?)")
            args.append(date_from)
        if date_to:
            where.append("date(ts) <= date(?)")
            args.append(date_to)
        if school:
            where.append("LOWER(TRIM(school)) = LOWER(?)")
            args.append(school.strip())
        if item_type:
            where.append("LOWER(TRIM(item_type)) = LOWER(?)")
            args.append(item_type.strip())
        if color:
            where.append("LOWER(TRIM(color)) = LOWER(?)")
            args.append(color.strip())
        cur = self.conn.execute(f"""
            SELECT m.item_type, m.school, m.color, m.size,
                SUM(CASE WHEN m.direction='IN' THEN m.qty ELSE 0 END) as received,
                SUM(
                    CASE
                        WHEN m.direction IN ('OUT','OUT_FACTORY')
                         AND (
                                m.bill_id IS NULL
                             OR COALESCE(b.bill_type,'SALE')='SALE'
                         )
                        THEN m.qty
                        ELSE 0
                    END
                ) as sold,
                SUM(CASE WHEN m.direction='RESERVE' THEN m.qty ELSE 0 END) as reserved,
                SUM(CASE WHEN m.direction='ADJUST_OUT' THEN m.qty ELSE 0 END) as adjusted
            FROM movements m
            LEFT JOIN bills b ON b.id = m.bill_id
            WHERE {' AND '.join(where)}
            GROUP BY m.item_type, m.school, m.color, m.size
            ORDER BY m.item_type, m.school, m.color, m.size
        """, args)
        return [dict(r) for r in cur.fetchall()]

    def reset_movement_counts(self):
        """Reset all movement records. Keep item definitions intact."""
        if not self.is_manager_feature_enabled("allow_reset_counts"):
            raise PermissionError(_feature_restricted_message("إعادة التعيين غير مسموح بها حالياً من نقطة البيع."))
        before = {
            "movements": int(self.conn.execute("SELECT COUNT(*) FROM movements").fetchone()[0] or 0),
            "reservations": int(self.conn.execute("SELECT COUNT(*) FROM reservations").fetchone()[0] or 0),
            "bills": int(self.conn.execute("SELECT COUNT(*) FROM bills").fetchone()[0] or 0),
            "bill_items": int(self.conn.execute("SELECT COUNT(*) FROM bill_items").fetchone()[0] or 0),
        }
        with self.conn:
            self.conn.execute("DELETE FROM movements")
            self.conn.execute("DELETE FROM reservations")
            self.conn.execute("DELETE FROM bills")
            self.conn.execute("DELETE FROM bill_items")
        self._audit("movement_counts_reset", {"deleted_counts": before}, actor="manager")

    # -------- Bill history APIs --------
    def list_bills(self) -> List[Dict[str, Any]]:
        cur = self.conn.cursor()
        cur.execute(
            "SELECT id,created_at,customer,customer_phone,total,"
            " COALESCE(status,'CONFIRMED') AS status,"
            " COALESCE(bill_type,'SALE') AS bill_type,"
            f" COALESCE(payment_method,'{PAYMENT_METHOD_CASH}') AS payment_method,"
            " void_reason, voided_at"
            " FROM bills ORDER BY id DESC"
        )
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows

    def list_bill_history(self) -> List[Dict[str, Any]]:
        rows: List[Dict[str, Any]] = []
        for b in self.list_bills():
            b = dict(b)
            b["history_kind"] = "bill"
            b["history_key"] = "bill:%s" % b["id"]
            rows.append(b)

        cur = self.conn.cursor()
        cur.execute(
            """
            SELECT
                group_key,
                MIN(id) AS id,
                MIN(created_at) AS created_at,
                COALESCE(MAX(NULLIF(TRIM(customer), '')), '') AS customer,
                COALESCE(MAX(NULLIF(TRIM(customer_phone), '')), '') AS customer_phone,
                COALESCE(SUM(CASE WHEN status != ? THEN total_amount ELSE 0 END), 0) AS total,
                COALESCE(SUM(CASE WHEN status != ? THEN paid_amount ELSE 0 END), 0) AS paid_amount,
                COALESCE(MAX(NULLIF(TRIM(payment_method), '')), ?) AS payment_method,
                SUM(CASE WHEN status NOT IN (?, ?) THEN 1 ELSE 0 END) AS pending_count,
                SUM(CASE WHEN status = ? THEN 1 ELSE 0 END) AS cancelled_count,
                COUNT(*) AS item_count
            FROM (
                SELECT *,
                       COALESCE(NULLIF(TRIM(reservation_group_uuid), ''), 'legacy-' || id) AS group_key
                FROM reservations
            )
            GROUP BY group_key
            """,
            (
                RESERVATION_STATUS_CANCELLED,
                RESERVATION_STATUS_CANCELLED,
                PAYMENT_METHOD_CASH,
                RESERVATION_STATUS_DELIVERED,
                RESERVATION_STATUS_CANCELLED,
                RESERVATION_STATUS_CANCELLED,
            ),
        )
        for r in cur.fetchall():
            pending_count = int(r["pending_count"] or 0)
            total_count = int(r["item_count"] or 0)
            cancelled_count = int(r["cancelled_count"] or 0)
            status = RESERVATION_STATUS_PENDING if pending_count else RESERVATION_STATUS_DELIVERED
            if cancelled_count and cancelled_count == total_count:
                status = RESERVATION_STATUS_CANCELLED
            if pending_count and pending_count < total_count:
                status = "تسليم جزئي"
            row = dict(r)
            row.update(
                {
                    "history_kind": "reservation",
                    "history_key": "reservation:%s" % row["group_key"],
                    "bill_type": "RESERVATION",
                    "status": status,
                    "void_reason": None,
                    "voided_at": None,
                }
            )
            rows.append(row)
        cur.close()
        rows.sort(key=lambda x: str(x.get("created_at") or ""), reverse=True)
        return rows

    def void_bill(self, bill_id: int, reason: str) -> None:
        self._require_shift()
        reason = (reason or "").strip()
        if not reason:
            raise ValueError("سبب الإلغاء مطلوب.")

        with self.conn:
            bill = self.conn.execute(
                "SELECT id, customer, total, COALESCE(status,'CONFIRMED') AS status, "
                "COALESCE(bill_type,'SALE') AS bill_type, "
                "COALESCE(payment_method, ?) AS payment_method, uuid "
                "FROM bills WHERE id = ?",
                (PAYMENT_METHOD_CASH, int(bill_id)),
            ).fetchone()
            if bill is None:
                raise ValueError("الفاتورة غير موجودة.")
            if str(bill["status"]).upper() == "VOID":
                raise ValueError("تم إلغاء هذه الفاتورة بالفعل.")
            if str(bill["bill_type"]).upper() != "SALE":
                raise ValueError("الإلغاء المباشر متاح لفواتير البيع فقط. استخدم المرتجع للأنواع الأخرى.")

            items = self.conn.execute(
                "SELECT item_type, school, color, size, unit_price, qty, origin "
                "FROM bill_items WHERE bill_id = ?",
                (int(bill_id),),
            ).fetchall()
            if not items:
                raise ValueError("لا توجد بنود في هذه الفاتورة.")

            for item in items:
                if str(item["origin"] or "STOCK").upper() != "STOCK":
                    continue
                stock_id = self.add_or_update_stock_row(
                    item["item_type"], item["school"], item["color"], item["size"],
                    float(item["unit_price"]), int(item["qty"]),
                )
                self.conn.execute(
                    """INSERT INTO movements
                       (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (
                        now_iso(),
                        "VOID_IN",
                        stock_id,
                        int(item["qty"]),
                        f"Bill void: {reason}",
                        int(bill_id),
                        item["item_type"],
                        item["school"],
                        item["color"],
                        item["size"],
                        float(item["unit_price"]),
                        self.active_shift_id,
                    ),
                )

            self.conn.execute(
                """INSERT INTO movements
                   (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id,payment_method)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    now_iso(), "VOID_PAY", None, 0, f"Void refund: {reason}",
                    int(bill_id), None, None, None, None, float(bill["total"] or 0.0),
                    self.active_shift_id,
                    bill["payment_method"] if "payment_method" in bill.keys() else PAYMENT_METHOD_CASH,
                ),
            )

            self.conn.execute(
                "UPDATE bills SET status='VOID', void_reason=?, voided_at=? WHERE id=?",
                (reason, now_iso(), int(bill_id)),
            )

            self._record_sync_event_or_raise("SALE_VOIDED", {
                "bill_uuid": bill["uuid"] if "uuid" in bill.keys() else None,
                "bill_id": int(bill_id),
                "customer": bill["customer"],
                "total": float(bill["total"] or 0),
                "payment_method": bill["payment_method"] if "payment_method" in bill.keys() else PAYMENT_METHOD_CASH,
                "reason": reason,
                "shift_id": self.active_shift_id,
            })
            self._audit("sale_voided", {
                "bill_id": int(bill_id),
                "bill_uuid": bill["uuid"] if "uuid" in bill.keys() else None,
                "customer": bill["customer"],
                "total": float(bill["total"] or 0),
                "reason": reason,
                "items": [dict(item) for item in items],
            }, actor="manager")

    def convert_exchange_bill_to_sale(self, bill_id: int, reason: str = "") -> Dict[str, Any]:
        """Admin correction for a bill that was entered as exchange but was really a sale."""
        self._require_shift()
        reason = (reason or "").strip() or "Admin corrected replacement bill to sale bill"

        with self.conn:
            bill = self.conn.execute(
                """
                SELECT id, uuid, customer, customer_phone, total,
                       COALESCE(status,'CONFIRMED') AS status,
                       COALESCE(bill_type,'SALE') AS bill_type,
                       shift_id
                  FROM bills
                 WHERE id = ?
                """,
                (int(bill_id),),
            ).fetchone()
            if bill is None:
                raise ValueError("الفاتورة غير موجودة.")
            if str(bill["status"]).upper() == "VOID":
                raise ValueError("لا يمكن تحويل فاتورة ملغاة.")
            if str(bill["bill_type"]).upper() != "EXCHANGE":
                raise ValueError("هذا التصحيح متاح لفواتير الاستبدال فقط.")

            item_rows = self.conn.execute(
                """
                SELECT rowid AS row_id, item_type, school, color, size,
                       unit_price, qty, line_total, COALESCE(origin,'STOCK') AS origin
                  FROM bill_items
                 WHERE bill_id = ?
                 ORDER BY rowid ASC
                """,
                (int(bill_id),),
            ).fetchall()
            if not item_rows:
                raise ValueError("لا توجد بنود في هذه الفاتورة.")

            return_lines: List[Dict[str, Any]] = []
            take_lines: List[Dict[str, Any]] = []
            sale_items: List[Dict[str, Any]] = []
            for item in item_rows:
                qty = int(item["qty"] or 0)
                price = float(item["unit_price"] or 0)
                line = {
                    "item_type": item["item_type"],
                    "school": item["school"],
                    "color": item["color"],
                    "size": item["size"],
                    "unit_price": price,
                    "qty": qty,
                    "line_total": float(item["line_total"] or (price * qty)),
                    "origin": str(item["origin"] or "STOCK").upper(),
                }
                sale_items.append(dict(line, origin="STOCK"))
                if line["origin"] == "RETURN":
                    return_lines.append(line)
                else:
                    take_lines.append(line)

            for line in return_lines:
                qty = int(line["qty"] or 0)
                move = self.conn.execute(
                    """
                    SELECT id, stock_id
                      FROM movements
                     WHERE bill_id = ?
                       AND direction = 'RETURN_IN'
                       AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                       AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                       AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                       AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                     ORDER BY id ASC
                     LIMIT 1
                    """,
                    (int(bill_id), line["item_type"], line["school"], line["color"], line["size"]),
                ).fetchone()
                stock_id = int(move["stock_id"]) if move and move["stock_id"] is not None else None
                if stock_id is None:
                    stock_id = self.add_or_update_stock_row(
                        line["item_type"], line["school"], line["color"], line["size"],
                        float(line["unit_price"] or 0), 0,
                    )
                # The mistaken exchange RETURN_IN already added qty. A sale should
                # have deducted qty instead, so current stock needs a 2x correction.
                self.conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?", (qty * 2, stock_id))
                if move:
                    self.conn.execute(
                        """
                        UPDATE movements
                           SET direction = 'OUT',
                               note = ?,
                               shift_id = COALESCE(shift_id, ?)
                         WHERE id = ?
                        """,
                        (f"Admin correction exchange->sale: {reason}", self.active_shift_id, int(move["id"])),
                    )
                else:
                    self.conn.execute(
                        """
                        INSERT INTO movements
                            (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price,shift_id)
                        VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            now_iso(), "OUT", stock_id, qty,
                            f"Admin correction exchange->sale: {reason}", int(bill_id),
                            line["item_type"], line["school"], line["color"], line["size"],
                            float(line["unit_price"] or 0), self.active_shift_id,
                        ),
                    )

            for line in take_lines:
                note = f"Admin correction exchange->sale: {reason}"
                self.conn.execute(
                    """
                    UPDATE movements
                       SET note = CASE
                            WHEN note IS NULL OR TRIM(note) = '' THEN ?
                            ELSE note || ' | ' || ?
                           END,
                           shift_id = COALESCE(shift_id, ?)
                     WHERE bill_id = ?
                       AND direction = 'OUT'
                       AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
                       AND LOWER(TRIM(school)) = LOWER(TRIM(?))
                       AND LOWER(TRIM(color)) = LOWER(TRIM(?))
                       AND LOWER(TRIM(size)) = LOWER(TRIM(?))
                    """,
                    (note, note, self.active_shift_id, int(bill_id), line["item_type"], line["school"], line["color"], line["size"]),
                )

            new_total = sum(float(line["line_total"] or 0.0) for line in sale_items)
            old_diff = float(bill["total"] or 0.0)
            return_total = sum(float(line["line_total"] or 0.0) for line in return_lines)
            take_total = sum(float(line["line_total"] or 0.0) for line in take_lines)

            self.conn.execute("UPDATE bill_items SET origin='STOCK' WHERE bill_id=?", (int(bill_id),))
            self.conn.execute(
                """
                UPDATE bills
                   SET bill_type='SALE',
                       total=?,
                       payment_method=COALESCE(payment_method, ?)
                 WHERE id=?
                """,
                (float(new_total), PAYMENT_METHOD_CASH, int(bill_id)),
            )

            self._record_sync_event_or_raise("SALE_BILL_TYPE_CORRECTED", {
                "bill_uuid": bill["uuid"] if "uuid" in bill.keys() else None,
                "bill_id": int(bill_id),
                "customer": bill["customer"],
                "customer_phone": bill["customer_phone"] if "customer_phone" in bill.keys() else None,
                "from_bill_type": "EXCHANGE",
                "to_bill_type": "SALE",
                "old_diff": old_diff,
                "old_return_total": float(return_total),
                "old_take_total": float(take_total),
                "new_total": float(new_total),
                "amount_delta": float(new_total - old_diff),
                "return_lines": return_lines,
                "take_lines": take_lines,
                "items": sale_items,
                "reason": reason,
                "shift_id": self.active_shift_id,
            })

            summary = {
                "bill_id": int(bill_id),
                "new_total": float(new_total),
                "amount_delta": float(new_total - old_diff),
                "corrected_return_qty": sum(int(line["qty"] or 0) for line in return_lines),
                "item_count": len(sale_items),
            }
            self._audit("bill_type_corrected", {
                **summary,
                "bill_uuid": bill["uuid"] if "uuid" in bill.keys() else None,
                "customer": bill["customer"],
                "from_bill_type": "EXCHANGE",
                "to_bill_type": "SALE",
                "old_diff": old_diff,
                "old_return_total": float(return_total),
                "old_take_total": float(take_total),
                "reason": reason,
                "items": sale_items,
            }, actor="manager")
            return summary

    def list_bill_items(self, bill_id: int) -> List[Dict[str, Any]]:
        cur = self.conn.cursor()
        cur.execute(
            """SELECT item_type,school,color,size,unit_price,qty,line_total,origin
            FROM bill_items WHERE bill_id=?""",
            (int(bill_id),),
        )
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows

    def list_reservation_bill_items(self, group_key: str) -> List[Dict[str, Any]]:
        key = str(group_key or "").strip()
        cur = self.conn.cursor()
        cur.execute(
            """
            SELECT item_type,school,color,size,unit_price,qty,
                   total_amount AS line_total,
                   'RESERVATION' AS origin,
                   status
            FROM (
                SELECT *,
                       COALESCE(NULLIF(TRIM(reservation_group_uuid), ''), 'legacy-' || id) AS group_key
                FROM reservations
            )
            WHERE group_key=?
              AND status != ?
            ORDER BY id ASC
            """,
            (key, RESERVATION_STATUS_CANCELLED),
        )
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows

    def export_inventory_excel(self, path: str, rows: Sequence[Dict[str, Any]]) -> None:
        headers = [
            "id","item_type","school","color","size","unit_price","count","value",
        ]
        table = []
        for r in rows:
            value = r.get("value", float(r["unit_price"]) * int(r["count"]))
            table.append(
                [
                    r["id"], r["item_type"], r["school"], r["color"], r["size"],
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
            cur = self.conn.execute(
                "UPDATE app_settings SET value = ? WHERE key = ?",
                (str(value), str(key)),
            )
            if cur.rowcount == 0:
                self.conn.execute(
                    "INSERT OR IGNORE INTO app_settings(key, value) VALUES(?, ?)",
                    (str(key), str(value)),
                )

    def is_manager_feature_enabled(self, feature_key: str) -> bool:
        default = "1" if POS_MANAGER_FEATURE_DEFAULTS.get(feature_key, False) else "0"
        raw = (self.get_app_setting(f"feature:{feature_key}", default) or default).strip().lower()
        return raw in ("1", "true", "yes", "on")

    def set_manager_feature_enabled(self, feature_key: str, enabled: bool) -> None:
        self.set_app_setting(f"feature:{feature_key}", "1" if enabled else "0")
        self._audit("manager_feature_changed", {
            "feature_key": str(feature_key),
            "enabled": bool(enabled),
        }, actor="manager")

    def verify_admin_password(self, plain: str) -> bool:
        stored = self.get_app_setting("admin_password", ADMIN_PASSWORD_PLAIN) or ADMIN_PASSWORD_PLAIN
        if str(stored).startswith(ADMIN_PASSWORD_HASH_PREFIX):
            digest = hashlib.sha256(str(plain).encode("utf-8")).hexdigest()
            return str(stored) == f"{ADMIN_PASSWORD_HASH_PREFIX}{digest}"
        ok = str(plain) == str(stored)
        # Backward-compatible auto-migration from plain-text storage.
        if ok:
            self.set_admin_password(str(plain))
        return ok

    def set_admin_password(self, plain: str) -> None:
        digest = hashlib.sha256(str(plain).encode("utf-8")).hexdigest()
        self.set_app_setting("admin_password", f"{ADMIN_PASSWORD_HASH_PREFIX}{digest}")
        self._audit("admin_password_changed", {}, actor="manager")

    def _hidden_definition_sql(self, field: str, column_sql: str) -> str:
        if field not in ("item_type", "school", "color", "size"):
            raise ValueError("Unsupported hidden-definition field")
        return (
            "NOT EXISTS ("
            "SELECT 1 FROM hidden_definitions hd "
            f"WHERE hd.field = '{field}' "
            f"AND LOWER(TRIM(hd.value)) = LOWER(TRIM({column_sql}))"
            ")"
        )

    def hide_definition(self, field: str, value: str) -> None:
        if field not in ("item_type", "school", "color", "size"):
            raise ValueError("Unsupported hidden-definition field")
        value = (value or "").strip()
        if not value:
            raise ValueError("Value is required")
        with self.conn:
            self.conn.execute(
                "INSERT OR REPLACE INTO hidden_definitions(field, value, hidden_at) VALUES(?, ?, ?)",
                (field, value, now_iso()),
            )
        self._audit("definition_hidden", {"field": field, "value": value}, actor="manager")

    def delete_school_from_ui(self, school: str) -> int:
        if not self.is_manager_feature_enabled("allow_inventory_delete"):
            raise PermissionError(_feature_restricted_message("حذف المدارس من واجهة نقطة البيع غير مسموح به حالياً."))

        school = (school or "").strip()
        if not school:
            raise ValueError("اسم المدرسة مطلوب.")

        cur = self.conn.cursor()
        try:
            cur.execute(
                """
                SELECT COUNT(*) AS row_count, COALESCE(SUM(count), 0) AS total_qty
                FROM stocks
                WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))
                """,
                (school,),
            )
            row = cur.fetchone()
            row_count = int(row["row_count"] or 0)
            total_qty = int(row["total_qty"] or 0)
        finally:
            cur.close()

        if total_qty > 0:
            raise ValueError("لا يمكن حذف المدرسة من الواجهة بينما ما زال لها رصيد في المخزون.")

        with self.conn:
            self.conn.execute(
                "DELETE FROM stocks WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))",
                (school,),
            )
            self.conn.execute(
                "DELETE FROM size_profiles WHERE LOWER(TRIM(school)) = LOWER(TRIM(?))",
                (school,),
            )
            self.conn.execute(
                "DELETE FROM spec_history WHERE field = 'school' AND LOWER(TRIM(value)) = LOWER(TRIM(?))",
                (school,),
            )
            self.hide_definition("school", school)
        self._audit("school_deleted_from_ui", {
            "school": school,
            "deleted_stock_rows": row_count,
        }, actor="manager")
        return row_count

    def remove_from_stock(self, stock_id: int, qty: Optional[int], note: str = "Admin remove") -> int:
        if not self.is_manager_feature_enabled("allow_inventory_delete"):
            raise PermissionError(_feature_restricted_message("الحذف اليدوي من مخزون نقطة البيع غير مسموح به حالياً."))
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
                   (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (
                    now_iso(), "ADJUST_OUT", int(stock_id), int(take), note, None,
                    s["item_type"], s["school"], s["color"], s["size"],
                    float(s["unit_price"]),
                ),
            )
            self._record_sync_event("STOCK_ADJUST", {
                "stock_uuid": s["uuid"] if "uuid" in s.keys() else None,
                "stock_id":   int(stock_id),
                "direction":  "OUT",
                "qty":        int(take),
                "note":       note,
                "item_type":  s["item_type"],
                "school":     s["school"],
                "color":      s["color"],
                "size":       s["size"],
                "unit_price": float(s["unit_price"]),
            })
            self._audit("stock_removed_manually", {
                "stock_id": int(stock_id),
                "qty": int(take),
                "note": note,
                "item_type": s["item_type"],
                "school": s["school"],
                "color": s["color"],
                "size": s["size"],
                "unit_price": float(s["unit_price"]),
            }, actor="manager")
            return int(take)

    def _stock_audit_line_values(self, line: Dict[str, Any]) -> Tuple[Any, ...]:
        expected = int(line.get("expected") or 0)
        actual = int(line.get("actual") or 0)
        diff = int(line.get("diff", actual - expected) or 0)
        price = float(line.get("unit_price") or 0)
        return (
            line.get("stock_id"),
            str(line.get("item_type") or "").strip(),
            str(line.get("school") or "").strip(),
            str(line.get("color") or "").strip(),
            _normalize_size_label(str(line.get("size") or "").strip()),
            expected,
            actual,
            diff,
            price,
            float(diff * price),
        )

    def _insert_stock_audit_report_lines(self, report_id: int, lines: Sequence[Dict[str, Any]]) -> None:
        for line in lines:
            self.conn.execute(
                """INSERT INTO stock_audit_report_lines
                   (report_id,stock_id,item_type,school,color,size,expected,actual,diff,unit_price,diff_value)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                (int(report_id), *self._stock_audit_line_values(line)),
            )

    def create_stock_audit_report(self, lines: Sequence[Dict[str, Any]], reason: str = "manual", bucket_key: Optional[str] = None) -> Optional[int]:
        diff_lines = [dict(line) for line in lines if int(line.get("diff") or 0) != 0]
        if not diff_lines:
            return None
        total_diff = sum(int(line.get("diff") or 0) for line in diff_lines)
        total_value = sum(float(int(line.get("diff") or 0) * float(line.get("unit_price") or 0)) for line in diff_lines)
        with self.conn:
            cur = self.conn.execute(
                """INSERT INTO stock_audit_reports
                   (created_at,reason,diff_count,total_diff,total_value,bucket_key)
                   VALUES(?,?,?,?,?,?)""",
                (now_iso(), reason, len(diff_lines), int(total_diff), float(total_value), bucket_key),
            )
            report_id = int(cur.lastrowid)
            self._insert_stock_audit_report_lines(report_id, diff_lines)
        return report_id

    def append_stock_audit_report_bucket(self, lines: Sequence[Dict[str, Any]], reason: str = "auto-equalization", bucket_key: Optional[str] = None) -> Optional[int]:
        diff_lines = [dict(line) for line in lines if int(line.get("diff") or 0) != 0]
        if not diff_lines:
            return None
        bucket = (bucket_key or now_iso()[:13]).strip()
        total_diff = sum(int(line.get("diff") or 0) for line in diff_lines)
        total_value = sum(float(int(line.get("diff") or 0) * float(line.get("unit_price") or 0)) for line in diff_lines)
        with self.conn:
            row = self.conn.execute(
                """SELECT id FROM stock_audit_reports
                   WHERE reason=? AND bucket_key=?
                   ORDER BY id DESC LIMIT 1""",
                (reason, bucket),
            ).fetchone()
            if row:
                report_id = int(row["id"])
                self.conn.execute(
                    """UPDATE stock_audit_reports
                       SET created_at=?, diff_count=diff_count+?, total_diff=total_diff+?, total_value=total_value+?
                       WHERE id=?""",
                    (now_iso(), len(diff_lines), int(total_diff), float(total_value), report_id),
                )
            else:
                cur = self.conn.execute(
                    """INSERT INTO stock_audit_reports
                       (created_at,reason,diff_count,total_diff,total_value,bucket_key)
                       VALUES(?,?,?,?,?,?)""",
                    (now_iso(), reason, len(diff_lines), int(total_diff), float(total_value), bucket),
                )
                report_id = int(cur.lastrowid)
            self._insert_stock_audit_report_lines(report_id, diff_lines)
        return report_id

    def normalize_auto_stock_audit_reports_by_hour(self) -> None:
        rows = self.conn.execute(
            """SELECT id, created_at FROM stock_audit_reports
               WHERE reason='auto-equalization'
               ORDER BY created_at ASC, id ASC"""
        ).fetchall()
        grouped: Dict[str, List[int]] = {}
        for r in rows:
            created = str(r["created_at"] or "")
            bucket = created[:13] if len(created) >= 13 else now_iso()[:13]
            grouped.setdefault(bucket, []).append(int(r["id"]))
        with self.conn:
            for bucket, ids in grouped.items():
                keep = ids[0]
                self.conn.execute("UPDATE stock_audit_reports SET bucket_key=? WHERE id=?", (bucket, keep))
                for old_id in ids[1:]:
                    self.conn.execute(
                        "UPDATE stock_audit_report_lines SET report_id=? WHERE report_id=?",
                        (keep, old_id),
                    )
                    self.conn.execute("DELETE FROM stock_audit_reports WHERE id=?", (old_id,))
                totals = self.conn.execute(
                    """SELECT COUNT(*) AS c, COALESCE(SUM(diff),0) AS d, COALESCE(SUM(diff_value),0) AS v
                       FROM stock_audit_report_lines WHERE report_id=?""",
                    (keep,),
                ).fetchone()
                self.conn.execute(
                    "UPDATE stock_audit_reports SET diff_count=?, total_diff=?, total_value=? WHERE id=?",
                    (int(totals["c"] or 0), int(totals["d"] or 0), float(totals["v"] or 0), keep),
                )

    def list_stock_audit_reports(self) -> List[Dict[str, Any]]:
        try:
            self.normalize_auto_stock_audit_reports_by_hour()
        except Exception:
            pass
        rows = self.conn.execute(
            """SELECT id,created_at,reason,diff_count,total_diff,total_value,bucket_key
               FROM stock_audit_reports
               ORDER BY created_at DESC, id DESC"""
        ).fetchall()
        return [dict(r) for r in rows]

    def get_stock_audit_report(self, report_id: int) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]]]:
        report = self.conn.execute(
            """SELECT id,created_at,reason,diff_count,total_diff,total_value,bucket_key
               FROM stock_audit_reports WHERE id=?""",
            (int(report_id),),
        ).fetchone()
        if not report:
            return None, []
        lines = self.conn.execute(
            """SELECT id,report_id,stock_id,item_type,school,color,size,expected,actual,diff,unit_price,diff_value
               FROM stock_audit_report_lines
               WHERE report_id=?
               ORDER BY item_type, school, color,
                   CASE WHEN TRIM(size) NOT GLOB '*[^0-9]*' THEN 0 ELSE 1 END,
                   CASE WHEN TRIM(size) NOT GLOB '*[^0-9]*' THEN CAST(size AS INTEGER) ELSE NULL END,
                   CASE UPPER(TRIM(size))
                       WHEN 'XXS' THEN 1 WHEN 'XS' THEN 2 WHEN 'S' THEN 3 WHEN 'M' THEN 4
                       WHEN 'L' THEN 5 WHEN 'XL' THEN 6 WHEN '2XL' THEN 7 WHEN '3XL' THEN 8
                       WHEN '4XL' THEN 9 WHEN '5XL' THEN 10 ELSE 99 END""",
            (int(report_id),),
        ).fetchall()
        return dict(report), [dict(r) for r in lines]

    def stock_audit_touched_keys(self) -> Set[Tuple[str, str, str, str]]:
        touched: Set[Tuple[str, str, str, str]] = set()
        for r in self.conn.execute("SELECT DISTINCT item_type,school,color,size FROM stock_audit_report_lines").fetchall():
            touched.add((
                str(r["item_type"] or "").strip().lower(),
                str(r["school"] or "").strip().lower(),
                str(r["color"] or "").strip().lower(),
                _normalize_size_label(str(r["size"] or "")).lower(),
            ))
        for r in self.conn.execute(
            """SELECT DISTINCT item_type,school,color,size FROM movements
               WHERE direction IN ('ADJUST_IN','ADJUST_OUT')
                 AND (note LIKE 'Physical count adjustment%' OR note LIKE 'POS stock audit%')"""
        ).fetchall():
            touched.add((
                str(r["item_type"] or "").strip().lower(),
                str(r["school"] or "").strip().lower(),
                str(r["color"] or "").strip().lower(),
                _normalize_size_label(str(r["size"] or "")).lower(),
            ))
        return touched

    def record_pos_stock_audit_applied(
        self,
        report_id: Optional[int],
        lines: Sequence[Dict[str, Any]],
        reason: str,
    ) -> Optional[str]:
        # Warehouse requested-stock reports treat local POS audit deltas as
        # part of branch income/baseline, so send the audit detail event in
        # addition to the regular POS stock snapshot.
        return self._record_pos_stock_audit_applied_legacy(report_id, lines, reason)

    def _record_pos_stock_audit_applied_legacy(
        self,
        report_id: Optional[int],
        lines: Sequence[Dict[str, Any]],
        reason: str,
    ) -> Optional[str]:
        diff_lines = [dict(line) for line in lines if int(line.get("diff") or 0) != 0]
        if not report_id or not diff_lines:
            return None
        source_device = self._current_device_name() or "pos"
        normalized = []
        for line in diff_lines:
            expected = int(line.get("expected") or 0)
            actual = int(line.get("actual") or 0)
            diff = int(line.get("diff", actual - expected) or 0)
            price = float(line.get("unit_price") or 0)
            normalized.append({
                "stock_id": line.get("stock_id"),
                "item_type": str(line.get("item_type") or "").strip(),
                "school": str(line.get("school") or "").strip(),
                "color": str(line.get("color") or "").strip(),
                "size": _normalize_size_label(str(line.get("size") or "").strip()),
                "expected": expected,
                "actual": actual,
                "diff": diff,
                "unit_price": price,
                "diff_value": float(diff * price),
            })
        payload = {
            "audit_uuid": f"{source_device}:{int(report_id)}:{uuid.uuid4().hex}",
            "source_device_name": source_device,
            "report_id": int(report_id),
            "reason": str(reason or ""),
            "created_at": now_iso(),
            "total_diff": sum(int(line["diff"]) for line in normalized),
            "total_value": sum(float(line["diff_value"]) for line in normalized),
            "lines": normalized,
        }
        try:
            return self._record_targeted_inventory_event(
                event_type="POS_STOCK_AUDIT_APPLIED",
                target_scope="warehouse",
                payload=payload,
            )
        except Exception:
            import traceback
            traceback.print_exc()
            return None

    @staticmethod
    def _parse_stock_id_list(value: Any) -> List[int]:
        ids: List[int] = []
        for part in str(value or "").replace(";", ",").split(","):
            part = part.strip()
            if not part:
                continue
            try:
                stock_id = int(part)
            except Exception:
                continue
            if stock_id > 0 and stock_id not in ids:
                ids.append(stock_id)
        return ids

    def apply_stock_adjustments(self, adjustments: Sequence[Dict[str, Any]], note: str = "Physical count adjustment") -> int:
        self._require_shift()
        applied = 0
        applied_lines: List[Dict[str, Any]] = []
        with self.conn:
            for adj in adjustments:
                diff = int(adj.get("diff") or 0)
                if diff == 0:
                    continue
                if "actual" in adj and int(adj.get("actual") or 0) < 0:
                    raise ValueError("Actual stock count cannot be negative")
                stock_id = adj.get("stock_id")
                stock_ids = self._parse_stock_id_list(stock_id)
                item_type = str(adj.get("item_type") or "").strip()
                school = str(adj.get("school") or "").strip()
                color = str(adj.get("color") or "").strip()
                size = _normalize_size_label(str(adj.get("size") or "").strip())
                price = float(adj.get("unit_price") or self.get_effective_price(item_type, school, color, size) or 0)
                row = None
                if stock_ids:
                    row = self.conn.execute("SELECT * FROM stocks WHERE id=?", (int(stock_ids[0]),)).fetchone()
                if row is None:
                    row = self.conn.execute(
                        """SELECT * FROM stocks
                           WHERE LOWER(TRIM(item_type))=LOWER(TRIM(?))
                             AND LOWER(TRIM(school))=LOWER(TRIM(?))
                             AND LOWER(TRIM(color))=LOWER(TRIM(?))
                             AND LOWER(TRIM(size))=LOWER(TRIM(?))
                           ORDER BY CASE WHEN unit_price=? THEN 0 ELSE 1 END, id ASC LIMIT 1""",
                        (item_type, school, color, size, price),
                    ).fetchone()
                    if row:
                        stock_id = int(row["id"])
                    else:
                        stock_id = self.add_or_update_stock_row(
                            item_type, school, color, size, float(price), 0,
                        )
                        row = self.conn.execute("SELECT * FROM stocks WHERE id=?", (stock_id,)).fetchone()
                        self._upsert_history({"item_type": item_type, "school": school, "color": color, "size": size})
                direction = "ADJUST_IN" if diff > 0 else "ADJUST_OUT"
                qty = abs(diff)
                movement_ts = now_iso()
                if diff > 0:
                    affected_rows = [(row, qty)]
                    self.conn.execute("UPDATE stocks SET count = count + ? WHERE id = ?", (int(qty), int(stock_id)))
                else:
                    affected_rows = []
                    remaining = int(qty)
                    candidates: List[sqlite3.Row] = []
                    seen_candidate_ids: Set[int] = set()
                    if stock_ids:
                        placeholders = ",".join("?" for _ in stock_ids)
                        preferred = self.conn.execute(
                            f"""SELECT * FROM stocks
                                WHERE id IN ({placeholders})
                                  AND LOWER(TRIM(item_type))=LOWER(TRIM(?))
                                  AND LOWER(TRIM(school))=LOWER(TRIM(?))
                                  AND LOWER(TRIM(color))=LOWER(TRIM(?))
                                  AND LOWER(TRIM(size))=LOWER(TRIM(?))
                                  AND count > 0
                                ORDER BY id ASC""",
                            (*stock_ids, item_type, school, color, size),
                        ).fetchall()
                        for candidate in preferred:
                            candidates.append(candidate)
                            seen_candidate_ids.add(int(candidate["id"]))

                    fallback = self.conn.execute(
                        """SELECT * FROM stocks
                           WHERE LOWER(TRIM(item_type))=LOWER(TRIM(?))
                             AND LOWER(TRIM(school))=LOWER(TRIM(?))
                             AND LOWER(TRIM(color))=LOWER(TRIM(?))
                             AND LOWER(TRIM(size))=LOWER(TRIM(?))
                             AND count > 0
                           ORDER BY CASE WHEN id=? THEN 0 ELSE 1 END, id ASC""",
                        (
                            item_type or row["item_type"],
                            school or row["school"],
                            color or row["color"],
                            size or row["size"],
                            int(row["id"]),
                        ),
                    ).fetchall()
                    for candidate in fallback:
                        cid = int(candidate["id"])
                        if cid not in seen_candidate_ids:
                            candidates.append(candidate)
                            seen_candidate_ids.add(cid)

                    for candidate in candidates:
                        if remaining <= 0:
                            break
                        take = min(int(candidate["count"] or 0), remaining)
                        if take <= 0:
                            continue
                        self.conn.execute("UPDATE stocks SET count = count - ? WHERE id = ?", (int(take), int(candidate["id"])))
                        affected_rows.append((candidate, take))
                        remaining -= take
                    if remaining > 0:
                        raise ValueError("لا يمكن تطبيق التسوية لأن الكمية المتاحة أقل من الفرق المطلوب.")

                for affected, take_qty in affected_rows:
                    self.conn.execute(
                        """INSERT INTO movements
                           (ts,direction,stock_id,qty,note,bill_id,item_type,school,color,size,unit_price)
                           VALUES(?,?,?,?,?,?,?,?,?,?,?)""",
                        (
                            movement_ts, direction, int(affected["id"]), int(take_qty), note, None,
                            affected["item_type"], affected["school"], affected["color"], affected["size"], float(affected["unit_price"]),
                        ),
                    )
                    self._record_sync_event("STOCK_ADJUST", {
                        "stock_id": int(affected["id"]),
                        "direction": "IN" if diff > 0 else "OUT",
                        "qty": int(take_qty),
                        "note": note,
                        "item_type": affected["item_type"],
                        "school": affected["school"],
                        "color": affected["color"],
                        "size": affected["size"],
                        "unit_price": float(affected["unit_price"]),
                    })
                    applied_lines.append({
                        "stock_id": int(affected["id"]),
                        "direction": "IN" if diff > 0 else "OUT",
                        "qty": int(take_qty),
                        "note": note,
                        "item_type": affected["item_type"],
                        "school": affected["school"],
                        "color": affected["color"],
                        "size": affected["size"],
                        "unit_price": float(affected["unit_price"]),
                    })
                applied += 1
        if applied:
            self._audit("stock_adjustment_applied", {
                "adjustment_count": applied,
                "movement_count": len(applied_lines),
                "note": note,
                "lines": applied_lines,
            }, actor="manager")
        return applied

    def list_movements(self, filters: Dict[str, Any]) -> List[Dict[str, Any]]:
        where = ["1=1"]
        args: List[Any] = []
        if filters.get("direction"):
            where.append("direction = ?")
            args.append(filters["direction"])
        for k in ("item_type", "school", "color", "size"):
            v = (filters.get(k) or "").strip()
            if v:
                where.append(f"LOWER(TRIM({k})) = LOWER(TRIM(?))")
                args.append(v)

        cust = (filters.get("customer") or "").strip()
        if cust:
            where.append("bill_id IN (SELECT id FROM bills WHERE LOWER(TRIM(customer)) = LOWER(TRIM(?)))")
            args.append(cust)

        txt = (filters.get("text") or "").strip()
        if txt:
            like = f"%{txt}%"
            where.append("""(
                LOWER(COALESCE(note,''))      LIKE LOWER(?)
            OR LOWER(COALESCE(item_type,'')) LIKE LOWER(?)
            OR LOWER(COALESCE(school,''))    LIKE LOWER(?)
            OR LOWER(COALESCE(color,''))     LIKE LOWER(?)
            OR LOWER(COALESCE(size,''))      LIKE LOWER(?)
            )""")
            args += [like, like, like, like, like]

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
            SELECT id, ts, direction, stock_id, qty, note, bill_id,
                item_type, school, color, size, unit_price
            FROM movements
            WHERE {' AND '.join(where)}
            ORDER BY ts DESC, id DESC
        """, args)
        rows = [dict(r) for r in cur.fetchall()]
        cur.close()
        return rows


# ------------------- UI helpers -------------------
class LabeledStaticCombo(ttk.Frame):
    def __init__(self, master, text: str, values, **kwargs):
        super().__init__(master)
        ttk.Label(self, text=text).pack(side=tk.TOP, anchor="w")
        self.var = tk.StringVar()
        self.cb = ttk.Combobox(self, textvariable=self.var, values=values, state="readonly", **kwargs)
        self.cb.pack(fill=tk.X)

    def get(self) -> str:
        return (self.var.get() or "").strip()

    def set(self, v: str) -> None:
        self.var.set(v)

class LabeledCombobox(ttk.Frame):
    FILTER_IDLE_MS = 60
    MATCH_MODE = "startswith"
    MIN_CHARS_TO_OPEN = 1
    OPEN_ALL_ON_BUTTON = True

    def __init__(self, master, text: str, db: 'SqliteDatabase', field: str, **kwargs):
        super().__init__(master)
        self.db = db
        self.field = field

        ttk.Label(self, text=text).pack(side=tk.TOP, anchor="w")

        row = ttk.Frame(self)
        row.pack(fill=tk.X)

        self.var = tk.StringVar()
        self.cb = ttk.Entry(row, textvariable=self.var, **kwargs)
        self.cb.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn = ttk.Button(row, text="▼", width=2, takefocus=0, command=self._on_caret_click)
        try:
            self.btn.configure(style="Toolbutton")
        except Exception:
            pass
        self.btn.pack(side=tk.RIGHT, padx=(4, 0))

        self._all_values: List[str] = []
        self._supplier = None
        self._debounce_job: Optional[str] = None
        self._suspend_set = False
        self._opened = False
        self._mouse_inside_popup = False

        self._popup: Optional[tk.Toplevel] = None
        self._list: Optional[tk.Listbox] = None
        self._ysb: Optional[ttk.Scrollbar] = None

        self.cb.bind("<FocusIn>", self._on_focus, add="+")
        self.cb.bind("<KeyRelease>", self._on_key_release, add="+")
        self.cb.bind("<Down>", self._on_down, add="+")
        self.cb.bind("<Escape>", self._on_escape, add="+")
        self.cb.bind("<FocusOut>", self._on_focus_out, add="+")

        self.winfo_toplevel().bind("<Configure>", self._reposition_popup_safely, add="+")

    def get(self) -> str:
        return (self.var.get() or "").strip()

    def set(self, v: str) -> None:
        self._suspend_set = True
        try:
            self.var.set(v or "")
        finally:
            self._suspend_set = False
        self._close_popup()

    def set_supplier(self, fn):
        self._supplier = fn
        self.refresh_values()

    def refresh_values(self):
        self._ensure_data(force=True)
        if self._opened:
            self._fill_list(self._filtered(self.get()))

    def _ensure_data(self, force: bool = False) -> None:
        if self._supplier:
            try:
                vals = self._supplier() or []
            except Exception:
                vals = []
            if force or (vals != self._all_values):
                self._all_values = vals[:]
            return

        if not self._all_values or force:
            try:
                self._all_values = self.db.get_distinct_filtered(self.field, {}) or []
            except Exception:
                self._all_values = []

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
        _bind_mousewheel(self._list)

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
            w = self.cb.winfo_width() + self.btn.winfo_width() + 4
            rows = min(8, max(1, self._list.size() if self._list else 0))
            row_h = max(18, int(self.cb.winfo_fpixels("1.2m")))
            h = rows * row_h + 2
            self._popup.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            pass

    def _reposition_popup_safely(self, *_):
        if self._opened:
            self.after_idle(self._position_popup)

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
        self.after(20, self._maybe_close_after_blur)

    def _maybe_close_after_blur(self):
        try:
            focused = self.focus_get()
        except (KeyError, tk.TclError):
            focused = None
        if self._opened and not self._mouse_inside_popup and focused is not self.cb:
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

    def _on_caret_click(self):
        if self._opened:
            self._close_popup()
        else:
            if self.OPEN_ALL_ON_BUTTON and not self.get():
                self._ensure_data()
                self._open_popup(items=self._all_values[:])
            else:
                self._open_popup()
        self.after_idle(self._restore_entry_focus)

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
    # Entry + calendar button; popup calendar; returns YYYY-MM-DD; empty allowed.
    def __init__(self, master, label: str):
        super().__init__(master)
        ttk.Label(self, text=label).pack(side=tk.TOP, anchor="w")
        row = ttk.Frame(self); row.pack(fill=tk.X)
        self.var = tk.StringVar()
        self.entry = ttk.Entry(row, textvariable=self.var, width=14)
        self.entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.btn = ttk.Button(row, text="...", width=3, command=self._open)
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
        tp.withdraw()
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
        ttk.Button(hdr, text="<", width=2, command=self._prev_month).pack(side=tk.LEFT)
        self._title = ttk.Label(hdr, font=("", 10, "bold")); self._title.pack(side=tk.LEFT, padx=6)
        ttk.Button(hdr, text=">", width=2, command=self._next_month).pack(side=tk.RIGHT)

        self._grid = ttk.Frame(frm); self._grid.pack()
        self._render()

        self._position_popup()
        tp.deiconify()

    def _position_popup(self):
        popup = self._popup
        if not popup or not popup.winfo_exists():
            return
        popup.update_idletasks()
        pop_w = max(popup.winfo_reqwidth(), popup.winfo_width())
        pop_h = max(popup.winfo_reqheight(), popup.winfo_height())
        anchor_left = self.entry.winfo_rootx()
        anchor_right = self.btn.winfo_rootx() + self.btn.winfo_width()
        x = anchor_right - pop_w
        if x < anchor_left - 8:
            x = anchor_left
        y = self.entry.winfo_rooty() + self.entry.winfo_height() + 2
        screen_w = popup.winfo_screenwidth()
        screen_h = popup.winfo_screenheight()
        margin = 8
        x = max(margin, min(x, screen_w - pop_w - margin))
        if y + pop_h + margin > screen_h:
            y = self.entry.winfo_rooty() - pop_h - 2
        y = max(margin, min(y, screen_h - pop_h - margin))
        popup.geometry(f"{pop_w}x{pop_h}+{int(x)}+{int(y)}")

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
    COLS = 6

    def __init__(self, master, sizes: List[str], **kwargs):
        super().__init__(master, **kwargs)

        self.fixed_sizes = sizes[:]
        self.extra_sizes = ["__custom1__", "__custom2__"]
        self.sizes = self.fixed_sizes + self.extra_sizes

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

            if is_custom:
                v_label = tk.StringVar(value="" if sz.startswith("__") else sz)
                ttk.Entry(cell, textvariable=v_label, width=7, justify="center").pack()
            else:
                v_label = tk.StringVar(value=sz)
                ttk.Label(cell, text=sz, anchor="center").pack()

            ttk.Label(cell, text="الكمية", font=("", 8)).pack()
            v_qty = tk.StringVar()
            ttk.Entry(cell, textvariable=v_qty, width=6).pack()

            ttk.Label(cell, text="السعر", font=("", 8)).pack()
            v_price = tk.StringVar()
            ttk.Entry(cell, textvariable=v_price, width=8).pack()

            self._vars[sz] = (v_qty, v_price, v_label, is_custom)

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
                continue

            try:
                qty = int(qty_txt)
            except Exception:
                continue

            if qty <= 0:
                continue

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
        """Set the price for a specific size in the grid."""
        for sz in self.sizes:
            v_qty, v_price, v_label, is_custom = self._vars[sz]
            label = v_label.get().strip() if is_custom else sz
            if label == size_label:
                v_price.set(str(price))
                break

    def set_price_for_all(self, price):
        """Set the same price for all sizes in the grid."""
        for sz in self.sizes:
            v_qty, v_price, v_label, is_custom = self._vars[sz]
            v_price.set(str(price))

    def clear_quantities(self):
        """Clear quantity fields in the grid, keeping price fields intact."""
        for sz in self.sizes:
            v_qty, v_price, v_label, is_custom = self._vars[sz]
            v_qty.set("")


# ------------------- Toast Notification -------------------

class ToastNotification:
    """Non-blocking floating notification that auto-dismisses."""
    _active: list = []

    @classmethod
    def show(cls, parent, message, duration=2500, toast_type="success"):
        return cls(parent, message, duration, toast_type)

    def __init__(self, parent, message, duration, toast_type):
        colors = {
            "success": ("#166534", "#dcfce7", "#16a34a"),
            "info":    ("#1e40af", "#dbeafe", "#2563eb"),
            "warning": ("#854d0e", "#fef9c3", "#f59e0b"),
            "error":   ("#991b1b", "#fee2e2", "#dc2626"),
        }
        fg, bg, border = colors.get(toast_type, colors["info"])
        icons = {"success": "\u2714", "info": "\u2139", "warning": "\u26A0", "error": "\u2716"}
        icon = icons.get(toast_type, "")

        y_off = 8
        for t in ToastNotification._active:
            try:
                if t._frame.winfo_exists():
                    y_off += t._frame.winfo_reqheight() + 6
            except Exception:
                pass

        self._frame = tk.Frame(parent, bg=bg, highlightbackground=border,
                               highlightthickness=2, padx=16, pady=10, cursor="hand2")
        tk.Label(self._frame, text=f"{icon}  {message}", bg=bg, fg=fg,
                 font=("Segoe UI", 10, "bold"), wraplength=450, justify="center").pack()
        self._frame.place(relx=0.5, y=y_off, anchor="n")
        self._frame.lift()
        self._frame.bind("<Button-1>", lambda e: self._dismiss())
        ToastNotification._active.append(self)
        parent.after(duration, self._dismiss)

    def _dismiss(self):
        if self in ToastNotification._active:
            ToastNotification._active.remove(self)
        try:
            self._frame.destroy()
        except Exception:
            pass


# ------------------- Tooltip -------------------

class ToolTip:
    """Hover tooltip for any widget."""
    def __init__(self, widget, text_func=None, text="", delay=400):
        self.widget = widget
        self.text_func = text_func
        self.text = text
        self.delay = delay
        self._tip: Optional[tk.Toplevel] = None
        self._job: Optional[str] = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, event=None):
        self._hide()
        self._job = self.widget.after(self.delay, self._show)

    def _show(self):
        txt = self.text_func() if callable(self.text_func) else (self.text_func if self.text_func is not None else self.text)
        if not txt:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.attributes("-topmost", True)
        self._tip.geometry(f"+{x}+{y}")
        frm = tk.Frame(self._tip, bg="#1e293b", padx=10, pady=8,
                        highlightbackground="#475569", highlightthickness=1)
        frm.pack()
        tk.Label(frm, text=txt, bg="#1e293b", fg="#f1f5f9",
                 font=("Segoe UI", 9), justify="right", wraplength=350).pack()

    def _hide(self, event=None):
        if self._job:
            try:
                self.widget.after_cancel(self._job)
            except Exception:
                pass
            self._job = None
        if self._tip:
            try:
                self._tip.destroy()
            except Exception:
                pass
            self._tip = None


def _add_tooltip(widget, text="", delay=400, text_func=None):
    return ToolTip(widget, text_func=text_func, text=text, delay=delay)


# ------------------- Treeview Helpers -------------------

def _bind_mousewheel(widget, scrollable=None):
    """Bind mouse-wheel to *widget* so it scrolls on hover (no click needed).

    Works on Windows (<MouseWheel>) and Linux (<Button-4/5>).
    *scrollable* defaults to *widget* if omitted.
    """
    target = scrollable or widget

    def _on_wheel(event):
        try:
            if hasattr(event, "delta") and event.delta:
                target.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif getattr(event, "num", 0) == 4:
                target.yview_scroll(-3, "units")
            elif getattr(event, "num", 0) == 5:
                target.yview_scroll(3, "units")
        except Exception:
            pass
        return "break"          # prevent double-scroll from native handler

    def _enter(_e=None):
        widget.bind_all("<MouseWheel>", _on_wheel)
        widget.bind_all("<Button-4>", _on_wheel)
        widget.bind_all("<Button-5>", _on_wheel)

    def _leave(_e=None):
        try:
            widget.unbind_all("<MouseWheel>")
            widget.unbind_all("<Button-4>")
            widget.unbind_all("<Button-5>")
        except Exception:
            pass

    widget.bind("<Enter>", _enter, add="+")
    widget.bind("<Leave>", _leave, add="+")


def _apply_zebra_tags(tree):
    """Apply alternating row colors to a treeview after insert."""
    semantic_tags = {
        "surplus", "deficit", "verified", "current_touched", "history_touched",
        "draft", "void", "error", "warning", "ok",
    }
    for i, iid in enumerate(tree.get_children()):
        existing = list(tree.item(iid, "tags") or ())
        existing = [t for t in existing if t not in ("even", "odd")]
        if any(t in semantic_tags for t in existing):
            tree.item(iid, tags=tuple(existing))
            continue
        existing.append("even" if i % 2 == 0 else "odd")
        tree.item(iid, tags=tuple(existing))


def _add_context_menu(tree, parent_widget, extra_items=None):
    """Add a right-click context menu to a treeview."""
    menu = tk.Menu(parent_widget, tearoff=0)

    def _copy_cell():
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        col = tree.identify_column(menu._click_x)
        try:
            idx = int(col.replace("#", "")) - 1
            text = str(vals[idx]) if 0 <= idx < len(vals) else ""
        except Exception:
            text = " | ".join(str(v) for v in vals)
        try:
            parent_widget.clipboard_clear()
            parent_widget.clipboard_append(text)
        except Exception:
            pass

    def _copy_row():
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        text = " | ".join(str(v) for v in vals)
        try:
            parent_widget.clipboard_clear()
            parent_widget.clipboard_append(text)
        except Exception:
            pass

    menu.add_command(label="نسخ الخلية", command=_copy_cell)
    menu.add_command(label="نسخ الصف", command=_copy_row)

    if extra_items:
        menu.add_separator()
        for label, cmd in extra_items:
            menu.add_command(label=label, command=cmd)

    def _on_right_click(event):
        iid = tree.identify_row(event.y)
        if iid:
            tree.selection_set(iid)
        menu._click_x = event.x
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    tree.bind("<Button-3>", _on_right_click)
    return menu


# ------------------- Dashboard Frame -------------------

class DashboardFrame(ttk.Frame):
    """Home dashboard with today's summary, recent activity, and alerts."""

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        self._build()

    def _build(self):
        ttk.Label(self, text="\u2302  \u0644\u0648\u062D\u0629 \u0627\u0644\u062A\u062D\u0643\u0645",
                  font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 12))

        # Stats cards row
        cards = ttk.Frame(self)
        cards.pack(fill=tk.X, pady=(0, 12))

        self._today_sales = tk.StringVar(value="0")
        self._today_bills = tk.StringVar(value="0")
        self._today_reservations = tk.StringVar(value="0")
        self._pending_reservations = tk.StringVar(value="0")
        self._low_stock_count = tk.StringVar(value="0")

        card_defs = [
            ("\u0645\u0628\u064A\u0639\u0627\u062A \u0627\u0644\u064A\u0648\u0645", self._today_sales, "#2563eb"),
            ("\u0641\u0648\u0627\u062A\u064A\u0631 \u0627\u0644\u064A\u0648\u0645", self._today_bills, "#16a34a"),
            ("\u062D\u062C\u0648\u0632\u0627\u062A \u0627\u0644\u064A\u0648\u0645", self._today_reservations, "#f59e0b"),
            ("\u062D\u062C\u0648\u0632\u0627\u062A \u0645\u0639\u0644\u0642\u0629", self._pending_reservations, "#dc2626"),
            ("\u0645\u062E\u0632\u0648\u0646 \u0645\u0646\u062E\u0641\u0636", self._low_stock_count, "#7c3aed"),
        ]
        for col, (title, var, color) in enumerate(card_defs):
            card = tk.Frame(cards, bg="white", highlightbackground=color,
                            highlightthickness=2, padx=18, pady=14, cursor="hand2")
            card.grid(row=0, column=col, padx=6, pady=4, sticky="nsew")
            tk.Label(card, text=title, bg="white", fg="#64748b",
                     font=("Segoe UI", 9)).pack(anchor="center")
            tk.Label(card, textvariable=var, bg="white", fg=color,
                     font=("Segoe UI", 20, "bold")).pack(anchor="center", pady=(4, 0))
        for c in range(5):
            cards.columnconfigure(c, weight=1)

        # Bottom: recent bills + low stock alerts
        bottom = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        bottom.pack(fill=tk.BOTH, expand=True)

        recent = ttk.LabelFrame(bottom, text="\u0622\u062E\u0631 \u0627\u0644\u0641\u0648\u0627\u062A\u064A\u0631")
        bottom.add(recent, weight=1)
        self._recent_tbl = ttk.Treeview(
            recent, columns=("id", "date", "customer", "total"),
            show="headings", height=8)
        for col, txt, w in [("id", "#", 50), ("date", "\u0627\u0644\u062A\u0627\u0631\u064A\u062E", 140),
                             ("customer", "\u0627\u0644\u0639\u0645\u064A\u0644", 150), ("total", "\u0627\u0644\u0625\u062C\u0645\u0627\u0644\u064A", 90)]:
            self._recent_tbl.heading(col, text=txt)
            self._recent_tbl.column(col, width=w, anchor="center")
        self._recent_tbl.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        _bind_mousewheel(self._recent_tbl)

        alerts = ttk.LabelFrame(bottom, text="\u062A\u0646\u0628\u064A\u0647\u0627\u062A \u0627\u0644\u0645\u062E\u0632\u0648\u0646 \u0627\u0644\u0645\u0646\u062E\u0641\u0636")
        bottom.add(alerts, weight=1)
        self._alerts_tbl = ttk.Treeview(
            alerts, columns=("item", "school", "color", "size", "count"),
            show="headings", height=8)
        for col, txt, w in [("item", "\u0627\u0644\u0646\u0648\u0639", 100), ("school", "\u0627\u0644\u0645\u062F\u0631\u0633\u0629", 100),
                             ("color", "\u0627\u0644\u0644\u0648\u0646", 70), ("size", "\u0627\u0644\u0645\u0642\u0627\u0633", 50),
                             ("count", "\u0627\u0644\u0643\u0645\u064A\u0629", 50)]:
            self._alerts_tbl.heading(col, text=txt)
            self._alerts_tbl.column(col, width=w, anchor="center")
        self._alerts_tbl.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)
        _bind_mousewheel(self._alerts_tbl)
        self._alerts_tbl.tag_configure("low", foreground="#b45309")
        self._alerts_tbl.tag_configure("critical", foreground="#ffffff", background="#dc2626")

    def refresh(self):
        today = date.today().isoformat()
        try:
            stats = self.db.get_sales_stats(date_from=today, date_to=today)
            self._today_sales.set(f"{format_money(stats['sales_total'])}")
            self._today_bills.set(str(stats["sales_count"]))
            self._today_reservations.set(str(stats["res_count"]))
        except Exception:
            pass
        try:
            pending_count = self.db.count_reservations(status="\u0645\u0639\u0644\u0642")
            self._pending_reservations.set(str(pending_count))
        except Exception:
            pass
        try:
            self._recent_tbl.delete(*self._recent_tbl.get_children())
            for b in self.db.list_bills()[:10]:
                self._recent_tbl.insert("", tk.END, values=(
                    b["id"], fmt_local_ts(b["created_at"], ""),
                    b.get("customer", ""), f"{format_money(float(b['total']))}"))
            _apply_zebra_tags(self._recent_tbl)
        except Exception:
            pass
        try:
            self._alerts_tbl.delete(*self._alerts_tbl.get_children())
            low_count, low_rows = self.db.low_stock_summary(threshold=5, limit=80)
            for r in low_rows:
                cnt = int(r.get("count", 0))
                tag = "critical" if cnt <= 2 else "low"
                self._alerts_tbl.insert("", tk.END, values=(
                    r["item_type"], r["school"], r["color"], r["size"], cnt),
                    tags=(tag,))
            self._low_stock_count.set(str(low_count))
        except Exception:
            pass


# ------------------- Income Frame -------------------
class IncomeFrame(ttk.Frame):

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        self._income_r1 = tk.StringVar()
        self._income_r2 = tk.StringVar()
        self._income_has_alpha = tk.BooleanVar(value=False)
        self._staging_lines: List[Dict[str, Any]] = []

        self._build()

    def _build(self):
        ttk.Label(self, text="وارد (إضافة أصناف جديدة)", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))

        self._income_r1.trace_add("write", lambda *_: self._rebuild_sizes_grid())
        self._income_r2.trace_add("write", lambda *_: self._rebuild_sizes_grid())
        self._income_has_alpha.trace_add("write", lambda *_: self._rebuild_sizes_grid())

        # Main horizontal split
        vsplit = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        vsplit.pack(fill=tk.BOTH, expand=True)

        # LEFT side: existing form
        left = ttk.Frame(vsplit)
        vsplit.add(left, weight=6)

        grid = ttk.LabelFrame(left, text="بيانات الصنف")
        grid.pack(fill=tk.BOTH, expand=True)

        self.item_type = LabeledCombobox(grid, "النوع", self.db, "item_type")
        self.school    = LabeledCombobox(grid, "المدرسة", self.db, "school")
        self.color     = LabeledCombobox(grid, "اللون", self.db, "color")

        self.item_type.set_supplier(lambda: self.db.get_distinct_filtered("item_type", self._build_income_filter_constraints("item_type")))
        self.school.set_supplier(lambda: self.db.get_distinct_filtered("school", self._build_income_filter_constraints("school")))
        self.color.set_supplier(lambda: self.db.get_distinct_filtered("color", self._build_income_filter_constraints("color")))

        ranges_box = ttk.LabelFrame(grid, text="نطاقات المقاسات (للوارد فقط)")
        ranges_box.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=6)

        ttk.Label(ranges_box, text="النطاق الأول").grid(row=0, column=0, sticky="w")
        ttk.Combobox(
            ranges_box,
            textvariable=self._income_r1,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=18
        ).grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(ranges_box, text="النطاق الثاني").grid(row=1, column=0, sticky="w")
        ttk.Combobox(
            ranges_box,
            textvariable=self._income_r2,
            values=[""] + NUMERIC_RANGE_LABELS,
            state="readonly",
            width=18
        ).grid(row=1, column=1, sticky="w", padx=4)

        ttk.Checkbutton(
            ranges_box,
            text="مقاسات حرفية (S إلى 5XL)",
            variable=self._income_has_alpha
        ).grid(row=2, column=0, columnspan=2, sticky="w")

        for w in (self.item_type.cb, self.school.cb, self.color.cb):
            w.bind("<<ComboboxSelected>>", lambda e: (self._on_income_filter_changed(), self._auto_load_size_profile(), self._auto_fill_price_for_grid()), add="+")
            w.bind("<FocusOut>",           lambda e: (self._auto_load_size_profile(), self._auto_fill_price_for_grid()), add="+")

        sizes_container = ttk.Frame(grid)
        sizes_container.grid(row=3, column=0, columnspan=2, sticky="nsew", padx=6, pady=6)

        sizes_canvas = tk.Canvas(sizes_container, highlightthickness=0)
        sizes_scroll = ttk.Scrollbar(sizes_container, orient="vertical", command=sizes_canvas.yview)
        sizes_canvas.configure(yscrollcommand=sizes_scroll.set)

        sizes_scroll.pack(side="right", fill="y")
        sizes_canvas.pack(side="left", fill="both", expand=True)

        self._sizes_inner = ttk.Frame(sizes_canvas)
        sizes_window = sizes_canvas.create_window((0, 0), window=self._sizes_inner, anchor="nw")

        self.sizes_grid = None

        def _on_sizes_config(event=None):
            try:
                sizes_canvas.configure(scrollregion=sizes_canvas.bbox("all"))
                sizes_canvas.itemconfigure(sizes_window, width=sizes_canvas.winfo_width())
            except Exception:
                pass

        self._sizes_inner.bind("<Configure>", _on_sizes_config)
        sizes_canvas.bind("<Configure>", _on_sizes_config)

        _bind_mousewheel(sizes_canvas)

        self.item_type.grid(row=0, column=0, padx=6, pady=6, sticky="ew")
        self.school.grid(   row=0, column=1, padx=6, pady=6, sticky="ew")
        self.color.grid(    row=1, column=0, padx=6, pady=6, sticky="ew")

        for c in range(2):
            grid.columnconfigure(c, weight=1)

        try:
            grid.rowconfigure(3, weight=1)
            grid.rowconfigure(4, weight=0)
            grid.rowconfigure(5, weight=0)
        except Exception:
            pass

        add_btns = ttk.Frame(grid)
        add_btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 6))

        ttk.Button(add_btns, text="إضافة للقائمة", command=self._add_to_staging).pack(side=tk.LEFT)
        ttk.Button(add_btns, text="تفريغ", command=self._on_reset_keep_pkg).pack(side=tk.LEFT, padx=8)

        # RIGHT side: staging area
        right = ttk.LabelFrame(vsplit, text="فاتورة الوارد")
        vsplit.add(right, weight=4)

        # Supplier field
        sup_row = ttk.Frame(right)
        sup_row.pack(fill=tk.X, padx=4, pady=(4, 2))
        ttk.Label(sup_row, text="المورد:").pack(side=tk.LEFT)
        self._supplier_var = tk.StringVar()
        ttk.Entry(sup_row, textvariable=self._supplier_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))

        # Staging treeview
        tree_wrap = ttk.Frame(right)
        tree_wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=2)
        self._staging_table = ttk.Treeview(
            tree_wrap,
            columns=("type", "school", "color", "size", "price", "qty", "total"),
            show="headings", height=10,
        )
        for col, txt, w in [
            ("type", "النوع", 100), ("school", "المدرسة", 100), ("color", "اللون", 70),
            ("size", "المقاس", 60), ("price", "السعر", 70), ("qty", "الكمية", 60), ("total", "الإجمالي", 80),
        ]:
            self._staging_table.heading(col, text=txt)
            self._staging_table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self._staging_table.yview)
        self._staging_table.configure(yscrollcommand=ysb.set)
        self._staging_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._staging_table)

        # Context menu on staging table
        _add_context_menu(self._staging_table, self, extra_items=[
            ("حذف السطر", self._remove_staging_line),
        ])

        # Controls
        ctrl_row = ttk.Frame(right)
        ctrl_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(ctrl_row, text="حذف السطر", command=self._remove_staging_line).pack(side=tk.LEFT, padx=2)
        ttk.Button(ctrl_row, text="تفريغ القائمة", command=self._clear_staging).pack(side=tk.LEFT, padx=2)

        # Total
        tot_row = ttk.Frame(right)
        tot_row.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(tot_row, text="الإجمالي:").pack(side=tk.RIGHT)
        self._staging_total_var = tk.StringVar(value="0")
        ttk.Label(tot_row, textvariable=self._staging_total_var, font=("Segoe UI", 12, "bold")).pack(side=tk.RIGHT, padx=(0, 6))

        self._staging_qty_var = tk.StringVar(value="0")
        ttk.Label(tot_row, text="القطع:").pack(side=tk.RIGHT, padx=(12, 0))
        ttk.Label(tot_row, textvariable=self._staging_qty_var, font=("Segoe UI", 10)).pack(side=tk.RIGHT, padx=(0, 6))

        # Finalize button
        ttk.Button(right, text="تأكيد فاتورة الوارد", command=self._finalize_income).pack(fill=tk.X, padx=4, pady=(4, 4))

    def _build_income_filter_constraints(self, exclude_field: str) -> Dict[str, Any]:
        c: Dict[str, Any] = {}
        if exclude_field != "school" and hasattr(self, 'school'):
            v = self.school.get()
            if v:
                c["school"] = v
        if exclude_field != "item_type" and hasattr(self, 'item_type'):
            v = self.item_type.get()
            if v:
                c["item_type"] = v
        if exclude_field != "color" and hasattr(self, 'color'):
            v = self.color.get()
            if v:
                c["color"] = v
        return c

    def _on_income_filter_changed(self, *_):
        self.item_type.refresh_values()
        self.school.refresh_values()
        self.color.refresh_values()

    def _auto_load_size_profile(self):
        """Auto-load size ranges from saved profile when all 3 specs are selected."""
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        if not (it and sc and cl):
            return
        profile = self.db.get_size_profile(it, sc, cl)
        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            r1_label = ""
            r2_label = ""
            for (a, b) in ALLOWED_NUMERIC_RANGES:
                if a == r1s and b == r1e:
                    r1_label = f"{a} \u2192 {b}"
                if a == r2s and b == r2e:
                    r2_label = f"{a} \u2192 {b}"
            self._income_r1.set(r1_label)
            self._income_r2.set(r2_label)
            self._income_has_alpha.set(bool(has_alpha))
        else:
            self._rebuild_sizes_grid()

    def _rebuild_sizes_grid(self):
        try:
            self.sizes_grid.destroy()
        except Exception:
            pass

        numeric_ranges = []

        def _parse_range(label: str):
            if not label:
                return None
            a, b = parse_numeric_range_label(label)
            if a is None or b is None:
                return None
            return a, b

        for lbl in (self._income_r1.get(), self._income_r2.get()):
            pair = _parse_range(lbl)
            if pair:
                numeric_ranges.append(pair)

        numeric_ranges.sort()
        merged = []

        for start, end in numeric_ranges:
            if not merged:
                merged.append([start, end])
            else:
                last_start, last_end = merged[-1]
                if start <= last_end:
                    merged[-1][1] = max(last_end, end)
                else:
                    merged.append([start, end])

        sizes = []
        for start, end in merged:
            sizes.extend(str(x) for x in range(start, end + 1, 2))

        if self._income_has_alpha.get():
            sizes.extend(ALPHA_SIZES)

        self.sizes_grid = SizesGrid(self._sizes_inner, sizes)
        self.sizes_grid.pack(fill="both", expand=False)

        self._sizes_inner.update_idletasks()

    def _auto_fill_price_for_grid(self):
        it = (self.item_type.get() or "").strip()
        sc = (self.school.get() or "").strip()
        cl = (self.color.get() or "").strip()
        if not (it and sc and cl):
            return
        try:
            p = self.db.last_price_for_specs(it, sc, cl, "")
            if p is None:
                try:
                    sizes = self.db.list_sizes_for_item(sc, it, cl)
                except Exception:
                    sizes = []
                if sizes:
                    for row in sizes:
                        size_label = str(row.get("size") or "")
                        last_p = row.get("last_price")
                        if last_p is not None:
                            self.sizes_grid.set_price_for_size(size_label, last_p)
                    return
            if p is not None:
                self.sizes_grid.set_price_for_all(p)
        except Exception:
            pass

    def _on_reset_keep_pkg(self):
        for w in (self.item_type, self.school, self.color):
            w.set("")
        self._income_r1.set("")
        self._income_r2.set("")
        self._income_has_alpha.set(False)

        if self.sizes_grid:
            self.sizes_grid.destroy()
            self.sizes_grid = None

        self._sizes_inner.update_idletasks()

    def _add_to_staging(self):
        """Add current form items to the staging list (not yet saved to DB)."""
        item_type = (self.item_type.get() or "").strip()
        school = (self.school.get() or "").strip()
        color = (self.color.get() or "").strip()
        if not item_type or not school or not color:
            messagebox.showwarning("بيانات ناقصة", "يجب اختيار النوع والمدرسة واللون.", parent=self)
            return

        if self.sizes_grid is None:
            messagebox.showwarning("فارغ", "لم تُدخل أي كميات في شبكة المقاسات.", parent=self)
            return

        rows = self.sizes_grid.get_rows()
        added = 0
        for r in rows:
            sz = str(r["size"]).strip()
            qty = int(r["qty"])
            if qty <= 0:
                continue

            price = r["price"]
            if price is None or float(price) <= 0:
                try:
                    price = float(self.db.get_effective_price(item_type, school, color, sz) or 0)
                except Exception:
                    price = 0.0
            else:
                price = float(price)

            # Merge with existing staging line if same specs
            merged = False
            for line in self._staging_lines:
                if (line["item_type"] == item_type and line["school"] == school
                        and line["color"] == color and line["size"] == sz
                        and abs(line["unit_price"] - price) < 0.001):
                    line["qty"] += qty
                    merged = True
                    break

            if not merged:
                self._staging_lines.append({
                    "item_type": item_type,
                    "school": school,
                    "color": color,
                    "size": sz,
                    "unit_price": price,
                    "qty": qty,
                })
            added += 1

        if added == 0:
            messagebox.showwarning("لا توجد كميات", "أدخل كمية واحدة على الأقل.", parent=self)
            return

        self._sync_staging_table()
        # Clear the size grid quantities but keep the prices
        self.sizes_grid.clear_quantities()

    def _sync_staging_table(self):
        """Refresh the staging treeview from in-memory list."""
        self._staging_table.delete(*self._staging_table.get_children())
        total_value = 0.0
        total_qty = 0
        for idx, ln in enumerate(self._staging_lines):
            line_total = float(ln["unit_price"]) * int(ln["qty"])
            total_value += line_total
            total_qty += int(ln["qty"])
            self._staging_table.insert(
                "", tk.END, iid=str(idx),
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"],
                        f"{format_money(float(ln['unit_price']))}", ln["qty"], f"{format_money(line_total)}")
            )
        self._staging_total_var.set(f"{format_money(total_value)}")
        self._staging_qty_var.set(str(total_qty))
        _apply_zebra_tags(self._staging_table)

    def _remove_staging_line(self):
        sel = self._staging_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self._staging_lines):
            self._staging_lines.pop(idx)
            self._sync_staging_table()

    def _clear_staging(self):
        self._staging_lines.clear()
        self._sync_staging_table()

    def _finalize_income(self):
        if not self.db.is_manager_feature_enabled("allow_manual_incoming"):
            messagebox.showwarning("مقيد", _feature_restricted_message("فاتورة الوارد اليدوية متوقفة حالياً في نقطة البيع."), parent=self)
            return
        if not self._staging_lines:
            messagebox.showwarning("فارغ", "لا توجد أصناف في فاتورة الوارد.", parent=self)
            return

        supplier = self._supplier_var.get().strip()

        try:
            bill_id = self.db.create_income_bill(supplier, self._staging_lines)
            total = self._staging_total_var.get()
            qty = self._staging_qty_var.get()
            ToastNotification.show(self.winfo_toplevel(),
                f"تم إنشاء فاتورة وارد #{bill_id} — القطع: {qty} — الإجمالي: {total}",
                toast_type="success")
            self._staging_lines.clear()
            self._sync_staging_table()
            self._supplier_var.set("")
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)

    @staticmethod
    def _err(msg: str) -> None:
        raise RuntimeError(msg)


# ------------------- POS Frame (replaces OutcomeFrame) -------------------

class POSFrame(ttk.Frame):
    """
    Multi-bill POS with step-by-step item navigation.
    Bills 1-3 are regular sales; bill 4 is returns; bills 5 and 7 are reservations; bill 6 is exchanges.
    """
    RESERVATION_BILLS = (5, 7)

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=6)
        self.db = db

        # In-memory bill state
        self.bills = {1: [], 2: [], 3: [], 4: [], 5: [], 6: [], 7: []}
        self.customers = {1: "", 2: "", 3: "", 4: "", 5: "", 6: "", 7: ""}
        self.customer_phones = {1: "", 2: "", 3: "", 4: "", 5: "", 6: "", 7: ""}
        self.active_bill = 1
        self._exchange_mode = "return"  # "return" or "take" — for bill 6

        # Navigation state
        self._sel_school: Optional[str] = None
        self._sel_item: Optional[str] = None
        self._sel_color: Optional[str] = None
        self._sel_size: Optional[str] = None
        self._sizes_cache: List[Dict[str, Any]] = []
        self._size_btns: Dict[str, Any] = {}
        self._selected_size_btn = None
        self._price_user_edited = False

        self._build()

    # ------------------------------------------------------------------ build
    def _build(self):
        # Define styles once
        try:
            s = ttk.Style(self)
            s.configure("Size.TButton", padding=6)
            s.configure("SizeZero.TButton",
                        background="#f8d7da", foreground="#000000",
                        bordercolor="#b91c1c", focusthickness=1, focuscolor="#b91c1c", padding=6)
            s.map("SizeZero.TButton",
                background=[("active", "#f1b0b7"), ("pressed", "#f1b0b7")],
                bordercolor=[("focus", "#b91c1c")])
            s.configure("SizeSelected.TButton",
                        background="#d1fae5", foreground="#000000",
                        bordercolor="#059669", focusthickness=1, focuscolor="#059669", padding=6)
            s.map("SizeSelected.TButton",
                background=[("active", "#a7f3d0"), ("pressed", "#a7f3d0")],
                bordercolor=[("focus", "#059669")])
            s.configure("SizeSelectedZero.TButton",
                        background="#fbd4a8", foreground="#000000",
                        bordercolor="#b45309", focusthickness=1, focuscolor="#b45309", padding=6)
            s.map("SizeSelectedZero.TButton",
                background=[("active", "#f8c58a"), ("pressed", "#f8c58a")],
                bordercolor=[("focus", "#b45309")])
            # Low stock style (count 1-5) — amber/yellow warning
            s.configure("SizeLow.TButton",
                        background="#fef3c7", foreground="#000000",
                        bordercolor="#d97706", focusthickness=1, focuscolor="#d97706", padding=6)
            s.map("SizeLow.TButton",
                background=[("active", "#fde68a"), ("pressed", "#fde68a")],
                bordercolor=[("focus", "#d97706")])
            s.configure("ActiveBill.TButton",
                        background="#059669", foreground="#ffffff",
                        bordercolor="#047857", padding=6)
            s.map("ActiveBill.TButton",
                background=[("active", "#047857"), ("pressed", "#047857")])
        except Exception:
            pass

        # Main horizontal split: left = selector, right = bill panel
        vsplit = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        vsplit.pack(fill=tk.BOTH, expand=True)

        # ---- LEFT: item selector ----
        left = ttk.Frame(vsplit)
        vsplit.add(left, weight=6)

        # Universal search bar
        search_row = ttk.Frame(left)
        search_row.pack(fill=tk.X, padx=4, pady=(4, 2))
        ttk.Label(search_row, text="\u0628\u062D\u062B \u0633\u0631\u064A\u0639:", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 4))
        self._search_var = tk.StringVar()
        self._search_entry = ttk.Entry(search_row, textvariable=self._search_var, width=30)
        self._search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self._search_entry.bind("<KeyRelease>", lambda e: self._on_search_changed())
        ttk.Button(search_row, text="\u2716", width=3, command=self._clear_search,
                   style="Secondary.TButton").pack(side=tk.LEFT, padx=(4, 0))

        # Quick filter bar
        fbar = ttk.LabelFrame(left, text="\u062A\u0635\u0641\u064A\u0629 \u0633\u0631\u064A\u0639\u0629")
        fbar.pack(fill=tk.X, padx=4, pady=(2, 2))

        self._flt_school = LabeledCombobox(fbar, "\u0627\u0644\u0645\u062F\u0631\u0633\u0629", self.db, "school")
        self._flt_school.set_supplier(lambda: self.db.get_distinct_filtered("school", self._build_filter_constraints("school")))
        self._flt_school.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._flt_item = LabeledCombobox(fbar, "\u0627\u0644\u0646\u0648\u0639", self.db, "item_type")
        self._flt_item.set_supplier(lambda: self.db.get_distinct_filtered("item_type", self._build_filter_constraints("item_type")))
        self._flt_item.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._flt_color = LabeledCombobox(fbar, "\u0627\u0644\u0644\u0648\u0646", self.db, "color")
        self._flt_color.set_supplier(lambda: self.db.get_distinct_filtered("color", self._build_filter_constraints("color")))
        self._flt_color.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(fbar, text="\u0645\u0633\u062D", command=self._clear_quick_filter,
                   style="Secondary.TButton").grid(row=0, column=3, padx=4, pady=4)

        for c in range(4):
            fbar.columnconfigure(c, weight=1 if c < 3 else 0)

        # Auto-apply filters when a value is selected from the dropdown
        self._flt_school.cb.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed(), add="+")
        self._flt_item.cb.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed(), add="+")
        self._flt_color.cb.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed(), add="+")

        # Breadcrumb
        crumb = ttk.Frame(left)
        crumb.pack(fill=tk.X, padx=4, pady=(2, 0))
        self._crumb_var = tk.StringVar(value="اختر المدرسة")
        ttk.Label(crumb, textvariable=self._crumb_var, font=("", 10, "bold")).pack(side=tk.LEFT, padx=(2, 6))

        # Scrollable button grid
        sc = ttk.Frame(left)
        sc.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        self._canvas = tk.Canvas(sc, highlightthickness=0, bd=0)
        self._scroll_y = ttk.Scrollbar(sc, orient="vertical", command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._scroll_y.set)

        try:
            bg = ttk.Style(self).lookup("TFrame", "background") or "#ffffff"
            self._canvas.configure(background=bg)
        except Exception:
            pass

        self._grid_host = ttk.Frame(self._canvas)
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

        # Scroll wheel support for the button grid
        _bind_mousewheel(self._canvas)

        # Favorites section (most-sold items)
        self._fav_frame = ttk.LabelFrame(left, text="\u2605 \u0627\u0644\u0623\u0643\u062B\u0631 \u0645\u0628\u064A\u0639\u0627\u064B")
        self._fav_frame.pack(fill=tk.X, padx=4, pady=(2, 0))
        self._fav_inner = ttk.Frame(self._fav_frame)
        self._fav_inner.pack(fill=tk.X, padx=4, pady=4)
        self._refresh_favorites()

        # ---- RIGHT: bill management ----
        right = ttk.LabelFrame(vsplit, text="\u0627\u0644\u0641\u0648\u0627\u062A\u064A\u0631")
        vsplit.add(right, weight=4)
        # Start close to the operator-tuned POS ratio from testing.
        def _set_bill_sash(_e=None):
            try:
                w = vsplit.winfo_width()
                if w > 100:
                    vsplit.sashpos(0, int(w * 0.56))
                    vsplit.unbind("<Map>")
            except Exception:
                pass
        vsplit.bind("<Map>", _set_bill_sash, add="+")

        # Bill switcher buttons (1-7)
        switcher = ttk.Frame(right)
        switcher.pack(fill=tk.X, padx=4, pady=(4, 2))
        self._bill_btns: Dict[int, ttk.Button] = {}
        _bill_labels = {
            1: "1", 2: "2", 3: "3",
            4: "4 (مرتجع)", 5: "5 (حجز)", 6: "6 (استبدال)", 7: "7 (حجز)",
        }
        for n in range(1, 8):
            btn = ttk.Button(switcher, text=_bill_labels[n], width=9,
                             command=lambda b=n: self._switch_bill(b))
            btn.pack(side=tk.LEFT, padx=2)
            self._bill_btns[n] = btn
        self._update_bill_btn_styles()

        # Customer field
        self._customer_var = tk.StringVar()
        self._customer_phone_var = tk.StringVar()
        cust_row = ttk.Frame(right)
        cust_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Label(cust_row, text="العميل:").pack(side=tk.LEFT)
        self._cust_entry = ttk.Entry(cust_row, textvariable=self._customer_var)
        self._cust_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        ttk.Label(cust_row, text="رقم:").pack(side=tk.LEFT, padx=(6, 0))
        self._cust_phone_entry = ttk.Entry(cust_row, textvariable=self._customer_phone_var, width=16)
        self._cust_phone_entry.pack(side=tk.LEFT, padx=(4, 0))
        _wh_btn = ttk.Button(
            cust_row,
            text=WAREHOUSE_RETURN_LABEL,
            command=lambda: self._customer_var.set(WAREHOUSE_RETURN_LABEL),
            style="Secondary.TButton",
        )
        _wh_btn.pack(side=tk.LEFT, padx=(4, 0))
        ToolTip(_wh_btn, "استخدم الفاتورة الحالية كإرجاع مخزون من الفرع إلى المصنع / المخزن الرئيسي")
        _branch_btn = ttk.Button(
            cust_row,
            text="طلب تحويل",
            command=self._choose_branch_target,
            style="Secondary.TButton",
        )
        _branch_btn.pack(side=tk.LEFT, padx=(4, 0))
        ToolTip(_branch_btn, "سجّل طلب تحويل إلى فرع آخر. التنفيذ الفعلي يتم من خلال المخزن فقط")
        self._cust_entry.bind("<FocusOut>", lambda e: self._save_customer(), add="+")
        self._cust_entry.bind("<KeyRelease>", lambda e: self._autocomplete_customer(), add="+")
        self._cust_phone_entry.bind("<FocusOut>", lambda e: self._save_customer(), add="+")
        self._cust_phone_entry.bind("<KeyRelease>", lambda e: self._save_customer(), add="+")
        self._cust_listbox = None

        # Bill items treeview
        tree_wrap = ttk.Frame(right)
        tree_wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=(2, 2))
        self.bill_table = ttk.Treeview(
            tree_wrap,
            columns=("type", "school", "color", "size", "price", "qty", "total"),
            show="headings",
            height=8,
        )
        for col, txt, w in [
            ("type", "النوع", 100), ("school", "المدرسة", 120), ("color", "اللون", 70),
            ("size", "المقاس", 60), ("price", "السعر", 70), ("qty", "الكمية", 60), ("total", "الإجمالي", 80),
        ]:
            self.bill_table.heading(col, text=txt)
            self.bill_table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.bill_table.yview)
        self.bill_table.configure(yscrollcommand=ysb.set)
        self.bill_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self.bill_table)

        # Context menu on bill table
        _add_context_menu(self.bill_table, self, extra_items=[
            ("حذف السطر", self._remove_line),
        ])

        # Qty controls
        qty_row = ttk.Frame(right)
        qty_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Button(qty_row, text="-", width=3, command=self._dec_qty).pack(side=tk.LEFT)
        ttk.Button(qty_row, text="+", width=3, command=self._inc_qty).pack(side=tk.LEFT, padx=2)
        ttk.Button(qty_row, text="حذف السطر", command=self._remove_line).pack(side=tk.LEFT, padx=4)
        ttk.Button(qty_row, text="تفريغ", command=self._clear_bill).pack(side=tk.LEFT, padx=4)

        # Total display
        tot_row = ttk.Frame(right)
        tot_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Label(tot_row, text="الإجمالي:").pack(side=tk.RIGHT)
        self.total_var = tk.StringVar(value="0")
        ttk.Label(tot_row, textvariable=self.total_var, font=("Segoe UI", 12, "bold")).pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Label(tot_row, text="عدد القطع:").pack(side=tk.RIGHT, padx=(12, 0))
        self.total_items_var = tk.StringVar(value="0")
        ttk.Label(tot_row, textvariable=self.total_items_var, font=("Segoe UI", 12, "bold")).pack(side=tk.RIGHT, padx=(0, 6))

        # Reservation extras (visible for reservation bill slots)
        self._res_frame = ttk.LabelFrame(right, text="بيانات الحجز")
        self._res_frame.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Label(self._res_frame, text="المبلغ المدفوع:").grid(row=0, column=0, padx=4, pady=4, sticky="e")
        self._paid_var = tk.StringVar(value="0")
        ttk.Entry(self._res_frame, textvariable=self._paid_var, width=12).grid(row=0, column=1, padx=4, pady=4, sticky="w")
        ttk.Label(self._res_frame, text="ملاحظة:").grid(row=1, column=0, padx=4, pady=4, sticky="e")
        self._res_note_var = tk.StringVar()
        ttk.Entry(self._res_frame, textvariable=self._res_note_var, width=24).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self._res_frame.columnconfigure(1, weight=1)

        # Return extras (only visible for bill 4)
        self._return_frame = ttk.LabelFrame(right, text="بيانات المرتجع")
        ttk.Label(self._return_frame, text="أضف الأصناف المراد إرجاعها",
                  font=("Segoe UI", 9, "italic")).pack(padx=8, pady=4, anchor="w")

        # Exchange extras (only visible for bill 6)
        self._exchange_frame = ttk.LabelFrame(right, text="بيانات الاستبدال")
        ex_mode_row = ttk.Frame(self._exchange_frame)
        ex_mode_row.pack(fill=tk.X, padx=10, pady=(8, 4))
        self._exchange_mode_var = tk.StringVar(value="return")
        ttk.Radiobutton(ex_mode_row, text="↩ إضافة مرتجع", variable=self._exchange_mode_var,
                        value="return", command=self._on_exchange_mode_changed).pack(side=tk.LEFT, padx=(0, 18), pady=2)
        ttk.Radiobutton(ex_mode_row, text="↪ إضافة مأخوذ", variable=self._exchange_mode_var,
                        value="take", command=self._on_exchange_mode_changed).pack(side=tk.LEFT, pady=2)
        self._exchange_diff_var = tk.StringVar(value="فرق السعر: 0")
        ttk.Label(self._exchange_frame, textvariable=self._exchange_diff_var,
                  font=("Segoe UI", 11, "bold")).pack(fill=tk.X, padx=10, pady=(2, 8), anchor="w")

        # Finalize button with shortcut hint
        ttk.Button(right, text="تأكيد الفاتورة / الحجز (F8)", command=self._finalize,
                   style="Success.TButton").pack(fill=tk.X, padx=4, pady=(4, 4))

        self._update_res_frame_visibility()

        # Initial render - start with schools
        self._render_schools()

    # ------------------------------------------------------------------ helpers
    def _clear_grid(self):
        for w in self._grid_host.winfo_children():
            w.destroy()
        self._size_btns.clear()
        self._selected_size_btn = None

    def _mk_grid_buttons(self, labels: List[str], on_click, *, cols: int = 5):
        if not labels:
            ttk.Label(self._grid_host, text="لا توجد بيانات.").pack(pady=8)
            return
        row = None
        for i, label in enumerate(labels):
            if i % cols == 0:
                row = ttk.Frame(self._grid_host)
                row.pack(fill=tk.X, padx=4, pady=2)
            btn = ttk.Button(row, text=label, command=lambda v=label: on_click(v))
            btn.pack(side=tk.LEFT, padx=3, pady=3)

    # ------------------------------------------------------------------ navigation
    # Flow: Schools -> Items -> Colors -> Sizes

    def _back_to_schools(self):
        self._sel_school = None
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._price_user_edited = False
        try:
            self._flt_school.set("")
            self._flt_item.set("")
            self._flt_color.set("")
            self._flt_school.refresh_values()
            self._flt_item.refresh_values()
            self._flt_color.refresh_values()
        except Exception:
            pass
        self._render_schools()

    def _back_to_items(self):
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._price_user_edited = False
        try:
            self._flt_item.set("")
            self._flt_color.set("")
            self._flt_item.refresh_values()
            self._flt_color.refresh_values()
        except Exception:
            pass
        self._render_items()

    def _back_to_colors(self):
        self._sel_color = None
        self._sel_size = None
        self._price_user_edited = False
        try:
            self._flt_color.set("")
            self._flt_color.refresh_values()
            self._flt_item.refresh_values()
            self._flt_school.refresh_values()
        except Exception:
            pass
        self._render_colors()

    def _render_schools(self):
        """Entry point: show school buttons."""
        self._sel_school = None
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._crumb_var.set("اختر المدرسة")
        self._clear_grid()

        school_f = self._flt_school.get() or None
        item_f = self._flt_item.get() or None
        color_f = self._flt_color.get() or None

        constraints: Dict[str, Any] = {}
        if item_f:
            constraints["item_type"] = item_f
        if color_f:
            constraints["color"] = color_f

        try:
            schools = sorted(
                self.db.get_distinct_filtered(
                    "school",
                    constraints,
                    available_only=False,
                )
            )
        except Exception:
            schools = []

        if school_f:
            schools = [s for s in schools if s == school_f]

        self._mk_grid_buttons(schools, self._select_school, cols=4)

    def _select_school(self, school: str):
        self._sel_school = school
        self._price_user_edited = False
        self._render_items()

    def _render_items(self):
        """Show item type buttons for the selected school."""
        self._sel_item = None
        self._sel_size = None
        if self._sel_color:
            self._crumb_var.set(
                f"المدرسة: {self._sel_school}  \u27f6  اللون: {self._sel_color}  \u27f6  اختر النوع"
            )
        else:
            self._crumb_var.set(f"المدرسة: {self._sel_school}  \u27f6  اختر النوع")
        self._clear_grid()

        item_f = self._flt_item.get() or None

        constraints: Dict[str, Any] = {"school": self._sel_school}
        if self._sel_color:
            constraints["color"] = self._sel_color

        try:
            items = self.db.get_distinct_filtered(
                "item_type",
                constraints,
                available_only=False,
            )
        except Exception:
            items = []

        items = sort_item_type_values(items)
        if item_f:
            items = [i for i in items if i == item_f]

        self._mk_grid_buttons(items, self._select_item, cols=4)
        ttk.Button(self._grid_host, text="\u25c4 رجوع إلى المدارس", command=self._back_to_schools)\
            .pack(anchor="w", padx=4, pady=4)

    def _select_item(self, item_type: str):
        self._sel_item = item_type
        self._price_user_edited = False
        self._render_colors()

    def _render_colors(self):
        """Show color buttons for selected school + item type."""
        self._sel_color = None
        self._sel_size = None
        self._crumb_var.set(
            f"المدرسة: {self._sel_school}  \u27f6  النوع: {self._sel_item}  \u27f6  اختر اللون"
        )
        self._clear_grid()

        color_f = self._flt_color.get() or None
        constraints: Dict[str, Any] = {"school": self._sel_school}
        if self._sel_item:
            constraints["item_type"] = self._sel_item
        try:
            colors = self.db.get_distinct_filtered(
                "color",
                constraints,
                available_only=False,
            )
        except Exception:
            colors = []

        if color_f:
            colors = [c for c in colors if c == color_f]

        self._mk_grid_buttons(colors, self._select_color, cols=4)
        ttk.Button(self._grid_host, text="\u25c4 رجوع إلى الأنواع", command=self._back_to_items)\
            .pack(anchor="w", padx=4, pady=4)

    def _select_color(self, color: str):
        self._sel_color = color
        self._price_user_edited = False
        self._render_sizes()

    @staticmethod
    def _spec_match(a: str, b: str) -> bool:
        return (a or "").strip().casefold() == (b or "").strip().casefold()

    def _pending_out_qty_for_specs(self, school: str, item: str, color: str, size: str) -> int:
        """Qty on open bills that will deduct shop stock when finalized (not yet in DB)."""
        total = 0
        for b, lines in (self.bills or {}).items():
            if b in (1, 2, 3):
                for ln in lines or []:
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
            elif b == 6:
                for ln in lines or []:
                    if ln.get("direction") != "take":
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

    def _active_add_requires_stock(self) -> bool:
        if self.active_bill in (1, 2, 3):
            return True
        return self.active_bill == 6 and self._exchange_mode == "take"

    def _rerender_selector_for_active_mode(self) -> None:
        """Rebuild the selector after bill mode changes its stock visibility rule."""
        try:
            self._flt_school.refresh_values()
            self._flt_item.refresh_values()
            self._flt_color.refresh_values()
        except Exception:
            pass

        query = ""
        try:
            query = (self._search_var.get() or "").strip()
        except Exception:
            query = ""
        if query:
            self._on_search_changed()
            return

        if self._sel_school and self._sel_item and self._sel_color:
            self._render_sizes(preserve_size=self._sel_size)
        elif self._sel_school and self._sel_item:
            self._render_colors()
        elif self._sel_school:
            self._render_items()
        else:
            self._render_schools()

    def _available_qty_for_active_add(self, size: str) -> int:
        return self._available_qty_for_specs(
            self._sel_school or "", self._sel_item or "", self._sel_color or "", str(size or ""))

    def _available_qty_for_specs(self, school: str, item: str, color: str, size: str) -> int:
        try:
            rows = self.db.current_inventory({
                "school": school,
                "item_type": item,
                "color": color,
                "size": size,
            })
            on_hand = sum(int(r.get("count") or 0) for r in rows)
        except Exception:
            on_hand = 0
        pending = self._pending_out_qty_for_specs(school, item, color, str(size or ""))
        return max(0, int(on_hand) - int(pending))

    def _line_requires_stock(self, bill_no: int, line: Dict[str, Any]) -> bool:
        return bill_no in (1, 2, 3)

    def _warn_unavailable_add(self, size: str, requested: int, available: int) -> None:
        messagebox.showwarning(
            "كمية غير متاحة",
            f"لا يمكن إضافة هذا المقاس إلى الفاتورة.\nالمطلوب: {requested}\nالمتاح: {available}",
            parent=self,
        )

    def _legacy_pending_out_qty_for_specs_active_only(self, school: str, item: str, color: str, size: str) -> int:
        """Kept unused for reference during the POS stock guard transition."""
        b = self.active_bill
        lines = self.bills.get(b) or []
        total = 0
        if b in (1, 2, 3):
            for ln in lines:
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
        elif b == 6:
            for ln in lines:
                if ln.get("direction") != "take":
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
        """Re-query stock and rebuild size buttons when the grid is visible."""
        if self._sel_school and self._sel_item and self._sel_color:
            try:
                self._render_sizes(preserve_size=preserve_size)
            except Exception:
                pass

    def _render_sizes(self, preserve_size: Optional[str] = None):
        self._sel_size = None
        self._crumb_var.set(
            f"المدرسة: {self._sel_school}  \u27f6  النوع: {self._sel_item}  \u27f6  اللون: {self._sel_color}  \u27f6  اختر المقاس"
        )
        self._clear_grid()

        raw_sizes = self._get_sizes_for_bill(
            self._sel_school or "", self._sel_item or "", self._sel_color or "")

        stock_rows: Dict[str, Any] = {}
        try:
            for r in self.db.current_inventory({
                "school": self._sel_school,
                "item_type": self._sel_item,
                "color": self._sel_color,
            }):
                sz = _normalize_size_label(r.get("size") or "")
                if not sz:
                    continue
                stock_rows.setdefault(sz, {"count": 0, "last_price": r.get("unit_price")})
                stock_rows[sz]["count"] += int(r.get("count") or 0)
                if r.get("unit_price") not in (None, ""):
                    stock_rows[sz]["last_price"] = r.get("unit_price")
        except Exception:
            pass

        self._sizes_cache = []
        seen_sizes: Set[str] = set()
        for sz in raw_sizes:
            sz = _normalize_size_label(sz)
            if not sz:
                continue
            size_key = sz.casefold()
            if size_key in seen_sizes:
                continue
            seen_sizes.add(size_key)
            r = stock_rows.get(sz, {})
            base = int(r.get("count", 0) or 0)
            pending = self._pending_out_qty_for_specs(
                self._sel_school or "", self._sel_item or "", self._sel_color or "", str(sz))
            eff = max(0, base - pending)
            self._sizes_cache.append({
                "size": sz,
                "count": eff,
                "last_price": r.get("last_price"),
            })

        cols = 4
        row = None
        for i, r in enumerate(self._sizes_cache):
            if i % cols == 0:
                row = ttk.Frame(self._grid_host)
                row.pack(fill=tk.X, padx=4, pady=2)

            sz = str(r.get("size") or "")
            cnt = int(r.get("count") or 0)
            label = f"{sz} ({cnt})"
            style = "SizeZero.TButton" if cnt == 0 else ("SizeLow.TButton" if cnt <= 5 else "Size.TButton")

            btn = ttk.Button(row, text=label, style=style)
            # Double-click adds to active bill directly
            btn.configure(command=lambda v=label, b=btn: self._on_size_click(v, b))
            btn.bind("<Double-Button-1>", lambda e, v=label, b=btn: self._on_size_double_click(v, b))
            btn.pack(side=tk.LEFT, padx=3, pady=3)

            self._size_btns[label] = (btn, cnt == 0)

        nav_row = ttk.Frame(self._grid_host)
        nav_row.pack(anchor="w", padx=4, pady=4)
        ttk.Button(nav_row, text="\u25c4 رجوع إلى الألوان", command=self._back_to_colors).pack(side=tk.LEFT, padx=(0, 6))
        can_add_any = (not self._active_add_requires_stock()) or any(int(r.get("count") or 0) > 0 for r in self._sizes_cache)
        add_all_btn = ttk.Button(nav_row, text="Add all to bill", command=self._add_all_visible_sizes_to_bill)
        add_all_btn.pack(side=tk.LEFT)
        if not can_add_any:
            try:
                add_all_btn.configure(state="disabled")
            except Exception:
                pass

        if preserve_size:
            psz = str(preserve_size).strip()
            for lbl in list(self._size_btns.keys()):
                if lbl.startswith(psz + " ("):
                    btn, _ = self._size_btns[lbl]
                    self._on_size_click(lbl, btn)
                    break

    def _on_size_click(self, label_with_count: str, btn):
        size = label_with_count.split(" (", 1)[0].strip()
        self._sel_size = size
        self._price_user_edited = False

        if self._selected_size_btn is not None:
            old_btn, was_zero = self._selected_size_btn
            if old_btn.winfo_exists():
                old_btn.configure(style="SizeZero.TButton" if was_zero else "Size.TButton")

        btn.configure(style="SizeSelected.TButton")
        was_zero = "(0)" in label_with_count
        self._selected_size_btn = (btn, was_zero)

    def _on_size_double_click(self, label_with_count: str, btn):
        size = label_with_count.split(" (", 1)[0].strip()
        self._sel_size = size
        self._on_size_click(label_with_count, btn)
        self._add_to_active_bill(size, qty=1)

    def _add_to_active_bill(self, size: str, qty: int = 1):
        if not all([self._sel_school, self._sel_item, self._sel_color]):
            messagebox.showwarning("اختر أولاً", "اختر المدرسة والنوع واللون أولاً.")
            return
        try:
            qty = int(qty)
        except Exception:
            qty = 0
        if qty <= 0:
            messagebox.showwarning("كمية غير صالحة", "الكمية يجب أن تكون أكبر من صفر.", parent=self)
            return
        if self._active_add_requires_stock():
            available = self._available_qty_for_active_add(size)
            if available <= 0 or qty > available:
                self._warn_unavailable_add(size, qty, available)
                self._refresh_size_grid_if_current(preserve_size=size)
                return

        # Find price
        price = self._compute_price_for_size(size)
        if price is None:
            price = 0.0

        # Determine direction tag for exchange bill
        direction = None
        if self.active_bill == 6:
            direction = self._exchange_mode  # "return" or "take"

        bill_lines = self.bills[self.active_bill]
        # Merge if same specs, price, and direction
        for existing in bill_lines:
            if (existing.get("item_type") == self._sel_item
                    and existing.get("school") == self._sel_school
                    and existing.get("color") == self._sel_color
                    and existing.get("size") == size
                    and abs(float(existing.get("unit_price", 0)) - price) < 0.001
                    and existing.get("direction") == direction):
                existing["qty"] = int(existing["qty"]) + qty
                self._sync_bill_table()
                self._refresh_size_grid_if_current(preserve_size=size)
                return

        line = {
            "item_type": self._sel_item,
            "school": self._sel_school,
            "color": self._sel_color,
            "size": size,
            "unit_price": price,
            "qty": qty,
        }
        if direction is not None:
            line["direction"] = direction
        bill_lines.append(line)
        self._sync_bill_table()
        self._refresh_size_grid_if_current(preserve_size=size)

    def _add_all_visible_sizes_to_bill(self):
        if not all([self._sel_school, self._sel_item, self._sel_color]):
            messagebox.showwarning("اختر أولاً", "اختر المدرسة والنوع واللون أولاً.", parent=self)
            return

        sizes = [str(r.get("size") or "").strip() for r in (self._sizes_cache or []) if str(r.get("size") or "").strip()]
        if not sizes:
            messagebox.showwarning("لا توجد مقاسات", "لا توجد مقاسات لإضافتها.", parent=self)
            return

        if self._active_add_requires_stock():
            sizes = [sz for sz in sizes if self._available_qty_for_active_add(sz) > 0]
            if not sizes:
                messagebox.showwarning(
                    "لا توجد كميات",
                    "كل المقاسات المعروضة كميتها صفر، لذلك لا يمكن إضافتها إلى الفاتورة.",
                    parent=self,
                )
                self._refresh_size_grid_if_current()
                return

        for sz in sizes:
            self._add_to_active_bill(sz, qty=1)

    def _compute_price_for_size(self, target_size: str) -> Optional[float]:
        import math
        try:
            def _ceil10(x):
                return float(int(math.ceil(x / 10.0) * 10))

            known: Dict[str, float] = {}
            for r in (self._sizes_cache or []):
                s = str(r.get("size") or "")
                lp = r.get("last_price")
                if lp is not None:
                    try:
                        known[s] = float(lp)
                    except Exception:
                        pass

            if target_size in known:
                return float(known[target_size])

            try:
                index_of = {lbl: i for i, lbl in enumerate([r.get("size") for r in (self._sizes_cache or [])])}
            except Exception:
                index_of = {}

            points = [(index_of[s], float(p)) for s, p in known.items() if s in index_of]

            if len(points) >= 2:
                xs = [float(x) for x, _ in points]
                ys = [float(y) for _, y in points]
                mean_x = sum(xs) / len(xs)
                mean_y = sum(ys) / len(ys)
                cov = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys))
                varx = sum((x - mean_x) ** 2 for x in xs)
                slope = 0.0 if varx == 0 else cov / varx
                intercept = mean_y - slope * mean_x
                if target_size in index_of:
                    tx = float(index_of[target_size])
                    pred = slope * tx + intercept
                    if pred is not None and pred == pred and pred >= 0:
                        return _ceil10(pred)

            if points and target_size in index_of:
                tgt_idx = index_of[target_size]
                nearest = min(points, key=lambda ip: abs(ip[0] - tgt_idx))
                return _ceil10(nearest[1])

            if points:
                prices = sorted([p for _, p in points])
                return _ceil10(prices[len(prices) // 2])

            try:
                p = self.db.last_price_for_specs(
                    self._sel_item or "", self._sel_school or "", self._sel_color or "", "")
                if p is not None:
                    return float(round(float(p), 2))
            except Exception:
                pass
        except Exception:
            pass
        return None

    def _get_sizes_for_bill(self, school: str, item: str, color: str) -> List[str]:
        profile = self.db.get_size_profile(item, school, color)
        sizes: List[str] = []
        seen: set = set()

        def _add_size(value: Any) -> None:
            sz = _normalize_size_label(str(value or "").strip())
            key = sz.casefold()
            if sz and key not in seen:
                seen.add(key)
                sizes.append(sz)

        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            for sz in merged_numeric_size_labels_from_profile(r1s, r1e, r2s, r2e):
                _add_size(sz)
            if has_alpha:
                for sz in ALPHA_SIZES:
                    _add_size(sz)

        try:
            rows = self.db.current_inventory({"school": school, "item_type": item, "color": color})
            for r in rows:
                _add_size(r.get("size"))
        except Exception:
            pass

        def _sort_key(label: str):
            text = str(label or "").strip()
            if text.isdigit():
                return (0, int(text))
            try:
                return (1, ALPHA_SIZES.index(text.upper()))
            except ValueError:
                return (2, text.casefold())

        return sorted(sizes, key=_sort_key)

    # ------------------------------------------------------------------ quick filter
    def _build_filter_constraints(self, exclude_field: str) -> Dict[str, Any]:
        c: Dict[str, Any] = {}
        if exclude_field != "school":
            v = self._flt_school.get()
            if v:
                c["school"] = v
        if exclude_field != "item_type":
            v = self._flt_item.get()
            if v:
                c["item_type"] = v
        if exclude_field != "color":
            v = self._flt_color.get()
            if v:
                c["color"] = v
        return c

    def _on_filter_changed(self):
        """Refresh all filter dropdowns (cascading) then apply."""
        self._flt_school.refresh_values()
        self._flt_item.refresh_values()
        self._flt_color.refresh_values()
        self._auto_apply_filter()

    def _auto_apply_filter(self):
        """Called automatically when any filter combobox value changes."""
        sc = self._flt_school.get()
        it = self._flt_item.get()
        cl = self._flt_color.get()

        if sc and it and cl:
            self._sel_school = sc
            self._sel_item = it
            self._sel_color = cl
            self._render_sizes()
        elif sc and it:
            self._sel_school = sc
            self._sel_item = it
            self._render_colors()
        elif sc and cl:
            self._sel_school = sc
            self._sel_color = cl
            self._render_items()
        elif sc:
            self._sel_school = sc
            self._render_items()
        else:
            self._render_schools()

    def _clear_quick_filter(self):
        self._flt_school.set("")
        self._flt_item.set("")
        self._flt_color.set("")
        self._flt_school.refresh_values()
        self._flt_item.refresh_values()
        self._flt_color.refresh_values()
        self._render_schools()

    # ------------------------------------------------------------------ search & favorites
    def _on_search_changed(self):
        """Filter the button grid based on search text."""
        query = self._search_var.get().strip()
        if not query:
            if self._sel_school and self._sel_color and self._sel_item:
                self._render_sizes()
            elif self._sel_school and self._sel_item:
                self._render_colors()
            elif self._sel_school:
                self._render_items()
            else:
                self._render_schools()
            return

        self._clear_grid()
        self._crumb_var.set(f"نتائج البحث: {query}")
        try:
            results = self.db.current_inventory({})
            seen = set()
            matches = []
            for r in results:
                key = (r.get("school", ""), r.get("item_type", ""), r.get("color", ""))
                if key in seen:
                    continue
                text = f"{r.get('school', '')} {r.get('item_type', '')} {r.get('color', '')}"
                if query in text:
                    seen.add(key)
                    matches.append(r)
            cols = 3
            row_frame = None
            for i, r in enumerate(matches[:24]):
                if i % cols == 0:
                    row_frame = ttk.Frame(self._grid_host)
                    row_frame.pack(fill=tk.X, padx=4, pady=2)
                label = f"{r['item_type']} - {r['school']} ({r['color']})"
                btn = ttk.Button(row_frame, text=label,
                                 command=lambda s=r.get("school"), it=r.get("item_type"), c=r.get("color"): self._jump_to(s, it, c))
                btn.pack(side=tk.LEFT, padx=3, pady=3)
        except Exception:
            pass

    def _jump_to(self, school, item_type, color):
        """Navigate directly to sizes view for a specific school/item/color combo."""
        self._sel_school = school
        self._sel_item = item_type
        self._sel_color = color
        self._price_user_edited = False
        self._search_var.set("")
        self._render_sizes()

    def _clear_search(self):
        """Clear the search bar and reset to schools view."""
        self._search_var.set("")
        self._sel_school = None
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._render_schools()

    def sync_refresh(self):
        """Refresh POS selector state after inbound sync changes stock."""
        try:
            self._flt_school.refresh_values()
            self._flt_item.refresh_values()
            self._flt_color.refresh_values()
        except Exception:
            pass

        try:
            self._refresh_favorites()
        except Exception:
            pass

        query = (self._search_var.get() or "").strip()
        if query:
            self._on_search_changed()
            return

        if self._sel_school and self._sel_color and self._sel_item:
            self._render_sizes()
        elif self._sel_school and self._sel_item:
            self._render_colors()
        elif self._sel_school:
            self._render_items()
        else:
            self._render_schools()

    def _refresh_favorites(self):
        """Show top 5 most-sold item combos as quick-access buttons."""
        for w in self._fav_inner.winfo_children():
            w.destroy()
        try:
            rows = self.db.conn.execute("""
                SELECT item_type, school, color, SUM(qty) as total_qty
                FROM bill_items
                GROUP BY item_type, school, color
                ORDER BY total_qty DESC
                LIMIT 5
            """).fetchall()
            if not rows:
                ttk.Label(self._fav_inner, text="لا توجد مبيعات بعد",
                          font=("Segoe UI", 8)).pack(side=tk.LEFT)
                return
            for r in rows:
                label = f"{r['item_type']}-{r['school']} ({r['color']})"
                btn = ttk.Button(self._fav_inner, text=label, style="Secondary.TButton",
                                 command=lambda s=r["school"], it=r["item_type"], c=r["color"]: self._jump_to(s, it, c))
                btn.pack(side=tk.LEFT, padx=2, pady=2)
        except Exception:
            pass

    # ------------------------------------------------------------------ bill management
    def _switch_bill(self, n: int):
        # Save current customer
        self.customers[self.active_bill] = self._customer_var.get()
        self.customer_phones[self.active_bill] = _normalize_customer_phone(self._customer_phone_var.get())
        previous_requires_stock = self._active_add_requires_stock()
        self.active_bill = n
        # Restore customer for this bill
        self._customer_var.set(self.customers[n])
        self._customer_phone_var.set(self.customer_phones[n])
        self._update_bill_btn_styles()
        self._sync_bill_table()
        self._update_res_frame_visibility()
        if previous_requires_stock != self._active_add_requires_stock():
            self._rerender_selector_for_active_mode()
        else:
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _sales_bill_for_next_entry(self) -> int:
        for n in (1, 2, 3):
            if not self.bills.get(n):
                return n
        return 1

    def _reset_to_sales_bill_after_special_finalize(self) -> None:
        target = self._sales_bill_for_next_entry()
        if self.active_bill != target:
            self._switch_bill(target)
        else:
            self._update_bill_btn_styles()
            self._update_res_frame_visibility()
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _update_bill_btn_styles(self):
        for n, btn in self._bill_btns.items():
            if n == self.active_bill:
                try:
                    btn.configure(style="ActiveBill.TButton")
                except Exception:
                    pass
            else:
                try:
                    btn.configure(style="TButton")
                except Exception:
                    pass

    def _save_customer(self):
        self.customers[self.active_bill] = self._customer_var.get()
        self.customer_phones[self.active_bill] = _normalize_customer_phone(self._customer_phone_var.get())

    def _clear_customer_for_bill(self, bill_no: Optional[int] = None) -> None:
        n = int(bill_no or self.active_bill)
        self.customers[n] = ""
        self.customer_phones[n] = ""
        if self.active_bill == n:
            self._customer_var.set("")
            self._customer_phone_var.set("")

    def _choose_branch_target(self):
        current_device = (self.db.conn.execute(
            "SELECT device_name FROM device_identity WHERE id = 1"
        ).fetchone() or [None])[0]
        choices = [n for n in DEFAULT_BRANCH_POS_NAMES if n != current_device]
        ui_choices = [_branch_display_name(n) for n in choices]
        ui_to_dev = {_branch_display_name(n): n for n in choices}
        dlg = tk.Toplevel(self.winfo_toplevel())
        dlg.title("طلب تحويل إلى فرع")
        dlg.geometry("320x140")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="اختر الفرع المطلوب لإرسال الطلب إلى المخزن:").pack(anchor="w", pady=(0, 8))
        var = tk.StringVar(value=ui_choices[0] if ui_choices else "")
        cb = ttk.Combobox(frm, textvariable=var, values=ui_choices, state="readonly")
        cb.pack(fill=tk.X)

        def _apply():
            picked = (var.get() or "").strip()
            dev = ui_to_dev.get(picked, picked)
            if dev:
                self._customer_var.set(f"{BRANCH_TARGET_PREFIX}{_branch_display_name(dev)}")
                self._save_customer()
            dlg.destroy()

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(btns, text="اختيار", command=_apply).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.LEFT)

    def _autocomplete_customer(self):
        """Show auto-complete suggestions for customer name."""
        query = self._customer_var.get().strip()
        # Destroy old listbox if exists
        if self._cust_listbox:
            try:
                self._cust_listbox.destroy()
            except Exception:
                pass
            self._cust_listbox = None

        if len(query) < 2:
            return
        try:
            rows = self.db.conn.execute(
                "SELECT DISTINCT customer FROM bills WHERE customer LIKE ? AND customer != '' LIMIT 8",
                (f"%{query}%",)
            ).fetchall()
            names = [r["customer"] for r in rows if r["customer"]]
            special = [WAREHOUSE_RETURN_LABEL] + [
                f"{BRANCH_TARGET_PREFIX}{_branch_display_name(n)}" for n in DEFAULT_BRANCH_POS_NAMES
            ]
            for label in reversed(special):
                if query.lower() in label.lower() and label not in names:
                    names.insert(0, label)
            if not names:
                return

            lb = tk.Listbox(self.winfo_toplevel(), height=min(len(names), 6),
                            font=("Segoe UI", 9), bg="#ffffff", fg="#1e293b",
                            selectbackground="#bfdbfe", bd=1, relief="solid")
            for n in names:
                lb.insert(tk.END, n)

            # Position below the entry widget
            x = self._cust_entry.winfo_rootx()
            y = self._cust_entry.winfo_rooty() + self._cust_entry.winfo_height()
            lb.place(x=x - self.winfo_toplevel().winfo_rootx(),
                     y=y - self.winfo_toplevel().winfo_rooty(),
                     width=self._cust_entry.winfo_width())
            lb.lift()

            def _select(event=None):
                sel = lb.curselection()
                if sel:
                    picked = lb.get(sel[0])
                    self._customer_var.set(picked)
                    try:
                        row = self.db.conn.execute(
                            """
                            SELECT customer_phone FROM (
                                SELECT customer_phone, created_at FROM bills
                                WHERE LOWER(TRIM(customer)) = LOWER(TRIM(?))
                                  AND COALESCE(TRIM(customer_phone), '') <> ''
                                UNION ALL
                                SELECT customer_phone, created_at FROM reservations
                                WHERE LOWER(TRIM(customer)) = LOWER(TRIM(?))
                                  AND COALESCE(TRIM(customer_phone), '') <> ''
                            )
                            ORDER BY created_at DESC
                            LIMIT 1
                            """,
                            (picked, picked),
                        ).fetchone()
                        if row:
                            self._customer_phone_var.set(row["customer_phone"] or "")
                    except Exception:
                        pass
                    self._save_customer()
                lb.destroy()
                self._cust_listbox = None

            lb.bind("<ButtonRelease-1>", _select)
            lb.bind("<Return>", _select)
            self._cust_listbox = lb

            # Auto-dismiss when clicking elsewhere
            def _dismiss(event=None):
                if self._cust_listbox:
                    try:
                        self._cust_listbox.destroy()
                    except Exception:
                        pass
                    self._cust_listbox = None
            self._cust_entry.bind("<Escape>", _dismiss, add="+")
        except Exception:
            pass

    def _update_res_frame_visibility(self):
        self._res_frame.pack_forget()
        self._return_frame.pack_forget()
        self._exchange_frame.pack_forget()
        if self.active_bill in self.RESERVATION_BILLS:
            self._res_frame.pack(fill=tk.X, padx=4, pady=(2, 2))
        elif self.active_bill == 4:
            self._return_frame.pack(fill=tk.X, padx=4, pady=(2, 2))
        elif self.active_bill == 6:
            self._exchange_frame.pack(fill=tk.X, padx=4, pady=(2, 2))

    def _on_exchange_mode_changed(self):
        previous_requires_stock = self._active_add_requires_stock()
        self._exchange_mode = self._exchange_mode_var.get()
        try:
            self._flt_school.refresh_values()
            self._flt_item.refresh_values()
            self._flt_color.refresh_values()
        except Exception:
            pass
        if previous_requires_stock != self._active_add_requires_stock():
            self._rerender_selector_for_active_mode()
        else:
            self.sync_refresh()

    def _sync_bill_table(self):
        self.bill_table.delete(*self.bill_table.get_children())
        total = 0.0
        return_total = 0.0
        take_total = 0.0
        total_qty = 0
        is_return = self.active_bill == 4
        is_exchange = self.active_bill == 6

        for idx, ln in enumerate(self.bills[self.active_bill]):
            qty = int(ln["qty"])
            line_total = float(ln["unit_price"]) * qty
            total_qty += qty
            display_type = ln["item_type"]

            if is_return:
                display_type = f"\u21a9 {display_type}"
                total += line_total
            elif is_exchange:
                direction = ln.get("direction", "return")
                if direction == "return":
                    display_type = f"\u21a9 {display_type}"
                    return_total += line_total
                else:
                    display_type = f"\u21aa {display_type}"
                    take_total += line_total
            else:
                total += line_total

            self.bill_table.insert(
                "", tk.END, iid=str(idx),
                values=(display_type, ln["school"], ln["color"], ln["size"],
                        f"{format_money(float(ln['unit_price']))}", qty, f"{format_money(line_total)}")
            )

        self.total_items_var.set(str(total_qty))
        if is_exchange:
            diff = take_total - return_total
            if diff > 0:
                diff_text = f"فرق السعر: {format_money(diff)} (يدفع العميل)"
            elif diff < 0:
                diff_text = f"فرق السعر: {format_money(abs(diff))} (يسترد العميل)"
            else:
                diff_text = "فرق السعر: 0 (متساوي)"
            self._exchange_diff_var.set(diff_text)
            self.total_var.set(f"{format_money(diff)}")
        else:
            self.total_var.set(f"{format_money(total)}")

        _apply_zebra_tags(self.bill_table)

    def _inc_qty(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        lines = self.bills[self.active_bill]
        if 0 <= idx < len(lines):
            line = lines[idx]
            if self._line_requires_stock(self.active_bill, line):
                available = self._available_qty_for_specs(
                    line.get("school") or "",
                    line.get("item_type") or "",
                    line.get("color") or "",
                    line.get("size") or "",
                )
                if available <= 0:
                    self._warn_unavailable_add(str(line.get("size") or ""), 1, available)
                    self._refresh_size_grid_if_current(preserve_size=self._sel_size)
                    return
            lines[idx]["qty"] = int(lines[idx]["qty"]) + 1
            self._sync_bill_table()
            self.bill_table.selection_set(str(idx))
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _dec_qty(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        lines = self.bills[self.active_bill]
        if 0 <= idx < len(lines):
            if int(lines[idx]["qty"]) > 1:
                lines[idx]["qty"] = int(lines[idx]["qty"]) - 1
            else:
                lines.pop(idx)
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _remove_line(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        lines = self.bills[self.active_bill]
        if 0 <= idx < len(lines):
            lines.pop(idx)
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    def _clear_bill(self):
        self.bills[self.active_bill].clear()
        self._sync_bill_table()
        self._refresh_size_grid_if_current(preserve_size=self._sel_size)

    # ------------------------------------------------------------------ finalize
    def _finalize(self):
        lines = self.bills[self.active_bill]
        if not lines:
            messagebox.showwarning("فارغ", "لا توجد أصناف في الفاتورة.")
            return

        self._save_customer()
        customer = (self.customers[self.active_bill] or "").strip()
        customer_phone = _normalize_customer_phone(self.customer_phones.get(self.active_bill, ""))
        total_qty = sum(int(ln["qty"]) for ln in lines)
        is_reservation = self.active_bill in self.RESERVATION_BILLS
        is_return = self.active_bill == 4
        is_exchange = self.active_bill == 6
        warehouse_target = _extract_warehouse_target(customer)

        # Determine title and total display
        if is_return:
            overlay_title = "تأكيد المرتجع"
            review_title = "مراجعة المرتجع"
            total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in lines)
            total_label = f"مبلغ الاسترداد: {format_money(total)}"
        elif is_exchange:
            overlay_title = "تأكيد الاستبدال"
            review_title = "مراجعة الاستبدال"
            ret_lines = [ln for ln in lines if ln.get("direction") == "return"]
            take_lines = [ln for ln in lines if ln.get("direction") == "take"]
            if not ret_lines and not take_lines:
                messagebox.showwarning("فارغ", "أضف أصناف مرتجعة ومأخوذة.")
                return
            ret_total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in ret_lines)
            take_total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in take_lines)
            total = take_total - ret_total
            if total > 0:
                total_label = f"يدفع العميل: {format_money(total)}"
            elif total < 0:
                total_label = f"يسترد العميل: {format_money(abs(total))}"
            else:
                total_label = "متساوي - لا فرق في السعر"
        elif is_reservation:
            overlay_title = "تأكيد الحجز"
            review_title = "مراجعة الحجز"
            total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in lines)
            total_label = f"الإجمالي: {format_money(total)}"
        elif warehouse_target:
            overlay_title = f"تأكيد {WAREHOUSE_RETURN_LABEL}"
            review_title = f"مراجعة {WAREHOUSE_RETURN_LABEL}"
            total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in lines)
            total_label = f"قيمة المرجع: {format_money(total)}"
        else:
            branch_target = _extract_branch_target(customer)
            if branch_target:
                overlay_title = "تأكيد تحويل إلى فرع"
                review_title = f"مراجعة تحويل إلى {branch_target}"
                total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in lines)
                total_label = f"قيمة التحويل: {format_money(total)}"
            else:
                overlay_title = "تأكيد الفاتورة"
                review_title = "مراجعة الفاتورة"
                total = sum(float(ln["unit_price"]) * int(ln["qty"]) for ln in lines)
                total_label = f"الإجمالي: {format_money(total)}"

        # Show confirmation overlay
        overlay = tk.Toplevel(self)
        overlay.title(overlay_title)
        overlay.geometry("640x520" if is_exchange else "600x480")
        overlay.minsize(520, 380)
        overlay.resizable(True, True)
        overlay.transient(self.winfo_toplevel())
        overlay.grab_set()

        frm = ttk.Frame(overlay, padding=16)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=review_title, font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 8))

        if customer:
            ttk.Label(frm, text=f"العميل: {customer}").pack(anchor="w")
        if customer_phone:
            ttk.Label(frm, text=f"رقم العميل: {customer_phone}").pack(anchor="w")

        ttk.Label(frm, text=f"عدد الأصناف: {len(lines)}   |   عدد القطع: {total_qty}",
                  font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 4))

        # Items summary table
        cols_frm = ttk.Frame(frm)
        cols_frm.pack(fill=tk.BOTH, expand=True, pady=(4, 8))

        if is_exchange:
            summary_tree = ttk.Treeview(cols_frm, columns=("dir", "spec", "qty", "total"),
                                         show="headings", height=min(len(lines), 6))
            summary_tree.heading("dir", text="نوع")
            summary_tree.heading("spec", text="الصنف")
            summary_tree.heading("qty", text="الكمية")
            summary_tree.heading("total", text="المبلغ")
            summary_tree.column("dir", width=50, anchor="center")
            summary_tree.column("spec", width=200)
            summary_tree.column("qty", width=60, anchor="center")
            summary_tree.column("total", width=80, anchor="center")
            for ln in lines:
                spec = f"{ln['item_type']} - {ln['school']} ({ln['color']}/{ln['size']})"
                lt = float(ln['unit_price']) * int(ln['qty'])
                d = "\u21a9" if ln.get("direction") == "return" else "\u21aa"
                summary_tree.insert("", tk.END, values=(d, spec, ln['qty'], f"{format_money(lt)}"))
        else:
            summary_tree = ttk.Treeview(cols_frm, columns=("spec", "qty", "total"),
                                         show="headings", height=min(len(lines), 6))
            summary_tree.heading("spec", text="الصنف")
            summary_tree.heading("qty", text="الكمية")
            summary_tree.heading("total", text="المبلغ")
            summary_tree.column("spec", width=220)
            summary_tree.column("qty", width=60, anchor="center")
            summary_tree.column("total", width=80, anchor="center")
            for ln in lines:
                prefix = "\u21a9 " if is_return else ""
                spec = f"{prefix}{ln['item_type']} - {ln['school']} ({ln['color']}/{ln['size']})"
                lt = float(ln['unit_price']) * int(ln['qty'])
                summary_tree.insert("", tk.END, values=(spec, ln['qty'], f"{format_money(lt)}"))
        ysb = ttk.Scrollbar(cols_frm, orient="vertical", command=summary_tree.yview)
        summary_tree.configure(yscrollcommand=ysb.set)
        summary_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(summary_tree)

        ttk.Label(frm, text=total_label,
                  font=("Segoe UI", 14, "bold")).pack(anchor="center", pady=(8, 12))

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=tk.X)

        def _do_finalize():
            overlay.destroy()
            self._execute_finalize(lines, customer, is_reservation, total, customer_phone)

        ttk.Button(btn_row, text="تأكيد", command=_do_finalize,
                   style="Success.TButton").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="إلغاء", command=overlay.destroy,
                   style="Secondary.TButton").pack(side=tk.LEFT)

    def _execute_finalize(self, lines, customer, is_reservation, total, customer_phone=""):
        """Execute the actual finalize after confirmation."""
        is_return = self.active_bill == 4
        is_exchange = self.active_bill == 6
        customer_phone = _normalize_customer_phone(customer_phone)

        if is_return:
            payment_method = self._choose_payment_method(total, "\u0637\u0631\u064a\u0642\u0629 \u0631\u062f \u0627\u0644\u0645\u0628\u0644\u063a")
            if not payment_method:
                return
            try:
                bill_id = self.db.create_return_bill(customer, lines, customer_phone=customer_phone, payment_method=payment_method)
            except Exception as ex:
                messagebox.showerror("فشل المرتجع", str(ex), parent=self)
                return
            self.bills[4].clear()
            self._clear_customer_for_bill(4)
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)
            ToastNotification.show(self.winfo_toplevel(),
                                   f"تم إنشاء فاتورة مرتجع #{bill_id} (استرداد: {format_money(total)})", toast_type="success")
            if self._ask_print_receipt():
                self._print_return_bill(bill_id, total)
            self._reset_to_sales_bill_after_special_finalize()

        elif is_exchange:
            # Copy lines before clearing since lines is a reference to self.bills[6]
            lines_copy = [dict(ln) for ln in lines]
            ret_lines = [ln for ln in lines_copy if ln.get("direction") == "return"]
            take_lines = [ln for ln in lines_copy if ln.get("direction") == "take"]
            payment_method = self._choose_payment_method(total, "\u0637\u0631\u064a\u0642\u0629 \u0641\u0631\u0642 \u0627\u0644\u0627\u0633\u062a\u0628\u062f\u0627\u0644")
            if not payment_method:
                return
            try:
                bill_id = self.db.create_exchange_bill(customer, ret_lines, take_lines, customer_phone=customer_phone, payment_method=payment_method)
            except Exception as ex:
                messagebox.showerror("فشل الاستبدال", str(ex), parent=self)
                return
            self.bills[6].clear()
            self._clear_customer_for_bill(6)
            self._exchange_mode = "return"
            self._exchange_mode_var.set("return")
            self._exchange_diff_var.set("فرق السعر: 0")
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)
            ToastNotification.show(self.winfo_toplevel(),
                                   f"تم إنشاء فاتورة استبدال #{bill_id}", toast_type="success")
            if self._ask_print_receipt():
                self._print_exchange_bill(bill_id, lines_copy, total)
            self._reset_to_sales_bill_after_special_finalize()

        elif is_reservation:
            try:
                paid = float((self._paid_var.get() or "0").strip())
            except Exception:
                paid = 0.0
            note = self._res_note_var.get().strip()
            for ln in lines:
                ln["note"] = note
            payment_method = PAYMENT_METHOD_CASH
            if paid > 1e-9:
                payment_method = self._choose_payment_method(paid, "\u0637\u0631\u064a\u0642\u0629 \u062f\u0641\u0639 \u0627\u0644\u062d\u062c\u0632")
                if not payment_method:
                    return
            try:
                ids = self.db.create_reservation(customer, lines, paid_amount=paid, customer_phone=customer_phone, payment_method=payment_method)
                if self._ask_print_receipt():
                    self._print_reservation_receipt(ids, lines, customer, paid, total, customer_phone)
                bill_no = min(int(x) for x in ids) if ids else ""
                ToastNotification.show(self.winfo_toplevel(),
                                       f"تم إنشاء فاتورة حجز #{bill_no} بنجاح", toast_type="success")
                self.bills[self.active_bill].clear()
                self._clear_customer_for_bill(self.active_bill)
                self._paid_var.set("0")
                self._res_note_var.set("")
                self._sync_bill_table()
                self._refresh_size_grid_if_current(preserve_size=self._sel_size)
                self._reset_to_sales_bill_after_special_finalize()
            except Exception as ex:
                messagebox.showerror("فشل الحجز", str(ex), parent=self)
        else:
            payment_method = self._choose_sale_payment_method(total)
            if not payment_method:
                return
            try:
                bill_id = self.db.create_bill(customer, lines, payment_method=payment_method, customer_phone=customer_phone)
            except Exception as ex:
                messagebox.showerror("فشل الحفظ", str(ex), parent=self)
                return

            self.bills[self.active_bill].clear()
            self._clear_customer_for_bill(self.active_bill)
            self._sync_bill_table()
            self._refresh_size_grid_if_current(preserve_size=self._sel_size)
            self._refresh_favorites()
            warehouse_target = _extract_warehouse_target(customer)
            branch_target = _extract_branch_target(customer)
            if warehouse_target:
                ToastNotification.show(
                    self.winfo_toplevel(),
                    f"تم إنشاء {WAREHOUSE_RETURN_LABEL} #{bill_id} وسيتم إرساله للمراجعة في المخزن",
                    toast_type="success",
                )
                try:
                    import sync_ui
                    sync_ui.run_sync_now(self.winfo_toplevel(), self.db.conn, reason=WAREHOUSE_RETURN_LABEL)
                except Exception:
                    pass
            elif branch_target:
                ToastNotification.show(
                    self.winfo_toplevel(),
                    f"تم إنشاء طلب تحويل #{bill_id} إلى {branch_target} عبر المخزن",
                    toast_type="success",
                )
                try:
                    import sync_ui
                    sync_ui.run_sync_now(self.winfo_toplevel(), self.db.conn, reason=f"تحويل إلى {branch_target}")
                except Exception:
                    pass
            else:
                ToastNotification.show(
                    self.winfo_toplevel(),
                    f"تم إنشاء الفاتورة #{bill_id} (الإجمالي: {format_money(total)})",
                    toast_type="success",
                )

            if self._ask_print_receipt():
                self._direct_print_bill(bill_id)

    def _choose_payment_method(self, total: float, title: str = "\u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u062f\u0641\u0639") -> Optional[str]:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="\u0627\u062e\u062a\u0631 \u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u062f\u0641\u0639:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, text=f"\u0627\u0644\u0645\u0628\u0644\u063a: {format_money(abs(float(total or 0.0)))}").pack(anchor="w", pady=(4, 10))

        method_var = tk.StringVar(value=PAYMENT_METHOD_CASH)
        ttk.Radiobutton(frm, text="\u0643\u0627\u0634", variable=method_var, value=PAYMENT_METHOD_CASH).pack(anchor="w", pady=2)
        ttk.Radiobutton(frm, text="\u0641\u064a\u0632\u0627", variable=method_var, value=PAYMENT_METHOD_VISA).pack(anchor="w", pady=2)

        result = {"value": None}

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(12, 0))

        def on_ok():
            result["value"] = method_var.get().strip().upper() or PAYMENT_METHOD_CASH
            dlg.destroy()

        ttk.Button(btns, text="\u0625\u0644\u063a\u0627\u0621", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="\u062a\u0623\u0643\u064a\u062f", command=on_ok).pack(side=tk.RIGHT, padx=6)

        dlg.wait_window()
        return result["value"]

    def _choose_sale_payment_method(self, total: float) -> Optional[str]:
        return self._choose_payment_method(total)

    def _ask_print_receipt(self) -> bool:
        return messagebox.askyesno(
            "طباعة الفاتورة",
            "تم حفظ الفاتورة بنجاح.\nهل تريد طباعة الفاتورة الآن؟",
            parent=self,
        )

    def _direct_print_bill(self, bill_id: int, copies: int = 2):
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            messagebox.showerror("فشل الطباعة", "لم يتم العثور على الفاتورة.")
            return
        self._print_bill_receipt(bill, items, copies=max(1, int(copies)))

    def _print_bill_receipt(self, bill: Dict[str, Any], items: List[Dict[str, Any]], copies: int = 1):
        tmp_dir = tempfile.gettempdir()
        bill_id = int(bill["id"])
        html_path = os.path.join(tmp_dir, "bill_%s.html" % bill_id)
        pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
        save_bill_as_html(
            html_path,
            bill,
            items,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            shift_id=getattr(self.db, "active_shift_id", "") or "",
        )
        _print_receipt_direct_or_fallback(
            html_path,
            bill,
            items,
            copies=max(1, int(copies)),
            parent=self,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
        )

    def _print_reservation_receipt(self, ids, lines, customer, paid, total, customer_phone="", bill_id: Optional[Any] = None):
        """Auto-print a receipt for a newly created reservation."""
        clean_ids = []
        for rid in ids or []:
            try:
                clean_ids.append(int(rid))
            except Exception:
                pass
        res_num = str(bill_id or (min(clean_ids) if clean_ids else ""))
        customer_phone = _normalize_customer_phone(customer_phone)
        remaining = max(0.0, float(total or 0.0) - float(paid or 0.0))
        path = os.path.join(tempfile.gettempdir(), f"reservation_{res_num}.html")
        pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
        receipt_items = []
        for ln in lines:
            qty = int(ln.get("qty") or 0)
            price = float(ln.get("unit_price") or 0)
            receipt_items.append({
                "item_type": ln.get("item_type") or "",
                "school": ln.get("school") or "",
                "color": ln.get("color") or "",
                "size": ln.get("size") or "",
                "qty": qty,
                "unit_price": price,
                "line_total": price * qty,
            })
        receipt_bill = {
            "id": res_num,
            "created_at": now_iso(),
            "total": float(total),
            "bill_type": "RESERVATION",
            "customer": customer,
            "customer_phone": customer_phone,
        }
        extra_rows = [
            ("المدفوع", paid),
            ("المتبقي للدفع لاحقاً", remaining),
        ]
        save_bill_as_html(
            path,
            receipt_bill,
            receipt_items,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            extra_summary_rows=extra_rows,
        )
        _print_receipt_direct_or_fallback(
            path,
            receipt_bill,
            receipt_items,
            copies=2,
            parent=self,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            extra_summary_rows=extra_rows,
        )

    def _print_pending_reservation_receipt(self, group_items: Sequence[Dict[str, Any]]):
        """Print a fresh reservation receipt for the still-pending items only."""
        pending = [dict(r) for r in group_items if _is_reservation_active(r.get("status"))]
        if not pending:
            return
        bill_id = _reservation_bill_id_from_rows([dict(r) for r in group_items])
        ids = [int(r.get("id") or 0) for r in pending]
        customer = str(pending[0].get("customer") or "")
        customer_phone = _normalize_customer_phone(pending[0].get("customer_phone") or "")
        total = sum(float(r.get("total_amount") or 0.0) for r in pending)
        paid = sum(float(r.get("paid_amount") or 0.0) for r in pending)
        lines = [
            {
                "item_type": r.get("item_type") or "",
                "school": r.get("school") or "",
                "color": r.get("color") or "",
                "size": r.get("size") or "",
                "qty": int(r.get("qty") or 0),
                "unit_price": float(r.get("unit_price") or 0.0),
            }
            for r in pending
        ]
        self._print_reservation_receipt(ids, lines, customer, paid, total, customer_phone, bill_id=bill_id)

    def _print_return_bill(self, bill_id: int, total: float):
        items = self.db.list_bill_items(bill_id)
        path = os.path.join(tempfile.gettempdir(), f"return_{bill_id}.html")
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
        except StopIteration:
            bill = {"id": bill_id, "created_at": now_iso(), "total": total, "bill_type": "RETURN"}
        pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
        save_bill_as_html(
            path,
            bill,
            items,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
        )
        _print_receipt_direct_or_fallback(
            path,
            bill,
            items,
            copies=2,
            parent=self,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
        )

    def _print_exchange_bill(self, bill_id: int, lines: list, diff: float):
        if diff > 0:
            diff_label = "يدفع العميل"
            diff_value = diff
        elif diff < 0:
            diff_label = "يسترد العميل"
            diff_value = abs(diff)
        else:
            diff_label = "فرق السعر"
            diff_value = 0
        path = os.path.join(tempfile.gettempdir(), f"exchange_{bill_id}.html")
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
            items = self.db.list_bill_items(bill_id)
        except StopIteration:
            bill = {"id": bill_id, "created_at": now_iso(), "total": diff, "bill_type": "EXCHANGE"}
            items = []
        receipt_items = []
        source_items = items or lines
        for ln in source_items:
            origin = str(ln.get("origin") or "").upper()
            direction = str(ln.get("direction") or "").lower()
            if origin == "RETURN" or direction == "return":
                prefix = "مرتجع: "
            else:
                prefix = "بديل: "
            qty = int(ln.get("qty") or 0)
            price = float(ln.get("unit_price") or 0)
            line_total = ln.get("line_total")
            if line_total is None:
                line_total = price * qty
            receipt_items.append({
                "item_type": prefix + str(ln.get("item_type") or ""),
                "school": ln.get("school") or "",
                "color": ln.get("color") or "",
                "size": ln.get("size") or "",
                "qty": qty,
                "unit_price": price,
                "line_total": float(line_total or 0),
            })
        pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
        save_bill_as_html(
            path,
            bill,
            receipt_items,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            extra_summary_rows=[(diff_label, diff_value)],
        )
        _print_receipt_direct_or_fallback(
            path,
            bill,
            receipt_items,
            copies=2,
            parent=self,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            extra_summary_rows=[(diff_label, diff_value)],
        )


# ------------------- Factory Item Dialog -------------------

class FactoryItemDialog(tk.Toplevel):
    """
    Create an ad-hoc bill line that ships directly from the factory.
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

        self.t = LabeledCombobox(frm, "النوع",   db, "item_type"); self.t.set(preset.get("item_type",""))
        self.s = LabeledCombobox(frm, "المدرسة", db, "school");    self.s.set(preset.get("school",""))
        self.c = LabeledCombobox(frm, "اللون",   db, "color");     self.c.set(preset.get("color",""))
        self.z = LabeledCombobox(frm, "المقاس",  db, "size");      self.z.set(preset.get("size",""))
        for i, w in enumerate((self.t, self.s, self.c, self.z)):
            w.grid(row=i//2, column=i%2, padx=6, pady=6, sticky="ew")
        for w in (self.t.cb, self.s.cb, self.c.cb, self.z.cb):
            w.bind("<<ComboboxSelected>>", lambda e: self._auto_fill_price(), add="+")
            w.bind("<FocusOut>",           lambda e: self._auto_fill_price(), add="+")
            w.bind("<KeyRelease>",         lambda e: self._auto_fill_price(), add="+")
        frm.columnconfigure(0, weight=1); frm.columnconfigure(1, weight=1)

        self.qv = tk.StringVar(value=preset.get("qty","1"))
        self.pv = tk.StringVar(value=preset.get("unit_price",""))
        grid2 = ttk.Frame(frm); grid2.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        ttk.Label(grid2, text="الكمية:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(grid2, textvariable=self.qv, width=12).grid(row=0, column=1, sticky="w")
        ttk.Label(grid2, text="سعر الوحدة:").grid(row=0, column=2, sticky="e", padx=12, pady=4)
        ttk.Entry(grid2, textvariable=self.pv, width=12).grid(row=0, column=3, sticky="w")

        btns = ttk.Frame(frm); btns.grid(row=3, column=0, columnspan=2, sticky="e", padx=6, pady=(10,0))
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="إضافة", command=self._ok).pack(side=tk.RIGHT, padx=8)

        self.after(10, lambda: (self.t.cb.focus_set(), self._auto_fill_price()))

        self._result: Optional[Dict[str, str]] = None

    def _auto_fill_price(self):
        if (self.pv.get() or "").strip():
            return
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
            pass

    def _ok(self):
        try:
            item_type = self.t.get() or self._err("النوع مطلوب")
            school    = self.s.get() or self._err("المدرسة مطلوبة")
            color     = self.c.get() or self._err("اللون مطلوب")
            size      = self.z.get() or self._err("المقاس مطلوب")
            qty       = int((self.qv.get() or "0").strip());   assert qty > 0
            price     = float((self.pv.get() or "").strip());  assert price >= 0.0
        except AssertionError:
            messagebox.showerror("بيانات غير صالحة", "تحقق من الكمية (>0) والسعر (>=0).", parent=self); return
        except Exception as ex:
            messagebox.showerror("بيانات ناقصة", str(ex), parent=self); return

        self._result = {
            "item_type": item_type, "school": school, "color": color, "size": size,
            "unit_price": price, "qty": qty,
            "allow_factory_fill": True,
        }
        self.destroy()

    @staticmethod
    def _err(msg: str) -> None:
        raise RuntimeError(msg)

    def run(self) -> Optional[Dict[str, str]]:
        self.wait_window()
        return self._result

# ------------------- Statistics Frame -------------------

class StatisticsFrame(ttk.Frame):
    """Statistics and reporting tab."""

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=6)
        self.db = db
        self._build()

    def _build(self):
        ttk.Label(self, text="الإحصائيات والتقارير", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 6))

        # Filter bar with dropdown filters + date range
        filter_bar = ttk.Frame(self)
        filter_bar.pack(fill=tk.X, pady=(0, 6))

        self._sf_school = LabeledCombobox(filter_bar, "المدرسة", self.db, "school")
        self._sf_school.set_supplier(lambda: self.db.get_distinct_filtered("school", self._build_stats_constraints("school")))
        self._sf_school.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._sf_item = LabeledCombobox(filter_bar, "النوع", self.db, "item_type")
        self._sf_item.set_supplier(lambda: self.db.get_distinct_filtered("item_type", self._build_stats_constraints("item_type")))
        self._sf_item.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._sf_color = LabeledCombobox(filter_bar, "اللون", self.db, "color")
        self._sf_color.set_supplier(lambda: self.db.get_distinct_filtered("color", self._build_stats_constraints("color")))
        self._sf_color.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        df_frame = ttk.Frame(filter_bar)
        df_frame.grid(row=0, column=3, padx=4, pady=4, sticky="ew")
        ttk.Label(df_frame, text="من:").pack(side=tk.LEFT)
        self._df = DateField(df_frame, "")
        self._df.pack(side=tk.LEFT, padx=(4, 0))

        dt_frame = ttk.Frame(filter_bar)
        dt_frame.grid(row=0, column=4, padx=4, pady=4, sticky="ew")
        ttk.Label(dt_frame, text="إلى:").pack(side=tk.LEFT)
        self._dt = DateField(dt_frame, "")
        self._dt.pack(side=tk.LEFT, padx=(4, 0))

        ttk.Button(filter_bar, text="تحديث", command=self._refresh_all).grid(row=0, column=5, padx=4, pady=4)
        ttk.Button(filter_bar, text="مسح", command=self._clear_stats_filters).grid(row=0, column=6, padx=4, pady=4)

        for c in range(7):
            filter_bar.columnconfigure(c, weight=1 if c < 5 else 0)

        self._sf_school.cb.bind("<<ComboboxSelected>>", lambda e: self._on_stats_filter_changed(), add="+")
        self._sf_item.cb.bind("<<ComboboxSelected>>", lambda e: self._on_stats_filter_changed(), add="+")
        self._sf_color.cb.bind("<<ComboboxSelected>>", lambda e: self._on_stats_filter_changed(), add="+")

        # ---- Section 1: Money flow ----
        mf = ttk.LabelFrame(self, text="التدفق المالي")
        mf.pack(fill=tk.X, padx=4, pady=(0, 8))

        self._sales_count_var = tk.StringVar(value="0")
        self._sales_total_var = tk.StringVar(value="0")
        self._res_count_var   = tk.StringVar(value="0")
        self._res_total_var   = tk.StringVar(value="0")
        self._res_paid_var    = tk.StringVar(value="0")

        r1 = ttk.Frame(mf); r1.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(r1, text="عدد الفواتير:").pack(side=tk.LEFT)
        ttk.Label(r1, textvariable=self._sales_count_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4,20))
        ttk.Label(r1, text="إجمالي المبيعات:").pack(side=tk.LEFT)
        ttk.Label(r1, textvariable=self._sales_total_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4,0))

        r2 = ttk.Frame(mf); r2.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(r2, text="عدد الحجوزات:").pack(side=tk.LEFT)
        ttk.Label(r2, textvariable=self._res_count_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4,20))
        ttk.Label(r2, text="إجمالي الحجوزات:").pack(side=tk.LEFT)
        ttk.Label(r2, textvariable=self._res_total_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4,20))
        ttk.Label(r2, text="المدفوع:").pack(side=tk.LEFT)
        ttk.Label(r2, textvariable=self._res_paid_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4,0))

        ttk.Button(mf, text="طباعة", command=self._print_money_flow).pack(anchor="e", padx=8, pady=4)

        # ---- Section 2: Item movement ----
        mv = ttk.LabelFrame(self, text="حركة الأصناف")
        mv.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 8))

        # Row filter checkboxes — only show items where checked column > 0
        mv_col_row = ttk.Frame(mv)
        mv_col_row.pack(fill=tk.X, padx=8, pady=(4, 0))
        ttk.Label(mv_col_row, text="إظهار فقط:").pack(side=tk.RIGHT, padx=(0, 4))
        self._mv_flt_received = tk.BooleanVar(value=False)
        self._mv_flt_sold = tk.BooleanVar(value=False)
        self._mv_flt_reserved = tk.BooleanVar(value=False)
        self._mv_flt_remaining = tk.BooleanVar(value=False)
        for var, text in [
            (self._mv_flt_received, "وارد"),
            (self._mv_flt_sold, "مباع"),
            (self._mv_flt_reserved, "محجوز"),
            (self._mv_flt_remaining, "متبقي"),
        ]:
            ttk.Checkbutton(mv_col_row, text=text, variable=var,
                            command=self._refresh_all).pack(side=tk.RIGHT, padx=4)

        mv_wrap = ttk.Frame(mv)
        mv_wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._mv_table = ttk.Treeview(
            mv_wrap,
            columns=("type", "school", "color", "size", "received", "sold", "reserved", "remaining"),
            show="headings",
            height=6,
        )
        for col, txt, w in [
            ("type", "النوع", 100), ("school", "المدرسة", 120), ("color", "اللون", 70),
            ("size", "المقاس", 60), ("received", "وارد", 70), ("sold", "مباع", 70),
            ("reserved", "محجوز", 70), ("remaining", "متبقي", 70),
        ]:
            self._mv_table.heading(col, text=txt)
            self._mv_table.column(col, width=w, anchor="center")
        mv_ysb = ttk.Scrollbar(mv_wrap, orient="vertical", command=self._mv_table.yview)
        self._mv_table.configure(yscrollcommand=mv_ysb.set)
        self._mv_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mv_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._mv_table)

        # Context menu on movement table
        _add_context_menu(self._mv_table, self)

        ttk.Button(mv, text="طباعة", command=self._print_movements).pack(anchor="e", padx=8, pady=4)

        # ---- Section 3: Income ----
        income_frame = ttk.LabelFrame(self, text="الوارد")
        income_frame.pack(fill=tk.X, padx=4, pady=4)

        self._income_count_var = tk.StringVar(value="0")
        self._income_qty_var = tk.StringVar(value="0")
        self._income_total_var = tk.StringVar(value="0")

        ir1 = ttk.Frame(income_frame)
        ir1.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(ir1, text="عدد فواتير الوارد:").pack(side=tk.RIGHT)
        ttk.Label(ir1, textvariable=self._income_count_var, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=(0, 6))

        ir2 = ttk.Frame(income_frame)
        ir2.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(ir2, text="إجمالي القطع الواردة:").pack(side=tk.RIGHT)
        ttk.Label(ir2, textvariable=self._income_qty_var, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=(0, 6))

        ir3 = ttk.Frame(income_frame)
        ir3.pack(fill=tk.X, padx=8, pady=4)
        ttk.Label(ir3, text="إجمالي قيمة الوارد:").pack(side=tk.RIGHT)
        ttk.Label(ir3, textvariable=self._income_total_var, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=(0, 6))

        ttk.Button(income_frame, text="طباعة", command=self._print_income).pack(anchor="w", padx=8, pady=4)

        self._refresh_all()

    def _refresh_all(self):
        df = self._df.get() or None
        dt = self._dt.get() or None
        school = self._sf_school.get() or None
        item_type = self._sf_item.get() or None
        color = self._sf_color.get() or None

        # Money flow - pass all filters
        try:
            stats = self.db.get_sales_stats(df, dt, school=school, item_type=item_type, color=color)
            self._sales_count_var.set(str(stats["sales_count"]))
            self._sales_total_var.set(f"{format_money(stats['sales_total'])}")
            self._res_count_var.set(str(stats["res_count"]))
            self._res_total_var.set(f"{format_money(stats['res_total'])}")
            self._res_paid_var.set(f"{format_money(stats['res_paid'])}")
        except Exception:
            pass

        # Item movements - pass all filters + row filtering by checkboxes
        flt_recv = self._mv_flt_received.get()
        flt_sold = self._mv_flt_sold.get()
        flt_resv = self._mv_flt_reserved.get()
        flt_rem  = self._mv_flt_remaining.get()
        any_flt  = flt_recv or flt_sold or flt_resv or flt_rem
        try:
            self._mv_table.delete(*self._mv_table.get_children())
            for r in self.db.get_item_movement_stats(date_from=df, date_to=dt, school=school, item_type=item_type, color=color):
                received = int(r.get("received") or 0)
                sold = int(r.get("sold") or 0)
                reserved = int(r.get("reserved") or 0)
                adjusted = int(r.get("adjusted") or 0)
                remaining = received - sold - reserved - adjusted
                # If any checkbox is active, only show rows matching at least one
                if any_flt:
                    if not ((flt_recv and received > 0) or
                            (flt_sold and sold > 0) or
                            (flt_resv and reserved > 0) or
                            (flt_rem and remaining > 0)):
                        continue
                self._mv_table.insert("", tk.END, values=(
                    r["item_type"], r["school"], r["color"], r["size"],
                    received, sold, reserved, remaining
                ))
            _apply_zebra_tags(self._mv_table)
        except Exception:
            pass

        # Income stats
        try:
            inc = self.db.get_income_stats(date_from=df, date_to=dt, school=school, item_type=item_type, color=color)
            self._income_count_var.set(str(inc["bill_count"]))
            self._income_qty_var.set(str(inc["total_qty"]))
            self._income_total_var.set(f"{format_money(inc['total_value'])}")
        except Exception:
            pass

    def _build_stats_constraints(self, exclude_field: str) -> Dict[str, Any]:
        c: Dict[str, Any] = {}
        if exclude_field != "school":
            v = self._sf_school.get()
            if v:
                c["school"] = v
        if exclude_field != "item_type":
            v = self._sf_item.get()
            if v:
                c["item_type"] = v
        if exclude_field != "color":
            v = self._sf_color.get()
            if v:
                c["color"] = v
        return c

    def _on_stats_filter_changed(self):
        self._sf_school.refresh_values()
        self._sf_item.refresh_values()
        self._sf_color.refresh_values()
        self._refresh_all()

    def _clear_stats_filters(self):
        self._sf_school.set("")
        self._sf_item.set("")
        self._sf_color.set("")
        self._df.set("")
        self._dt.set("")
        self._sf_school.refresh_values()
        self._sf_item.refresh_values()
        self._sf_color.refresh_values()
        self._refresh_all()

    def _print_money_flow(self):
        try:
            df = self._df.get() or None
            dt = self._dt.get() or None
            stats = self.db.get_sales_stats(df, dt)
            period = ""
            if df or dt:
                period = f"من {df or '---'} إلى {dt or '---'}"
            html = f"""<!DOCTYPE html>
<html lang='ar' dir='rtl'>
<head><meta charset='utf-8'><title>التدفق المالي</title>
<style>body{{font-family:Tahoma,Arial;margin:20px}} table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #777;padding:8px;text-align:right}}</style>
</head><body>
<h2>التدفق المالي {period}</h2>
<table><thead><tr><th>البيان</th><th>القيمة</th></tr></thead><tbody>
<tr><td>عدد الفواتير</td><td>{stats['sales_count']}</td></tr>
<tr><td>إجمالي المبيعات</td><td>{format_money(stats['sales_total'])}</td></tr>
<tr><td>عدد الحجوزات</td><td>{stats['res_count']}</td></tr>
<tr><td>إجمالي الحجوزات</td><td>{format_money(stats['res_total'])}</td></tr>
<tr><td>المبلغ المدفوع من الحجوزات</td><td>{format_money(stats['res_paid'])}</td></tr>
</tbody></table>
</body></html>"""
            import tempfile, os
            path = os.path.join(tempfile.gettempdir(), "money_flow.html")
            _write_html_file(path, html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)

    def _print_movements(self):
        try:
            rows_html = ""
            for iid in self._mv_table.get_children():
                vals = self._mv_table.item(iid, "values")
                rows_html += "<tr>" + "".join(f"<td>{_html(str(v))}</td>" for v in vals) + "</tr>"
            html = f"""<!DOCTYPE html>
<html lang='ar' dir='rtl'>
<head><meta charset='utf-8'><title>حركة الأصناف</title>
<style>body{{font-family:Tahoma,Arial;margin:20px}} table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #777;padding:6px;text-align:center}}</style>
</head><body>
<h2>حركة الأصناف</h2>
<table><thead><tr><th>النوع</th><th>المدرسة</th><th>اللون</th><th>المقاس</th><th>وارد</th><th>مباع</th><th>محجوز</th><th>متبقي</th></tr></thead>
<tbody>{rows_html}</tbody></table>
</body></html>"""
            import tempfile, os
            path = os.path.join(tempfile.gettempdir(), "item_movements.html")
            _write_html_file(path, html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)

    def _print_income(self):
        html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>إحصائيات الوارد</title>
<style>{_receipt_font_face_css()}@page {{ size: 80mm auto; margin: 2mm; }}
body {{ font-family: {RECEIPT_FONT_STACK}; font-size: 11px; width: 76mm; direction: rtl; margin:0; padding:2mm; }}
h2 {{ font-size: 14px; text-align: center; margin: 4px 0; }}
.sep {{ border:none; border-top:1px dashed #000; margin:4px 0; }}
.row {{ margin: 3px 0; }}
</style></head><body>
<h2>إحصائيات الوارد</h2>
<hr class="sep">
<div class="row">عدد فواتير الوارد: {self._income_count_var.get()}</div>
<div class="row">إجمالي القطع الواردة: {self._income_qty_var.get()}</div>
<div class="row">إجمالي قيمة الوارد: {self._income_total_var.get()}</div>
</body></html>"""
        tmp = os.path.join(tempfile.gettempdir(), "income_stats.html")
        _write_html_file(tmp, html)
        _print_html_auto(tmp, copies=1, parent=self)


class SchoolAccountsFrame(ttk.Frame):
    """School cash-flow report without automatic percentage calculations."""

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master, padding=10)
        self.db = db
        self._selected_schools: List[str] = []
        self._build()
        self._refresh()

    def _update_selected_schools_label(self) -> None:
        if self._selected_schools:
            self._schools_var.set("المدارس المختارة: " + " - ".join(self._selected_schools))
        else:
            self._schools_var.set("كل المدارس")

    def _build(self):
        ttk.Label(self, text="حسابات المدارس", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 12))

        filters = ttk.Frame(self)
        filters.pack(fill=tk.X, pady=(0, 8))

        self.df = DateField(filters, "من")
        self.df.pack(side=tk.LEFT, padx=(0, 8))
        self.dt = DateField(filters, "إلى")
        self.dt.pack(side=tk.LEFT, padx=(0, 8))
        today = date.today().strftime("%Y-%m-%d")
        self.df.set(today)
        self.dt.set(today)
        ttk.Button(filters, text="اختيار المدارس", command=self._pick_schools).pack(side=tk.LEFT, padx=8)
        ttk.Button(filters, text="تحديث", command=self._refresh).pack(side=tk.LEFT, padx=8)
        ttk.Button(filters, text="طباعة الإجمالي", command=self._print_day_total).pack(side=tk.LEFT, padx=8)
        ttk.Button(filters, text="مسح الاختيار", command=self._clear_schools).pack(side=tk.LEFT, padx=8)

        self._schools_var = tk.StringVar(value="كل المدارس")
        ttk.Label(self, textvariable=self._schools_var, foreground="#64748b").pack(anchor="w", pady=(0, 8))

        summary = ttk.LabelFrame(self, text="ملخص اليوم")
        summary.pack(fill=tk.X, pady=(0, 8))
        self._sum_day = tk.StringVar(value="0")
        self._sum_cash = tk.StringVar(value="0")
        self._sum_visa = tk.StringVar(value="0")
        for i, (label, var) in enumerate([
            ("إجمالي اليوم:", self._sum_day),
            ("إجمالي كاش:", self._sum_cash),
            ("إجمالي فيزا:", self._sum_visa),
        ]):
            ttk.Label(summary, text=label).grid(row=0, column=i * 2, padx=6, pady=8, sticky="e")
            ttk.Label(summary, textvariable=var, font=("Segoe UI", 12, "bold")).grid(
                row=0, column=i * 2 + 1, padx=(0, 16), pady=8, sticky="w"
            )

        table_wrap = ttk.LabelFrame(self, text="فواتير اليوم")
        table_wrap.pack(fill=tk.BOTH, expand=True)
        cols = (
            "event_at", "bill_no", "bill_type", "schools", "customer",
            "payment_method", "cash_total", "visa_total", "total_paid",
        )
        self._tree = ttk.Treeview(table_wrap, columns=cols, show="headings", height=16)
        for col, txt, w in [
            ("event_at", "التاريخ", 150),
            ("bill_no", "رقم الفاتورة", 160),
            ("bill_type", "النوع", 130),
            ("schools", "المدارس", 180),
            ("customer", "العميل", 170),
            ("payment_method", "طريقة الدفع", 100),
            ("cash_total", "كاش", 110),
            ("visa_total", "فيزا", 110),
            ("total_paid", "الإجمالي", 120),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self._tree.yview)
        xsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self._tree.xview)
        self._tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4, pady=4)
        ysb.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 4), pady=4)
        xsb.pack(side=tk.BOTTOM, fill=tk.X, padx=4, pady=(0, 4))
        _add_context_menu(self._tree, self)
        _bind_mousewheel(self._tree)

    def _pick_schools(self):
        values = self.db.list_schools_all()
        dlg = MultiSelectDialog(self, title="اختيار المدارس", values=values, preselected=self._selected_schools)
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
        date_from = self.df.get() or None
        date_to = self.dt.get() or None
        if date_from and not date_to:
            date_to = date_from
        elif date_to and not date_from:
            date_from = date_to
        report = self.db.get_school_accounts_day_report(
            self._selected_schools or self.db.list_schools_all(),
            date_from=date_from,
            date_to=date_to,
        )
        rows = report.get("rows", [])
        self._tree.delete(*self._tree.get_children())
        for row in rows:
            self._tree.insert(
                "",
                tk.END,
                values=(
                    fmt_local_ts(row.get("event_at"), ""),
                    row.get("bill_no") or "",
                    self._bill_type_label(row.get("bill_type")),
                    row.get("schools") or "",
                    row.get("customer") or "",
                    _payment_method_label(row.get("payment_method")),
                    f"{format_money(float(row.get('cash_total') or 0.0))}",
                    f"{format_money(float(row.get('visa_total') or 0.0))}",
                    f"{format_money(float(row.get('total_paid') or 0.0))}",
                ),
            )
        _apply_zebra_tags(self._tree)
        self._update_selected_schools_label()
        self._sum_day.set(f"{format_money(float(report.get('total_day') or 0.0))}")
        self._sum_cash.set(f"{format_money(float(report.get('total_cash') or 0.0))}")
        self._sum_visa.set(f"{format_money(float(report.get('total_visa') or 0.0))}")

    @staticmethod
    def _bill_type_label(value: Any) -> str:
        code = str(value or "").strip().upper()
        if code == "VOID":
            return "إلغاء"
        if code == "RESERVATION":
            return "عربون حجز"
        if code == "RESERVATION_DELIVERY":
            return "تسليم حجز"
        if code == "RESERVATION_REFUND":
            return "استرداد حجز"
        return _receipt_bill_type_label(code)

    def _print_day_total(self):
        try:
            date_from = self.df.get() or None
            date_to = self.dt.get() or None
            if date_from and not date_to:
                date_to = date_from
            elif date_to and not date_from:
                date_from = date_to
            schools = self._selected_schools or self.db.list_schools_all()
            report = self.db.get_school_accounts_day_report(
                schools,
                date_from=date_from,
                date_to=date_to,
            )
            schools_text = " - ".join(schools) if schools else "كل المدارس"
            range_text = _strip_digit_marks(f"{date_from or '---'} إلى {date_to or '---'}")
            total = _strip_digit_marks(format_money(float(report.get("total_day") or 0.0)))
            html = f"""<!DOCTYPE html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>إجمالي حسابات المدارس</title>
<style>{_receipt_font_face_css()}@page {{ size: 80mm auto; margin: 2mm; }}
body {{ font-family: {RECEIPT_FONT_STACK}; font-size: 12px; width: 76mm; direction: rtl; margin:0; padding:2mm; }}
h2 {{ font-size: 15px; text-align: center; margin: 4px 0 6px; }}
.sep {{ border:none; border-top:1px dashed #000; margin:6px 0; }}
.row {{ margin: 5px 0; }}
.label {{ font-weight: bold; }}
.total {{ font-size: 18px; font-weight: bold; text-align: center; margin-top: 8px; }}
.amount {{ font-family: Tahoma, Arial, "Segoe UI", sans-serif; direction: ltr; unicode-bidi: isolate; font-weight: bold; }}
</style></head><body>
<h2>إجمالي حسابات المدارس</h2>
<hr class="sep">
<div class="row"><span class="label">المدارس:</span> {_html(schools_text)}</div>
<div class="row"><span class="label">الفترة:</span> <span class="amount">{_html(range_text)}</span></div>
<hr class="sep">
<div class="total">الإجمالي: <span class="amount">{_html(total)}</span></div>
</body></html>"""
            path = os.path.join(tempfile.gettempdir(), "school_accounts_day_total.html")
            _write_html_file(path, html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("فشل الطباعة", str(ex), parent=self)


# ------------------- Shifts Summary Frame -------------------

class ShiftsSummaryFrame(ttk.Frame):
    """Tab showing history and details of all shifts."""

    def __init__(self, master, db: "SqliteDatabase"):
        super().__init__(master, padding=6)
        self.db = db
        self._build()

    def _build(self):
        ttk.Label(self, text="ملخص الورديات", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 6))

        # ---- Filter bar ----
        filter_bar = ttk.Frame(self)
        filter_bar.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(filter_bar, text="من:").pack(side=tk.LEFT)
        self._df = DateField(filter_bar, "")
        self._df.pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(filter_bar, text="إلى:").pack(side=tk.LEFT)
        self._dt = DateField(filter_bar, "")
        self._dt.pack(side=tk.LEFT, padx=(4, 12))

        ttk.Button(filter_bar, text="تحديث", command=self._refresh_all).pack(side=tk.LEFT, padx=4)
        ttk.Button(filter_bar, text="مسح", command=self._clear_filters).pack(side=tk.LEFT, padx=4)
        ttk.Button(filter_bar, text="طباعة الكل", command=self._print_all_shifts).pack(side=tk.RIGHT, padx=4)

        # ---- Shifts table ----
        table_frame = ttk.Frame(self)
        table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        cols = ("id", "started", "ended", "status", "day_total",
                "return_total", "void_total", "exchange_total", "expense_total", "cash", "visa")
        self._tree = ttk.Treeview(table_frame, columns=cols, show="headings", height=10)
        for col, txt, w in [
            ("id", "#", 40), ("started", "البداية", 130), ("ended", "النهاية", 130),
            ("status", "الحالة", 70), ("day_total", "اجمالي اليوم", 110),
            ("return_total", "مرتجعات", 90),
            ("void_total", "إلغاء", 80), ("exchange_total", "استبدالات", 90),
            ("expense_total", "مصروفات", 90),
            ("cash", "كاش صافي", 100), ("visa", "فيزا صافي", 100),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=ysb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._tree)
        _add_context_menu(self._tree, self)

        self._tree.bind("<<TreeviewSelect>>", self._on_shift_selected)

        # ---- Detail panel ----
        self._detail_frame = ttk.LabelFrame(self, text="تفاصيل الوردية")
        self._detail_frame.pack(fill=tk.X, pady=(0, 4))

        detail_top = ttk.Frame(self._detail_frame)
        detail_top.pack(fill=tk.X, padx=8, pady=4)

        self._det_id_var = tk.StringVar(value="-")
        self._det_period_var = tk.StringVar(value="-")
        self._det_status_var = tk.StringVar(value="-")

        ttk.Label(detail_top, text="الوردية:").pack(side=tk.LEFT)
        ttk.Label(detail_top, textvariable=self._det_id_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 16))
        ttk.Label(detail_top, text="الفترة:").pack(side=tk.LEFT)
        ttk.Label(detail_top, textvariable=self._det_period_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 16))
        ttk.Label(detail_top, text="الحالة:").pack(side=tk.LEFT)
        ttk.Label(detail_top, textvariable=self._det_status_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 0))

        detail_nums = ttk.Frame(self._detail_frame)
        detail_nums.pack(fill=tk.X, padx=8, pady=4)

        self._det_sales_var = tk.StringVar(value="-")
        self._det_return_var = tk.StringVar(value="-")
        self._det_void_var = tk.StringVar(value="-")
        self._det_exchange_var = tk.StringVar(value="-")
        self._det_expense_var = tk.StringVar(value="-")
        self._det_cash_var = tk.StringVar(value="-")
        self._det_visa_var = tk.StringVar(value="-")

        for lbl, var in [
            ("اجمالي اليوم:", self._det_sales_var),
            ("المرتجعات:", self._det_return_var),
            ("الإلغاء:", self._det_void_var),
            ("الاستبدالات:", self._det_exchange_var),
            ("المصروفات:", self._det_expense_var),
        ]:
            ttk.Label(detail_nums, text=lbl).pack(side=tk.LEFT)
            ttk.Label(detail_nums, textvariable=var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 16))

        cash_row = ttk.Frame(self._detail_frame)
        cash_row.pack(fill=tk.X, padx=8, pady=(0, 4))
        ttk.Label(cash_row, text="كاش صافي:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT)
        ttk.Label(cash_row, textvariable=self._det_cash_var, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(cash_row, text=" | فيزا صافي:", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(cash_row, textvariable=self._det_visa_var, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(4, 0))

        detail_btns = ttk.Frame(self._detail_frame)
        detail_btns.pack(fill=tk.X, padx=8, pady=(0, 6))
        ttk.Button(detail_btns, text="طباعة ملخص الوردية", command=self._print_selected_shift).pack(side=tk.LEFT)

        # ---- Totals bar ----
        totals = ttk.LabelFrame(self, text="الإجماليات")
        totals.pack(fill=tk.X)

        totals_row = ttk.Frame(totals)
        totals_row.pack(fill=tk.X, padx=8, pady=4)

        self._tot_shifts_var = tk.StringVar(value="0")
        self._tot_day_var = tk.StringVar(value="0")
        self._tot_expense_var = tk.StringVar(value="0")
        self._tot_cash_var = tk.StringVar(value="0")
        self._tot_visa_var = tk.StringVar(value="0")

        ttk.Label(totals_row, text="عدد الورديات:").pack(side=tk.LEFT)
        ttk.Label(totals_row, textvariable=self._tot_shifts_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 20))
        ttk.Label(totals_row, text="اجمالي اليوم:").pack(side=tk.LEFT)
        ttk.Label(totals_row, textvariable=self._tot_day_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 20))
        ttk.Label(totals_row, text="مصروفات:").pack(side=tk.LEFT)
        ttk.Label(totals_row, textvariable=self._tot_expense_var, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 20))
        ttk.Label(totals_row, text="كاش صافي:").pack(side=tk.LEFT)
        ttk.Label(totals_row, textvariable=self._tot_cash_var, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(totals_row, text=" | فيزا صافي:").pack(side=tk.LEFT, padx=(16, 0))
        ttk.Label(totals_row, textvariable=self._tot_visa_var, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=(4, 0))

        self._shifts_data: List[Dict[str, Any]] = []
        self._refresh_all()

    def _refresh_all(self):
        df = self._df.get() or None
        dt = self._dt.get() or None
        try:
            self._shifts_data = self.db.get_all_shifts(date_from=df, date_to=dt)
        except Exception:
            self._shifts_data = []

        self._tree.delete(*self._tree.get_children())
        total_day = 0.0
        total_expense = 0.0
        total_cash = 0.0
        total_visa = 0.0
        for s in self._shifts_data:
            started = fmt_local_ts(s["started_at"], "")
            ended = fmt_local_ts(s["ended_at"], "") if s["ended_at"] else "-"
            status = "مفتوحة" if s["status"] == "OPEN" else "مغلقة"
            day_total = float(s.get("cash_collected") or 0.0) + float(s.get("visa_collected") or 0.0)
            self._tree.insert("", tk.END, values=(
                s["id"], started, ended, status,
                f"{format_money(day_total)}",
                f"{format_money(s.get('return_total', 0.0))}",
                f"{format_money(s.get('void_total', 0.0))}",
                f"{format_money(s.get('exchange_total', 0.0))}",
                f"{format_money(s.get('expense_total', 0.0))}",
                f"{format_money(s['cash_collected'])}",
                f"{format_money(s.get('visa_collected', 0.0))}",
            ))
            total_day += day_total
            total_expense += float(s.get("expense_total") or 0.0)
            total_cash += s["cash_collected"]
            total_visa += float(s.get("visa_collected", 0.0))
        _apply_zebra_tags(self._tree)

        self._tot_shifts_var.set(str(len(self._shifts_data)))
        self._tot_day_var.set(f"{format_money(total_day)}")
        self._tot_expense_var.set(f"{format_money(total_expense)}")
        self._tot_cash_var.set(f"{format_money(total_cash)}")
        self._tot_visa_var.set(f"{format_money(total_visa)}")

        # Clear detail
        self._det_id_var.set("-")
        self._det_period_var.set("-")
        self._det_status_var.set("-")
        self._det_sales_var.set("-")
        self._det_return_var.set("-")
        self._det_void_var.set("-")
        self._det_exchange_var.set("-")
        self._det_expense_var.set("-")
        self._det_cash_var.set("-")
        self._det_visa_var.set("-")

    def _on_shift_selected(self, _event=None):
        sel = self._tree.selection()
        if not sel:
            return
        vals = self._tree.item(sel[0], "values")
        shift_id = parse_int_text(vals[0])
        s = next((x for x in self._shifts_data if x["id"] == shift_id), None)
        if not s:
            return

        started = s["started_at"][:16].replace("T", " ")
        ended = s["ended_at"][:16].replace("T", " ") if s["ended_at"] else "الآن"
        status = "مفتوحة" if s["status"] == "OPEN" else "مغلقة"

        self._det_id_var.set(f"#{s['id']}")
        self._det_period_var.set(f"{started}  →  {ended}")
        self._det_status_var.set(status)
        day_total = float(s.get("cash_collected") or 0.0) + float(s.get("visa_collected") or 0.0)
        self._det_sales_var.set(f"{format_money(day_total)}")
        self._det_return_var.set(f"{s.get('return_count', 0)} فاتورة - {format_money(s.get('return_total', 0.0))}")
        self._det_void_var.set(f"{s.get('void_count', 0)} فاتورة - {format_money(s.get('void_total', 0.0))}")
        self._det_exchange_var.set(f"{s.get('exchange_count', 0)} فاتورة - صافي {format_money(s.get('exchange_total', 0.0))}")
        self._det_expense_var.set(f"{s.get('expense_count', 0)} عملية - {format_money(s.get('expense_total', 0.0))}")
        self._det_cash_var.set(f"{format_money(s['cash_collected'])}")
        self._det_visa_var.set(f"{format_money(float(s.get('visa_collected', 0.0)))}")

    def _clear_filters(self):
        self._df.set("")
        self._dt.set("")
        self._refresh_all()

    def _get_selected_shift(self) -> Optional[Dict[str, Any]]:
        sel = self._tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر وردية أولاً", parent=self)
            return None
        shift_id = parse_int_text(self._tree.item(sel[0], "values")[0])
        return next((x for x in self._shifts_data if x["id"] == shift_id), None)

    def _print_selected_shift(self):
        s = self._get_selected_shift()
        if not s:
            return
        started = s["started_at"][:16].replace("T", " ")
        ended = s["ended_at"][:16].replace("T", " ") if s["ended_at"] else "الآن"
        status = "مفتوحة" if s["status"] == "OPEN" else "مغلقة"
        void_html = ""
        if int(s.get("void_count", 0) or 0) > 0 or float(s.get("void_total", 0.0) or 0.0) > 1e-9:
            void_html = f'<div><b>الإلغاء:</b> {_num_html(s.get("void_count", 0))} فاتورة - {_money_html(s.get("void_total", 0.0))}</div>\n'
        customer_money = float(s.get("res_paid", 0.0)) + float(s.get("deliver_total", 0.0))
        customer_money_html = ""
        if customer_money > 1e-9 or int(s.get("res_count", 0) or 0) > 0 or int(s.get("deliver_count", 0) or 0) > 0:
            customer_money_html = (
                f'<div><b>مبالغ واردة من العملاء:</b></div>\n'
                f'<div>حجوزات جديدة: {_num_html(s.get("res_count", 0))} - إجمالي {_money_html(s.get("res_total", 0.0))}</div>\n'
                f'<div>عربون الحجوزات: {_money_html(s.get("res_paid", 0.0))}</div>\n'
                f'<div>تحصيل تسليم الحجوزات: {_money_html(s.get("deliver_total", 0.0))}</div>\n'
                f'<hr class="sep">'
            )
        expense_html = (
            f'<div><b>مصروفات:</b> {_num_html(s.get("expense_count", 0))} عملية - {_money_html(s.get("expense_total", 0.0))}</div>\n'
            f'<div>للعرض فقط ولا تدخل في اجمالي اليوم.</div>\n'
            f'<hr class="sep">'
        )
        html = f"""<!DOCTYPE html>
<html lang="ar" dir="ltr"><head><meta charset="utf-8"><title>ملخص وردية</title>
<style>
{_receipt_font_face_css()}
@page {{ size: 80mm auto; margin: 2mm; }}
body {{ font-family: {RECEIPT_FONT_STACK}; font-size: 11px; width: 76mm; min-height: 0; direction: ltr; margin:0; padding:2mm; }}
.receipt {{ direction: rtl; unicode-bidi: isolate; min-height: auto; height: auto; max-height: none; overflow: visible; }}
.digits {{ direction: ltr; unicode-bidi: isolate; font-family: Tahoma, Arial, "Segoe UI", sans-serif; }}
h2 {{ font-size: 14px; text-align: center; margin: 4px 0; }}
.sep {{ border:none; border-top:1px dashed #000; margin:4px 0; }}
.total {{ font-size: 13px; font-weight: bold; text-align: center; margin: 6px 0; }}
</style></head><body><div class="receipt">
<h2>ملخص الوردية #{_num_html(s['id'])}</h2>
<hr class="sep">
<div>الحالة: {_html(status)}</div>
<div>بداية: {_num_html(started)}</div>
<div>نهاية: {_num_html(ended)}</div>
<hr class="sep">
<div><b>اجمالي اليوم:</b> {_money_html(float(s.get('cash_collected') or 0.0) + float(s.get('visa_collected') or 0.0))}</div>
<hr class="sep">
<div><b>المرتجعات:</b> {_num_html(s.get('return_count', 0))} فاتورة - {_money_html(s.get('return_total', 0.0))}</div>
{void_html}
<div><b>الاستبدالات:</b> {_num_html(s.get('exchange_count', 0))} فاتورة - صافي {_money_html(s.get('exchange_total', 0.0))}</div>
<hr class="sep">
{customer_money_html}{expense_html}<div class="total">كاش صافي: {_money_html(s['cash_collected'])} | فيزا صافي: {_money_html(float(s.get('visa_collected', 0.0)))}</div>
</div></body></html>"""
        tmp = os.path.join(tempfile.gettempdir(), f"shift_summary_{s['id']}.html")
        _write_html_file(tmp, html)
        _print_html_auto(tmp, copies=1, parent=self)

    def _print_all_shifts(self):
        if not self._shifts_data:
            messagebox.showinfo("تنبيه", "لا توجد ورديات لطباعتها", parent=self)
            return
        rows_html = ""
        total_day = 0.0
        total_expense = 0.0
        total_cash = 0.0
        total_visa = 0.0
        for s in self._shifts_data:
            started = s["started_at"][:16].replace("T", " ")
            ended = s["ended_at"][:16].replace("T", " ") if s["ended_at"] else "-"
            status = "مفتوحة" if s["status"] == "OPEN" else "مغلقة"
            day_total = float(s.get("cash_collected") or 0.0) + float(s.get("visa_collected") or 0.0)
            rows_html += f"<tr><td>{_num_html(s['id'])}</td><td>{_num_html(started)}</td><td>{_num_html(ended)}</td><td>{_html(status)}</td>"
            rows_html += f"<td>{_money_html(day_total)}</td>"
            rows_html += f"<td>{_money_html(s.get('return_total', 0.0))}</td><td>{_money_html(s.get('void_total', 0.0))}</td><td>{_money_html(s.get('exchange_total', 0.0))}</td>"
            rows_html += f"<td>{_money_html(s.get('expense_total', 0.0))}</td>"
            rows_html += f"<td>{_money_html(s['cash_collected'])}</td><td>{_money_html(float(s.get('visa_collected', 0.0)))}</td></tr>\n"
            total_day += day_total
            total_expense += float(s.get("expense_total") or 0.0)
            total_cash += s["cash_collected"]
            total_visa += float(s.get("visa_collected", 0.0))
        html = f"""<!DOCTYPE html>
<html lang="ar" dir="ltr"><head><meta charset="utf-8"><title>ملخص الورديات</title>
<style>
{_receipt_font_face_css()}
body {{ font-family: {RECEIPT_FONT_STACK}; margin: 20px; direction: ltr; }}
.report {{ direction: rtl; unicode-bidi: isolate; }}
.digits {{ direction: ltr; unicode-bidi: isolate; font-family: Tahoma, Arial, "Segoe UI", sans-serif; }}
h2 {{ text-align: center; }}
table {{ border-collapse: collapse; width: 100%; font-size: 11px; }}
th {{ border-bottom: 2px solid #000; padding: 6px; text-align: center; background: #f0f0f0; }}
td {{ padding: 5px; border-bottom: 1px solid #ccc; text-align: center; }}
.totals {{ font-weight: bold; font-size: 13px; margin-top: 12px; }}
</style></head><body><div class="report">
<h2>ملخص جميع الورديات ({_num_html(len(self._shifts_data))} وردية)</h2>
<table><thead><tr>
<th>#</th><th>البداية</th><th>النهاية</th><th>الحالة</th>
<th>اجمالي اليوم</th><th>مرتجعات</th><th>إلغاء</th><th>استبدالات</th><th>مصروفات</th>
<th>كاش صافي</th><th>فيزا صافي</th>
</tr></thead><tbody>
{rows_html}
</tbody></table>
<div class="totals">اجمالي اليوم: {_money_html(total_day)} | مصروفات: {_money_html(total_expense)} | كاش صافي: {_money_html(total_cash)} | فيزا صافي: {_money_html(total_visa)}</div>
</div></body></html>"""
        tmp = os.path.join(tempfile.gettempdir(), "all_shifts_summary.html")
        _write_html_file(tmp, html)
        _print_html_auto(tmp, copies=1, parent=self)


# ------------------- Reservations Frame -------------------

class ReservationsFrame(ttk.Frame):
    """Standalone tab for managing reservations."""
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=6)
        self.db = db
        self._build()

    def _build(self):
        # ---- Filter row ----
        flt = ttk.Frame(self)
        flt.pack(fill=tk.X, padx=4, pady=(0, 6))

        self._rf_school = LabeledCombobox(flt, "المدرسة", self.db, "school")
        self._rf_school.pack(side=tk.RIGHT, padx=4)
        self._rf_school.set_supplier(lambda: self.db.get_distinct_filtered("school", self._build_res_filter_constraints("school")))

        self._rf_item = LabeledCombobox(flt, "النوع", self.db, "item_type")
        self._rf_item.pack(side=tk.RIGHT, padx=4)
        self._rf_item.set_supplier(lambda: self.db.get_distinct_filtered("item_type", self._build_res_filter_constraints("item_type")))

        self._rf_color = LabeledCombobox(flt, "اللون", self.db, "color")
        self._rf_color.pack(side=tk.RIGHT, padx=4)
        self._rf_color.set_supplier(lambda: self.db.get_distinct_filtered("color", self._build_res_filter_constraints("color")))

        # Status filter
        ttk.Label(flt, text="الحالة:").pack(side=tk.RIGHT, padx=(4, 0))
        self._rf_status = ttk.Combobox(flt, values=["", "معلق", "تم التسليم"], state="readonly", width=12)
        self._rf_status.pack(side=tk.RIGHT, padx=4)
        self._rf_status.bind("<<ComboboxSelected>>", lambda e: self._refresh())

        # Date filters
        self._rf_df = DateField(flt, "من:")
        self._rf_df.pack(side=tk.RIGHT, padx=4)
        self._rf_dt = DateField(flt, "إلى:")
        self._rf_dt.pack(side=tk.RIGHT, padx=4)

        ttk.Button(flt, text="تحديث", command=self._refresh).pack(side=tk.RIGHT, padx=4)
        ttk.Button(flt, text="مسح", command=self._clear_filters).pack(side=tk.RIGHT, padx=4)

        for w in (self._rf_school.cb, self._rf_item.cb, self._rf_color.cb):
            w.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed(), add="+")

        # ---- Reservations table ----
        tbl_wrap = ttk.Frame(self)
        tbl_wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._res_table = ttk.Treeview(
            tbl_wrap,
            columns=("bill", "created", "customer", "phone", "item", "school", "color", "size", "qty", "total", "paid", "status"),
            show="headings",
            height=12,
        )
        for col, txt, w in [
            ("bill", "الفاتورة", 88),
            ("created", "التاريخ", 100), ("customer", "العميل", 100),
            ("phone", "رقم العميل", 110),
            ("item", "النوع", 90), ("school", "المدرسة", 110), ("color", "اللون", 70),
            ("size", "المقاس", 55), ("qty", "الكمية", 55), ("total", "الإجمالي", 80),
            ("paid", "المدفوع", 80), ("status", "الحالة", 80),
        ]:
            self._res_table.heading(col, text=txt)
            self._res_table.column(col, width=w, anchor="center")
        rv_ysb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self._res_table.yview)
        self._res_table.configure(yscrollcommand=rv_ysb.set)
        self._res_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rv_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(self._res_table)

        # Context menu
        _add_context_menu(self._res_table, self)

        # ---- Buttons row ----
        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, padx=4, pady=4)
        ttk.Button(btns, text="تسليم الحجز", command=self._deliver_reservation_partial).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btns, text="طباعة اجمالي الحجوزات", command=self._print_reservation_totals).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btns, text="طباعة القائمة", command=self._print_reservations).pack(side=tk.RIGHT)

        self._refresh()

    def _build_res_filter_constraints(self, exclude_field: str) -> Dict[str, Any]:
        c: Dict[str, Any] = {}
        if exclude_field != "school":
            v = self._rf_school.get()
            if v: c["school"] = v
        if exclude_field != "item_type":
            v = self._rf_item.get()
            if v: c["item_type"] = v
        if exclude_field != "color":
            v = self._rf_color.get()
            if v: c["color"] = v
        return c

    def _on_filter_changed(self):
        self._rf_school.refresh_values()
        self._rf_item.refresh_values()
        self._rf_color.refresh_values()
        self._refresh()

    def _clear_filters(self):
        self._rf_school.set("")
        self._rf_item.set("")
        self._rf_color.set("")
        self._rf_status.set("")
        self._rf_df.set("")
        self._rf_dt.set("")
        self._rf_school.refresh_values()
        self._rf_item.refresh_values()
        self._rf_color.refresh_values()
        self._refresh()

    def _refresh(self):
        df = self._rf_df.get() or None
        dt = self._rf_dt.get() or None
        school = self._rf_school.get() or None
        item_type = self._rf_item.get() or None
        color = self._rf_color.get() or None
        status = self._rf_status.get() or None
        try:
            self._res_table.delete(*self._res_table.get_children())
            for r in self.db.list_reservations(status=status, date_from=df, date_to=dt,
                                                school=school, item_type=item_type, color=color):
                try:
                    group_rows = self.db.list_reservation_group_items(int(r["id"]))
                    bill_code = _reservation_bill_id_from_rows(group_rows) or str(r["id"])
                except Exception:
                    bill_code = str(r["id"])
                self._res_table.insert("", tk.END, iid=str(r["id"]), values=(
                    bill_code, fmt_local_ts(r["created_at"], "")[:10], r.get("customer", ""),
                    r.get("customer_phone", "") or "",
                    r["item_type"], r["school"], r["color"], r["size"],
                    r["qty"], f"{format_money(float(r['total_amount']))}",
                    f"{format_money(float(r['paid_amount']))}", r["status"]
                ))
            _apply_zebra_tags(self._res_table)
        except Exception:
            pass

    def _choose_payment_method(self, total: float, title: str = "\u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u062f\u0641\u0639") -> Optional[str]:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text="\u0627\u062e\u062a\u0631 \u0637\u0631\u064a\u0642\u0629 \u0627\u0644\u062f\u0641\u0639:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, text=f"\u0627\u0644\u0645\u0628\u0644\u063a: {format_money(abs(float(total or 0.0)))}").pack(anchor="w", pady=(4, 10))
        method_var = tk.StringVar(value=PAYMENT_METHOD_CASH)
        ttk.Radiobutton(frm, text="\u0643\u0627\u0634", variable=method_var, value=PAYMENT_METHOD_CASH).pack(anchor="w", pady=2)
        ttk.Radiobutton(frm, text="\u0641\u064a\u0632\u0627", variable=method_var, value=PAYMENT_METHOD_VISA).pack(anchor="w", pady=2)
        result = {"value": None}
        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(12, 0))

        def on_ok():
            result["value"] = method_var.get().strip().upper() or PAYMENT_METHOD_CASH
            dlg.destroy()

        ttk.Button(btns, text="\u0625\u0644\u063a\u0627\u0621", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="\u062a\u0623\u0643\u064a\u062f", command=on_ok).pack(side=tk.RIGHT, padx=6)
        dlg.wait_window()
        return result["value"]

    def _deliver_reservation(self):
        """Mark selected reservation as delivered and collect remaining payment."""
        sel = self._res_table.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر حجزاً من الجدول أولاً.", parent=self)
            return
        vals = self._res_table.item(sel[0], "values")
        res_id = parse_int_text(sel[0])
        status = vals[11]
        if status == "تم التسليم":
            messagebox.showinfo("تنبيه", "تم تسليم هذا الحجز مسبقاً.", parent=self)
            return
        total = _parse_money_amount(vals[9])
        paid = _parse_money_amount(vals[10])
        remaining = total - paid

        dlg = tk.Toplevel(self)
        dlg.title("تسليم حجز")
        dlg.geometry("350x220")
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=f"حجز رقم: {res_id}", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 4))
        ttk.Label(frm, text=f"الإجمالي: {format_money(total)}").pack(anchor="w")
        ttk.Label(frm, text=f"المدفوع سابقاً: {format_money(paid)}").pack(anchor="w")
        ttk.Label(frm, text=f"المتبقي: {format_money(remaining)}", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 8))

        ttk.Label(frm, text="المبلغ المحصّل الآن:").pack(anchor="w")
        collect_var = tk.StringVar(value=f"{format_money(remaining)}" if remaining > 0 else "0")
        ttk.Entry(frm, textvariable=collect_var, width=15).pack(anchor="w", pady=(0, 8))

        def _confirm():
            try:
                collected = _parse_money_amount(collect_var.get())
            except ValueError:
                messagebox.showerror("خطأ", "أدخل مبلغاً صحيحاً.", parent=dlg)
                return
            try:
                payment_method = PAYMENT_METHOD_CASH
                if collected > 1e-9:
                    payment_method = self._choose_payment_method(collected, "\u0637\u0631\u064a\u0642\u0629 \u062a\u062d\u0635\u064a\u0644 \u0628\u0627\u0642\u064a \u0627\u0644\u062d\u062c\u0632")
                    if not payment_method:
                        return
                self.db.deliver_reservation(res_id, collected, payment_method=payment_method)
                dlg.destroy()
                ToastNotification.show(self.winfo_toplevel(), "تم تسليم الحجز بنجاح", toast_type="success")
                self._refresh()
                try:
                    ac = getattr(self.winfo_toplevel(), "_app_controller", None)
                    if ac is not None and hasattr(ac, "refresh_dashboard"):
                        ac.refresh_dashboard()
                except Exception:
                    pass
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        ttk.Button(frm, text="تأكيد التسليم", command=_confirm).pack(anchor="e")

    def _deliver_reservation_partial(self):
        """Deliver selected items from a reservation bill (partial by default)."""
        sel = self._res_table.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر حجزاً من الجدول أولاً.", parent=self)
            return
        vals = self._res_table.item(sel[0], "values")
        res_id = parse_int_text(sel[0])
        status = vals[11]
        if status == "تم التسليم":
            messagebox.showinfo("تنبيه", "هذا السطر تم تسليمه مسبقاً.", parent=self)
            return

        group_items = self.db.list_reservation_group_items(res_id)
        if not group_items:
            messagebox.showerror("خطأ", "تعذر تحميل عناصر فاتورة الحجز.", parent=self)
            return
        pending_items = [r for r in group_items if _is_reservation_active(r.get("status"))]
        if not pending_items:
            messagebox.showinfo("تنبيه", "كل عناصر الفاتورة تم تسليمها مسبقاً.", parent=self)
            return
        group_code = _reservation_bill_id_from_rows(group_items) or str(res_id)
        remaining = sum(max(0.0, float(r.get("total_amount") or 0.0) - float(r.get("paid_amount") or 0.0)) for r in pending_items)

        dlg = tk.Toplevel(self)
        dlg.title("تسليم عناصر الحجز")
        # Start larger so bottom action controls are visible immediately.
        dlg.geometry("900x620")
        try:
            dlg.minsize(860, 560)
        except Exception:
            pass
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text=f"فاتورة الحجز #{group_code}", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 4))
        customer_name = str(group_items[0].get("customer") or "").strip()
        customer_phone = _normalize_customer_phone(group_items[0].get("customer_phone") or "")
        if customer_name or customer_phone:
            details = []
            if customer_name:
                details.append(f"العميل: {customer_name}")
            if customer_phone:
                details.append(f"رقم العميل: {customer_phone}")
            ttk.Label(frm, text=" | ".join(details), font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 4))
        ttk.Label(frm, text="اختر العناصر التي تم تسليمها الآن (تسليم جزئي افتراضياً)").pack(anchor="w")
        ttk.Label(frm, text="مهم: في التسليم الجزئي يجب تحصيل قيمة العناصر المسلمة بالكامل، والعربون يبقى للعناصر المعلقة.", foreground="#7a3e00").pack(anchor="w")
        ttk.Label(frm, text=f"المتبقي لكل العناصر المعلقة: {format_money(remaining)}", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 4))
        required_var = tk.StringVar(value="المطلوب الآن: 0")
        selected_total_var = tk.StringVar(value="إجمالي العناصر المحددة: 0")
        ttk.Label(frm, textvariable=selected_total_var, font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(frm, textvariable=required_var, font=("Segoe UI", 10, "bold"), foreground="#0f5132").pack(anchor="w", pady=(0, 8))

        table_wrap = ttk.Frame(frm)
        table_wrap.pack(fill=tk.BOTH, expand=True, pady=(0, 8))
        tv = ttk.Treeview(table_wrap, columns=("sel", "id", "item", "school", "color", "size", "qty", "total", "paid", "remain"), show="headings", height=8)
        for col, txt, w in [("sel", "تحديد", 56), ("id", "رقم", 48), ("item", "النوع", 120), ("school", "المدرسة", 110), ("color", "اللون", 90), ("size", "المقاس", 70), ("qty", "كمية", 60), ("total", "الإجمالي", 90), ("paid", "العربون", 90), ("remain", "متبقي", 90)]:
            tv.heading(col, text=txt)
            tv.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=ysb.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

        selected_ids: Set[int] = set()
        pending_by_id = {int(r["id"]): r for r in pending_items}
        for r in pending_items:
            rid = int(r["id"])
            total = float(r.get("total_amount") or 0.0)
            paid = float(r.get("paid_amount") or 0.0)
            rem = max(0.0, float(r.get("total_amount") or 0.0) - float(r.get("paid_amount") or 0.0))
            tv.insert("", tk.END, iid=str(rid), values=("☐", group_code, r.get("item_type", ""), r.get("school", ""), r.get("color", ""), r.get("size", ""), int(r.get("qty") or 0), f"{format_money(total)}", f"{format_money(paid)}", f"{format_money(rem)}"))

        ttk.Label(frm, text="المبلغ المحصّل الآن:").pack(anchor="w")
        collect_var = tk.StringVar(value="0")
        ttk.Entry(frm, textvariable=collect_var, width=15).pack(anchor="w", pady=(0, 8))

        def _selected_delivery_required() -> Tuple[float, float]:
            selected_rows = [pending_by_id[rid] for rid in sorted(selected_ids) if rid in pending_by_id]
            selected_total = sum(float(r.get("total_amount") or 0.0) for r in selected_rows)
            selected_paid = sum(float(r.get("paid_amount") or 0.0) for r in selected_rows)
            required = max(0.0, selected_total - selected_paid)
            return selected_total, required

        def _refresh_selected_payment():
            selected_total, required = _selected_delivery_required()
            selected_total_var.set(f"إجمالي العناصر المحددة: {format_money(selected_total)}")
            required_var.set(f"المطلوب الآن: {format_money(required)}")
            collect_var.set(f"{format_money(required)}")

        def _toggle_selected(_e=None):
            cur = tv.focus() or (tv.selection()[0] if tv.selection() else None)
            if not cur:
                return
            rid = int(cur)
            vals2 = list(tv.item(cur, "values"))
            if rid in selected_ids:
                selected_ids.remove(rid)
                vals2[0] = "☐"
            else:
                selected_ids.add(rid)
                vals2[0] = "☑"
            tv.item(cur, values=vals2)
            _refresh_selected_payment()
        tv.bind("<Double-1>", _toggle_selected, add="+")

        def _confirm():
            if not selected_ids:
                messagebox.showwarning("تنبيه", "اختر عنصراً واحداً على الأقل للتسليم.", parent=dlg)
                return
            try:
                collected = _parse_money_amount(collect_var.get())
            except ValueError:
                messagebox.showerror("خطأ", "أدخل مبلغاً صحيحاً.", parent=dlg)
                return
            try:
                payment_method = PAYMENT_METHOD_CASH
                if collected > 1e-9:
                    payment_method = self._choose_payment_method(collected, "\u0637\u0631\u064a\u0642\u0629 \u062a\u062d\u0635\u064a\u0644 \u0628\u0627\u0642\u064a \u0627\u0644\u062d\u062c\u0632")
                    if not payment_method:
                        return
                summary = self.db.deliver_reservation_items(sorted(selected_ids), collected, payment_method=payment_method)
                if not summary.get("group_completed"):
                    refreshed_group = self.db.list_reservation_group_items(res_id)
                    self._print_pending_reservation_receipt(refreshed_group)
                dlg.destroy()
                msg = f"تم تسليم {int(summary.get('delivered_items') or 0)} عنصر."
                msg += " وتم إغلاق الفاتورة بالكامل." if summary.get("group_completed") else " وباقي العناصر لا تزال معلقة."
                ToastNotification.show(self.winfo_toplevel(), msg, toast_type="success")
                self._refresh()
                try:
                    ac = getattr(self.winfo_toplevel(), "_app_controller", None)
                    if ac is not None and hasattr(ac, "refresh_dashboard"):
                        ac.refresh_dashboard()
                except Exception:
                    pass
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        def _cancel_selected():
            if not selected_ids:
                messagebox.showwarning("تنبيه", "اختر عنصراً واحداً على الأقل للإلغاء.", parent=dlg)
                return
            selected_total, _required = _selected_delivery_required()
            selected_paid = sum(float(pending_by_id[rid].get("paid_amount") or 0.0) for rid in selected_ids if rid in pending_by_id)
            remaining_after_cancel = [
                r for rid, r in pending_by_id.items()
                if rid not in selected_ids and _is_reservation_active(r.get("status"))
            ]
            refund_now = 0.0 if remaining_after_cancel else selected_paid
            redistributed = selected_paid if remaining_after_cancel else 0.0
            money_line = (
                f"المبلغ الذي يجب رده للعميل الآن: {format_money(refund_now)}"
                if refund_now > 1e-9
                else f"لا يوجد رد نقدي الآن. سيتم توزيع العربون على باقي العناصر: {format_money(redistributed)}"
            )
            if not messagebox.askyesno(
                "تأكيد إلغاء بند حجز",
                "سيتم إلغاء العناصر المحددة من فاتورة الحجز.\n"
                f"إجمالي العناصر: {format_money(selected_total)}\n"
                f"العربون المخصص لها: {format_money(selected_paid)}\n\n"
                f"{money_line}\n\n"
                "هل تريد المتابعة؟",
                parent=dlg,
            ):
                return
            refund_method = PAYMENT_METHOD_CASH
            if refund_now > 1e-9:
                refund_method = self._choose_payment_method(refund_now, f"طريقة رد مبلغ {format_money(refund_now)}")
                if not refund_method:
                    return
            try:
                summary = self.db.cancel_reservation_items(
                    sorted(selected_ids),
                    reason="إلغاء بند من فاتورة الحجز",
                    refund_payment_method=refund_method,
                )
                refreshed_group = self.db.list_reservation_group_items(res_id)
                self._print_pending_reservation_receipt(refreshed_group)
                dlg.destroy()
                msg = f"تم إلغاء {int(summary.get('cancelled_items') or 0)} عنصر."
                refund_amount = float(summary.get("refund_amount") or 0.0)
                if refund_amount > 1e-9:
                    msg += f" المبلغ المطلوب رده للعميل: {format_money(refund_amount)}"
                else:
                    msg += f" لا يوجد رد نقدي؛ تم توزيع العربون: {format_money(float(summary.get('redistributed_amount') or 0.0))}"
                ToastNotification.show(self.winfo_toplevel(), msg, toast_type="success")
                self._refresh()
                try:
                    ac = getattr(self.winfo_toplevel(), "_app_controller", None)
                    if ac is not None and hasattr(ac, "refresh_dashboard"):
                        ac.refresh_dashboard()
                except Exception:
                    pass
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        action_row = ttk.Frame(frm)
        action_row.pack(fill=tk.X)
        ttk.Button(action_row, text="إلغاء العناصر المحددة", command=_cancel_selected).pack(side=tk.LEFT)
        ttk.Button(action_row, text="تأكيد التسليم", command=_confirm).pack(side=tk.RIGHT)

    def _print_reservation_receipt(self, ids, lines, customer, paid, total, customer_phone="", bill_id: Optional[Any] = None):
        clean_ids = []
        for rid in ids or []:
            try:
                clean_ids.append(int(rid))
            except Exception:
                pass
        res_num = str(bill_id or (min(clean_ids) if clean_ids else ""))
        customer_phone = _normalize_customer_phone(customer_phone)
        remaining = max(0.0, float(total or 0.0) - float(paid or 0.0))
        path = os.path.join(tempfile.gettempdir(), f"reservation_{res_num}.html")
        pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
        receipt_items = []
        for ln in lines:
            qty = int(ln.get("qty") or 0)
            price = float(ln.get("unit_price") or 0)
            receipt_items.append({
                "item_type": ln.get("item_type") or "",
                "school": ln.get("school") or "",
                "color": ln.get("color") or "",
                "size": ln.get("size") or "",
                "qty": qty,
                "unit_price": price,
                "line_total": price * qty,
            })
        receipt_bill = {
            "id": res_num,
            "created_at": now_iso(),
            "total": float(total or 0.0),
            "bill_type": "RESERVATION",
            "customer": customer,
            "customer_phone": customer_phone,
        }
        extra_rows = [
            ("المدفوع", paid),
            ("المتبقي للدفع لاحقاً", remaining),
        ]
        save_bill_as_html(
            path,
            receipt_bill,
            receipt_items,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            extra_summary_rows=extra_rows,
        )
        _print_receipt_direct_or_fallback(
            path,
            receipt_bill,
            receipt_items,
            copies=2,
            parent=self,
            pos_name=pos_name,
            support_call=support_call,
            support_whatsapp=support_whatsapp,
            extra_summary_rows=extra_rows,
        )

    def _print_pending_reservation_receipt(self, group_items: Sequence[Dict[str, Any]]):
        pending = [dict(r) for r in group_items if _is_reservation_active(r.get("status"))]
        if not pending:
            return
        bill_id = _reservation_bill_id_from_rows([dict(r) for r in group_items])
        ids = [int(r.get("id") or 0) for r in pending]
        customer = str(pending[0].get("customer") or "")
        customer_phone = _normalize_customer_phone(pending[0].get("customer_phone") or "")
        total = sum(float(r.get("total_amount") or 0.0) for r in pending)
        paid = sum(float(r.get("paid_amount") or 0.0) for r in pending)
        lines = [
            {
                "item_type": r.get("item_type") or "",
                "school": r.get("school") or "",
                "color": r.get("color") or "",
                "size": r.get("size") or "",
                "qty": int(r.get("qty") or 0),
                "unit_price": float(r.get("unit_price") or 0.0),
            }
            for r in pending
        ]
        self._print_reservation_receipt(ids, lines, customer, paid, total, customer_phone, bill_id=bill_id)

    def _print_reservations(self):
        try:
            rows_html = ""
            for iid in self._res_table.get_children():
                vals = self._res_table.item(iid, "values")
                rows_html += "<tr>" + "".join(f"<td>{_html(str(v))}</td>" for v in vals) + "</tr>"
            html = f"""<!DOCTYPE html>
<html lang='ar' dir='rtl'>
<head><meta charset='utf-8'><title>الحجوزات</title>
<style>body{{font-family:Tahoma,Arial;margin:20px}} table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #777;padding:6px;text-align:center}}</style>
</head><body>
<h2>الحجوزات</h2>
<table><thead><tr><th>الفاتورة</th><th>التاريخ</th><th>العميل</th><th>رقم العميل</th><th>النوع</th><th>المدرسة</th><th>اللون</th><th>المقاس</th><th>الكمية</th><th>الإجمالي</th><th>المدفوع</th><th>الحالة</th></tr></thead>
<tbody>{rows_html}</tbody></table>
</body></html>"""
            import tempfile, os
            path = os.path.join(tempfile.gettempdir(), "reservations.html")
            _write_html_file(path, html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)

    def _print_reservation_totals(self):
        try:
            df = self._rf_df.get() or None
            dt = self._rf_dt.get() or None
            school = self._rf_school.get() or None
            item_type = self._rf_item.get() or None
            color = self._rf_color.get() or None
            status = self._rf_status.get() or None
            rows = self.db.list_reservation_totals(
                status=status,
                date_from=df,
                date_to=dt,
                school=school,
                item_type=item_type,
                color=color,
            )
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)
            return
        if not rows:
            messagebox.showinfo("لا توجد بيانات", "لا توجد حجوزات مطابقة للطباعة.", parent=self)
            return

        from collections import defaultdict, OrderedDict

        groups = OrderedDict()
        total_qty = 0
        total_value = 0.0
        total_paid = 0.0
        for r in rows:
            sch = str(r.get("school") or "").strip()
            clr = str(r.get("color") or "").strip()
            typ = str(r.get("item_type") or "").strip()
            groups.setdefault(sch, OrderedDict())
            groups[sch].setdefault(clr, OrderedDict())
            groups[sch][clr].setdefault(typ, []).append(r)
            total_qty += int(r.get("qty") or 0)
            total_value += float(r.get("total_amount") or 0.0)
            total_paid += float(r.get("paid_amount") or 0.0)

        def size_tables_for_profile(profile, existing_sizes):
            numeric_tables = []
            alpha_labels = []
            if profile is not None:
                r1_start, r1_end, r2_start, r2_end, has_alpha = profile
                merged = merged_numeric_size_labels_from_profile(r1_start, r1_end, r2_start, r2_end)
                if merged:
                    numeric_tables.append(merged[:])
                if has_alpha:
                    alpha_labels = ALPHA_SIZES[:]
            if not numeric_tables and not alpha_labels:
                all_sizes = sorted({_normalize_size_label(str(s or "")) for s in existing_sizes if str(s or "").strip()})
                numeric = [s for s in all_sizes if s.isdigit()]
                alpha = [s for s in all_sizes if not s.isdigit()]
                if numeric:
                    numeric_tables = [numeric]
                if alpha:
                    alpha_labels = alpha
            return numeric_tables, alpha_labels

        sections = []
        for sch, color_groups in groups.items():
            for clr, item_groups in sorted(color_groups.items(), key=lambda kv: kv[0].casefold()):
                for typ, items in sorted(item_groups.items(), key=lambda kv: item_type_sort_key(kv[0])):
                    size_counts = defaultdict(int)
                    group_qty = 0
                    group_value = 0.0
                    group_paid = 0.0
                    for r in items:
                        size = _normalize_size_label(str(r.get("size") or ""))
                        qty = int(r.get("qty") or 0)
                        size_counts[size] += qty
                        group_qty += qty
                        group_value += float(r.get("total_amount") or 0.0)
                        group_paid += float(r.get("paid_amount") or 0.0)
                    profile = self.db.get_size_profile(typ, sch, clr)
                    numeric_tables, alpha_labels = size_tables_for_profile(profile, [r.get("size") for r in items])

                    def row_counts(labels):
                        out = []
                        for lbl in labels:
                            v = int(size_counts.get(_normalize_size_label(str(lbl or "")), 0))
                            out.append("" if v == 0 else str(v))
                        return out

                    def build_table(labels):
                        return (
                            "<table class='grid'><tbody>"
                            f"<tr>{''.join(f'<th>{_html(str(x))}</th>' for x in labels)}</tr>"
                            f"<tr>{''.join(f'<td class=num>{_html(v)}</td>' for v in row_counts(labels))}</tr>"
                            "</tbody></table>"
                        )

                    tables = []
                    for labels in numeric_tables:
                        for i in range(0, len(labels), 15):
                            tables.append(build_table(labels[i:i + 15]))
                    if alpha_labels:
                        tables.append("<div class='subhdr'>المقاسات بالحروف</div>" + build_table(alpha_labels))
                    header = (
                        f"<div class='hdr'><span>المدرسة: {_html(sch)}</span>"
                        f"<span>اللون: {_html(clr)}</span><span>النوع: {_html(typ)}</span></div>"
                        f"<div class='meta'>الكمية: {_num_html(group_qty)} | الإجمالي: {_money_html(group_value)} | المدفوع: {_money_html(group_paid)}</div>"
                    )
                    sections.append(f"<section class='sheet'>{header}{''.join(tables)}</section>")

        title_bits = []
        if status:
            title_bits.append(f"الحالة: {_html(status)}")
        if df or dt:
            title_bits.append(f"الفترة: {_html(df or '')} - {_html(dt or '')}")
        filters_html = (" | ".join(title_bits)) if title_bits else "كل الحجوزات المطابقة للفلاتر"
        html = f"""<!DOCTYPE html>
<html lang='ar' dir='rtl'><head><meta charset='utf-8'><title>اجمالي الحجوزات</title>
<style>@page{{size:A4;margin:12mm}}*{{box-sizing:border-box}}body{{font-family:'Segoe UI',Tahoma,Arial,'Noto Sans Arabic',sans-serif;margin:0;direction:rtl;color:#111}}
h2{{margin:0 0 6px;text-align:center}}.summary{{margin:0 0 10px;text-align:center;font-weight:600}}.sheet{{page-break-inside:avoid;margin-bottom:10mm}}
.hdr{{display:flex;justify-content:space-between;font-weight:700;margin:6px 2px 4px;gap:8px}}.meta{{margin:0 2px 8px;font-weight:600}}
.grid{{border-collapse:collapse;width:100%;table-layout:fixed;margin-bottom:6px}}.grid th,.grid td{{border:1px solid #555;padding:6px 4px;text-align:center}}
.grid th{{background:#eee}}.num{{font-variant-numeric:tabular-nums}}.subhdr{{margin:6px 0 4px;font-weight:700}}</style></head><body>
<h2>اجمالي الحجوزات</h2>
<div class='summary'>{filters_html}<br>إجمالي الكمية: {_num_html(total_qty)} | إجمالي القيمة: {_money_html(total_value)} | إجمالي المدفوع: {_money_html(total_paid)}</div>
{''.join(sections)}
<script>window.onload=function(){{try{{window.print();}}catch(e){{}}}};</script></body></html>"""
        import tempfile, os
        path = os.path.join(tempfile.gettempdir(), f"reservation_totals_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        _write_html_file(path, html)
        _print_html_auto(path, copies=1, parent=self)


# ------------------- Stock Audit Window -------------------

def _audit_key(row: Dict[str, Any]) -> Tuple[str, str, str, str]:
    return (
        str(row.get("item_type") or "").strip().lower(),
        str(row.get("school") or "").strip().lower(),
        str(row.get("color") or "").strip().lower(),
        _normalize_size_label(str(row.get("size") or "")).lower(),
    )


def _audit_rows_to_export(lines: Sequence[Dict[str, Any]]) -> List[List[Any]]:
    return [[
        r.get("item_type", ""),
        r.get("school", ""),
        r.get("color", ""),
        r.get("size", ""),
        int(r.get("expected") or 0),
        int(r.get("actual") or 0),
        int(r.get("diff") or 0),
        float(r.get("unit_price") or 0),
        float(r.get("diff_value", int(r.get("diff") or 0) * float(r.get("unit_price") or 0)) or 0),
    ] for r in lines]


def _stock_audit_report_html(report: Dict[str, Any], lines: Sequence[Dict[str, Any]]) -> str:
    total_value = float(report.get("total_value") or sum(float(r.get("diff_value") or 0) for r in lines))
    total_diff = int(report.get("total_diff") or sum(int(r.get("diff") or 0) for r in lines))
    body = []
    for r in lines:
        diff = int(r.get("diff") or 0)
        cls = "plus" if diff > 0 else "minus" if diff < 0 else ""
        body.append(
            "<tr class='%s'><td>%s</td><td>%s</td><td>%s</td><td>%s</td>"
            "<td>%s</td><td>%s</td><td>%+d</td><td>%s</td><td>%s</td></tr>" % (
                cls,
                _html(r.get("item_type", "")),
                _html(r.get("school", "")),
                _html(r.get("color", "")),
                _html(r.get("size", "")),
                int(r.get("expected") or 0),
                int(r.get("actual") or 0),
                diff,
                _html(format_money(float(r.get("unit_price") or 0))),
                _html(format_money(float(r.get("diff_value") or 0))),
            )
        )
    return """<!DOCTYPE html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>تقرير فروق الجرد</title>
<style>@page{size:A4;margin:12mm}body{font-family:'Segoe UI',Tahoma,Arial,sans-serif;direction:rtl}
h1{font-size:18px;margin:0 0 8px}.meta{display:flex;gap:24px;margin-bottom:12px;font-weight:600}
table{width:100%;border-collapse:collapse}th,td{border:1px solid #777;padding:5px;text-align:center}
th{background:#eee}.plus{background:#dcfce7}.minus{background:#fee2e2}.total{margin-top:12px;font-weight:700}</style></head>
<body><h1>تقرير فروق الجرد - POS</h1>
<div class="meta"><span>رقم التقرير: %s</span><span>التاريخ: %s</span><span>السبب: %s</span></div>
<table><thead><tr><th>النوع</th><th>المدرسة</th><th>اللون</th><th>المقاس</th><th>المتوقع</th><th>الفعلي</th><th>الفرق</th><th>السعر</th><th>قيمة الفرق</th></tr></thead>
<tbody>%s</tbody></table>
<div class="total">إجمالي الفرق: %+d &nbsp;&nbsp; إجمالي القيمة: %s</div>
<script>window.onload=function(){try{window.print();}catch(e){}}</script></body></html>""" % (
        _html(report.get("id", "")),
        _html(report.get("created_at", "")),
        _html(report.get("reason", "")),
        "".join(body),
        total_diff,
        _html(format_money(total_value)),
    )


class StockAuditWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("جرد المخزون - POS")
        self.geometry("1220x680")
        self._rows: List[Dict[str, Any]] = []
        self._touched_keys: Set[Tuple[str, str, str, str]] = set()
        self._recent_keys: Set[Tuple[str, str, str, str]] = set()
        self._verified_keys: Set[Tuple[str, str, str, str]] = set()
        self._audit_sync_job = None
        self._build()

    def _build(self):
        filters = ttk.LabelFrame(self, text="تصفية الجرد")
        filters.pack(fill=tk.X, padx=8, pady=8)

        self.f_type = LabeledCombobox(filters, "النوع", self.db, "item_type")
        self.f_school = LabeledCombobox(filters, "المدرسة", self.db, "school")
        self.f_color = LabeledCombobox(filters, "اللون", self.db, "color")

        def _constraints(exclude=None):
            d = {"item_type": self.f_type.get(), "school": self.f_school.get(), "color": self.f_color.get()}
            if exclude:
                d.pop(exclude, None)
            return d

        self.f_type.set_supplier(lambda: self.db.get_distinct_filtered("item_type", _constraints("item_type")))
        self.f_school.set_supplier(lambda: self.db.get_distinct_filtered("school", _constraints("school")))
        self.f_color.set_supplier(lambda: self.db.get_distinct_filtered("color", _constraints("color")))

        for i, w in enumerate((self.f_type, self.f_school, self.f_color)):
            w.grid(row=0, column=i, sticky="ew", padx=6, pady=6)
            filters.columnconfigure(i, weight=1)
            for ev in ("<KeyRelease>", "<<ComboboxSelected>>"):
                w.cb.bind(ev, lambda e: (self._refresh_filter_values(), self._schedule_refresh()), add="+")

        ttk.Button(filters, text="تحميل المخزون", command=self._refresh).grid(row=0, column=3, padx=8, pady=6)
        ttk.Button(filters, text="سجل تقارير الجرد", command=self._open_history).grid(row=0, column=4, padx=8, pady=6)

        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.table = ttk.Treeview(
            wrap,
            columns=("id", "type", "school", "color", "size", "expected", "actual", "diff", "price", "value"),
            show="headings",
            selectmode="browse",
        )
        for col, txt, w in [
            ("id", "ID", 70), ("type", "النوع", 150), ("school", "المدرسة", 150),
            ("color", "اللون", 140), ("size", "المقاس", 80), ("expected", "الكمية المتوقعة", 120),
            ("actual", "الكمية الفعلية", 120), ("diff", "الفرق", 90), ("price", "السعر", 100), ("value", "قيمة الفرق", 120),
        ]:
            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self.table.yview)
        xsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        _bind_mousewheel(self.table)
        _add_context_menu(self.table, self)
        self.table.tag_configure("surplus", background="#dcfce7")
        self.table.tag_configure("deficit", background="#fee2e2")
        self.table.tag_configure("verified", background="#bbf7d0")
        self.table.tag_configure("current_touched", background="#fed7aa")
        self.table.tag_configure("history_touched", background="#dbeafe")
        self.table.bind("<Double-1>", lambda e: self._edit_actual(), add="+")

        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=8, pady=(0, 8))
        ttk.Label(bar, text="أخضر: مطابق | برتقالي: تم تعديل العدد | أزرق: تم لمسه سابقاً").pack(side=tk.RIGHT, padx=8)
        ttk.Label(bar, text="الكمية الفعلية للبند المحدد:").pack(side=tk.LEFT)
        self.actual_var = tk.StringVar(value="")
        ttk.Entry(bar, textvariable=self.actual_var, width=10).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="تعيين", command=self._assign_selected_actual).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="مطابق", command=self._mark_selected_verified).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="تطبيق التسويات", command=self._apply_all_mismatches).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="تصدير تقرير الفروق", command=self._export_current_report).pack(side=tk.LEFT, padx=12)
        ttk.Button(bar, text="طباعة تقرير الفروق", command=self._print_current_report).pack(side=tk.LEFT, padx=4)
        self.summary_var = tk.StringVar(value="")
        ttk.Label(bar, textvariable=self.summary_var, font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=8)

        self._refresh()

    def _refresh_filter_values(self):
        for w in (self.f_type, self.f_school, self.f_color):
            w.refresh_values()

    def _schedule_refresh(self, delay_ms: int = 250):
        if hasattr(self, "_audit_job") and self._audit_job:
            self.after_cancel(self._audit_job)
        self._audit_job = self.after(delay_ms, self._refresh)

    def _schedule_audit_sync(self) -> None:
        if self._audit_sync_job:
            try:
                self.after_cancel(self._audit_sync_job)
            except Exception:
                pass

        def _run() -> None:
            self._audit_sync_job = None
            try:
                import sync_ui
                sync_ui.run_sync_now(self.winfo_toplevel(), self.db.conn, reason="POS stock audit")
            except Exception:
                pass

        self._audit_sync_job = self.after(1500, _run)

    def _filters(self) -> Dict[str, Any]:
        return {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "color": self.f_color.get() or None,
        }

    def _size_labels_for_group(self, item_type: str, school: str, color: str, existing_sizes: Sequence[str]) -> List[str]:
        labels: List[str] = []
        profile = self.db.get_size_profile(item_type, school, color)
        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            labels.extend(merged_numeric_size_labels_from_profile(r1s, r1e, r2s, r2e))
            if has_alpha:
                labels.extend(ALPHA_SIZES[:])
        norm_existing = [_normalize_size_label(s) for s in existing_sizes if str(s or "").strip()]
        if any(s.upper() in ALPHA_SIZES for s in norm_existing):
            for s in ALPHA_SIZES:
                if s not in labels:
                    labels.append(s)
        for s in norm_existing:
            if s not in labels:
                labels.append(s)

        def sort_key(s):
            if str(s).isdigit():
                return (0, int(s))
            try:
                return (1, ALPHA_SIZES.index(str(s).upper()))
            except ValueError:
                return (2, str(s).lower())
        return sorted(labels, key=sort_key)

    def _expanded_rows(self) -> List[Dict[str, Any]]:
        stock_rows = self.db.current_inventory(self._filters())
        groups: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
        for r in stock_rows:
            item_type = str(r.get("item_type") or "").strip()
            school = str(r.get("school") or "").strip()
            color = str(r.get("color") or "").strip()
            if not (item_type and school and color):
                continue
            key = (item_type, school, color)
            g = groups.setdefault(key, {"rows": {}, "price": float(r.get("unit_price") or 0)})
            size = _normalize_size_label(r.get("size") or "")
            if not size:
                continue
            existing = g["rows"].get(size)
            if existing is None:
                item = dict(r, size=size)
                item["stock_ids"] = [int(r.get("id"))] if r.get("id") not in (None, "") else []
                item["value"] = float(r.get("value") or (float(r.get("unit_price") or 0) * int(r.get("count") or 0)))
                g["rows"][size] = item
            else:
                existing["count"] = int(existing.get("count") or 0) + int(r.get("count") or 0)
                existing["value"] = float(existing.get("value") or 0) + float(
                    r.get("value") or (float(r.get("unit_price") or 0) * int(r.get("count") or 0))
                )
                if r.get("id") not in (None, ""):
                    rid = int(r.get("id"))
                    ids = existing.setdefault("stock_ids", [])
                    if rid not in ids:
                        ids.append(rid)
                existing["id"] = ", ".join(str(x) for x in existing.get("stock_ids") or [])
                if int(existing.get("count") or 0) > 0:
                    existing["unit_price"] = round(float(existing.get("value") or 0) / int(existing.get("count") or 1), 2)
            if float(r.get("unit_price") or 0) > 0:
                g["price"] = float(r.get("unit_price") or 0)
        filters = self._filters()
        if filters.get("item_type") and filters.get("school") and filters.get("color"):
            groups.setdefault((filters["item_type"], filters["school"], filters["color"]), {"rows": {}, "price": 0.0})

        out: List[Dict[str, Any]] = []
        for (item_type, school, color), data in sorted(groups.items(), key=lambda x: tuple(str(v).lower() for v in x[0])):
            labels = self._size_labels_for_group(item_type, school, color, list(data["rows"].keys()))
            if not labels:
                labels = list(data["rows"].keys())
            for size in labels:
                row = data["rows"].get(size)
                expected = int(row.get("count") or 0) if row else 0
                row_price = float(row.get("unit_price") or 0) if row else 0.0
                price = float(row_price or data.get("price") or self.db.get_effective_price(item_type, school, color, size) or 0)
                out.append({
                    "stock_id": row.get("id") if row else None,
                    "item_type": item_type,
                    "school": school,
                    "color": color,
                    "size": size,
                    "expected": expected,
                    "actual": expected,
                    "diff": 0,
                    "unit_price": price,
                    "diff_value": 0.0,
                })
        return out

    def _refresh(self):
        try:
            self.db.ensure_branch_catalog_stock_rows_lightweight()
            self._touched_keys = self.db.stock_audit_touched_keys()
            self._rows = self._expanded_rows()
        except Exception as ex:
            messagebox.showerror("فشل تحميل الجرد", str(ex), parent=self)
            return
        self._render()

    def _render(self):
        self.table.delete(*self.table.get_children())
        total_diff = 0
        total_value = 0.0
        for idx, r in enumerate(self._rows):
            diff = int(r.get("diff") or 0)
            value = float(diff * float(r.get("unit_price") or 0))
            r["diff_value"] = value
            total_diff += diff
            total_value += value
            key = _audit_key(r)
            tags = []
            if key in self._recent_keys:
                tags.append("current_touched")
            elif key in self._verified_keys:
                tags.append("verified")
            elif diff > 0:
                tags.append("surplus")
            elif diff < 0:
                tags.append("deficit")
            elif key in self._touched_keys:
                tags.append("history_touched")
            self.table.insert(
                "", tk.END, iid=str(idx),
                values=(
                    r.get("stock_id") or "",
                    r["item_type"], r["school"], r["color"], r["size"],
                    int(r["expected"]), int(r["actual"]), f"{diff:+d}",
                    format_money(float(r["unit_price"])), format_money(value),
                ),
                tags=tuple(tags),
            )
        visible_verified = sum(1 for r in self._rows if _audit_key(r) in self._verified_keys)
        self.summary_var.set(
            f"تم التأكد: {visible_verified}/{len(self._rows)} | "
            f"إجمالي الفرق: {total_diff:+d} | إجمالي القيمة: {format_money(total_value)}"
        )
        _apply_zebra_tags(self.table)

    def _selected_index(self) -> Optional[int]:
        sel = self.table.selection()
        if not sel:
            return None
        return int(sel[0])

    def _edit_actual(self):
        idx = self._selected_index()
        if idx is None:
            return
        current = int(self._rows[idx].get("actual") or 0)
        val = simpledialog.askinteger("الكمية الفعلية", "أدخل الكمية الفعلية:", initialvalue=current, minvalue=0, parent=self)
        if val is None:
            return
        self.actual_var.set(str(val))
        self._assign_selected_actual()

    def _assign_selected_actual(self):
        idx = self._selected_index()
        if idx is None:
            messagebox.showwarning("لم يتم التحديد", "اختر بندا من الجدول أولا.", parent=self)
            return
        try:
            actual = int(western_digits(self.actual_var.get()).strip())
        except Exception:
            messagebox.showerror("قيمة غير صالحة", "أدخل رقما صحيحا للكمية الفعلية.", parent=self)
            return
        row = self._rows[idx]
        row["actual"] = actual
        row["diff"] = int(actual) - int(row.get("expected") or 0)
        key = _audit_key(row)
        if int(row["diff"]) == 0:
            self._verified_keys.add(key)
        else:
            self._verified_keys.discard(key)
        if int(row["diff"]) != 0:
            try:
                self._recent_keys.add(key)
                report_id = self.db.append_stock_audit_report_bucket([row], reason="auto-equalization", bucket_key=now_iso()[:13])
                self.db.apply_stock_adjustments([row], note="POS stock audit auto-equalization")
                self.db.record_pos_stock_audit_applied(report_id, [row], reason="auto-equalization")
                self._schedule_audit_sync()
                ToastNotification.show(self.winfo_toplevel(), "تم حفظ التسوية في تقرير الساعة وتطبيقها.", toast_type="success")
                self._refresh()
                return
            except Exception as ex:
                messagebox.showerror("فشل تطبيق التسوية", str(ex), parent=self)
                return
        self._render()

    def _mark_selected_verified(self):
        idx = self._selected_index()
        if idx is None:
            messagebox.showwarning("لم يتم التحديد", "اختر بندا من الجدول أولا.", parent=self)
            return
        row = self._rows[idx]
        expected = int(row.get("expected") or 0)
        row["actual"] = expected
        row["diff"] = 0
        self.actual_var.set(str(expected))
        self._verified_keys.add(_audit_key(row))
        self._render()

    def _mismatch_rows(self) -> List[Dict[str, Any]]:
        return [dict(r) for r in self._rows if int(r.get("diff") or 0) != 0]

    def _apply_all_mismatches(self):
        rows = self._mismatch_rows()
        if not rows:
            messagebox.showinfo("لا توجد فروق", "لا توجد فروق لتطبيقها.", parent=self)
            return
        if not messagebox.askyesno("تطبيق التسويات", f"سيتم تطبيق {len(rows)} فرق وحفظ تقرير. هل تريد المتابعة؟", parent=self):
            return
        try:
            report_id = self.db.create_stock_audit_report(rows, reason="manual")
            self.db.apply_stock_adjustments(rows, note="POS stock audit manual equalization")
            self.db.record_pos_stock_audit_applied(report_id, rows, reason="manual")
            self._schedule_audit_sync()
            self._recent_keys.update(_audit_key(r) for r in rows)
            ToastNotification.show(self.winfo_toplevel(), "تم حفظ التقرير وتطبيق التسويات.", toast_type="success")
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل تطبيق التسويات", str(ex), parent=self)

    def _current_report_payload(self) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
        rows = self._mismatch_rows()
        report = {
            "id": "غير محفوظ",
            "created_at": now_iso(),
            "reason": "current",
            "total_diff": sum(int(r.get("diff") or 0) for r in rows),
            "total_value": sum(float(r.get("diff_value") or 0) for r in rows),
        }
        return report, rows

    def _export_current_report(self):
        report, rows = self._current_report_payload()
        if not rows:
            messagebox.showinfo("لا توجد فروق", "لا توجد فروق للتصدير.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="تصدير تقرير فروق الجرد",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"pos_stock_audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        export_to_excel(path, ["النوع", "المدرسة", "اللون", "المقاس", "المتوقع", "الفعلي", "الفرق", "السعر", "قيمة الفرق"], _audit_rows_to_export(rows))
        ToastNotification.show(self.winfo_toplevel(), f"تم تصدير التقرير إلى: {path}", toast_type="success")

    def _print_current_report(self):
        report, rows = self._current_report_payload()
        if not rows:
            messagebox.showinfo("لا توجد فروق", "لا توجد فروق للطباعة.", parent=self)
            return
        path = os.path.join(tempfile.gettempdir(), f"pos_stock_audit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        _write_html_file(path, _stock_audit_report_html(report, rows))
        _print_html_auto(path, copies=1, parent=self)

    def _open_history(self):
        StockAuditReportHistoryWindow(self, self.db)


class StockAuditReportHistoryWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("سجل تقارير الجرد - POS")
        self.geometry("1180x640")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Button(top, text="تحديث", command=self._refresh).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="طباعة المحدد", command=self._print_selected).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="تصدير المحدد", command=self._export_selected).pack(side=tk.LEFT, padx=4)

        self.reports = ttk.Treeview(
            self,
            columns=("id", "created", "reason", "count", "diff", "value"),
            show="headings",
            height=12,
        )
        for col, txt, w in [
            ("id", "رقم التقرير", 100), ("created", "التاريخ", 180), ("reason", "السبب", 160),
            ("count", "عدد الفروق", 120), ("diff", "إجمالي الفرق", 140), ("value", "إجمالي القيمة", 160),
        ]:
            self.reports.heading(col, text=txt)
            self.reports.column(col, width=w, anchor="center")
        self.reports.pack(fill=tk.BOTH, expand=True, padx=8)
        self.reports.bind("<<TreeviewSelect>>", lambda e: self._load_details(), add="+")

        ttk.Label(self, text="تفاصيل التقرير المحدد").pack(fill=tk.X, padx=8, pady=(8, 0), anchor="w")
        self.lines = ttk.Treeview(
            self,
            columns=("type", "school", "color", "size", "expected", "actual", "diff", "price", "value"),
            show="headings",
            height=8,
        )
        for col, txt, w in [
            ("type", "النوع", 150), ("school", "المدرسة", 150), ("color", "اللون", 140),
            ("size", "المقاس", 80), ("expected", "المتوقع", 100), ("actual", "الفعلي", 100),
            ("diff", "الفرق", 90), ("price", "السعر", 100), ("value", "قيمة الفرق", 120),
        ]:
            self.lines.heading(col, text=txt)
            self.lines.column(col, width=w, anchor="center")
        self.lines.tag_configure("surplus", background="#dcfce7")
        self.lines.tag_configure("deficit", background="#fee2e2")
        self.lines.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self._refresh()

    def _refresh(self):
        self.reports.delete(*self.reports.get_children())
        for r in self.db.list_stock_audit_reports():
            self.reports.insert(
                "", tk.END, iid=str(r["id"]),
                values=(r["id"], r["created_at"], r.get("reason") or "", r["diff_count"], f"{int(r['total_diff'] or 0):+d}", format_money(float(r["total_value"] or 0))),
            )
        _apply_zebra_tags(self.reports)
        self._load_details()

    def _selected_report_id(self) -> Optional[int]:
        sel = self.reports.selection()
        if not sel:
            return None
        return int(sel[0])

    def _load_details(self):
        self.lines.delete(*self.lines.get_children())
        rid = self._selected_report_id()
        if rid is None:
            return
        _report, lines = self.db.get_stock_audit_report(rid)
        for idx, r in enumerate(lines):
            diff = int(r.get("diff") or 0)
            tags = ("surplus",) if diff > 0 else ("deficit",) if diff < 0 else ()
            self.lines.insert(
                "", tk.END, iid=str(idx),
                values=(r["item_type"], r["school"], r["color"], r["size"], int(r["expected"]), int(r["actual"]), f"{diff:+d}", format_money(float(r["unit_price"] or 0)), format_money(float(r["diff_value"] or 0))),
                tags=tags,
            )
        _apply_zebra_tags(self.lines)

    def _selected_payload(self) -> Tuple[Optional[Dict[str, Any]], List[Dict[str, Any]]]:
        rid = self._selected_report_id()
        if rid is None:
            messagebox.showwarning("لم يتم التحديد", "اختر تقريرا أولا.", parent=self)
            return None, []
        return self.db.get_stock_audit_report(rid)

    def _export_selected(self):
        report, lines = self._selected_payload()
        if not report:
            return
        path = filedialog.asksaveasfilename(
            title="تصدير تقرير الجرد",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"pos_stock_audit_report_{report['id']}.xlsx",
            parent=self,
        )
        if not path:
            return
        export_to_excel(path, ["النوع", "المدرسة", "اللون", "المقاس", "المتوقع", "الفعلي", "الفرق", "السعر", "قيمة الفرق"], _audit_rows_to_export(lines))
        ToastNotification.show(self.winfo_toplevel(), f"تم تصدير التقرير إلى: {path}", toast_type="success")

    def _print_selected(self):
        report, lines = self._selected_payload()
        if not report:
            return
        path = os.path.join(tempfile.gettempdir(), f"pos_stock_audit_report_{report['id']}.html")
        _write_html_file(path, _stock_audit_report_html(report, lines))
        _print_html_auto(path, copies=1, parent=self)


# ------------------- Inventory Window -------------------

class InventoryWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("المخزون")
        self.geometry("1000x560")

        self.multi: Dict[str, List[Any]] = {
            "item_type": [], "school": [], "color": [], "size": [],
        }

        self._multi_btns: Dict[str, ttk.Button] = {}
        self._field_widgets: Dict[str, tk.Widget] = {}
        self.show_zero_var = tk.BooleanVar(value=True)
        self._inventory_ready = False
        self._inv_job = None

        self._build()

    def _build(self):
        restricted = []
        if not self.db.is_manager_feature_enabled("allow_inventory_delete"):
            restricted.append("الحذف")
        if not self.db.is_manager_feature_enabled("allow_inventory_price_edit"):
            restricted.append("السعر")
        if not self.db.is_manager_feature_enabled("allow_inventory_specs_edit"):
            restricted.append("المواصفات")
        if restricted:
            ttk.Label(
                self,
                text=f"صلاحيات مقيدة من المدير في هذه النافذة: {', '.join(restricted)}",
                foreground="#b91c1c",
                font=("Segoe UI", 10, "bold"),
            ).pack(fill=tk.X, padx=8, pady=(8, 0))

        filters = ttk.LabelFrame(self, text="تصنيف")
        filters.pack(fill=tk.X, padx=8, pady=8)

        self.f_type   = LabeledCombobox(filters, "النوع",    self.db, "item_type")
        self.f_school = LabeledCombobox(filters, "المدرسة", self.db, "school")
        self.f_color  = LabeledCombobox(filters, "اللون",    self.db, "color")
        self.f_size   = LabeledCombobox(filters, "المقاس",   self.db, "size")

        def _constraints(exclude=None):
            d = {
                "item_type": self.f_type.get(),
                "school":    self.f_school.get(),
                "color":     self.f_color.get(),
                "size":      self.f_size.get(),
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

        widgets = [self.f_type, self.f_school, self.f_color, self.f_size]
        fields  = ["item_type","school","color","size"]

        for i, (w, fld) in enumerate(zip(widgets, fields)):
            r, c = divmod(i, 3)
            w.grid(row=r*2, column=c, padx=6, pady=(6, 0), sticky="ew")
            filters.columnconfigure(c, weight=1)

            from functools import partial

            def _safe_open(field):
                try:
                    self._open_multi_dialog(field)
                except Exception as ex:
                    messagebox.showerror("خطأ", str(ex), parent=self)

            btn = ttk.Button(filters, text="اختيار متعدد...", command=partial(_safe_open, fld))
            btn.grid(row=r*2+1, column=c, sticky="w", padx=6, pady=(2, 8))
            self._multi_btns[fld] = btn

            if hasattr(w, "cb"):
                self._field_widgets[fld] = w.cb
            elif hasattr(w, "entry"):
                self._field_widgets[fld] = w.entry
            else:
                self._field_widgets[fld] = w

        for cb in (self.f_type.cb, self.f_school.cb, self.f_color.cb, self.f_size.cb):
            cb.bind("<<ComboboxSelected>>", lambda e: self._schedule_refresh(), add="+")
            cb.bind("<KeyRelease>",         lambda e: self._schedule_refresh(), add="+")

        btns = ttk.Frame(filters)
        btns.grid(row=4, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))
        ttk.Button(btns, text="بحث", command=self._refresh).pack(side=tk.LEFT)
        ttk.Button(btns, text="مسح", command=self._clear_all).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="تصدير إلى إكسل", command=self._export_excel).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="طباعة جداول المقاسات", command=self._print_size_sheets).pack(side=tk.LEFT, padx=8)
        size_ranges_btn = ttk.Button(btns, text="تعديل نطاقات المقاسات...", command=self._edit_size_ranges_dialog)
        size_ranges_btn.pack(side=tk.LEFT, padx=8)
        if not self.db.is_manager_feature_enabled("allow_size_profile_edit"):
            size_ranges_btn.configure(state="disabled")
        show_zero_cb = ttk.Checkbutton(btns, text="إظهار الكميات الصفرية", variable=self.show_zero_var, command=self._refresh)
        show_zero_cb.pack(side=tk.LEFT, padx=8)
        _add_tooltip(show_zero_cb, "إظهار أو إخفاء الأصناف التي كميتها صفر")

        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "price", "count", "value"),
            show="headings",
            selectmode="extended",
        )
        for col, txt, w in [
            ("id","المعرّف",60), ("type","النوع",140), ("school","المدرسة",160),
            ("color","اللون",80), ("size","المقاس",70),
            ("price","السعر",80), ("count","الكمية",70), ("value","القيمة",90),
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
        _bind_mousewheel(self.table)

        # Context menu
        _add_context_menu(self.table, self)

        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.sum_qty = tk.StringVar(value="0")
        self.sum_val = tk.StringVar(value="0")
        ttk.Label(bar, text="إجمالي الكمية:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_qty, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(bar, text="إجمالي القيمة:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_val, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        edit_price_btn = ttk.Button(bar, text="تعديل السعر...", command=self._edit_price_dialog)
        edit_price_btn.pack(side=tk.RIGHT)
        edit_specs_btn = ttk.Button(bar, text="تعديل المواصفات...", command=self._edit_specs_dialog)
        edit_specs_btn.pack(side=tk.RIGHT, padx=(8, 0))
        remove_btn = ttk.Button(bar, text="حذف المحدد...", command=self._remove_selected_dialog)
        remove_btn.pack(side=tk.RIGHT, padx=(8, 0))
        delete_school_btn = ttk.Button(bar, text="حذف مدرسة...", command=self._delete_school_dialog)
        delete_school_btn.pack(side=tk.RIGHT, padx=(8, 0))
        if not self.db.is_manager_feature_enabled("allow_inventory_price_edit"):
            edit_price_btn.configure(state="disabled")
        if not self.db.is_manager_feature_enabled("allow_inventory_specs_edit"):
            edit_specs_btn.configure(state="disabled")
        if not self.db.is_manager_feature_enabled("allow_inventory_delete"):
            remove_btn.configure(state="disabled")
            delete_school_btn.configure(state="disabled")

        self._inventory_ready = True
        self._refresh()

    def _get_selected_profile_keys(self):
        sel = self.table.selection()
        if not sel:
            return None
        keys = set()
        for iid in sel:
            vals = self.table.item(iid, "values")
            keys.add((vals[1], vals[2], vals[3]))
        if len(keys) != 1:
            messagebox.showwarning("تحديد غير صالح",
                "يجب أن تكون الصفوف المحددة من نفس (النوع، المدرسة، اللون).", parent=self)
            return None
        return keys.pop()

    def _edit_size_ranges_dialog(self):
        if not self.db.is_manager_feature_enabled("allow_size_profile_edit"):
            messagebox.showwarning("مقيد", _feature_restricted_message("تعديل نطاقات المقاسات من نقطة البيع غير مسموح به حالياً."), parent=self)
            return
        picked = self._get_selected_profile_keys()
        if picked:
            item_type, school, color = picked
        else:
            item_type = (self.f_type.get() or "").strip()
            school    = (self.f_school.get() or "").strip()
            color     = (self.f_color.get() or "").strip()
            if not (item_type and school and color):
                messagebox.showwarning("حدد الصنف",
                    "حدد صفوفاً من الجدول أو اختر (النوع، المدرسة، اللون) أولاً.", parent=self)
                return

        dlg = tk.Toplevel(self)
        dlg.title("تعديل نطاقات المقاسات")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=f"{item_type} / {school} / {color}", font=("Segoe UI", 10, "bold"))\
            .grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        r1_var = tk.StringVar()
        r2_var = tk.StringVar()
        has_alpha_var = tk.BooleanVar(value=False)

        ttk.Label(frm, text="النطاق الأول").grid(row=0, column=0, sticky="w")
        r1_combo = ttk.Combobox(frm, textvariable=r1_var, values=[""] + NUMERIC_RANGE_LABELS,
                                state="readonly", width=12)
        r1_combo.grid(row=0, column=1, columnspan=2, sticky="w")

        ttk.Label(frm, text="النطاق الثاني (اختياري)").grid(row=1, column=0, sticky="w")
        r2_combo = ttk.Combobox(frm, textvariable=r2_var, values=[""] + NUMERIC_RANGE_LABELS,
                                state="readonly", width=12)
        r2_combo.grid(row=1, column=1, columnspan=2, sticky="w")

        profile = self.db.get_size_profile(item_type, school, color)
        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            if r1s is not None and r1e is not None:
                r1_var.set(f"{r1s} \\u2192 {r1e}")
            if r2s is not None and r2e is not None:
                r2_var.set(f"{r2s} \\u2192 {r2e}")
            has_alpha_var.set(bool(has_alpha))

        ttk.Checkbutton(frm, text="تفعيل المقاسات بالحروف (S / M / L ...)",
                        variable=has_alpha_var).grid(row=2, column=0, columnspan=3, sticky="w")

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=4, sticky="e", pady=(10,0))

        def _parse_range_label(v: str):
            return parse_numeric_range_label(v)

        def on_save():
            try:
                r1s, r1e = _parse_range_label(r1_var.get())
                r2s, r2e = _parse_range_label(r2_var.get())
                self.db.upsert_size_profile(item_type, school, color,
                    r1_start=r1s, r1_end=r1e, r2_start=r2s, r2_end=r2e,
                    has_alpha=has_alpha_var.get())
                dlg.destroy()
                ToastNotification.show(self.winfo_toplevel(), "تم حفظ نطاقات المقاسات", toast_type="success")
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=on_save).pack(side=tk.RIGHT, padx=6)

    def _open_multi_dialog(self, field: str):
        try:
            values = self.db.get_distinct(field)
        except Exception:
            values = []

        dlg = MultiSelectDialog(self, title="اختيار متعدد", values=values,
                                preselected=[str(x) for x in self.multi[field]])
        picked = dlg.run()
        if picked is None:
            return

        self.multi[field] = [str(x) for x in picked]

        btn = self._multi_btns.get(field)
        if btn:
            btn.configure(text=f"اختيار متعدد... ({len(self.multi[field])})" if self.multi[field] else "اختيار متعدد...")

        self._enforce_single_active_field()
        self._refresh()

    def _enforce_single_active_field(self):
        active_fields = [k for k, v in self.multi.items() if v]
        if len(active_fields) >= 1:
            active = active_fields[0]
            for fld, w in self._field_widgets.items():
                is_active = (fld == active)
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

            if self.multi[active]:
                for fld in self._field_widgets:
                    if fld != active:
                        self._clear_single_field_text(fld)
        else:
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

    def _print_size_sheets(self):
        try:
            rows = self.db.current_inventory(self._filters())
        except Exception as ex:
            messagebox.showerror("فشل", str(ex), parent=self)
            return
        if not rows:
            messagebox.showinfo("لا توجد بيانات", "لا توجد صفوف مطابقة للطباعة.", parent=self)
            return
        from collections import defaultdict, OrderedDict

        school_groups = OrderedDict()
        for r in rows:
            school = (r.get("school") or "").strip()
            item   = (r.get("item_type") or "").strip()
            color  = (r.get("color") or "").strip()
            school_groups.setdefault(school, OrderedDict())
            school_groups[school].setdefault(color, OrderedDict())
            school_groups[school][color].setdefault(item, []).append(r)

        def build_size_ranges_from_profile(profile):
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
            if has_alpha:
                alpha_labels = ALPHA_SIZES[:]
            return numeric_tables, alpha_labels

        tables_html = []

        for sch, color_groups in school_groups.items():
            for clr, item_groups in sorted(color_groups.items(), key=lambda kv: kv[0].casefold()):
                for t, items in sorted(item_groups.items(), key=lambda kv: item_type_sort_key(kv[0])):
                    size_counts = defaultdict(int)
                    for r in items:
                        sz = _normalize_size_label(r.get("size") or "")
                        size_counts[sz] += int(r.get("count") or 0)

                    profile = self.db.get_size_profile(t, sch, clr)
                    numeric_tables, alpha_labels = build_size_ranges_from_profile(profile)

                    if not numeric_tables and not alpha_labels:
                        all_sizes = sorted({_normalize_size_label(r.get("size") or "") for r in items})
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

                    head = f"<div class='hdr'><span>المدرسة: {_html(sch)}</span><span>اللون: {_html(clr)}</span><span>النوع: {_html(t)}</span></div>"

                    def build_table(chunk):
                        return f"<table class='grid'><tbody><tr>{''.join(f'<th>{_html(x)}</th>' for x in chunk)}</tr><tr>{''.join(f'<td class=num>{v}</td>' for v in row_counts(chunk))}</tr><tr>{''.join('<td>&nbsp;</td>' for _ in chunk)}</tr></tbody></table>"

                    tables = []
                    for numeric_labels in numeric_tables:
                        for i in range(0, len(numeric_labels), 15):
                            tables.append(build_table(numeric_labels[i:i + 15]))
                    if alpha_labels:
                        tables.append("<div style='margin-top:6px;font-weight:600'>المقاسات بالحروف</div>" + build_table(alpha_labels))

                    tables_html.append(f'<section class="sheet">{head}{"".join(tables)}</section>')

        html = f"""<!DOCTYPE html>
<html lang='ar' dir='rtl'><head><meta charset='utf-8'><title>جداول المقاسات</title>
<style>@page{{size:A4;margin:12mm}}*{{box-sizing:border-box}}body{{font-family:'Segoe UI',Tahoma,Arial,'Noto Sans Arabic',sans-serif;margin:0;direction:rtl}}
.sheet{{page-break-inside:avoid;margin-bottom:10mm}}.hdr{{display:flex;justify-content:space-between;font-weight:600;margin:6px 2px 8px}}
.grid{{border-collapse:collapse;width:100%;table-layout:fixed;margin-bottom:6px}}.grid th,.grid td{{border:1px solid #555;padding:6px 4px;text-align:center}}.grid th{{background:#eee}}.num{{font-variant-numeric:tabular-nums}}
</style></head><body>{''.join(tables_html)}<script>window.onload=function(){{try{{window.print();}}catch(e){{}}}};</script></body></html>"""

        import tempfile, os
        path = os.path.join(tempfile.gettempdir(), f"size_sheets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        _write_html_file(path, html)
        _print_html_auto(path, copies=1, parent=self)

    def _schedule_refresh(self, delay_ms: int = 250):
        if not getattr(self, "_inventory_ready", False):
            return
        if hasattr(self, "_inv_job") and self._inv_job:
            self.after_cancel(self._inv_job)
        self._inv_job = self.after(delay_ms, self._refresh)

    def _filters(self) -> Dict[str, Any]:
        active_multi = [k for k, v in self.multi.items() if v]
        if active_multi:
            fld = active_multi[0]
            f: Dict[str, Any] = {
                "item_type": None, "school": None, "color": None, "size": None,
            }
            f[fld] = self.multi[fld][:]
            f["hide_zero"] = not bool(self.show_zero_var.get())
            return f

        return {
            "item_type": self.f_type.get() or None,
            "school": self.f_school.get() or None,
            "color": self.f_color.get() or None,
            "size": self.f_size.get() or None,
            "hide_zero": not bool(self.show_zero_var.get()),
        }

    def _refresh(self):
        if not getattr(self, "_inventory_ready", False) or not hasattr(self, "table"):
            return
        try:
            rows = self.db.current_inventory(self._filters())
            rows = self._with_profile_zero_rows(rows)
        except Exception as ex:
            messagebox.showerror("فشل البحث", str(ex), parent=self)
            return

        self.table.delete(*self.table.get_children())
        total_qty = 0
        total_value = 0.0
        for r in rows:
            tags = ("zero_stock",) if int(r.get("count") or 0) == 0 else ()
            self.table.insert(
                "", tk.END,
                values=(r["id"], r["item_type"], r["school"], r["color"], r["size"],
                        f"{format_money(float(r['unit_price']))}",
                        r["count"], f"{format_money(float(r['value']))}"),
                tags=tags,
            )
            total_qty += int(r["count"])
            total_value += float(r["value"])
        self.sum_qty.set(str(total_qty))
        self.sum_val.set(f"{format_money(total_value)}")
        _apply_zebra_tags(self.table)

    def _with_profile_zero_rows(self, rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        if not bool(self.show_zero_var.get()):
            return rows
        filters = self._filters()
        item_type = str(filters.get("item_type") or "").strip()
        school = str(filters.get("school") or "").strip()
        color = str(filters.get("color") or "").strip()
        size_filter = filters.get("size")
        if isinstance(size_filter, (list, tuple, set)):
            wanted_sizes = {_normalize_size_label(str(x or "").strip()).casefold() for x in size_filter if str(x or "").strip()}
            size_text = ""
        else:
            size_text = str(size_filter or "").strip()
            wanted_sizes = {_normalize_size_label(size_text).casefold()} if size_text else set()
        if not (item_type and school and color):
            return rows

        existing = {
            (
                str(r.get("item_type") or "").strip().casefold(),
                str(r.get("school") or "").strip().casefold(),
                str(r.get("color") or "").strip().casefold(),
                _normalize_size_label(str(r.get("size") or "").strip()).casefold(),
            )
            for r in rows
        }
        out = list(rows)
        try:
            profile_sizes = self.db.list_sizes_for_item(school, item_type, color)
        except Exception:
            profile_sizes = []
        for sr in profile_sizes:
            size = _normalize_size_label(str(sr.get("size") or "").strip())
            if not size:
                continue
            if wanted_sizes and size.casefold() not in wanted_sizes:
                continue
            key = (item_type.casefold(), school.casefold(), color.casefold(), size.casefold())
            if key in existing:
                continue
            price = sr.get("last_price")
            if price is None:
                price = self.db.get_effective_price(item_type, school, color, size) or 0
            out.append({
                "id": "",
                "item_type": item_type,
                "school": school,
                "color": color,
                "size": size,
                "unit_price": float(price or 0),
                "count": 0,
                "value": 0.0,
            })
            existing.add(key)

        def _sort_key(r: Dict[str, Any]):
            size = _normalize_size_label(str(r.get("size") or "").strip())
            if size.isdigit():
                rank = (0, int(size), "")
            elif size.upper() in ALPHA_SIZES:
                rank = (1, ALPHA_SIZES.index(size.upper()), "")
            else:
                rank = (2, 9999, size.casefold())
            return (
                str(r.get("item_type") or "").casefold(),
                str(r.get("school") or "").casefold(),
                str(r.get("color") or "").casefold(),
                rank,
                float(r.get("unit_price") or 0),
            )

        out.sort(key=_sort_key)
        return out

    def _clear_all(self):
        for w in (self.f_type, self.f_school, self.f_color, self.f_size):
            w.set("")
        for k in self.multi.keys():
            self.multi[k] = []
        for b in self._multi_btns.values():
            b.configure(text="اختيار متعدد...")
        for w in (self.f_type, self.f_school, self.f_color, self.f_size):
            w.refresh_values()
        self._enforce_single_active_field()
        self._refresh()

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
            ToastNotification.show(self.winfo_toplevel(), f"تم حفظ المخزون إلى: {path}", toast_type="success")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)

    def _remove_selected_dialog(self):
        if not self.db.is_manager_feature_enabled("allow_inventory_delete"):
            messagebox.showwarning("مقيد", _feature_restricted_message("حذف المخزون من نقطة البيع غير مسموح به حالياً."), parent=self)
            return
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("لم يتم التحديد", "اختر صفاً واحداً أو أكثر من المخزون أولاً.", parent=self)
            return

        rows = [self.table.item(i, "values") for i in sel]
        ids = [parse_int_text(r[0]) for r in rows]

        if len(rows) == 1:
            r = rows[0]
            label = f"{r[1]} / {r[2]} / {r[3]} / {r[4]}"
            available = parse_int_text(r[6])
            title = "حذف من المخزون (يتطلب كلمة مرور)"
        else:
            label = f"عدد الصفوف المحددة: {len(rows)}"
            available = sum(parse_int_text(r[6]) for r in rows)
            title = "حذف متعدد من المخزون (يتطلب كلمة مرور)"

        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=label, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(frm, text=f"إجمالي المتوفر: {available}").grid(row=1, column=0, columnspan=2, sticky="w")
        ttk.Label(frm, text="الكمية المطلوب حذفها (فارغ = الكل):").grid(row=2, column=0, sticky="e", padx=4, pady=8)
        qty_var = tk.StringVar()
        ttk.Entry(frm, textvariable=qty_var, width=12).grid(row=2, column=1, sticky="w", padx=4, pady=8)
        ttk.Label(frm, text="كلمة المرور:").grid(row=3, column=0, sticky="e", padx=4, pady=4)
        pw_var = tk.StringVar()
        ttk.Entry(frm, textvariable=pw_var, show="*").grid(row=3, column=1, sticky="w", padx=4, pady=4)

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
                    messagebox.showerror("كمية غير صالحة", "أدخل عدداً موجباً أو اتركه فارغاً للكل.", parent=dlg)
                    return
            try:
                total_removed = 0
                for stock_id in ids:
                    total_removed += self.db.remove_from_stock(stock_id, qty, note="Admin remove (bulk)")
                dlg.destroy()
                ToastNotification.show(self.winfo_toplevel(), f"تم حذف {total_removed} وحدة من {len(ids)} صف(وف)", toast_type="success")
                self._refresh()
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حذف", command=on_ok).pack(side=tk.RIGHT, padx=6)

    def _delete_school_dialog(self):
        if not self.db.is_manager_feature_enabled("allow_inventory_delete"):
            messagebox.showwarning("مقيد", _feature_restricted_message("حذف المدارس من نقطة البيع غير مسموح به حالياً."), parent=self)
            return

        selected_schools = {
            str(self.table.item(iid, "values")[2]).strip()
            for iid in self.table.selection()
            if self.table.item(iid, "values")
        }
        initial_school = ""
        if len(selected_schools) == 1:
            initial_school = next(iter(selected_schools))
        elif (self.f_school.get() or "").strip():
            initial_school = (self.f_school.get() or "").strip()

        schools = self.db.list_schools_all()
        if initial_school and initial_school not in schools:
            schools = [initial_school] + schools

        dlg = tk.Toplevel(self)
        dlg.title("حذف مدرسة من الواجهة")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(
            frm,
            text="احذف المدرسة من واجهة نقطة البيع بعد التأكد أن رصيدها أصبح صفراً.",
            justify="right",
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        ttk.Label(frm, text="المدرسة:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        school_var = tk.StringVar(value=initial_school)
        ttk.Combobox(frm, textvariable=school_var, values=schools, width=28).grid(
            row=1, column=1, sticky="ew", padx=6, pady=4
        )
        ttk.Label(frm, text="كلمة مرور المدير:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        pw_var = tk.StringVar()
        ttk.Entry(frm, textvariable=pw_var, show="*").grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        frm.columnconfigure(1, weight=1)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(10, 0))

        def on_ok():
            school = (school_var.get() or "").strip()
            if not school:
                messagebox.showerror("بيانات ناقصة", "اختر المدرسة المطلوب حذفها من الواجهة.", parent=dlg)
                return
            if not self.db.verify_admin_password(pw_var.get()):
                messagebox.showerror("مرفوض", "كلمة مرور المدير غير صحيحة.", parent=dlg)
                return
            if not messagebox.askyesno(
                "تأكيد الحذف",
                f"سيتم حذف المدرسة '{school}' من واجهة نقطة البيع.\nلن يتم ذلك إذا كان لها رصيد متبقٍ.\n\nهل تريد المتابعة؟",
                parent=dlg,
            ):
                return
            try:
                removed_rows = self.db.delete_school_from_ui(school)
                dlg.destroy()
                if school in self.multi.get("school", []):
                    self.multi["school"] = [s for s in self.multi["school"] if str(s).strip().lower() != school.lower()]
                    self._multi_btns["school"].configure(
                        text=f"اختيار متعدد... ({len(self.multi['school'])})" if self.multi["school"] else "اختيار متعدد..."
                    )
                if (self.f_school.get() or "").strip().lower() == school.lower():
                    self.f_school.set("")
                self.f_school.refresh_values()
                self._enforce_single_active_field()
                self._refresh()
                ToastNotification.show(
                    self.winfo_toplevel(),
                    f"تم حذف المدرسة من الواجهة وتنظيف {removed_rows} صف مخزون",
                    toast_type="success",
                )
            except Exception as ex:
                messagebox.showerror("فشل الحذف", str(ex), parent=dlg)

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حذف", command=on_ok).pack(side=tk.RIGHT, padx=6)

    def _edit_price_dialog(self):
        if not self.db.is_manager_feature_enabled("allow_inventory_price_edit"):
            messagebox.showwarning("مقيد", _feature_restricted_message("تعديل الأسعار من نافذة المخزون غير مسموح به حالياً."), parent=self)
            return
        row = 0
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("لم يتم التحديد", "اختر صفاً من المخزون أولاً.", parent=self)
            return

        rows = [self.table.item(i, "values") for i in sel]
        first = rows[0]

        current_price = float(first[6])
        multi = len(rows) > 1

        dlg = tk.Toplevel(self)
        dlg.title("تعديل السعر")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        if multi:
            ttk.Label(frm, text=f"عدد الصفوف المحددة: {len(rows)}", font=("Segoe UI", 10, "bold"))\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
        else:
            ttk.Label(frm, text=f"الصنف: {first[1]} / {first[2]} / {first[3]} / {first[4]}")\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
            ttk.Label(frm, text=f"السعر الحالي: {format_money(current_price)}")\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 10))
            row += 1

        ttk.Label(frm, text="السعر الجديد:").grid(row=row, column=0, sticky="e", padx=4, pady=6)
        price_var = tk.StringVar(value=f"{format_money(current_price)}")
        ttk.Entry(frm, textvariable=price_var, width=16).grid(row=row, column=1, sticky="w", padx=4, pady=6)
        row += 1

        scope_var = tk.StringVar(value="row")

        if not multi:
            scope_box = ttk.LabelFrame(frm, text="النطاق")
            scope_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 10))
            ttk.Radiobutton(scope_box, text="الصف المحدد فقط", variable=scope_var, value="row").pack(anchor="w", padx=8, pady=4)
            ttk.Radiobutton(scope_box, text="كل الصفوف بنفس (النوع/المدرسة/المقاس)",
                            variable=scope_var, value="same_specs").pack(anchor="w", padx=8, pady=4)
            row += 1

        def on_ok():
            try:
                new_price = float(price_var.get())
                if new_price < 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("قيمة غير صالحة", "أدخل سعراً رقمياً غير سالب.", parent=dlg)
                return

            try:
                updated_total = 0
                if multi:
                    for vals in rows:
                        updated_total += self.db.update_prices({"id": parse_int_text(vals[0])}, new_price, note="Price update (multi-selection)")
                else:
                    if scope_var.get() == "row":
                        updated_total = self.db.update_prices({"id": int(first[0])}, new_price, note="Price update (single row)")
                    else:
                        updated_total = self.db.update_prices({
                            "item_type": first[1], "school": first[2], "size": first[4],
                        }, new_price, note="Price update (same type/school/size)")

                dlg.destroy()
                if updated_total == 0:
                    ToastNotification.show(self.winfo_toplevel(), "لم يتم العثور على صفوف مطابقة", toast_type="warning")
                else:
                    ToastNotification.show(self.winfo_toplevel(), f"تم تحديث السعر في {updated_total} صف(وف)", toast_type="success")
                self._refresh()
            except Exception as ex:
                messagebox.showerror("فشل التحديث", str(ex), parent=dlg)

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=on_ok).pack(side=tk.RIGHT, padx=6)
        dlg.update_idletasks()

    def _edit_specs_dialog(self):
        if not self.db.is_manager_feature_enabled("allow_inventory_specs_edit"):
            messagebox.showwarning("مقيد", _feature_restricted_message("تعديل المواصفات من نقطة البيع غير مسموح به حالياً."), parent=self)
            return
        sel = self.table.selection()
        ids: List[int] = []
        if not sel:
            messagebox.showwarning("حدد النطاق",
                "اختر صفوفاً من الجدول لتطبيق التعديل.",
                parent=self)
            return
        for iid in sel:
            vals = self.table.item(iid, "values")
            ids.append(parse_int_text(vals[0]))
        scope_text = f"عدد الصفوف المحددة: {len(ids)}"

        dlg = tk.Toplevel(self)
        dlg.title("تعديل المواصفات")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=scope_text, font=("", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,8))
        ttk.Label(frm, text="اترك الحقل فارغاً إذا كنت لا تريد تغييره.").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,10))

        it_box = LabeledCombobox(frm, "النوع (جديد):", self.db, "item_type", width=32)
        sc_box = LabeledCombobox(frm, "المدرسة (جديد):", self.db, "school", width=32)
        cl_box = LabeledCombobox(frm, "اللون (جديد):", self.db, "color", width=32)
        sz_box = LabeledCombobox(frm, "المقاس (جديد):", self.db, "size", width=32)
        it_box.set_supplier(lambda: self.db.get_distinct_filtered("item_type", {}))
        sc_box.set_supplier(lambda: self.db.get_distinct_filtered("school", {}))
        cl_box.set_supplier(lambda: self.db.get_distinct_filtered("color", {}))
        sz_box.set_supplier(lambda: self.db.get_distinct_filtered("size", {}))
        for row, widget in enumerate((it_box, sc_box, cl_box, sz_box), start=2):
            widget.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=4)

        btns = ttk.Frame(frm); btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=(10,0))
        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)

        def on_ok():
            changes = {}
            if it_box.get().strip(): changes["item_type"] = it_box.get().strip()
            if sc_box.get().strip(): changes["school"]    = sc_box.get().strip()
            if cl_box.get().strip(): changes["color"]     = cl_box.get().strip()
            if sz_box.get().strip(): changes["size"]      = sz_box.get().strip()

            if not changes:
                messagebox.showwarning("لا تغييرات", "لم تُدخل أي قيم جديدة.", parent=dlg)
                return
            try:
                updated = self.db.update_specs_by_ids(ids, **changes)

                dlg.destroy()
                if updated == 0:
                    ToastNotification.show(self.winfo_toplevel(), "لم يتم العثور على صفوف مطابقة", toast_type="warning")
                else:
                    ToastNotification.show(self.winfo_toplevel(), f"تم تعديل المواصفات في {updated} صف(وف)", toast_type="success")
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

        self._values = list(dict.fromkeys(values or []))
        self._ok = False

        frm = ttk.Frame(self, padding=10)
        frm.pack(fill=tk.BOTH, expand=True)

        frow = ttk.Frame(frm)
        frow.pack(fill=tk.X, pady=(0,6))
        ttk.Label(frow, text="بحث:").pack(side=tk.LEFT)
        self.q = tk.StringVar()
        ent = ttk.Entry(frow, textvariable=self.q)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6,0))
        ent.bind("<KeyRelease>", lambda e: self._refill())

        self.listbox = tk.Listbox(frm, selectmode="extended", activestyle="none")
        self.listbox.pack(fill=tk.BOTH, expand=True)
        ysb = ttk.Scrollbar(frm, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=ysb.set)
        _bind_mousewheel(self.listbox)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(8,0))
        ttk.Button(btns, text="مسح التحديد", command=self._clear_sel).pack(side=tk.LEFT)
        ttk.Button(btns, text="إلغاء", command=self._on_cancel).pack(side=tk.RIGHT)
        ttk.Button(btns, text="موافق", command=self._on_ok).pack(side=tk.RIGHT, padx=6)

        self._refill()

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
        self.geometry("1120x560")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="سجل الفواتير", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="تحديث", command=self._refresh).pack(side=tk.RIGHT)
        ttk.Button(top, text="طباعة", command=self._print_selected).pack(side=tk.RIGHT, padx=8)
        ttk.Button(top, text="معاينة", command=self._preview_selected).pack(side=tk.RIGHT)
        ttk.Button(top, text="تصدير المحدد إلى إكسل", command=self._export_selected).pack(side=tk.RIGHT)
        ttk.Button(top, text="VOID مع سبب", command=self._void_selected).pack(side=tk.RIGHT, padx=8)
        ttk.Button(top, text="تحويل استبدال إلى بيع", command=self._convert_exchange_to_sale_selected).pack(side=tk.RIGHT, padx=8)
        ttk.Label(top, text="نوع الفاتورة:").pack(side=tk.RIGHT, padx=(16, 4))
        self._type_filter = tk.StringVar(value="الكل")
        _types_cb = ttk.Combobox(
            top, textvariable=self._type_filter,
            values=("الكل", "بيع", "مرتجع", "استبدال", "حجز", "إلى المصنع"),
            width=14, state="readonly",
        )
        _types_cb.pack(side=tk.RIGHT)
        _types_cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh())

        bills_wrap = ttk.Frame(self)
        bills_wrap.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 6))
        self.bills_table = ttk.Treeview(
            bills_wrap,
            columns=("id", "created_at", "bill_type", "customer", "phone", "total", "payment_method", "status"),
            show="headings",
            height=10,
        )
        for col, txt, w in [
            ("id", "المعرّف", 70),
            ("created_at", "التاريخ", 160),
            ("bill_type", "النوع", 90),
            ("customer", "العميل", 200),
            ("phone", "رقم العميل", 120),
            ("total", "الإجمالي", 100),
            ("payment_method", "طريقة الدفع", 100),
            ("status", "الحالة", 100),
        ]:
            self.bills_table.heading(col, text=txt)
            self.bills_table.column(col, width=w, anchor="center")
        bills_ysb = ttk.Scrollbar(bills_wrap, orient="vertical", command=self.bills_table.yview)
        bills_xsb = ttk.Scrollbar(bills_wrap, orient="horizontal", command=self.bills_table.xview)
        self.bills_table.configure(yscrollcommand=bills_ysb.set, xscrollcommand=bills_xsb.set)
        self.bills_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bills_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        bills_xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.bills_table.bind("<<TreeviewSelect>>", lambda e: self._load_items())
        _bind_mousewheel(self.bills_table)

        # Context menu on bills table
        _add_context_menu(self.bills_table, self)

        ttk.Label(self, text="بنود الفاتورة").pack(anchor="w", padx=8)

        items_wrap = ttk.Frame(self)
        items_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.items_table = ttk.Treeview(
            items_wrap,
            columns=("type", "school", "color", "size", "origin", "price", "qty", "total"),
            show="headings",
            height=12,
        )
        for col, txt, w in [
            ("type","النوع",140), ("school","المدرسة",160), ("color","اللون",80), ("size","المقاس",70),
            ("origin","المصدر",90), ("price","السعر",80),
            ("qty","الكمية",60), ("total","إجمالي",100),
        ]:
            self.items_table.heading(col, text=txt)
            self.items_table.column(col, width=w, anchor="center")
        items_ysb = ttk.Scrollbar(items_wrap, orient="vertical", command=self.items_table.yview)
        items_xsb = ttk.Scrollbar(items_wrap, orient="horizontal", command=self.items_table.xview)
        self.items_table.configure(yscrollcommand=items_ysb.set, xscrollcommand=items_xsb.set)
        self.items_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        items_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        items_xsb.pack(side=tk.BOTTOM, fill=tk.X)
        _bind_mousewheel(self.items_table)

        # Context menu on items table
        _add_context_menu(self.items_table, self)

        # Tooltip preview on bills table
        def _bill_tooltip():
            sel = self.bills_table.selection()
            if not sel:
                return ""
            try:
                items = self._selected_items()
                if not items:
                    return "لا توجد بنود"
                lines = []
                for ln in items[:6]:
                    lines.append(f"{ln['item_type']} - {ln['school']} ({ln['color']}/{ln['size']}) x{ln['qty']}")
                if len(items) > 6:
                    lines.append(f"... و {len(items) - 6} أصناف أخرى")
                return "\n".join(lines)
            except Exception:
                return ""
        ToolTip(self.bills_table, text_func=_bill_tooltip, delay=600)

        self._refresh()

    @staticmethod
    def _bill_type_ar(bt: Optional[str]) -> str:
        return _receipt_bill_type_label(bt)

    def _selected_bill_type_code(self) -> str:
        m = {
            "الكل": "",
            "بيع": "SALE",
            "مرتجع": "RETURN",
            "استبدال": "EXCHANGE",
            "حجز": "RESERVATION",
            "إلى المصنع": "WAREHOUSE_RETURN",
        }
        return m.get(self._type_filter.get() or "الكل", "")

    def _refresh(self):
        self.bills_table.delete(*self.bills_table.get_children())
        want = self._selected_bill_type_code()
        self._history_rows: Dict[str, Dict[str, Any]] = {}
        for b in self.db.list_bill_history():
            bt = str(b.get("bill_type") or "SALE").upper()
            if want and bt != want:
                continue
            key = str(b.get("history_key") or ("bill:%s" % b.get("id")))
            self._history_rows[key] = b
            status_text = str(b.get("status") or "")
            if bt != "RESERVATION":
                status_text = "ملغاة" if status_text.upper() == "VOID" else "مؤكدة"
            self.bills_table.insert(
                "", tk.END, iid=key,
                values=(
                    b.get("id", ""),
                    fmt_local_ts(b["created_at"], ""),
                    self._bill_type_ar(b.get("bill_type")),
                    b.get("customer") or "",
                    b.get("customer_phone") or "",
                    f"{format_money(float(b['total']))}",
                    _payment_method_label(str(b.get("payment_method") or PAYMENT_METHOD_CASH)),
                    status_text,
                ),
            )
        self.items_table.delete(*self.items_table.get_children())
        _apply_zebra_tags(self.bills_table)

    def _get_selected_key(self) -> Optional[str]:
        sel = self.bills_table.selection()
        if not sel:
            return None
        return str(sel[0])

    def _selected_entry(self) -> Optional[Dict[str, Any]]:
        key = self._get_selected_key()
        if key is None:
            return None
        return getattr(self, "_history_rows", {}).get(key)

    def _selected_items(self) -> List[Dict[str, Any]]:
        entry = self._selected_entry()
        if not entry:
            return []
        if str(entry.get("history_kind") or "") == "reservation":
            return self.db.list_reservation_bill_items(str(entry.get("group_key") or ""))
        return self.db.list_bill_items(int(entry["id"]))

    def _get_selected_bill_id(self) -> Optional[int]:
        entry = self._selected_entry()
        if not entry or str(entry.get("history_kind") or "") != "bill":
            return None
        return int(entry["id"])

    def _load_items(self):
        entry = self._selected_entry()
        self.items_table.delete(*self.items_table.get_children())
        if not entry:
            return
        items = self._selected_items()
        for ln in items:
            origin = str(ln.get("origin") or "").upper()
            origin_txt = "حجز" if origin == "RESERVATION" else ("من المخزون" if origin == "STOCK" else ("من المصنع" if origin == "FACTORY" else ""))
            self.items_table.insert(
                "", tk.END,
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"], origin_txt,
                        f"{format_money(float(ln['unit_price']))}",
                        ln["qty"], f"{format_money(float(ln['line_total']))}")
            )
        _apply_zebra_tags(self.items_table)

    def _export_selected(self):
        entry = self._selected_entry()
        if not entry:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.")
            return
        items = self._selected_items()
        if not items:
            messagebox.showwarning("فارغ", "لا تحتوي هذه الفاتورة على بنود.")
            return
        bill_id = str(entry.get("id") or "").strip()
        prefix = "reservation" if str(entry.get("history_kind") or "") == "reservation" else "bill"
        path = filedialog.asksaveasfilename(
            title="تصدير الفاتورة إلى إكسل",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"{prefix}_{bill_id}.xlsx",
        )
        if not path:
            return
        headers = ["type", "school", "color", "size", "origin", "unit_price", "qty", "line_total"]
        def _origin_txt(o: Optional[str]) -> str:
            o = str(o or "").upper()
            return "حجز" if o == "RESERVATION" else ("من المخزون" if o == "STOCK" else ("من المصنع" if o == "FACTORY" else ""))
        rows = [
            [
                ln["item_type"], ln["school"], ln["color"], ln["size"],
                _origin_txt(ln.get("origin")),
                float(ln["unit_price"]), int(ln["qty"]), float(ln["line_total"]),
            ]
            for ln in items
        ]
        try:
            export_to_excel(path, headers, rows)
            ToastNotification.show(self.winfo_toplevel(), f"تم تصدير الفاتورة إلى: {path}", toast_type="success")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex))

    def _print_selected(self):
        bill = self._selected_entry()
        if not bill:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.")
            return
        items = self._selected_items()
        try:
            tmp_dir = tempfile.gettempdir()
            prefix = "reservation" if str(bill.get("history_kind") or "") == "reservation" else "bill"
            path = os.path.join(tmp_dir, "%s_%s.html" % (prefix, bill.get("id")))
            pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
            save_bill_as_html(
                path,
                bill,
                items,
                pos_name=pos_name,
                support_call=support_call,
                support_whatsapp=support_whatsapp,
                shift_id=getattr(self.db, "active_shift_id", "") or "",
            )
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("فشل الطباعة", str(ex), parent=self)

    def _preview_selected(self):
        bill = self._selected_entry()
        if not bill:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.", parent=self)
            return
        items = self._selected_items()
        try:
            tmp_dir = tempfile.gettempdir()
            prefix = "reservation" if str(bill.get("history_kind") or "") == "reservation" else "bill"
            path = os.path.join(tmp_dir, "%s_preview_%s.html" % (prefix, bill.get("id")))
            pos_name, support_call, support_whatsapp = _lookup_receipt_branding(self.db.conn)
            save_bill_as_html(
                path,
                bill,
                items,
                pos_name=pos_name,
                support_call=support_call,
                support_whatsapp=support_whatsapp,
            )
            webbrowser.open_new_tab(_file_url(path))
        except Exception as ex:
            messagebox.showerror("فشل المعاينة", str(ex), parent=self)

    def _void_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("غير مناسب", "اختر فاتورة بيع عادية للإلغاء. الحجوزات تدار من شاشة الحجوزات.", parent=self)
            return
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
        except StopIteration:
            messagebox.showerror("خطأ", "لم يتم العثور على الفاتورة.", parent=self)
            return
        refund_amount = float(bill.get("total") or 0.0)
        payment_method = _payment_label(str(bill.get("payment_method") or PAYMENT_METHOD_CASH))
        pw = simpledialog.askstring("كلمة مرور المدير", "أدخل كلمة مرور المدير:", show="*", parent=self)
        if not pw:
            return
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return
        if not messagebox.askyesno(
            "تأكيد إلغاء الفاتورة",
            f"سيتم إلغاء الفاتورة #{bill_id}.\n\n"
            f"المبلغ الذي يجب رده للعميل: {format_money(refund_amount)}\n"
            f"طريقة الدفع الأصلية: {payment_method}\n\n"
            "هذا المبلغ سيتم خصمه من ملخص الوردية. هل تريد المتابعة؟",
            parent=self,
        ):
            return
        reason = simpledialog.askstring("سبب الإلغاء", "أدخل سبب الـ VOID:", parent=self)
        if reason is None:
            return
        try:
            self.db.void_bill(bill_id, reason)
            ToastNotification.show(
                self.winfo_toplevel(),
                f"تم إلغاء الفاتورة #{bill_id}. المبلغ المطلوب رده: {format_money(refund_amount)}",
                toast_type="success",
            )
            self._refresh()
        except Exception as ex:
            messagebox.showerror("فشل الإلغاء", str(ex), parent=self)

    def _convert_exchange_to_sale_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("غير مناسب", "اختر فاتورة استبدال عادية.", parent=self)
            return
        try:
            bill = next(b for b in self.db.list_bills() if int(b["id"]) == int(bill_id))
        except StopIteration:
            messagebox.showerror("خطأ", "لم يتم العثور على الفاتورة.", parent=self)
            return
        if str(bill.get("bill_type") or "SALE").upper() != "EXCHANGE":
            messagebox.showwarning("غير مناسب", "اختر فاتورة استبدال فقط.", parent=self)
            return
        if str(bill.get("status") or "").upper() == "VOID":
            messagebox.showwarning("غير مناسب", "لا يمكن تحويل فاتورة ملغاة.", parent=self)
            return

        pw = simpledialog.askstring("كلمة مرور المدير", "أدخل كلمة مرور المدير:", show="*", parent=self)
        if not pw:
            return
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return
        reason = simpledialog.askstring(
            "سبب التصحيح",
            "اكتب سبب تحويل فاتورة الاستبدال إلى بيع:",
            initialvalue="تم إدخالها كاستبدال بالخطأ",
            parent=self,
        )
        if reason is None:
            return
        if not messagebox.askyesno(
            "تأكيد التصحيح",
            "سيتم تحويل الفاتورة إلى بيع وتصحيح المخزون والتقارير. هل أنت متأكد؟",
            parent=self,
        ):
            return
        try:
            out = self.db.convert_exchange_bill_to_sale(bill_id, reason)
            ToastNotification.show(
                self.winfo_toplevel(),
                f"تم تحويل الفاتورة #{bill_id} إلى بيع - الإجمالي {format_money(out['new_total'])}",
                toast_type="success",
            )
            self._refresh()
            try:
                self.bills_table.selection_set("bill:%s" % bill_id)
                self._load_items()
            except Exception:
                pass
        except Exception as ex:
            messagebox.showerror("فشل التصحيح", str(ex), parent=self)


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
        self._build()

    def _build(self):
        top = ttk.LabelFrame(self, text="تصنيف")
        top.pack(fill=tk.X, padx=8, pady=8)

        self.customer = LabeledStaticCombo(top, "العميل", values=[""])
        self.customer.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        self.customer.cb.bind("<Button-1>", lambda e: self._reload_customers(), add="+")
        self.customer.cb.bind("<FocusIn>", lambda e: self._reload_customers(), add="+")

        self.ftype  = LabeledCombobox(top, "النوع",   self.db, "item_type");  self.ftype.grid(row=0, column=2, padx=6, pady=4, sticky="ew")
        self.fsch   = LabeledCombobox(top, "المدرسة", self.db, "school");     self.fsch.grid(row=0, column=3, padx=6, pady=4, sticky="ew")
        self.fclr   = LabeledCombobox(top, "اللون",   self.db, "color");      self.fclr.grid(row=0, column=4, padx=6, pady=4, sticky="ew")
        self.fsiz   = LabeledCombobox(top, "المقاس",  self.db, "size");       self.fsiz.grid(row=0, column=5, padx=6, pady=4, sticky="ew")

        self.df = DateField(top, "من (YYYY-MM-DD)"); self.df.grid(row=1, column=0, padx=6, pady=4, sticky="w")
        self.dt = DateField(top, "إلى");             self.dt.grid(row=1, column=1, padx=6, pady=4, sticky="w")

        ttk.Label(top, text="نص").grid(row=1, column=4, sticky="e", padx=4, pady=4)
        self.txt = ttk.Entry(top); self.txt.grid(row=1, column=5, sticky="ew", padx=6, pady=4)
        top.columnconfigure(5, weight=1)

        btns = ttk.Frame(top)
        btns.grid(row=0, column=6, rowspan=2, sticky="e", padx=6, pady=4)
        ttk.Button(btns, text="تحديث", command=self._refresh).pack(side=tk.LEFT)
        ttk.Button(btns, text="مسح", command=self._clear).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="تصدير إلى إكسل", command=self._export).pack(side=tk.LEFT)

        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0,8))
        self.table = ttk.Treeview(
            table_wrap,
            columns=("id","ts","direction","type","school","color","size","qty","note","bill_id","stock_id"),
            show="headings",
            height=14,
        )
        for col, txt, w in [
            ("id","المعرّف",70), ("ts","الوقت",170), ("direction","الاتجاه",100), ("type","النوع",150),
            ("school","المدرسة",170), ("color","اللون",100), ("size","المقاس",80), ("qty","الكمية",70),
            ("note","ملاحظة",200), ("bill_id","الفاتورة",70), ("stock_id","المخزون",70),
        ]:
            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        xsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)
        _bind_mousewheel(self.table)

        # Context menu
        _add_context_menu(self.table, self)

        sum_fr = ttk.LabelFrame(self, text="ملخص النتائج المعروضة (حسب الفلاتر الحالية)")
        sum_fr.pack(fill=tk.X, padx=8, pady=(0, 8))
        self._summary_var = tk.StringVar(value="")
        ttk.Label(
            sum_fr, textvariable=self._summary_var, justify="right", wraplength=1120,
            font=("Segoe UI", 9),
        ).pack(anchor="e", padx=10, pady=8)

        self._refresh()
        self._reload_customers()

    def _reload_customers(self):
        vals = [""] + self.db.list_customers()
        self.customer.cb["values"] = vals

    def _filters(self) -> Dict[str, Any]:
        return {
            "customer":  self.customer.get() or None,
            "item_type": self.ftype.get(),
            "school":    self.fsch.get(),
            "color":     self.fclr.get(),
            "size":      self.fsiz.get(),
            "date_from": self.df.get() or None,
            "date_to":   self.dt.get() or None,
            "text":      (self.txt.get().strip() or None),
        }

    def _refresh(self):
        try:
            rows = self.db.list_movements(self._filters())
        except Exception as ex:
            messagebox.showerror("فشل البحث", str(ex), parent=self)
            self._summary_var.set("")
            return

        self.table.delete(*self.table.get_children())
        for m in rows:
            self.table.insert(
                "", tk.END,
                values=(m.get("id"), m.get("ts"), m.get("direction"),
                        m.get("item_type",""), m.get("school",""), m.get("color",""), m.get("size",""),
                        m.get("qty"), m.get("note",""),
                        m.get("bill_id") if m.get("bill_id") is not None else "",
                        m.get("stock_id") if m.get("stock_id") is not None else "")
            )
        _apply_zebra_tags(self.table)
        s = _summarize_movement_rows(rows)
        self._summary_var.set(
            f"عدد الحركات: {s['n']}  |  "
            f"إجمالي كمية الوارد: {s['qty_in']}  |  "
            f"إجمالي كمية المنصرف (بيع/مصنع): {s['qty_out']}  |  "
            f"تسويات سلبية (كمية): {s['qty_adj_out']}  |  "
            f"مرتجعات واردة (كمية): {s['qty_return_in']}  |  "
            f"حجوزات (كمية): {s['qty_reserve']}\n"
            f"قيمة المنصرف (كمية×سعر الحركة): {format_money(s['val_out'])}  |  "
            f"تحصيل تسليم حجوزات: {format_money(s['deliver_cash'])}  |  "
            f"إجمالي الدخل (منصرف + تحصيل تسليم): {format_money(s['income_moves'])}"
        )

    def _clear(self):
        self.customer.set("")
        for w in (self.ftype, self.fsch, self.fclr, self.fsiz):
            w.set("")
        self.df.set("")
        self.dt.set("")
        self.txt.delete(0, tk.END)
        self._refresh()

    def _export(self):
        rows = self.db.list_movements(self._filters())
        if not rows:
            messagebox.showwarning("فارغ", "لا توجد صفوف للتصدير.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            title="تصدير الحركات إلى إكسل",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx"), ("Excel 97-2003 XML", "*.xls"), ("All files", "*.*")],
            initialfile=f"movements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            parent=self,
        )
        if not path:
            return
        headers = ["id","ts","direction","item_type","school","color","size","unit_price","qty","note","bill_id","stock_id"]
        table = [[
            m.get("id"), m.get("ts"), m.get("direction"),
            m.get("item_type",""), m.get("school",""), m.get("color",""), m.get("size",""),
            m.get("unit_price",""),
            m.get("qty"), m.get("note",""), m.get("bill_id"), m.get("stock_id"),
        ] for m in rows]
        try:
            export_to_excel(path, headers, table)
            ToastNotification.show(self.winfo_toplevel(), f"تم حفظ الحركات إلى: {path}", toast_type="success")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)


# ------------------- Admin Window -------------------

class AdminWindow(tk.Toplevel):
    """Admin configuration window."""

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("الإعدادات والإدارة")
        self.geometry("560x620")
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()
        self._build()

    def _build(self):
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # ---- Tab 1: Password ----
        pw_tab = ttk.Frame(nb, padding=12)
        nb.add(pw_tab, text="كلمة المرور")

        ttk.Label(pw_tab, text="تغيير كلمة المرور", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(pw_tab, text="كلمة المرور الحالية:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self._pw_cur = ttk.Entry(pw_tab, show="*")
        self._pw_cur.grid(row=1, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(pw_tab, text="كلمة المرور الجديدة:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        self._pw_new = ttk.Entry(pw_tab, show="*")
        self._pw_new.grid(row=2, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(pw_tab, text="تأكيد كلمة المرور:").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        self._pw_confirm = ttk.Entry(pw_tab, show="*")
        self._pw_confirm.grid(row=3, column=1, sticky="ew", padx=6, pady=4)

        pw_tab.columnconfigure(1, weight=1)

        pw_btns = ttk.Frame(pw_tab)
        pw_btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(pw_btns, text="حفظ كلمة المرور", command=self._change_password).pack(side=tk.RIGHT)

        # ---- Tab 2: Feature Permissions ----
        perms_tab = ttk.Frame(nb, padding=12)
        nb.add(perms_tab, text="صلاحيات المدير")

        ttk.Label(perms_tab, text="الميزات المسموحة في نقطة البيع", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(
            perms_tab,
            text="أدخل كلمة مرور المدير ثم اختر ما تريد السماح به أو تقييده.",
            justify="right",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 8))

        self._feature_vars: Dict[str, tk.BooleanVar] = {}
        feature_rows = [
            ("allow_inventory_window", "السماح بفتح نافذة المخزون"),
            ("allow_manual_incoming", "السماح بفاتورة الوارد اليدوية"),
            ("allow_bulk_price", "السماح بتعديل الأسعار الجماعي"),
            ("allow_excel_import", "السماح بالاستيراد من إكسل"),
            ("allow_manual_adjustment", "السماح بالتعديل اليدوي للكميات"),
            ("allow_reset_counts", "السماح بإعادة التعيين"),
            ("allow_inventory_delete", "السماح بحذف المخزون يدوياً"),
            ("allow_inventory_price_edit", "السماح بتعديل الأسعار من نافذة المخزون"),
            ("allow_inventory_specs_edit", "السماح بتعديل المواصفات من نافذة المخزون"),
            ("allow_size_profile_edit", "السماح بتعديل نطاقات المقاسات"),
            ("allow_stock_audit", "السماح بفتح نافذة الجرد"),
        ]
        for row_idx, (key, label) in enumerate(feature_rows, start=2):
            var = tk.BooleanVar(value=self.db.is_manager_feature_enabled(key))
            self._feature_vars[key] = var
            ttk.Checkbutton(perms_tab, text=label, variable=var).grid(
                row=row_idx, column=0, columnspan=2, sticky="w", pady=2
            )

        ttk.Label(perms_tab, text="كلمة مرور المدير:").grid(
            row=2 + len(feature_rows), column=0, sticky="e", padx=6, pady=(10, 4)
        )
        self._perm_pw = ttk.Entry(perms_tab, show="*", width=20)
        self._perm_pw.grid(row=2 + len(feature_rows), column=1, sticky="w", padx=6, pady=(10, 4))
        perms_tab.columnconfigure(1, weight=1)

        perms_btns = ttk.Frame(perms_tab)
        perms_btns.grid(row=3 + len(feature_rows), column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(perms_btns, text="تحميل الحالي", command=self._reload_feature_permissions).pack(side=tk.RIGHT)
        ttk.Button(perms_btns, text="حفظ الصلاحيات", command=self._save_feature_permissions).pack(side=tk.RIGHT, padx=6)

        # ---- Tab 3: Import ----
        receipt_tab = ttk.Frame(nb, padding=12)
        nb.add(receipt_tab, text="إيصال الطباعة")

        ttk.Label(receipt_tab, text="بيانات الفاتورة المطبوعة", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))
        ttk.Label(
            receipt_tab,
            text="اسم نقطة البيع يتم أخذه تلقائيا من اسم الجهاز الحالي.",
            justify="right",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 10))

        ttk.Label(receipt_tab, text="رقم الاتصال:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        self._receipt_call_var = tk.StringVar()
        ttk.Entry(receipt_tab, textvariable=self._receipt_call_var, width=28).grid(
            row=2, column=1, sticky="ew", padx=6, pady=4
        )

        ttk.Label(receipt_tab, text="رقم واتساب:").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        self._receipt_whatsapp_var = tk.StringVar()
        ttk.Entry(receipt_tab, textvariable=self._receipt_whatsapp_var, width=28).grid(
            row=3, column=1, sticky="ew", padx=6, pady=4
        )
        receipt_tab.columnconfigure(1, weight=1)

        receipt_btns = ttk.Frame(receipt_tab)
        receipt_btns.grid(row=4, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(receipt_btns, text="تحميل الحالي", command=self._load_receipt_settings).pack(side=tk.RIGHT)
        ttk.Button(receipt_btns, text="حفظ بيانات الإيصال", command=self._save_receipt_settings).pack(side=tk.RIGHT, padx=6)
        self._load_receipt_settings()

        imp_tab = ttk.Frame(nb, padding=12)
        nb.add(imp_tab, text="استيراد")

        ttk.Label(imp_tab, text="استيراد من ملف إكسل", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(imp_tab, text="الملف:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self._imp_path = tk.StringVar()
        ttk.Entry(imp_tab, textvariable=self._imp_path, width=40).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(imp_tab, text="استعراض...", command=self._browse_import).grid(row=2, column=1, sticky="w", padx=6, pady=2)

        imp_tab.columnconfigure(1, weight=1)

        imp_btns = ttk.Frame(imp_tab)
        imp_btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        self._import_btn = ttk.Button(imp_btns, text="استيراد", command=self._do_import)
        self._import_btn.pack(side=tk.RIGHT)

        # ---- Tab 4: Adjustment ----
        adj_tab = ttk.Frame(nb, padding=12)
        nb.add(adj_tab, text="تعديل يدوي")

        ttk.Label(adj_tab, text="تعديل الكميات يدوياً", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        fields = [
            ("النوع:", "_adj_type"),
            ("المدرسة:", "_adj_school"),
            ("اللون:", "_adj_color"),
            ("المقاس:", "_adj_size"),
        ]
        for r, (lbl, attr) in enumerate(fields, 1):
            ttk.Label(adj_tab, text=lbl).grid(row=r, column=0, sticky="e", padx=6, pady=4)
            var = ttk.Entry(adj_tab, width=24)
            var.grid(row=r, column=1, sticky="ew", padx=6, pady=4)
            setattr(self, attr, var)

        ttk.Label(adj_tab, text="الكمية (موجب=إضافة، سالب=حذف):").grid(row=5, column=0, sticky="e", padx=6, pady=4)
        self._adj_qty = ttk.Entry(adj_tab, width=12)
        self._adj_qty.grid(row=5, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(adj_tab, text="السعر:").grid(row=6, column=0, sticky="e", padx=6, pady=4)
        self._adj_price = ttk.Entry(adj_tab, width=12)
        self._adj_price.grid(row=6, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(adj_tab, text="ملاحظة:").grid(row=7, column=0, sticky="e", padx=6, pady=4)
        self._adj_note = ttk.Entry(adj_tab, width=28)
        self._adj_note.grid(row=7, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(adj_tab, text="كلمة المرور:").grid(row=8, column=0, sticky="e", padx=6, pady=4)
        self._adj_pw = ttk.Entry(adj_tab, show="*", width=16)
        self._adj_pw.grid(row=8, column=1, sticky="w", padx=6, pady=4)

        adj_tab.columnconfigure(1, weight=1)

        adj_btns = ttk.Frame(adj_tab)
        adj_btns.grid(row=9, column=0, columnspan=2, sticky="e", pady=(12, 0))
        self._adjust_btn = ttk.Button(adj_btns, text="تطبيق التعديل", command=self._do_adjustment)
        self._adjust_btn.pack(side=tk.RIGHT)

        # ---- Tab 5: Reset Counts ----
        reset_tab = ttk.Frame(nb, padding=12)
        nb.add(reset_tab, text="إعادة تعيين")

        ttk.Label(reset_tab, text="إعادة تعيين العدادات", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(reset_tab, text=(
            "تحذير: هذا الإجراء سيحذف جميع سجلات الحركات والفواتير والحجوزات\n"
            "ولا يمكن التراجع عنه. يُستخدم عادةً في بداية موسم جديد."
        ), foreground="red", justify="right").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(reset_tab, text="كلمة المرور للتأكيد:").grid(row=2, column=0, sticky="e", padx=6, pady=4)
        self._reset_pw = ttk.Entry(reset_tab, show="*", width=20)
        self._reset_pw.grid(row=2, column=1, sticky="w", padx=6, pady=4)

        reset_tab.columnconfigure(1, weight=1)

        reset_btns = ttk.Frame(reset_tab)
        reset_btns.grid(row=3, column=0, columnspan=2, sticky="e", pady=(12, 0))
        self._reset_btn = ttk.Button(reset_btns, text="إعادة التعيين", command=self._reset_counts)
        self._reset_btn.pack(side=tk.RIGHT)

        # ---- Close button ----
        ttk.Button(self, text="إغلاق", command=self.destroy).pack(side=tk.BOTTOM, pady=8)
        self._apply_feature_button_states()

    # ---- Actions ----

    def _change_password(self):
        cur = self._pw_cur.get()
        new = self._pw_new.get().strip()
        confirm = self._pw_confirm.get().strip()
        if not self.db.verify_admin_password(cur):
            messagebox.showerror("مرفوض", "كلمة المرور الحالية غير صحيحة.", parent=self)
            return
        if not new:
            messagebox.showwarning("فارغة", "كلمة المرور الجديدة لا يمكن أن تكون فارغة.", parent=self)
            return
        if new != confirm:
            messagebox.showwarning("غير متطابقة", "كلمتا المرور غير متطابقتين.", parent=self)
            return
        self.db.set_admin_password(new)
        ToastNotification.show(self.winfo_toplevel(), "تم تغيير كلمة المرور بنجاح", toast_type="success")
        self._pw_cur.delete(0, tk.END)
        self._pw_new.delete(0, tk.END)
        self._pw_confirm.delete(0, tk.END)

    def _load_receipt_settings(self):
        self._receipt_call_var.set(self.db.get_app_setting(RECEIPT_SUPPORT_CALL_SETTING, "") or "")
        self._receipt_whatsapp_var.set(self.db.get_app_setting(RECEIPT_SUPPORT_WHATSAPP_SETTING, "") or "")

    def _save_receipt_settings(self):
        call_number = western_digits(self._receipt_call_var.get().strip())
        whatsapp_number = western_digits(self._receipt_whatsapp_var.get().strip())
        self.db.set_app_setting(RECEIPT_SUPPORT_CALL_SETTING, call_number)
        self.db.set_app_setting(RECEIPT_SUPPORT_WHATSAPP_SETTING, whatsapp_number)
        self._receipt_call_var.set(call_number)
        self._receipt_whatsapp_var.set(whatsapp_number)
        ToastNotification.show(self.winfo_toplevel(), "تم حفظ بيانات الإيصال بنجاح", toast_type="success")

    def _browse_import(self):
        path = filedialog.askopenfilename(
            title="اختر ملف إكسل",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            parent=self,
        )
        if path:
            self._imp_path.set(path)

    def _reload_feature_permissions(self):
        for key, var in self._feature_vars.items():
            var.set(self.db.is_manager_feature_enabled(key))
        self._apply_feature_button_states()

    def _save_feature_permissions(self):
        pw = self._perm_pw.get()
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return
        for key, var in self._feature_vars.items():
            self.db.set_manager_feature_enabled(key, bool(var.get()))
        self._apply_feature_button_states()
        app = getattr(self.winfo_toplevel(), "_app_controller", None)
        if app is not None and hasattr(app, "_refresh_feature_states"):
            try:
                app._refresh_feature_states()
            except Exception:
                pass
        ToastNotification.show(self.winfo_toplevel(), "تم حفظ صلاحيات المدير بنجاح", toast_type="success")
        self._perm_pw.delete(0, tk.END)

    def _apply_feature_button_states(self):
        self._import_btn.configure(
            state=("normal" if self.db.is_manager_feature_enabled("allow_excel_import") else "disabled")
        )
        self._adjust_btn.configure(
            state=("normal" if self.db.is_manager_feature_enabled("allow_manual_adjustment") else "disabled")
        )
        self._reset_btn.configure(
            state=("normal" if self.db.is_manager_feature_enabled("allow_reset_counts") else "disabled")
        )

    def _do_import(self):
        if not self.db.is_manager_feature_enabled("allow_excel_import"):
            messagebox.showwarning("مقيد", _feature_restricted_message("الاستيراد إلى نقطة البيع غير مسموح به حالياً."), parent=self)
            return
        path = self._imp_path.get().strip()

        if not path:
            messagebox.showwarning("مسار ناقص", "اختر ملف إكسل أولاً.", parent=self)
            return

        try:
            count = self.db.import_from_excel(path)
            ToastNotification.show(self.winfo_toplevel(), f"تم استيراد {count} صف بنجاح", toast_type="success")
        except Exception as ex:
            messagebox.showerror("فشل الاستيراد", str(ex), parent=self)

    def _do_adjustment(self):
        if not self.db.is_manager_feature_enabled("allow_manual_adjustment"):
            messagebox.showwarning("مقيد", _feature_restricted_message("التعديل اليدوي على مخزون نقطة البيع غير مسموح به حالياً."), parent=self)
            return
        pw = self._adj_pw.get()
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return

        item_type = self._adj_type.get().strip()
        school    = self._adj_school.get().strip()
        color     = self._adj_color.get().strip()
        size      = self._adj_size.get().strip()
        qty_txt   = self._adj_qty.get().strip()
        price_txt = self._adj_price.get().strip()
        note      = self._adj_note.get().strip() or "Manual adjustment"

        if not (item_type and school and color and size):
            messagebox.showwarning("بيانات ناقصة", "أدخل النوع والمدرسة واللون والمقاس.", parent=self)
            return
        try:
            qty = int(qty_txt)
        except ValueError:
            messagebox.showerror("كمية غير صالحة", "أدخل كمية صحيحة.", parent=self)
            return
        try:
            price = float(price_txt) if price_txt else 0.0
            if price < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("سعر غير صالح", "أدخل سعراً صحيحاً.", parent=self)
            return

        try:
            self.db.manual_adjustment(
                item_type=item_type, school=school, color=color, size=size,
                qty=qty, unit_price=price, note=note
            )
            ToastNotification.show(self.winfo_toplevel(), f"تم تطبيق التعديل ({qty:+d}) على المخزون", toast_type="success")
        except Exception as ex:
            messagebox.showerror("فشل التعديل", str(ex), parent=self)

    def _reset_counts(self):
        if not self.db.is_manager_feature_enabled("allow_reset_counts"):
            messagebox.showwarning("مقيد", _feature_restricted_message("إعادة التعيين من نقطة البيع غير مسموح بها حالياً."), parent=self)
            return
        pw = self._reset_pw.get()
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return
        confirm = messagebox.askyesno(
            "تأكيد إعادة التعيين",
            "هل أنت متأكد من حذف جميع سجلات الحركات والفواتير والحجوزات؟\n"
            "هذا الإجراء لا يمكن التراجع عنه.",
            parent=self,
        )
        if not confirm:
            return
        try:
            self.db.reset_movement_counts()
            ToastNotification.show(self.winfo_toplevel(), "تم إعادة تعيين جميع العدادات بنجاح", toast_type="success")
            self._reset_pw.delete(0, tk.END)
        except Exception as ex:
            messagebox.showerror("فشل", str(ex), parent=self)


# ------------------- Shift Summary Dialog -------------------

class ExpenseDialog(tk.Toplevel):
    """Capture visual-only cashier expenses for the active shift."""
    def __init__(self, master, db: SqliteDatabase, on_saved=None):
        super().__init__(master)
        self.db = db
        self._on_saved = on_saved
        self.title("تسجيل مصروف")
        self.geometry("420x220")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self._build()

    def _build(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="تسجيل مصروف يومي", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 10))

        row1 = ttk.Frame(frm)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="المبلغ:", width=10).pack(side=tk.LEFT)
        self._amount_var = tk.StringVar()
        amount_ent = ttk.Entry(row1, textvariable=self._amount_var, width=18)
        amount_ent.pack(side=tk.LEFT, fill=tk.X, expand=True)

        row2 = ttk.Frame(frm)
        row2.pack(fill=tk.X, pady=4)
        ttk.Label(row2, text="ملاحظة:", width=10).pack(side=tk.LEFT)
        self._note_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self._note_var).pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(
            frm,
            text="المصروفات للعرض فقط ولا تغير اجمالي اليوم أو الكاش/الفيزا.",
            foreground="#6b7280",
        ).pack(anchor="w", pady=(8, 4))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=self._save).pack(side=tk.RIGHT, padx=8)

        amount_ent.focus_set()
        amount_ent.bind("<Return>", lambda _e: self._save())

    def _save(self):
        try:
            amount = _parse_money_amount(self._amount_var.get())
            self.db.add_expense(amount, self._note_var.get())
            if self._on_saved:
                self._on_saved()
            ToastNotification.show(self.master, "تم تسجيل المصروف", toast_type="success")
            self.destroy()
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)


class ShiftSummaryDialog(tk.Toplevel):
    """Show shift summary and confirm closure."""
    def __init__(self, master, db: SqliteDatabase, shift_id: int, on_closed=None):
        super().__init__(master)
        self.db = db
        self.shift_id = shift_id
        self._on_closed = on_closed
        self.title("ملخص الوردية")
        self.geometry("650x700")
        self.minsize(550, 500)
        self.resizable(True, True)
        self.transient(master)
        self.grab_set()
        self._build()

    def _build(self):
        summary = self.db.get_shift_summary(self.shift_id)

        # Scrollable container so content is never clipped
        outer = ttk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(outer, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        _bind_mousewheel(canvas)

        main = ttk.Frame(canvas, padding=10)
        main_win = canvas.create_window((0, 0), window=main, anchor="nw")

        def _on_cfg(_e=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(main_win, width=canvas.winfo_width())
        main.bind("<Configure>", _on_cfg)
        canvas.bind("<Configure>", _on_cfg)

        ttk.Label(main, text="ملخص الوردية", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))

        info = ttk.LabelFrame(main, text="معلومات الوردية")
        info.pack(fill=tk.X, pady=4)
        started = fmt_local_ts(summary["started_at"], "")
        ttk.Label(info, text=f"بداية: {started}").pack(anchor="w", padx=8, pady=2)
        ttk.Label(info, text="نهاية: الآن").pack(anchor="w", padx=8, pady=2)

        sal = ttk.LabelFrame(main, text="المبيعات")
        sal.pack(fill=tk.X, pady=4)
        ttk.Label(sal, text=f"عدد الفواتير: {summary['sales_count']}").pack(anchor="w", padx=8, pady=2)
        ttk.Label(sal, text=f"إجمالي المبيعات: {format_money(summary['sales_total'])}").pack(anchor="w", padx=8, pady=2)

        return_count = summary.get("return_count", 0)
        return_total = summary.get("return_total", 0.0)
        if return_count > 0:
            ret = ttk.LabelFrame(main, text="المرتجعات")
            ret.pack(fill=tk.X, pady=4)
            ttk.Label(ret, text=f"عدد فواتير المرتجع: {return_count}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(ret, text=f"إجمالي المرتجعات: {format_money(return_total)}").pack(anchor="w", padx=8, pady=2)

        void_count = summary.get("void_count", 0)
        void_total = summary.get("void_total", 0.0)
        if void_count > 0:
            void_fr = ttk.LabelFrame(main, text="الإلغاء")
            void_fr.pack(fill=tk.X, pady=4)
            ttk.Label(void_fr, text=f"عدد فواتير الإلغاء: {void_count}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(void_fr, text=f"إجمالي الإلغاء: {format_money(void_total)}").pack(anchor="w", padx=8, pady=2)

        exchange_count = summary.get("exchange_count", 0)
        exchange_total = summary.get("exchange_total", 0.0)
        if exchange_count > 0:
            exc = ttk.LabelFrame(main, text="الاستبدالات")
            exc.pack(fill=tk.X, pady=4)
            ttk.Label(exc, text=f"عدد فواتير الاستبدال: {exchange_count}").pack(anchor="w", padx=8, pady=2)
            if exchange_total >= 0:
                ttk.Label(exc, text=f"صافي الاستبدال (محصّل): {format_money(exchange_total)}").pack(anchor="w", padx=8, pady=2)
            else:
                ttk.Label(exc, text=f"صافي الاستبدال (مسترد): {format_money(abs(exchange_total))}").pack(anchor="w", padx=8, pady=2)

        customer_money = float(summary.get("res_paid", 0.0)) + float(summary.get("deliver_total", 0.0))
        if customer_money > 1e-9 or int(summary.get("res_count", 0) or 0) > 0 or int(summary.get("deliver_count", 0) or 0) > 0:
            cm = ttk.LabelFrame(main, text="مبالغ واردة من العملاء")
            cm.pack(fill=tk.X, pady=4)
            ttk.Label(cm, text=f"عدد الحجوزات الجديدة: {summary.get('res_count', 0)}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(cm, text=f"إجمالي الحجوزات الجديدة: {format_money(summary.get('res_total', 0.0))}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(cm, text=f"عربون الحجوزات: {format_money(summary.get('res_paid', 0.0))}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(cm, text=f"تحصيل تسليم الحجوزات: {format_money(summary.get('deliver_total', 0.0))}").pack(anchor="w", padx=8, pady=2)

        res_refund = float(summary.get("res_refund", 0.0) or 0.0)
        if res_refund > 1e-9:
            rr = ttk.LabelFrame(main, text="استرداد حجوزات")
            rr.pack(fill=tk.X, pady=4)
            ttk.Label(rr, text=f"مبلغ مردود للعميل: {format_money(res_refund)}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(rr, text="هذا المبلغ مخصوم من صافي الكاش/الفيزا حسب طريقة الرد.").pack(anchor="w", padx=8, pady=2)

        expense_total = float(summary.get("expense_total", 0.0) or 0.0)
        expense_count = int(summary.get("expense_count", 0) or 0)
        if expense_count > 0 or expense_total > 1e-9:
            ex_fr = ttk.LabelFrame(main, text="المصروفات")
            ex_fr.pack(fill=tk.X, pady=4)
            ttk.Label(ex_fr, text=f"عدد المصروفات: {expense_count}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(ex_fr, text=f"إجمالي المصروفات: {format_money(expense_total)}").pack(anchor="w", padx=8, pady=2)
            ttk.Label(ex_fr, text="للعرض فقط ولا تدخل في اجمالي اليوم أو صافي الكاش/الفيزا.").pack(anchor="w", padx=8, pady=2)

        grand = float(summary.get("cash_collected", 0.0))
        visa_total = float(summary.get("visa_collected", 0.0))
        ttk.Label(main, text=f"إجمالي الكاش: {format_money(grand)} | إجمالي الفيزا: {format_money(visa_total)}", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=8)

        btns = ttk.Frame(main)
        btns.pack(fill=tk.X, pady=(4, 0))
        ttk.Button(btns, text="طباعة الملخص", command=lambda: self._print_summary(summary)).pack(side=tk.LEFT)
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="إنهاء الوردية", command=lambda: self._confirm_close(summary)).pack(side=tk.RIGHT, padx=8)

    def _confirm_close(self, summary):
        if messagebox.askyesno("تأكيد", "هل أنت متأكد من إنهاء الوردية؟", parent=self):
            try:
                self.db.end_shift(self.shift_id, summary_json=json.dumps(summary, ensure_ascii=False, default=str))
                root = self.master
                try:
                    if self.winfo_exists():
                        self.destroy()
                except tk.TclError:
                    pass
                if self._on_closed:
                    self._on_closed()
                try:
                    if root is not None and root.winfo_exists():
                        ToastNotification.show(root, "تم إنهاء الوردية بنجاح", toast_type="success")
                except tk.TclError:
                    pass
            except Exception as ex:
                try:
                    if self.winfo_exists():
                        messagebox.showerror("خطأ", str(ex), parent=self)
                except tk.TclError:
                    pass

    def _print_summary(self, summary):
        started = fmt_local_ts(summary["started_at"], "")
        return_count = summary.get("return_count", 0)
        return_total = summary.get("return_total", 0.0)
        void_count = summary.get("void_count", 0)
        void_total = summary.get("void_total", 0.0)
        exchange_count = summary.get("exchange_count", 0)
        exchange_total = summary.get("exchange_total", 0.0)
        grand = float(summary.get("cash_collected", 0.0))
        visa_total = float(summary.get("visa_collected", 0.0))
        return_html = f'<div><b>مرتجعات:</b> {_num_html(return_count)} فاتورة - {_money_html(return_total)}</div>\n<hr class="sep">' if return_count > 0 else ""
        void_html = f'<div><b>إلغاء:</b> {_num_html(void_count)} فاتورة - {_money_html(void_total)}</div>\n<hr class="sep">' if void_count > 0 else ""
        exchange_html = f'<div><b>استبدالات:</b> {_num_html(exchange_count)} فاتورة - صافي {_money_html(exchange_total)}</div>\n<hr class="sep">' if exchange_count > 0 else ""
        customer_money = float(summary.get("res_paid", 0.0)) + float(summary.get("deliver_total", 0.0))
        customer_money_html = ""
        if customer_money > 1e-9 or int(summary.get("res_count", 0) or 0) > 0 or int(summary.get("deliver_count", 0) or 0) > 0:
            customer_money_html = (
                f'<div><b>مبالغ واردة من العملاء:</b></div>\n'
                f'<div>حجوزات جديدة: {_num_html(summary.get("res_count", 0))} - إجمالي {_money_html(summary.get("res_total", 0.0))}</div>\n'
                f'<div>عربون الحجوزات: {_money_html(summary.get("res_paid", 0.0))}</div>\n'
                f'<div>تحصيل تسليم الحجوزات: {_money_html(summary.get("deliver_total", 0.0))}</div>\n'
                f'<hr class="sep">'
            )
        res_refund_html = ""
        if float(summary.get("res_refund", 0.0) or 0.0) > 1e-9:
            res_refund_html = (
                f'<div><b>استرداد حجوزات:</b> -{_money_html(summary.get("res_refund", 0.0))}</div>\n'
                f'<hr class="sep">'
            )
        expense_html = (
            f'<div><b>مصروفات:</b> {_num_html(summary.get("expense_count", 0))} عملية - {_money_html(summary.get("expense_total", 0.0))}</div>\n'
            f'<div>للعرض فقط ولا تدخل في اجمالي اليوم.</div>\n'
            f'<hr class="sep">'
        )

        html = f"""<!DOCTYPE html>
<html lang="ar" dir="ltr"><head><meta charset="utf-8"><title>ملخص الوردية</title>
<style>
{_receipt_font_face_css()}
@page {{ size: 80mm auto; margin: 2mm; }}
body {{ font-family: {RECEIPT_FONT_STACK}; font-size: 11px; width: 76mm; min-height: 0; direction: ltr; margin: 0; padding: 2mm; }}
.receipt {{ direction: rtl; unicode-bidi: isolate; min-height: auto; height: auto; max-height: none; overflow: visible; }}
.digits {{ direction: ltr; unicode-bidi: isolate; font-family: Tahoma, Arial, "Segoe UI", sans-serif; }}
h2 {{ font-size: 14px; text-align: center; margin: 4px 0; }}
.sep {{ border: none; border-top: 1px dashed #000; margin: 4px 0; }}
table {{ border-collapse: collapse; width: 100%; font-size: 10px; }}
th {{ border-bottom: 1px solid #000; padding: 2px; text-align: right; }}
td {{ padding: 2px; border-bottom: 1px dotted #ccc; }}
.total {{ font-size: 13px; font-weight: bold; text-align: center; margin: 6px 0; }}
</style></head><body><div class="receipt">
<h2>ملخص الوردية #{_num_html(summary['shift_id'])}</h2>
<hr class="sep">
<div>بداية: {_num_html(started)}</div>
<div>نهاية: الآن</div>
<hr class="sep">
<div><b>اجمالي اليوم الصافي:</b> {_money_html(float(summary.get('cash_collected') or 0.0) + float(summary.get('visa_collected') or 0.0))}</div>
<hr class="sep">
{return_html}{void_html}{exchange_html}{customer_money_html}{res_refund_html}{expense_html}<div class="total">كاش صافي: {_money_html(grand)} | فيزا صافي: {_money_html(visa_total)}</div>
</div></body></html>"""

        tmp = os.path.join(tempfile.gettempdir(), f"shift_{summary['shift_id']}.html")
        _write_html_file(tmp, html)
        _print_html_auto(tmp, copies=1, parent=self)


class IncomingShipmentsFrame(ttk.Frame):
    """Manual queue for incoming warehouse shipments waiting for cashier confirmation."""
    def __init__(self, master, db: SqliteDatabase, app):
        super().__init__(master, padding=10)
        self.db = db
        self.app = app
        self._rows: List[Dict[str, Any]] = []
        self._build()
        self._refresh()

    def _build(self):
        ttk.Label(self, text="شحنات الفواتير غير المؤكدة", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 8))
        ttk.Label(
            self,
            text="افتح الشحنة عندما تصل المنتجات فعلياً ثم راجع الكميات بنداً بنداً قبل التأكيد.",
            foreground="#6b7280",
        ).pack(anchor="w", pady=(0, 8))

        top = ttk.Frame(self)
        top.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(top, text="تحديث", command=self._refresh).pack(side=tk.LEFT)
        ttk.Button(top, text="فتح الشحنة المحددة", command=self._open_selected).pack(side=tk.LEFT, padx=8)

        self._summary_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self._summary_var).pack(side=tk.RIGHT)

        wrap = ttk.Frame(self)
        wrap.pack(fill=tk.BOTH, expand=True)
        cols = ("shipment", "source", "lines", "qty", "created", "note")
        self._tree = ttk.Treeview(wrap, columns=cols, show="headings", height=14)
        for col, txt, w in [
            ("shipment", "مرجع الشحنة", 130),
            ("source", "المرسل", 120),
            ("lines", "عدد البنود", 90),
            ("qty", "إجمالي الكمية", 100),
            ("created", "تاريخ الإنشاء", 160),
            ("note", "ملاحظة", 260),
        ]:
            self._tree.heading(col, text=txt)
            self._tree.column(col, width=w, anchor="center" if col != "note" else "w")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=ysb.set)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        self._tree.bind("<Double-1>", lambda _e: self._open_selected(), add="+")

    def _selected_row(self) -> Optional[Dict[str, Any]]:
        sel = self._tree.selection()
        if not sel:
            return None
        iid = sel[0]
        try:
            idx = int(iid)
        except Exception:
            return None
        if 0 <= idx < len(self._rows):
            return self._rows[idx]
        return None

    def _open_selected(self):
        row = self._selected_row()
        if not row:
            messagebox.showwarning("اختر شحنة", "اختر شحنة غير مؤكدة أولاً.", parent=self.winfo_toplevel())
            return
        alert_id = row.get("id")
        shipment_uuid = str(row.get("shipment_uuid") or "").strip()
        if shipment_uuid in getattr(self.app, "_open_incoming_shipments", set()):
            messagebox.showinfo("مفتوحة بالفعل", "هذه الشحنة مفتوحة حالياً في نافذة مراجعة أخرى.", parent=self.winfo_toplevel())
            return
        try:
            if alert_id is not None:
                self.db.mark_incoming_shipment_alert_shown(int(alert_id))
        except Exception:
            pass
        opened = self.app._open_incoming_shipment_checklist(row)
        if opened is False and alert_id is not None:
            try:
                self.db.reset_incoming_shipment_alert_shown(int(alert_id))
            except Exception:
                pass
        self.after(300, self._refresh)

    def _refresh(self):
        self._rows = self.db.list_pending_incoming_shipments()
        self._tree.delete(*self._tree.get_children())
        total_qty = 0
        for idx, row in enumerate(self._rows):
            qty = int(row.get("pending_qty") or 0)
            total_qty += qty
            ship = str(row.get("shipment_uuid") or "").strip()
            self._tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(
                    ship[:8] if ship else "-",
                    row.get("from_device") or "WAREHOUSE",
                    int(row.get("pending_lines") or 0),
                    qty,
                    str(row.get("created_at") or "").replace("T", " "),
                    row.get("note") or "",
                ),
            )
        try:
            _apply_zebra_tags(self._tree)
        except Exception:
            pass
        self._summary_var.set(f"عدد الشحنات المعلقة: {len(self._rows)} | إجمالي القطع: {total_qty}")
        try:
            self.app._update_incoming_shipments_badge(len(self._rows))
        except Exception:
            pass


# ------------------- Bulk Price Dialog -------------------

class BulkPriceDialog(tk.Toplevel):
    """Dialog for bulk price update by percentage or fixed value."""
    def __init__(self, master, db: SqliteDatabase, on_done=None):
        super().__init__(master)
        self.db = db
        self._on_done = on_done
        self.title("تعديل الأسعار جماعياً")
        self.geometry("450x350")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self._build()

    def _build(self):
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="تعديل الأسعار جماعياً", font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 8))

        # Mode selection
        mode_frm = ttk.LabelFrame(frm, text="طريقة التعديل")
        mode_frm.pack(fill=tk.X, pady=4)
        self._mode_var = tk.StringVar(value="percentage")
        ttk.Radiobutton(mode_frm, text="نسبة مئوية (يتم التقريب لأقرب 5)", variable=self._mode_var, value="percentage").pack(anchor="w", padx=8, pady=2)
        ttk.Radiobutton(mode_frm, text="قيمة ثابتة (تُضاف للسعر الحالي)", variable=self._mode_var, value="fixed").pack(anchor="w", padx=8, pady=2)

        val_frm = ttk.Frame(frm)
        val_frm.pack(fill=tk.X, pady=4)
        ttk.Label(val_frm, text="القيمة:").pack(side=tk.LEFT)
        self._value_var = tk.StringVar(value="10")
        ttk.Entry(val_frm, textvariable=self._value_var, width=12).pack(side=tk.LEFT, padx=8)
        ttk.Label(val_frm, text="(نسبة % أو مبلغ حسب الاختيار)").pack(side=tk.LEFT)

        # Optional filters
        filt = ttk.LabelFrame(frm, text="تطبيق على (اتركها فارغة لتعديل الكل)")
        filt.pack(fill=tk.X, pady=4)
        self._f_school = tk.StringVar()
        self._f_item = tk.StringVar()
        r0 = ttk.Frame(filt)
        r0.pack(fill=tk.X, padx=8, pady=2)
        ttk.Label(r0, text="المدرسة:").pack(side=tk.LEFT)
        ttk.Entry(r0, textvariable=self._f_school, width=20).pack(side=tk.LEFT, padx=4)
        ttk.Label(r0, text="النوع:").pack(side=tk.LEFT, padx=(12, 0))
        ttk.Entry(r0, textvariable=self._f_item, width=20).pack(side=tk.LEFT, padx=4)

        # Preview
        ttk.Label(frm, text="مثال: سعر 231 بنسبة 10% = 255 (231×1.1=254.1 → 255)", font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))

        # Buttons
        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(btns, text="تطبيق", command=self._apply).pack(side=tk.LEFT)
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.RIGHT)

    def _apply(self):
        try:
            value = float(self._value_var.get().strip())
        except ValueError:
            messagebox.showwarning("خطأ", "أدخل قيمة رقمية صحيحة.", parent=self)
            return

        mode = self._mode_var.get()
        constraints: Dict[str, Any] = {}
        sc = self._f_school.get().strip()
        it = self._f_item.get().strip()
        if sc:
            constraints["school"] = sc
        if it:
            constraints["item_type"] = it

        scope_text = "جميع الأصناف" if not constraints else f"الأصناف المحددة ({', '.join(constraints.values())})"
        if mode == "percentage":
            desc = f"زيادة {value}% على {scope_text}"
        else:
            desc = f"إضافة {value} على {scope_text}"

        if not messagebox.askyesno("تأكيد", f"هل تريد تطبيق: {desc}؟", parent=self):
            return

        try:
            count = self.db.bulk_update_prices(mode, value, constraints if constraints else None)
            ToastNotification.show(self.winfo_toplevel(), f"تم تعديل أسعار {count} صنف بنجاح", toast_type="success")
            if self._on_done:
                self._on_done()
            self.destroy()
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)


# ------------------- Main Application -------------------

class WarehouseApp:
    """Main application controller."""

    def __init__(self, root: tk.Tk, db: SqliteDatabase):
        self.root = root
        self.db = db
        self._current_shift_id: Optional[int] = None
        self._open_incoming_shipments: Set[str] = set()
        self.root._app_controller = self
        root.title(f"{APP_TITLE} - v{APP_VERSION}")
        root.geometry("1280x760")
        root.minsize(900, 600)
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._check_shift_state()
        self._bind_shortcuts()
        self._update_clock()
        self.root.after(1500, self._poll_reservation_alerts)
        self.root.after(2200, self._poll_incoming_shipment_alerts)
        try:
            import sync_periodic

            sync_periodic.attach_periodic_sync(self.root, self.db.path)
        except Exception:
            pass
        try:
            import pos_db_backup

            pos_db_backup.attach_hourly_db_backup(self.root, self.db.path)
        except Exception:
            pass

    def _build(self):
        # ---- Branded Header Bar ----
        T = getattr(self.root, "_theme", THEME_LIGHT)
        self._header = tk.Frame(self.root, bg=T["HEADER_BG"], height=48)
        self._header.pack(fill=tk.X, side=tk.TOP)
        self._header.pack_propagate(False)

        tk.Label(self._header, text=f"{APP_TITLE}  v{APP_VERSION}",
                 bg=T["HEADER_BG"], fg=T["HEADER_FG"],
                 font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT, padx=16)

        # Shift status in header
        self._shift_status_var = tk.StringVar(value="")
        tk.Label(self._header, textvariable=self._shift_status_var,
                 bg=T["HEADER_BG"], fg="#93c5fd",
                 font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=20)

        # Clock on right
        self._clock_var = tk.StringVar(value="")
        tk.Label(self._header, textvariable=self._clock_var,
                 bg=T["HEADER_BG"], fg=T["HEADER_FG"],
                 font=("Segoe UI", 11)).pack(side=tk.RIGHT, padx=16)

        # ---- Toolbar ----
        self._toolbar = ttk.Frame(self.root)
        self._toolbar.pack(fill=tk.X, side=tk.TOP, padx=6, pady=(4, 0))

        # Old fully-disabled toolbar buttons were commented during the first lockdown pass.
        self._inventory_btn = ttk.Button(self._toolbar, text="\u0627\u0644\u0645\u062E\u0632\u0648\u0646", command=self._open_inventory)
        self._inventory_btn.pack(side=tk.LEFT, padx=2)
        self._stock_audit_btn = ttk.Button(self._toolbar, text="جرد", command=self._open_stock_audit)
        self._stock_audit_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(self._toolbar, text="\u0633\u062C\u0644 \u0627\u0644\u0641\u0648\u0627\u062A\u064A\u0631", command=self._open_bills_history).pack(side=tk.LEFT, padx=2)
        ttk.Button(self._toolbar, text="\u0633\u062C\u0644 \u0627\u0644\u062D\u0631\u0643\u0627\u062A", command=self._open_movements).pack(side=tk.LEFT, padx=2)
        self._bulk_price_btn = ttk.Button(self._toolbar, text="\u062A\u0639\u062F\u064A\u0644 \u0627\u0644\u0623\u0633\u0639\u0627\u0631", command=self._open_bulk_price)
        self._bulk_price_btn.pack(side=tk.LEFT, padx=2)

        # Right side toolbar
        ttk.Button(self._toolbar, text="\u0627\u0644\u0625\u0639\u062F\u0627\u062F\u0627\u062A", command=self._open_admin).pack(side=tk.RIGHT, padx=2)
        ttk.Button(self._toolbar, text="\u0627\u0644\u0645\u0632\u0627\u0645\u0646\u0629", command=self._open_sync_dialog).pack(side=tk.RIGHT, padx=2)

        # Dark mode toggle
        self._dark_mode_var = tk.BooleanVar(value=False)
        self._dark_btn = ttk.Button(self._toolbar, text="\u263E \u0627\u0644\u0648\u0636\u0639 \u0627\u0644\u0644\u064A\u0644\u064A",
                                    command=self._toggle_dark_mode, style="Secondary.TButton")
        self._dark_btn.pack(side=tk.RIGHT, padx=4)

        # Shift button
        self._shift_btn = ttk.Button(self._toolbar, text="\u0628\u062F\u0621 \u0648\u0631\u062F\u064A\u0629", command=self._toggle_shift)
        self._shift_btn.pack(side=tk.LEFT, padx=(16, 2))
        self._expense_btn = ttk.Button(self._toolbar, text="مصروفات", command=self._open_expense_dialog)
        self._expense_btn.pack(side=tk.LEFT, padx=2)

        # ---- Main notebook ----
        self._nb = ttk.Notebook(self.root)
        self._nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=(4, 0))

        # Tab 0: Dashboard (Home)
        dash_tab = ttk.Frame(self._nb)
        self._nb.add(dash_tab, text="\u2302 \u0627\u0644\u0631\u0626\u064A\u0633\u064A\u0629")
        self._dashboard = DashboardFrame(dash_tab, self.db)
        self._dashboard.pack(fill=tk.BOTH, expand=True)

        # Tab 1: Income
        income_tab = ttk.Frame(self._nb)
        self._nb.add(income_tab, text="\u0627\u0644\u0648\u0627\u0631\u062F")
        self._income_frame = IncomeFrame(income_tab, self.db)
        self._income_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 2: POS
        pos_tab = ttk.Frame(self._nb)
        self._nb.add(pos_tab, text="\u0646\u0642\u0637\u0629 \u0627\u0644\u0628\u064A\u0639")
        self._pos_frame = POSFrame(pos_tab, self.db)
        self._pos_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 3: Reservations
        res_tab = ttk.Frame(self._nb)
        self._nb.add(res_tab, text="\u0627\u0644\u062D\u062C\u0648\u0632\u0627\u062A")
        self._res_frame = ReservationsFrame(res_tab, self.db)
        self._res_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 4: Statistics
        stats_tab = ttk.Frame(self._nb)
        self._nb.add(stats_tab, text="\u0627\u0644\u0625\u062D\u0635\u0627\u0626\u064A\u0627\u062A")
        self._stats_frame = StatisticsFrame(stats_tab, self.db)
        self._stats_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 5: Incoming shipments waiting for manual receipt confirmation
        incoming_ship_tab = ttk.Frame(self._nb)
        self._incoming_ship_tab = incoming_ship_tab
        self._incoming_shipments_base_tab_text = "شحنات غير مؤكدة"
        self._nb.add(incoming_ship_tab, text=self._incoming_shipments_base_tab_text)
        self._incoming_shipments_frame = IncomingShipmentsFrame(incoming_ship_tab, self.db, self)
        self._incoming_shipments_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 6: Shifts Summary
        shifts_tab = ttk.Frame(self._nb)
        self._nb.add(shifts_tab, text="ملخص الورديات")
        self._shifts_frame = ShiftsSummaryFrame(shifts_tab, self.db)
        self._shifts_frame.pack(fill=tk.BOTH, expand=True)

        school_accounts_tab = ttk.Frame(self._nb)
        self._nb.add(school_accounts_tab, text="حسابات المدارس")
        self._school_accounts_frame = SchoolAccountsFrame(school_accounts_tab, self.db)
        self._school_accounts_frame.pack(fill=tk.BOTH, expand=True)

        self._nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._refresh_feature_states()

        # ---- Status Bar ----
        self._statusbar = tk.Frame(self.root, bg=T.get("EDGE", "#cbd5e1"), height=26)
        self._statusbar.pack(fill=tk.X, side=tk.BOTTOM)
        self._statusbar.pack_propagate(False)
        self._status_var = tk.StringVar(value="\u062C\u0627\u0647\u0632")
        tk.Label(self._statusbar, textvariable=self._status_var,
                 bg=T.get("EDGE", "#cbd5e1"), fg=T.get("TEXT", "#0f172a"),
                 font=("Segoe UI", 8), anchor="w").pack(side=tk.LEFT, padx=8)
        self._status_items_var = tk.StringVar(value="")
        tk.Label(self._statusbar, textvariable=self._status_items_var,
                 bg=T.get("EDGE", "#cbd5e1"), fg=T.get("TEXT", "#0f172a"),
                 font=("Segoe UI", 8), anchor="e").pack(side=tk.RIGHT, padx=8)

        try:
            self._dashboard.refresh()
        except Exception:
            pass

    def refresh_dashboard(self):
        try:
            self._dashboard.refresh()
        except Exception:
            pass

    def _update_incoming_shipments_badge(self, pending_count: Optional[int] = None):
        try:
            if pending_count is None:
                pending_count = len(self.db.list_pending_incoming_shipments())
            tab_text = self._incoming_shipments_base_tab_text
            if int(pending_count or 0) > 0:
                tab_text = f"! {tab_text} ({int(pending_count)})"
            self._nb.tab(self._incoming_ship_tab, text=tab_text)
        except Exception:
            pass

    def _open_expense_dialog(self):
        if not self._current_shift_id:
            messagebox.showwarning("تنبيه", "يجب فتح وردية قبل تسجيل المصروفات.", parent=self.root)
            return

        def _after_saved():
            try:
                self._shifts_frame._refresh_all()
            except Exception:
                pass

        ExpenseDialog(self.root, self.db, on_saved=_after_saved)

    def _update_status(self, msg):
        self._status_var.set(msg)
        self.root.after(5000, lambda: self._status_var.set("\u062C\u0627\u0647\u0632"))

    def _update_clock(self):
        try:
            now = datetime.now()
            self._clock_var.set(now.strftime("%Y-%m-%d  %H:%M:%S"))
        except Exception:
            pass
        self.root.after(1000, self._update_clock)

    def _poll_reservation_alerts(self):
        try:
            alert = self.db.get_next_reservation_alert()
            if alert:
                hold_txt = str(alert.get("hold_until") or "").strip()
                msg = (
                    "طلب حجز جديد من الكول سنتر:\n\n"
                    f"العميل: {alert.get('customer')}\n"
                    f"الصنف: {alert.get('item_type')} | {alert.get('school')} | {alert.get('color')} | {alert.get('size')}\n"
                    f"الكمية: {alert.get('qty')}"
                )
                if hold_txt:
                    msg += f"\nمدة الحجز حتى: {hold_txt}"
                note = str(alert.get("note") or "").strip()
                if note:
                    msg += f"\nملاحظة: {note}"
                messagebox.showinfo("طلب حجز جديد", msg, parent=self.root)
                self.db.mark_reservation_alert_shown(int(alert["id"]))
        except Exception:
            pass
        self.root.after(4000, self._poll_reservation_alerts)

    def _poll_incoming_shipment_alerts(self):
        try:
            alert = self.db.get_next_incoming_shipment_alert()
            if alert:
                self._update_incoming_shipments_badge()
                shipment_uuid = str(alert.get("shipment_uuid") or "").strip()
                if shipment_uuid not in self._open_incoming_shipments:
                    self.db.mark_incoming_shipment_alert_shown(int(alert["id"]))
                    opened = self._open_incoming_shipment_checklist(alert)
                    if opened is False:
                        self.db.reset_incoming_shipment_alert_shown(int(alert["id"]))
                self._update_incoming_shipments_badge()
        except Exception:
            pass
        self.root.after(4500, self._poll_incoming_shipment_alerts)

    def _open_incoming_shipment_checklist(self, alert: Dict[str, Any]) -> bool:
        shipment_uuid = str(alert.get("shipment_uuid") or "").strip()
        rows = self.db.list_grouped_pending_shipment_items(shipment_uuid)
        if not shipment_uuid or not rows:
            return False
        self._open_incoming_shipments.add(shipment_uuid)

        dlg = tk.Toplevel(self.root)
        dlg.title("تأكيد استلام شحنة فرع")
        dlg.geometry("920x500")
        dlg.minsize(760, 420)
        dlg.resizable(True, True)
        dlg.transient(self.root)
        dlg.grab_set()
        confirmed = False

        def _dismiss_receipt():
            self._open_incoming_shipments.discard(shipment_uuid)
            if not confirmed:
                ToastNotification.show(
                    self.root,
                    "لم يتم تأكيد الشحنة بعد. ستظل موجودة في تبويب شحنات غير مؤكدة ويمكن فتحها لاحقاً.",
                    toast_type="warning",
                )
            try:
                if hasattr(self, "_incoming_shipments_frame"):
                    self._incoming_shipments_frame._refresh()
            except Exception:
                pass
            dlg.destroy()

        dlg.protocol("WM_DELETE_WINDOW", _dismiss_receipt)

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text=f"شحنة #{shipment_uuid[:8]} من {alert.get('from_device') or 'WAREHOUSE'}", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(frm, text="تحقق من الكمية بنداً بنداً. عند الاختلاف، اكتب الكمية المستلمة الفعلية.").pack(anchor="w", pady=(0, 8))

        wrap = ttk.Frame(frm)
        wrap.pack(fill=tk.BOTH, expand=True)
        tv = ttk.Treeview(
            wrap,
            columns=("idx", "item", "school", "color", "size", "expected", "received"),
            show="headings",
            height=12,
        )
        for col, txt, w in [
            ("idx", "#", 40),
            ("item", "النوع", 160),
            ("school", "المدرسة", 150),
            ("color", "اللون", 100),
            ("size", "المقاس", 80),
            ("expected", "المرسل", 90),
            ("received", "المستلم", 90),
        ]:
            tv.heading(col, text=txt)
            tv.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(wrap, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=ysb.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

        received_map: Dict[int, int] = {}
        for r in rows:
            idx = int(r["group_key"])
            exp = int(r["expected_qty"] or 0)
            received_map[idx] = exp
            tv.insert("", tk.END, iid=str(idx), values=(idx + 1, r["item_type"], r["school"], r["color"], r["size"], exp, exp))

        def _edit_received(_e=None):
            cur = tv.focus() or (tv.selection()[0] if tv.selection() else None)
            if not cur:
                return
            idx = int(cur)
            val = simpledialog.askinteger("تعديل الكمية المستلمة", "أدخل الكمية المستلمة:", initialvalue=int(received_map.get(idx, 0)), minvalue=0, parent=dlg)
            if val is None:
                return
            received_map[idx] = int(val)
            row_vals = list(tv.item(cur, "values"))
            row_vals[6] = int(val)
            tv.item(cur, values=row_vals)

        tv.bind("<Double-1>", _edit_received, add="+")
        ttk.Label(frm, text="انقر نقراً مزدوجاً على الصف لتعديل الكمية المستلمة.", foreground="#6b7280").pack(anchor="w", pady=(6, 0))

        note_var = tk.StringVar(value="")
        note_row = ttk.Frame(frm)
        note_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(note_row, text="ملاحظة (اختياري):").pack(side=tk.RIGHT)
        ttk.Entry(note_row, textvariable=note_var).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(8, 0))

        def _confirm_receipt():
            nonlocal confirmed
            payload = [{"line_indexes": [int(x) for x in (r.get("line_indexes") or [])], "received_qty": int(received_map.get(int(r["group_key"]), 0))} for r in rows]
            try:
                out = self.db.confirm_grouped_incoming_shipment(shipment_uuid, payload, note_var.get())
                confirmed = True
                self._open_incoming_shipments.discard(shipment_uuid)
                dlg.destroy()
                msg = "تم استلام الشحنة وتأكيد البنود."
                if out.get("has_diff"):
                    msg += " تم إرسال فروقات للـ Warehouse للمراجعة."
                ToastNotification.show(self.root, msg, toast_type="success")
                try:
                    if hasattr(self, "_incoming_shipments_frame"):
                        self._incoming_shipments_frame._refresh()
                except Exception:
                    pass
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        btn_row = ttk.Frame(frm)
        btn_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(btn_row, text="فتح لاحقاً", command=_dismiss_receipt).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="تأكيد الاستلام", command=_confirm_receipt).pack(side=tk.RIGHT)

        return True

    def _on_sync_completed(self):
        try:
            if hasattr(self, "_incoming_shipments_frame"):
                self._incoming_shipments_frame._refresh()
            else:
                self._update_incoming_shipments_badge()
        except Exception:
            pass
        self._repair_reclassification_counts_after_sync()
        self._on_tab_changed(None)

    def _on_background_sync_completed(self, summary=None):
        try:
            if hasattr(self, "_incoming_shipments_frame"):
                self._incoming_shipments_frame._refresh()
            else:
                self._update_incoming_shipments_badge()
        except Exception:
            pass
        self._repair_reclassification_counts_after_sync()
        self._on_tab_changed(None)

    def _repair_reclassification_counts_after_sync(self):
        try:
            result = self.db.repair_missing_branch_reclassification_counts()
            qty = int(result.get("qty_restored") or 0)
            if qty > 0:
                self._update_status(f"تم إصلاح كميات تصنيف الفروع ({qty} قطعة)")
                try:
                    ToastNotification.show(
                        self.root,
                        f"تم إصلاح كميات تصنيف الفروع ({qty} قطعة).",
                        toast_type="success",
                    )
                except Exception:
                    pass
        except Exception:
            try:
                logging_setup.log_exception("branch_reclassification_repair.after_sync_failed")  # type: ignore[union-attr]
            except Exception:
                pass

    def _bind_shortcuts(self):
        self.root.bind("<F5>", lambda e: self._shortcut_refresh())
        self.root.bind("<F8>", lambda e: self._shortcut_finalize())
        self.root.bind("<Escape>", lambda e: self._shortcut_escape())
        self.root.bind("<Control-p>", lambda e: self._shortcut_print())
        self.root.bind("<Delete>", lambda e: self._shortcut_delete())

    def _shortcut_refresh(self):
        try:
            tab = self._nb.tab(self._nb.select(), "text")
            if "\u0627\u0644\u0625\u062D\u0635\u0627\u0626\u064A\u0627\u062A" in tab:
                self._stats_frame._refresh_all()
            elif "\u0627\u0644\u062D\u062C\u0648\u0632\u0627\u062A" in tab:
                self._res_frame._refresh()
            elif "\u0627\u0644\u0631\u0626\u064A\u0633\u064A\u0629" in tab:
                self._dashboard.refresh()
            elif "حسابات المدارس" in tab:
                self._school_accounts_frame._refresh()
            self._update_status("F5: \u062A\u0645 \u0627\u0644\u062A\u062D\u062F\u064A\u062B")
        except Exception:
            pass

    def _shortcut_finalize(self):
        try:
            tab = self._nb.tab(self._nb.select(), "text")
            if "\u0646\u0642\u0637\u0629 \u0627\u0644\u0628\u064A\u0639" in tab:
                self._pos_frame._finalize()
            elif "\u0627\u0644\u0648\u0627\u0631\u062F" in tab and self._income_frame is not None:
                self._income_frame._finalize_income()
        except Exception:
            pass

    def _shortcut_escape(self):
        try:
            tab = self._nb.tab(self._nb.select(), "text")
            if "\u0646\u0642\u0637\u0629 \u0627\u0644\u0628\u064A\u0639" in tab:
                self._pos_frame._clear_bill()
        except Exception:
            pass

    def _shortcut_print(self):
        try:
            tab = self._nb.tab(self._nb.select(), "text")
            if "\u0627\u0644\u0625\u062D\u0635\u0627\u0626\u064A\u0627\u062A" in tab:
                self._stats_frame._print_movements()
        except Exception:
            pass

    def _shortcut_delete(self):
        try:
            tab = self._nb.tab(self._nb.select(), "text")
            if "\u0646\u0642\u0637\u0629 \u0627\u0644\u0628\u064A\u0639" in tab:
                self._pos_frame._remove_line()
        except Exception:
            pass

    def _toggle_dark_mode(self):
        is_dark = not getattr(self.root, "_theme_dark", False)
        _apply_colorful_theme_to_root(self.root, dark=is_dark)
        T = self.root._theme
        # Update header
        for w in (self._header,):
            w.configure(bg=T["HEADER_BG"])
        for w in self._header.winfo_children():
            try:
                w.configure(bg=T["HEADER_BG"])
            except Exception:
                pass
        # Update status bar
        self._statusbar.configure(bg=T["EDGE"])
        for w in self._statusbar.winfo_children():
            try:
                w.configure(bg=T["EDGE"], fg=T["TEXT"])
            except Exception:
                pass
        self._dark_btn.configure(
            text="\u2600 \u0627\u0644\u0648\u0636\u0639 \u0627\u0644\u0639\u0627\u062F\u064A" if is_dark else "\u263E \u0627\u0644\u0648\u0636\u0639 \u0627\u0644\u0644\u064A\u0644\u064A")
        self._update_status("\u062A\u0645 \u062A\u063A\u064A\u064A\u0631 \u0627\u0644\u0633\u0645\u0629" + (" \u0627\u0644\u0644\u064A\u0644\u064A\u0629" if is_dark else " \u0627\u0644\u0639\u0627\u062F\u064A\u0629"))

    def _on_tab_changed(self, event=None):
        try:
            current = self._nb.select()
            tab_text = self._nb.tab(current, "text")
            if "\u0627\u0644\u0625\u062D\u0635\u0627\u0626\u064A\u0627\u062A" in tab_text:
                self._stats_frame._refresh_all()
            elif "\u0627\u0644\u0631\u0626\u064A\u0633\u064A\u0629" in tab_text:
                self._dashboard.refresh()
            elif "\u0627\u0644\u062D\u062C\u0648\u0632\u0627\u062A" in tab_text:
                self._res_frame._refresh()
            elif "شحنات غير مؤكدة" in tab_text:
                self._incoming_shipments_frame._refresh()
            elif "ملخص الورديات" in tab_text:
                self._shifts_frame._refresh_all()
            elif "حسابات المدارس" in tab_text:
                self._school_accounts_frame._refresh()
        except Exception:
            pass

    def _open_inventory(self):
        if not self.db.is_manager_feature_enabled("allow_inventory_window"):
            messagebox.showwarning("مقيد", _feature_restricted_message("نافذة المخزون مقيدة حالياً في نقطة البيع."), parent=self.root)
            return
        try:
            InventoryWindow(self.root, self.db)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def _open_stock_audit(self):
        if not self.db.is_manager_feature_enabled("allow_stock_audit"):
            messagebox.showwarning("مقيد", _feature_restricted_message("نافذة الجرد مقيدة حاليا في نقطة البيع."), parent=self.root)
            return
        pw = simpledialog.askstring(
            "كلمة مرور المدير",
            "أدخل كلمة مرور المدير لفتح الجرد:",
            show="*",
            parent=self.root,
        )
        if pw is None:
            return
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة مرور المدير غير صحيحة.", parent=self.root)
            return
        try:
            StockAuditWindow(self.root, self.db)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def _open_bills_history(self):
        try:
            BillsHistoryWindow(self.root, self.db)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def _open_movements(self):
        try:
            MovementsWindow(self.root, self.db)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def _open_admin(self):
        try:
            AdminWindow(self.root, self.db)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex))

    def _open_sync_dialog(self):
        try:
            import sync_ui
            sync_ui.open_sync_dialog(self.root, self.db.conn)
        except Exception as ex:
            messagebox.showerror("المزامنة", f"تعذّر فتح نافذة المزامنة:\n{ex}")

    def _open_bulk_price(self):
        if not self.db.is_manager_feature_enabled("allow_bulk_price"):
            messagebox.showwarning("مقيد", _feature_restricted_message("تعديل الأسعار من نقطة البيع غير مسموح به حالياً."), parent=self.root)
            return
        pw = simpledialog.askstring("كلمة مرور المدير", "أدخل كلمة مرور المدير:", show="*", parent=self.root)
        if not pw:
            return
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self.root)
            return
        BulkPriceDialog(self.root, self.db)

    def _refresh_feature_states(self):
        try:
            self._inventory_btn.configure(
                state=("normal" if self.db.is_manager_feature_enabled("allow_inventory_window") else "disabled")
            )
        except Exception:
            pass
        try:
            self._stock_audit_btn.configure(
                state=("normal" if self.db.is_manager_feature_enabled("allow_stock_audit") else "disabled")
            )
        except Exception:
            pass
        try:
            self._bulk_price_btn.configure(
                state=("normal" if self.db.is_manager_feature_enabled("allow_bulk_price") else "disabled")
            )
        except Exception:
            pass

    def _check_shift_state(self):
        shift = self.db.get_open_shift()
        if shift:
            self._current_shift_id = shift["id"]
            self.db.active_shift_id = shift["id"]
            self._update_shift_ui(shift)
            self._set_app_enabled(True)
            return
        total = self.db.conn.execute("SELECT COUNT(*) AS c FROM shifts").fetchone()["c"]
        if total == 0:
            self._begin_shift_with_update_check(show_first_run_notice=True)
        else:
            self._current_shift_id = None
            self.db.active_shift_id = None
            self._shift_status_var.set("لا توجد وردية مفتوحة")
            self._shift_btn.configure(text="بدء وردية")
            self._set_app_enabled(False)
            self._prompt_start_shift()

    def _update_shift_ui(self, shift):
        started = fmt_local_ts(shift["started_at"], "")
        self._shift_status_var.set(f"الوردية مفتوحة منذ {started}")
        self._shift_btn.configure(text="إنهاء الوردية")

    def _toggle_shift(self):
        if self._current_shift_id:
            self._prompt_end_shift()
        else:
            self._prompt_start_shift()

    def _start_shift_now(self, show_first_run_notice: bool = False):
        try:
            sid = self.db.start_shift()
            self._current_shift_id = sid
            self.db.active_shift_id = sid
            shift = self.db.get_open_shift()
            self._update_shift_ui(shift)
            self._set_app_enabled(True)
            if show_first_run_notice:
                messagebox.showinfo("نظام الورديات", "تم تفعيل نظام الورديات.\nتم فتح وردية جديدة تلقائياً.", parent=self.root)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self.root)

    def _begin_shift_with_update_check(self, show_first_run_notice: bool = False):
        self._shift_btn.configure(state="disabled")

        def _continue_start():
            try:
                self._shift_btn.configure(state="normal")
            except Exception:
                pass
            self._start_shift_now(show_first_run_notice=show_first_run_notice)

        try:
            import sync_ui
            sync_ui.check_pos_update_before_shift(
                self.root,
                self.db.conn,
                on_continue=_continue_start,
            )
        except Exception:
            _continue_start()

    def _prompt_start_shift(self):
        if messagebox.askyesno("بدء وردية جديدة", "هل تريد بدء وردية جديدة؟\nيجب فتح وردية قبل البدء بالعمل."):
            self._begin_shift_with_update_check()

    def _prompt_end_shift(self):
        ShiftSummaryDialog(self.root, self.db, self._current_shift_id, on_closed=self._on_shift_closed)

    def _on_shift_closed(self):
        self._current_shift_id = None
        self.db.active_shift_id = None
        self._shift_status_var.set("لا توجد وردية مفتوحة")
        self._shift_btn.configure(text="بدء وردية")
        self._set_app_enabled(False)
        try:
            import sync_ui
            sync_ui.run_sync_now(self.root, self.db.conn, reason="إغلاق الوردية")
        except Exception:
            pass

    def _set_app_enabled(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        for tab_id in self._nb.tabs():
            try:
                self._nb.tab(tab_id, state=state)
            except Exception:
                pass
        for child in self._toolbar.winfo_children():
            if child is not self._shift_btn:
                try:
                    child.configure(state=state)
                except Exception:
                    pass

    def _on_close(self):
        if self._current_shift_id:
            result = messagebox.askyesnocancel(
                "إغلاق البرنامج",
                "يوجد وردية مفتوحة.\nهل تريد إنهاء الوردية قبل الإغلاق؟\n\nنعم = إنهاء الوردية ثم إغلاق\nلا = إغلاق بدون إنهاء\nإلغاء = العودة")
            if result is True:
                ShiftSummaryDialog(self.root, self.db, self._current_shift_id,
                                   on_closed=self._sync_after_shift_close_then_destroy)
            elif result is False:
                self.root.destroy()
        else:
            self.root.destroy()

    def _sync_after_shift_close_then_destroy(self):
        """Push the final SHIFT_CLOSED event before the POS process exits."""
        self._current_shift_id = None
        self.db.active_shift_id = None
        try:
            db_path = self.db.conn.execute("PRAGMA database_list").fetchone()[2]
        except Exception:
            self.root.destroy()
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("مزامنة إغلاق الوردية")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        ttk.Label(
            dlg,
            text="جارٍ إرسال إغلاق الوردية قبل إغلاق البرنامج...",
            padding=16,
        ).pack(fill=tk.X)
        status_var = tk.StringVar(value="يرجى الانتظار")
        ttk.Label(dlg, textvariable=status_var, padding=(16, 0, 16, 12)).pack(fill=tk.X)
        try:
            dlg.protocol("WM_DELETE_WINDOW", lambda: None)
            dlg.geometry("+%d+%d" % (self.root.winfo_rootx() + 420, self.root.winfo_rooty() + 260))
        except Exception:
            pass

        result: Dict[str, Any] = {"done": False, "error": ""}

        def worker() -> None:
            try:
                import sync_client
                conn = sqlite3.connect(db_path, timeout=8.0, isolation_level=None, check_same_thread=False)
                conn.row_factory = sqlite3.Row
                conn.execute("PRAGMA busy_timeout=8000;")
                conn.execute("PRAGMA synchronous=NORMAL;")
                conn.execute("PRAGMA temp_store=MEMORY;")
                conn.execute("PRAGMA cache_size=-12000;")
                try:
                    sync_client.SyncClient(conn).run_cycle(progress=None)
                finally:
                    conn.close()
            except Exception as ex:
                result["error"] = str(ex)
            finally:
                result["done"] = True

        def pump() -> None:
            if not result.get("done"):
                try:
                    self.root.after(120, pump)
                except Exception:
                    pass
                return
            try:
                if dlg.winfo_exists():
                    dlg.destroy()
            except Exception:
                pass
            if result.get("error"):
                try:
                    messagebox.showwarning(
                        "مزامنة إغلاق الوردية",
                        "تم حفظ إغلاق الوردية محلياً، لكن تعذر إرسالها الآن.\n"
                        "سيتم إرسالها في أول مزامنة قادمة.\n\n"
                        f"{result.get('error')}",
                        parent=self.root,
                    )
                except Exception:
                    pass
            self.root.destroy()

        threading.Thread(target=worker, daemon=True).start()
        self.root.after(120, pump)


# ------------------- Entry Point -------------------

def main():
    db = SqliteDatabase(DB_PATH)
    root = tk.Tk()
    # Apply the colorful theme
    _apply_colorful_theme_to_root(root)
    _app = WarehouseApp(root, db)
    root.mainloop()


_DARK_MODE = False  # Global toggle

THEME_LIGHT = {
    "BG": "#f8fafc", "SURFACE": "#e2e8f0", "CARD": "#ffffff",
    "ACCENT": "#2563eb", "ACCENT2": "#1d4ed8", "SUCCESS": "#16a34a", "SUCCESS2": "#15803d",
    "DANGER": "#dc2626", "DANGER2": "#b91c1c", "WARNING": "#f59e0b",
    "TEXT": "#0f172a", "TEXT2": "#475569", "SUBTLE": "#94a3b8",
    "ROW": "#ffffff", "ROW_ALT": "#f1f5f9", "SELBG": "#bfdbfe",
    "EDGE": "#cbd5e1", "HEADER_BG": "#1e3a5f", "HEADER_FG": "#ffffff",
}

THEME_DARK = {
    "BG": "#0f172a", "SURFACE": "#1e293b", "CARD": "#334155",
    "ACCENT": "#3b82f6", "ACCENT2": "#2563eb", "SUCCESS": "#22c55e", "SUCCESS2": "#16a34a",
    "DANGER": "#ef4444", "DANGER2": "#dc2626", "WARNING": "#fbbf24",
    "TEXT": "#f1f5f9", "TEXT2": "#94a3b8", "SUBTLE": "#64748b",
    "ROW": "#1e293b", "ROW_ALT": "#334155", "SELBG": "#1e40af",
    "EDGE": "#475569", "HEADER_BG": "#020617", "HEADER_FG": "#e2e8f0",
}


def _apply_colorful_theme_to_root(root, dark=False):
    """Apply a modern colour palette with semantic accent colours."""
    global _DARK_MODE
    _DARK_MODE = dark
    T = THEME_DARK if dark else THEME_LIGHT
    s = ttk.Style(root)
    try:
        s.theme_use("clam")
    except Exception:
        pass

    BG, SURFACE, ACCENT, ACCENT2 = T["BG"], T["SURFACE"], T["ACCENT"], T["ACCENT2"]
    TEXT, SUBTLE, ROW, SELBG, EDGE = T["TEXT"], T["SUBTLE"], T["ROW"], T["SELBG"], T["EDGE"]
    ROW_ALT = T["ROW_ALT"]
    SUCCESS, SUCCESS2 = T["SUCCESS"], T["SUCCESS2"]
    DANGER, DANGER2 = T["DANGER"], T["DANGER2"]

    try:
        root.configure(bg=BG)
    except Exception:
        pass

    # Base widgets
    s.configure("TFrame", background=SURFACE)
    s.configure("TLabelframe", background=SURFACE, bordercolor=EDGE, relief="groove")
    s.configure("TLabelframe.Label", background=SURFACE, foreground=TEXT,
                font=("Segoe UI", 10, "bold"))
    s.configure("TLabel", background=SURFACE, foreground=TEXT)
    s.configure("TPanedwindow", background=SURFACE)
    s.configure("TSeparator", background=EDGE)

    # Primary button — blue
    s.configure("TButton", background=ACCENT, foreground="#ffffff",
                bordercolor=ACCENT, focusthickness=1, focuscolor=ACCENT2,
                padding=(12, 6), font=("Segoe UI", 9, "bold"))
    s.map("TButton",
          background=[("active", ACCENT2), ("pressed", ACCENT2), ("disabled", SUBTLE)],
          foreground=[("disabled", "#94a3b8")],
          bordercolor=[("focus", ACCENT2)])

    # Success button — green
    s.configure("Success.TButton", background=SUCCESS, foreground="#ffffff",
                bordercolor=SUCCESS, focuscolor=SUCCESS2, padding=(12, 6),
                font=("Segoe UI", 10, "bold"))
    s.map("Success.TButton",
          background=[("active", SUCCESS2), ("pressed", SUCCESS2)],
          bordercolor=[("focus", SUCCESS2)])

    # Danger button — red
    s.configure("Danger.TButton", background=DANGER, foreground="#ffffff",
                bordercolor=DANGER, focuscolor=DANGER2, padding=(12, 6),
                font=("Segoe UI", 9, "bold"))
    s.map("Danger.TButton",
          background=[("active", DANGER2), ("pressed", DANGER2)],
          bordercolor=[("focus", DANGER2)])

    # Secondary / subtle button
    s.configure("Secondary.TButton", background=SURFACE, foreground=TEXT,
                bordercolor=EDGE, padding=(10, 5), font=("Segoe UI", 9))
    s.map("Secondary.TButton",
          background=[("active", EDGE), ("pressed", EDGE)],
          foreground=[("active", "#ffffff")])

    # Inputs
    for sty in ("TEntry", "TCombobox", "TSpinbox"):
        s.configure(sty, fieldbackground=ROW, background=ROW, foreground=TEXT,
                    bordercolor=EDGE, lightcolor=EDGE, darkcolor=EDGE,
                    padding=4)
        s.map(sty, fieldbackground=[("focus", "#ffffff" if not dark else "#475569")],
              bordercolor=[("focus", ACCENT)])

    s.configure("TCheckbutton", background=SURFACE, foreground=TEXT)
    s.map("TCheckbutton", background=[("active", SURFACE)])
    s.configure("TRadiobutton", background=SURFACE, foreground=TEXT)
    s.map("TRadiobutton", background=[("active", SURFACE)])

    # Notebook tabs
    s.configure("TNotebook", background=BG, bordercolor=EDGE)
    s.configure("TNotebook.Tab", background=SURFACE, foreground=SUBTLE,
                padding=[14, 6], font=("Segoe UI", 9, "bold"))
    s.map("TNotebook.Tab",
          background=[("selected", BG)],
          foreground=[("selected", ACCENT)])

    # Treeview — with alternating row support
    s.configure("Treeview",
                background=ROW, fieldbackground=ROW, foreground=TEXT,
                bordercolor=EDGE, rowheight=26,
                font=("Segoe UI", 9))
    s.configure("Treeview.Heading",
                background=SURFACE, foreground=TEXT, bordercolor=EDGE,
                font=("Segoe UI", 9, "bold"), padding=4)
    s.map("Treeview",
          background=[("selected", SELBG)],
          foreground=[("selected", TEXT)])

    # Scrollbar
    s.configure("TScrollbar", background=SURFACE, troughcolor=BG,
                bordercolor=EDGE, arrowcolor=TEXT)

    # Store theme in root for access
    root._theme = T
    root._theme_dark = dark


if __name__ == "__main__":
    main()
