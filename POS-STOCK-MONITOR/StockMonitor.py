# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import json
import sqlite3
import sys

try:
    import logging_setup
    logging_setup.install_crash_logging("StockMonitor")
except Exception:
    pass

import tkinter as tk
from datetime import datetime, timezone
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Sequence, Tuple


# In source mode, keep assets/DB beside this script folder.
# In PyInstaller onefile mode, use the executable directory (persistent),
# not the temp extraction folder, so settings survive restarts.
SOURCE_DIR = os.path.dirname(os.path.abspath(__file__))
RUNTIME_DIR = os.path.dirname(os.path.abspath(sys.executable)) if getattr(sys, "frozen", False) else SOURCE_DIR
APP_DATA_DIR = os.path.join(RUNTIME_DIR, "StockMonitorData")
DB_FILE_NAME = "stock_monitor_data.sqlite3"
DB_PATH = os.path.join(APP_DATA_DIR, DB_FILE_NAME)
STALE_MINUTES_WARNING = 15
STALE_MINUTES_CRITICAL = 60
ALL_BRANCHES_PICK = "كل الفروع"
BRANCH_UI_NAME_BY_DEVICE = {
    "POS-ZAY": "فرع زايد",
    "POS-OCT": "فرع اكتوبر",
    "POS-BAH": "فرع بهتيم",
    "POS-CEN": "فرع السنتر",
    "POS-OBO": "فرع العبور",
    "POS-GESR": "فرع جسر السويس",
}
WAREHOUSE_DIR = os.path.abspath(os.path.join(SOURCE_DIR, "..", "ادارة المخازن"))
if not getattr(sys, "frozen", False) and WAREHOUSE_DIR not in sys.path:
    sys.path.insert(0, WAREHOUSE_DIR)

import sync_core  # type: ignore
import sync_periodic  # type: ignore
import sync_ui  # type: ignore

_UI = {
    "BG": "#f8fafc",
    "SURFACE": "#e2e8f0",
    "SURFACE2": "#f1f5f9",
    "BORDER": "#cbd5e1",
    "ACCENT": "#2563eb",
    "ACCENT_H": "#1d4ed8",
    "TEXT": "#0f172a",
    "TEXT_SEC": "#475569",
    "TEXT_DIM": "#94a3b8",
    "SEL_BG": "#bfdbfe",
    "SEL_FG": "#0f172a",
    "ROW_EVEN": "#ffffff",
    "ROW_ODD": "#f1f5f9",
}


def _bind_mousewheel(widget: tk.Widget) -> None:
    def _on_wheel(e: tk.Event) -> str:
        try:
            delta = int(-1 * (e.delta / 120))
        except Exception:
            delta = 0
        try:
            widget.yview_scroll(delta, "units")
        except Exception:
            pass
        return "break"

    widget.bind_all("<MouseWheel>", _on_wheel, add="+")


def _apply_zebra_tags(tree: ttk.Treeview) -> None:
    tree.tag_configure("odd", background=_UI["ROW_EVEN"])
    tree.tag_configure("even", background=_UI["ROW_ODD"])
    for i, iid in enumerate(tree.get_children()):
        tree.item(iid, tags=("even" if i % 2 else "odd",))


def _parse_iso_ts(value: Any) -> datetime:
    raw = str(value or "").strip()
    if not raw:
        return datetime.min.replace(tzinfo=timezone.utc)
    try:
        if raw.endswith("Z"):
            raw = raw[:-1] + "+00:00"
        dt = datetime.fromisoformat(raw)
        if dt.tzinfo is None:
            return dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc)
    except Exception:
        return datetime.min.replace(tzinfo=timezone.utc)


def _fmt_local_ts(value: Any, empty: str = "—") -> str:
    raw = str(value or "").strip()
    if not raw:
        return empty
    try:
        txt = raw[:-1] + "+00:00" if raw.endswith("Z") else raw
        dt = datetime.fromisoformat(txt)
        if dt.tzinfo is None:
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return raw.replace("T", " ")


def _branch_display_name(value: Any) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    return BRANCH_UI_NAME_BY_DEVICE.get(raw, raw)


class MonitorDatabase:
    def __init__(self, path: str):
        self.path = os.path.abspath(path)
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        self.conn = sqlite3.connect(self.path, isolation_level=None, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys=ON;")
        self.conn.execute("PRAGMA busy_timeout=8000;")
        self.conn.execute("PRAGMA journal_mode=WAL;")
        self.conn.execute("PRAGMA synchronous=NORMAL;")
        sync_core.apply_sync_migration(self.conn)
        sync_core.ensure_device_identity(
            self.conn,
            default_name="STOCK-MONITOR",
            default_role="warehouse",
        )

    def close(self) -> None:
        try:
            self.conn.close()
        except Exception:
            pass

    def _all_snapshot_sources(self) -> List[str]:
        try:
            rows = self.conn.execute(
                "SELECT DISTINCT source_device FROM pos_stocks_mirror WHERE TRIM(COALESCE(source_device,'')) <> ''"
            ).fetchall()
            return sorted(str(r["source_device"]) for r in rows if str(r["source_device"]).strip())
        except Exception:
            return []

    def _sources_for_pick(self, pick: str) -> List[str]:
        name = (pick or "").strip()
        if not name:
            return []
        if name == ALL_BRANCHES_PICK:
            return self._all_snapshot_sources()
        vals = {name}
        try:
            rows = self.conn.execute(
                """
                SELECT device_name, device_uuid
                FROM known_devices
                WHERE role = 'pos' AND (device_name = ? OR device_uuid = ?)
                """,
                (name, name),
            ).fetchall()
            for r in rows:
                if r["device_name"]:
                    vals.add(str(r["device_name"]))
                if r["device_uuid"]:
                    vals.add(str(r["device_uuid"]))
        except Exception:
            pass
        return sorted(vals)

    def list_device_picks(self) -> Tuple[List[str], Dict[str, Dict[str, Any]]]:
        rows = []
        known = []
        try:
            rows = self.conn.execute(
                """
                SELECT source_device, snapshot_at, row_count, total_value
                FROM pos_stocks_snapshot_meta
                ORDER BY source_device
                """
            ).fetchall()
        except Exception:
            rows = []

        try:
            known = self.conn.execute(
                "SELECT device_name FROM known_devices WHERE role='pos' ORDER BY device_name"
            ).fetchall()
        except Exception:
            known = []

        names: List[str] = []
        metas: Dict[str, Dict[str, Any]] = {}
        for r in rows:
            source = str(r["source_device"] or "").strip()
            if not source:
                continue
            names.append(source)
            metas[source] = {
                "snapshot_at": r["snapshot_at"],
                "row_count": int(r["row_count"] or 0),
                "total_value": float(r["total_value"] or 0.0),
            }
        for r in known:
            dev_name = str(r["device_name"] or "").strip()
            if dev_name and dev_name not in names:
                names.append(dev_name)
        if names:
            names = [ALL_BRANCHES_PICK] + names
        return names, metas

    def list_pos_stock_rows(self, pick: str) -> List[Tuple[str, str, str, str, str, float, int]]:
        sources = self._sources_for_pick(pick)
        if not sources:
            return []
        ph = ",".join("?" for _ in sources)
        sql = f"""
            SELECT source_device, item_type, school, color, size, unit_price, count
            FROM pos_stocks_mirror
            WHERE source_device IN ({ph})
            ORDER BY source_device, item_type, school, color, size
        """
        rows = self.conn.execute(sql, tuple(sources)).fetchall()
        return [
            (
                str(r["source_device"] or ""),
                str(r["item_type"] or ""),
                str(r["school"] or ""),
                str(r["color"] or ""),
                str(r["size"] or ""),
                float(r["unit_price"] or 0),
                int(r["count"] or 0),
            )
            for r in rows
        ]

    def get_sync_diagnostics(self) -> Dict[str, int]:
        def _count(sql: str) -> int:
            try:
                row = self.conn.execute(sql).fetchone()
                return int(row[0] or 0) if row else 0
            except Exception:
                return 0

        return {
            "outbox_pending": _count("SELECT COUNT(*) FROM sync_outbox WHERE status='pending'"),
            "outbox_error": _count("SELECT COUNT(*) FROM sync_outbox WHERE COALESCE(last_error,'')<>''"),
            "inbox_error": _count("SELECT COUNT(*) FROM sync_inbox WHERE apply_status='error'"),
            "inbox_skipped": _count("SELECT COUNT(*) FROM sync_inbox WHERE apply_status='skipped'"),
        }

    def _resolve_branch_device(self, value: str) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        for dev, ui in BRANCH_UI_NAME_BY_DEVICE.items():
            if raw == dev or raw == ui:
                return dev
        return raw

    def get_branch_available_qty(self, branch: str, item_type: str, school: str, color: str, size: str) -> int:
        dev = self._resolve_branch_device(branch)
        if not dev:
            return 0
        row = self.conn.execute(
            """
            SELECT COALESCE(SUM(count), 0)
            FROM pos_stocks_mirror
            WHERE source_device = ?
              AND LOWER(TRIM(item_type)) = LOWER(TRIM(?))
              AND LOWER(TRIM(school)) = LOWER(TRIM(?))
              AND LOWER(TRIM(color)) = LOWER(TRIM(?))
              AND LOWER(TRIM(size)) = LOWER(TRIM(?))
            """,
            (dev, item_type, school, color, size),
        ).fetchone()
        return int((row[0] if row else 0) or 0)

    def create_reservation_request(
        self,
        *,
        branch: str,
        customer: str,
        item_type: str,
        school: str,
        color: str,
        size: str,
        qty: int,
        unit_price: float,
    ) -> str:
        dev = self._resolve_branch_device(branch)
        if not dev:
            raise ValueError("الفرع غير صالح.")
        customer = str(customer or "").strip()
        if not customer:
            raise ValueError("اسم العميل مطلوب.")
        item_type = str(item_type or "").strip()
        school = str(school or "").strip()
        color = str(color or "").strip()
        size = str(size or "").strip()
        if not (item_type and school and color and size):
            raise ValueError("مواصفات الصنف غير مكتملة.")
        try:
            qty = int(qty)
        except Exception:
            raise ValueError("الكمية غير صالحة.")
        if qty <= 0:
            raise ValueError("الكمية يجب أن تكون أكبر من صفر.")
        available = self.get_branch_available_qty(dev, item_type, school, color, size)
        if available <= 0:
            raise ValueError("الصنف غير متاح في هذا الفرع.")
        if qty > available:
            raise ValueError(f"الكمية المطلوبة أكبر من المتاح ({available}).")
        if (available - qty) < 4:
            raise ValueError(
                f"لا يمكن الحجز لأن المتبقي سيكون أقل من 4 (المتاح: {available}, المطلوب: {qty})."
            )
        try:
            unit_price = float(unit_price or 0)
        except Exception:
            unit_price = 0.0
        event_uuid = sync_core.new_uuid()
        request_uuid = sync_core.new_uuid()
        created_at = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.%fZ")
        payload = {
            "request_uuid": request_uuid,
            "customer": customer,
            "hold_days": 2,
            "line": {
                "item_type": item_type,
                "school": school,
                "color": color,
                "size": size,
                "qty": qty,
                "unit_price": unit_price,
            },
            "requested_by": "STOCK-MONITOR",
        }
        with self.conn:
            self.conn.execute(
                """
                INSERT INTO sync_outbox
                    (event_uuid, event_type, payload_json, created_at, status, attempts, target_scope)
                VALUES (?, 'REMOTE_RESERVATION_REQUEST', ?, ?, 'pending', 0, ?)
                """,
                (
                    event_uuid,
                    json.dumps(payload, ensure_ascii=False, default=str),
                    created_at,
                    f"pos:{dev}",
                ),
            )
        return event_uuid

    def list_branch_health_rows(self) -> List[Dict[str, Any]]:
        now = datetime.now(timezone.utc)
        snaps: Dict[str, Dict[str, Any]] = {}
        try:
            rows = self.conn.execute(
                "SELECT source_device, snapshot_at, row_count, total_value FROM pos_stocks_snapshot_meta"
            ).fetchall()
            for r in rows:
                key = str(r["source_device"] or "").strip()
                if not key:
                    continue
                snaps[key] = {
                    "snapshot_at": r["snapshot_at"],
                    "row_count": int(r["row_count"] or 0),
                    "total_value": float(r["total_value"] or 0.0),
                }
        except Exception:
            snaps = {}

        devices: List[Tuple[str, str]] = []
        try:
            rows = self.conn.execute(
                "SELECT device_name, device_uuid FROM known_devices WHERE role='pos' ORDER BY device_name"
            ).fetchall()
            devices = [(str(r["device_name"] or "").strip(), str(r["device_uuid"] or "").strip()) for r in rows]
        except Exception:
            devices = []

        # Include snapshot-only devices that are not yet in known_devices.
        seen = set()
        health_rows: List[Dict[str, Any]] = []
        for name, uid in devices:
            if not name:
                continue
            meta = snaps.get(name) or (snaps.get(uid) if uid else None)
            snap_at = (meta or {}).get("snapshot_at")
            dt = _parse_iso_ts(snap_at)
            age_min = int((now - dt).total_seconds() // 60) if dt > datetime.min.replace(tzinfo=timezone.utc) else -1
            if age_min < 0:
                status = "NO_SNAPSHOT"
                alert = "No stock snapshot yet"
            elif age_min >= STALE_MINUTES_CRITICAL:
                status = "CRITICAL"
                alert = "Snapshot is stale (critical)"
            elif age_min >= STALE_MINUTES_WARNING:
                status = "WARNING"
                alert = "Snapshot is aging"
            else:
                status = "OK"
                alert = ""
            health_rows.append(
                {
                    "branch": name,
                    "status": status,
                    "snapshot_at": snap_at or "",
                    "age_min": age_min if age_min >= 0 else "",
                    "rows": int((meta or {}).get("row_count") or 0),
                    "value": float((meta or {}).get("total_value") or 0.0),
                    "alert": alert,
                }
            )
            seen.add(name)
            if uid:
                seen.add(uid)

        for source, meta in snaps.items():
            if source in seen:
                continue
            dt = _parse_iso_ts(meta.get("snapshot_at"))
            age_min = int((now - dt).total_seconds() // 60) if dt > datetime.min.replace(tzinfo=timezone.utc) else -1
            health_rows.append(
                {
                    "branch": source,
                    "status": "UNKNOWN",
                    "snapshot_at": meta.get("snapshot_at") or "",
                    "age_min": age_min if age_min >= 0 else "",
                    "rows": int(meta.get("row_count") or 0),
                    "value": float(meta.get("total_value") or 0.0),
                    "alert": "Snapshot exists but branch is not in known_devices",
                }
            )

        def _status_rank(s: str) -> int:
            order = {"CRITICAL": 0, "WARNING": 1, "NO_SNAPSHOT": 2, "UNKNOWN": 3, "OK": 4}
            return order.get(s, 9)

        health_rows.sort(key=lambda x: (_status_rank(str(x["status"])), str(x["branch"])))
        return health_rows


class StockMonitorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Call Center Stock Finder")
        self.geometry("1020x660")
        self.minsize(920, 560)
        self.option_add("*Font", "{Segoe UI} 9")
        self.option_add("*Listbox.Font", "{Segoe UI} 9")
        self.option_add("*Menu.Font", "{Segoe UI} 9")
        self.option_add("*Message.font", "{Segoe UI} 9")
        self._apply_pos_like_theme()
        self.db = MonitorDatabase(DB_PATH)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._all_rows: List[Tuple[str, str, str, str, str, float, int]] = []
        self._metas: Dict[str, Dict[str, Any]] = {}
        self._build()
        self._reload_devices()
        try:
            sync_periodic.attach_periodic_sync(self, self.db.path)
        except Exception:
            pass

    def _build(self) -> None:
        top = ttk.Frame(self, padding=8)
        top.pack(fill=tk.X)
        top_main = ttk.Frame(top)
        top_main.pack(fill=tk.X)
        ttk.Label(
            top_main,
            text="Read-only lookup tool for call center availability checks.",
            foreground=_UI["TEXT_SEC"],
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Label(top_main, text="الفرع:").pack(side=tk.RIGHT, padx=(0, 4))
        self._device_var = tk.StringVar()
        self._device_cb = ttk.Combobox(top_main, textvariable=self._device_var, state="readonly", width=30)
        self._device_cb.pack(side=tk.RIGHT, padx=4)
        self._device_cb.bind("<<ComboboxSelected>>", lambda _e: self._reload_stock())

        self._meta_var = tk.StringVar(value="—")
        ttk.Label(
            top,
            textvariable=self._meta_var,
            foreground=_UI["TEXT_DIM"],
            anchor="e",
            justify="right",
        ).pack(fill=tk.X, pady=(4, 0))

        ttk.Button(top_main, text="تحديث", command=self._reload_devices).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_main, text="مزامنة الآن", command=self._run_sync_and_reload).pack(side=tk.LEFT, padx=4)
        ttk.Button(top_main, text="طلب حجز", command=self._open_reservation_request).pack(side=tk.LEFT, padx=4)
        self._alert_var = tk.StringVar(value="")
        ttk.Label(
            self,
            textvariable=self._alert_var,
            foreground="#d97706",
            anchor="w",
            justify="left",
        ).pack(
            fill=tk.X, padx=10, pady=(0, 6)
        )

        filt = ttk.LabelFrame(self, text="تصنيف")
        filt.pack(fill=tk.X, padx=8, pady=(0, 6))

        ttk.Label(filt, text="النوع").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        ttk.Label(filt, text="المدرسة").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        ttk.Label(filt, text="اللون").grid(row=0, column=4, sticky="e", padx=6, pady=6)
        ttk.Label(filt, text="المقاس").grid(row=0, column=6, sticky="e", padx=6, pady=6)

        self._type_var = tk.StringVar()
        self._school_var = tk.StringVar()
        self._color_var = tk.StringVar()
        self._size_var = tk.StringVar()
        self._search_var = tk.StringVar()

        self._type_cb = ttk.Combobox(filt, textvariable=self._type_var, state="readonly", width=16)
        self._school_cb = ttk.Combobox(filt, textvariable=self._school_var, state="readonly", width=16)
        self._color_cb = ttk.Combobox(filt, textvariable=self._color_var, state="readonly", width=16)
        self._size_cb = ttk.Combobox(filt, textvariable=self._size_var, state="readonly", width=12)

        self._type_cb.grid(row=0, column=1, sticky="ew", padx=6, pady=6)
        self._school_cb.grid(row=0, column=3, sticky="ew", padx=6, pady=6)
        self._color_cb.grid(row=0, column=5, sticky="ew", padx=6, pady=6)
        self._size_cb.grid(row=0, column=7, sticky="ew", padx=6, pady=6)

        ttk.Label(filt, text="بحث").grid(row=1, column=6, sticky="e", padx=6, pady=6)
        ent = ttk.Entry(filt, textvariable=self._search_var, width=28)
        ent.grid(row=1, column=7, sticky="ew", padx=6, pady=6)
        self._search_var.trace_add("write", lambda *_: self._apply_filters())

        btns = ttk.Frame(filt)
        btns.grid(row=1, column=0, columnspan=4, sticky="w", padx=6, pady=6)
        ttk.Button(btns, text="بحث", command=self._apply_filters).pack(side=tk.LEFT)
        ttk.Button(btns, text="مسح", command=self._clear_filters).pack(side=tk.LEFT, padx=8)

        for c in (1, 3, 5, 7):
            filt.columnconfigure(c, weight=1)

        for cb in (self._type_cb, self._school_cb, self._color_cb, self._size_cb):
            cb.bind("<<ComboboxSelected>>", lambda _e: self._apply_filters(), add="+")

        cols = ("branch", "item_type", "school", "color", "size", "unit_price", "count", "value")
        self._tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        headers = {
            "branch": "الفرع",
            "item_type": "النوع",
            "school": "المدرسة",
            "color": "اللون",
            "size": "المقاس",
            "unit_price": "السعر",
            "count": "الكمية",
            "value": "القيمة",
        }
        widths = {
            "branch": 180, "item_type": 140, "school": 120, "color": 110, "size": 90,
            "unit_price": 90, "count": 90, "value": 110,
        }
        for c in cols:
            self._tree.heading(c, text=headers[c])
            self._tree.column(c, width=widths[c], anchor="center")
        self._tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        _bind_mousewheel(self._tree)

        health_fr = ttk.LabelFrame(self, text="Branch Health")
        health_fr.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 6))
        health_cols = ("branch", "status", "snapshot_at", "age_min", "rows", "value", "alert")
        self._health_tree = ttk.Treeview(health_fr, columns=health_cols, show="headings", height=7)
        health_headers = {
            "branch": "Branch",
            "status": "Status",
            "snapshot_at": "Last Snapshot",
            "age_min": "Age (min)",
            "rows": "Rows",
            "value": "Value",
            "alert": "Alert",
        }
        health_widths = {"branch": 160, "status": 90, "snapshot_at": 170, "age_min": 90, "rows": 70, "value": 90, "alert": 320}
        for c in health_cols:
            self._health_tree.heading(c, text=health_headers[c])
            self._health_tree.column(c, width=health_widths[c], anchor="center")
        self._health_tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self._status_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._status_var, anchor="e").pack(fill=tk.X, padx=8, pady=(0, 8))

    def _apply_pos_like_theme(self) -> None:
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass
        try:
            self.configure(bg=_UI["BG"])
        except Exception:
            pass

        s.configure("TFrame", background=_UI["SURFACE"])
        s.configure("TLabelframe", background=_UI["SURFACE"], bordercolor=_UI["BORDER"], relief="groove")
        s.configure("TLabelframe.Label", background=_UI["SURFACE"], foreground=_UI["TEXT"], font=("Segoe UI", 10, "bold"))
        s.configure("TLabel", background=_UI["SURFACE"], foreground=_UI["TEXT"])

        s.configure("TButton", background=_UI["ACCENT"], foreground="#ffffff",
                    bordercolor=_UI["ACCENT"], focusthickness=1, focuscolor=_UI["ACCENT_H"],
                    padding=(12, 6), font=("Segoe UI", 9, "bold"))
        s.map("TButton",
              background=[("active", _UI["ACCENT_H"]), ("pressed", _UI["ACCENT_H"]), ("disabled", _UI["TEXT_DIM"])],
              foreground=[("disabled", "#94a3b8")],
              bordercolor=[("focus", _UI["ACCENT_H"])])

        for sty in ("TEntry", "TCombobox", "TSpinbox"):
            s.configure(sty, fieldbackground="#ffffff", background="#ffffff", foreground=_UI["TEXT"],
                        bordercolor=_UI["BORDER"], lightcolor=_UI["BORDER"], darkcolor=_UI["BORDER"], padding=4)
            s.map(sty, fieldbackground=[("focus", "#ffffff"), ("readonly", _UI["SURFACE2"])],
                  bordercolor=[("focus", _UI["ACCENT"])],
                  lightcolor=[("focus", _UI["ACCENT"])],
                  darkcolor=[("focus", _UI["ACCENT"])])

        s.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground=_UI["TEXT"],
                    bordercolor=_UI["BORDER"], rowheight=26, font=("Segoe UI", 9))
        s.configure("Treeview.Heading", background=_UI["SURFACE"], foreground=_UI["TEXT"],
                    bordercolor=_UI["BORDER"], font=("Segoe UI", 9, "bold"), padding=4)
        s.map("Treeview", background=[("selected", _UI["SEL_BG"])], foreground=[("selected", _UI["SEL_FG"])])
        s.configure("TScrollbar", background=_UI["SURFACE"], troughcolor=_UI["BG"],
                    bordercolor=_UI["BORDER"], arrowcolor=_UI["TEXT"])

    def _refresh_filter_values(self) -> None:
        values = {"type": set(), "school": set(), "color": set(), "size": set()}
        for _br, it, sc, cl, sz, _p, _c in self._all_rows:
            if it:
                values["type"].add(it)
            if sc:
                values["school"].add(sc)
            if cl:
                values["color"].add(cl)
            if sz:
                values["size"].add(sz)

        self._type_cb["values"] = [""] + sorted(values["type"])
        self._school_cb["values"] = [""] + sorted(values["school"])
        self._color_cb["values"] = [""] + sorted(values["color"])
        self._size_cb["values"] = [""] + sorted(values["size"])

    def _reload_devices(self) -> None:
        names, metas = self.db.list_device_picks()
        self._metas = metas
        self._device_ui_to_raw = {ALL_BRANCHES_PICK: ALL_BRANCHES_PICK}
        for n in names:
            if n != ALL_BRANCHES_PICK:
                self._device_ui_to_raw[_branch_display_name(n)] = n
        ui_names = [ALL_BRANCHES_PICK if n == ALL_BRANCHES_PICK else _branch_display_name(n) for n in names]
        self._device_cb["values"] = ui_names
        self._refresh_observability()
        if not ui_names:
            self._device_var.set("")
            self._meta_var.set("لا توجد بيانات فروع بعد — شغّل المزامنة أولاً")
            self._all_rows = []
            self._apply_filters()
            return
        if self._device_var.get() not in ui_names:
            self._device_var.set(ui_names[0])
        self._reload_stock()

    def _reload_stock(self) -> None:
        pick_ui = (self._device_var.get() or "").strip()
        pick = getattr(self, "_device_ui_to_raw", {}).get(pick_ui, pick_ui)
        if not pick:
            return
        if pick == ALL_BRANCHES_PICK:
            self._meta_var.set("عرض مجمّع لكل الفروع المتاحة في المرآة")
        else:
            meta = self._metas.get(pick)
            if meta:
                self._meta_var.set(
                    f"{_branch_display_name(pick)} | آخر لقطة: {_fmt_local_ts(meta['snapshot_at'])}  |  عدد الصفوف: {meta['row_count']}  |  القيمة: {meta['total_value']:.2f}"
                )
            else:
                self._meta_var.set("لا توجد لقطة مخزون مباشرة لهذا الاسم (قد يكون محفوظًا بالـ UUID)")
        self._all_rows = self.db.list_pos_stock_rows(pick)
        self._refresh_filter_values()
        self._apply_filters()
        self._refresh_observability()

    def _refresh_observability(self) -> None:
        rows = self.db.list_branch_health_rows()
        self._health_tree.delete(*self._health_tree.get_children())
        for r in rows:
            self._health_tree.insert(
                "",
                tk.END,
                values=(
                    _branch_display_name(r["branch"]),
                    r["status"],
                    _fmt_local_ts(r["snapshot_at"], ""),
                    r["age_min"],
                    r["rows"],
                    "{:.2f}".format(float(r["value"])),
                    r["alert"],
                ),
            )
        _apply_zebra_tags(self._health_tree)

        critical = [r for r in rows if str(r["status"]) in ("CRITICAL", "NO_SNAPSHOT")]
        warning = [r for r in rows if str(r["status"]) == "WARNING"]
        if critical:
            self._alert_var.set(
                "Critical branches: " + ", ".join(_branch_display_name(str(r["branch"])) for r in critical[:6])
            )
        elif warning:
            self._alert_var.set(
                "Warning branches: " + ", ".join(_branch_display_name(str(r["branch"])) for r in warning[:6])
            )
        else:
            self._alert_var.set("All monitored branches are healthy.")

    def _apply_filters(self) -> None:
        q = (self._search_var.get() or "").strip().lower()
        ft = (self._type_var.get() or "").strip()
        fs = (self._school_var.get() or "").strip()
        fc = (self._color_var.get() or "").strip()
        fz = (self._size_var.get() or "").strip()
        self._tree.delete(*self._tree.get_children())
        shown = 0
        total_qty = 0
        total_val = 0.0
        for br, it, sc, cl, sz, price, count in self._all_rows:
            if ft and it != ft:
                continue
            if fs and sc != fs:
                continue
            if fc and cl != fc:
                continue
            if fz and sz != fz:
                continue
            br_ui = _branch_display_name(br)
            if q:
                blob = f"{br_ui} {it} {sc} {cl} {sz}".lower()
                if q not in blob:
                    continue
            value = float(price) * int(count)
            self._tree.insert("", tk.END, values=(br_ui, it, sc, cl, sz, f"{price:.2f}", int(count), f"{value:.2f}"))
            shown += 1
            total_qty += int(count)
            total_val += value
        _apply_zebra_tags(self._tree)
        self._status_var.set(f"يُعرض {shown} صف  |  إجمالي الكمية: {total_qty}  |  إجمالي القيمة: {total_val:.2f}")

    def _clear_filters(self) -> None:
        self._type_var.set("")
        self._school_var.set("")
        self._color_var.set("")
        self._size_var.set("")
        self._search_var.set("")
        self._apply_filters()

    def _open_sync_setup(self) -> None:
        try:
            sync_ui.open_sync_setup(self, self.db.conn)
            self.after(300, self._reload_devices)
        except Exception as ex:
            messagebox.showerror("المزامنة", str(ex), parent=self)

    def _run_sync_and_reload(self) -> None:
        try:
            sync_ui.open_sync_dialog(self, self.db.conn)
            self.after(500, self._reload_devices)
        except Exception as ex:
            messagebox.showerror("المزامنة", str(ex), parent=self)

    def _open_reservation_request(self) -> None:
        sel = self._tree.selection()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر صنفاً من الجدول أولاً.", parent=self)
            return
        vals = self._tree.item(sel[0], "values")
        if not vals or len(vals) < 8:
            messagebox.showwarning("تنبيه", "تعذر قراءة بيانات الصنف المحدد.", parent=self)
            return
        branch_ui, item_type, school, color, size, unit_price_txt, count_txt, _ = vals
        try:
            available = int(float(count_txt))
        except Exception:
            available = 0
        if available <= 0:
            messagebox.showwarning("تنبيه", "هذا الصنف غير متاح للحجز.", parent=self)
            return
        dlg = tk.Toplevel(self)
        dlg.title("طلب حجز للفرع")
        dlg.transient(self)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frm, text=f"الفرع: {branch_ui}", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(frm, text=f"{item_type} | {school} | {color} | {size}").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(frm, text=f"المتاح حالياً: {available}").grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 8))
        ttk.Label(frm, text="اسم العميل:").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        customer_var = tk.StringVar()
        ttk.Entry(frm, textvariable=customer_var, width=28).grid(row=3, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(frm, text="الكمية المطلوبة:").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        qty_var = tk.StringVar(value="1")
        ttk.Entry(frm, textvariable=qty_var, width=12).grid(row=4, column=1, sticky="w", padx=6, pady=4)
        frm.columnconfigure(1, weight=1)

        def _submit() -> None:
            try:
                qty = int((qty_var.get() or "").strip())
            except Exception:
                messagebox.showwarning("تنبيه", "الكمية غير صالحة.", parent=dlg)
                return
            try:
                event_uuid = self.db.create_reservation_request(
                    branch=str(branch_ui),
                    customer=customer_var.get(),
                    item_type=str(item_type),
                    school=str(school),
                    color=str(color),
                    size=str(size),
                    qty=qty,
                    unit_price=float(unit_price_txt or 0),
                )
                dlg.destroy()
                try:
                    sync_ui.run_sync_now(self, self.db.conn, reason="طلب حجز")
                except Exception:
                    pass
                messagebox.showinfo(
                    "تم الإرسال",
                    f"تم إنشاء طلب الحجز بنجاح.\nرقم الحدث: {event_uuid[:8]}\nسيتم إرسال الطلب فورًا عبر مزامنة تلقائية.",
                    parent=self,
                )
            except Exception as ex:
                messagebox.showerror("فشل طلب الحجز", str(ex), parent=dlg)

        btns = ttk.Frame(frm)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="إرسال الطلب", command=_submit).pack(side=tk.RIGHT, padx=6)

    def _on_close(self) -> None:
        self.db.close()
        self.destroy()


def main() -> None:
    app = StockMonitorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
