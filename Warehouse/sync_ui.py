# -*- coding: utf-8 -*-
"""Sync dialogs for the warehouse and POS apps.

Two Toplevel dialogs:

- SyncSetupDialog  — one-time setup: server URL, device name, API key.
- SyncDialog       — status + manual "مزامنة الآن" button, with a live
                     log fed by a background worker thread.

Both dialogs are RTL-friendly (pure ttk, no custom fonts) and follow
the existing app's dialog conventions: transient + grab_set + modal.

Threading contract
------------------
The sync cycle must never run on the Tk main loop (it does blocking
HTTP). SyncDialog spawns a threading.Thread that calls
SyncClient.run_cycle(progress=...) and the progress callback drops
messages into a thread-safe queue.Queue. The main loop polls the
queue via root.after(100, ...) and appends messages to the log.
"""

from __future__ import annotations

import queue
import sqlite3
import threading
import tkinter as tk
from datetime import datetime
from tkinter import messagebox, ttk
from typing import Any, Dict, List, Optional

import sync_client
import sync_core


# ------------------------------- helpers -------------------------------- #

def _fmt(value: Optional[str], empty: str = "—") -> str:
    raw = str(value) if value not in (None, "", "None") else ""
    if not raw:
        return empty
    try:
        txt = raw.strip()
        if txt.endswith("Z"):
            dt = datetime.fromisoformat(txt[:-1] + "+00:00")
            return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
        if "T" in txt:
            dt = datetime.fromisoformat(txt)
            if dt.tzinfo is not None:
                return dt.astimezone().strftime("%Y-%m-%d %H:%M:%S")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    return raw.replace("T", " ")


def _notify_host_synced(master: tk.Misc) -> None:
    host = getattr(master, "_app_controller", None) or None
    cur = master
    for _ in range(20):
        if cur is None:
            break
        if hasattr(cur, "_on_sync_completed") or hasattr(cur, "_refresh_current_tab"):
            host = cur
            break
        cur = getattr(cur, "master", None)
    if host is None:
        try:
            host = master.winfo_toplevel()
        except Exception:
            host = master

    for name in ("_on_sync_completed", "_refresh_current_tab", "_shortcut_refresh", "_on_tab_changed"):
        fn = getattr(host, name, None)
        if not callable(fn):
            continue
        try:
            fn()
            break
        except TypeError:
            try:
                fn(None)
                break
            except Exception:
                pass
        except Exception:
            pass
    try:
        host.event_generate("<<HosnySyncCompleted>>", when="tail")
    except Exception:
        pass


def open_sync_received_details(master: tk.Misc, summary: Dict[str, Any]) -> None:
    """Modal list of inbound / apply results from the last sync cycle."""
    dlg = tk.Toplevel(master)
    dlg.title("تفاصيل ما وصل من المزامنة")
    dlg.transient(master.winfo_toplevel())
    dlg.grab_set()
    dlg.geometry("900x520")

    frm = ttk.Frame(dlg, padding=10)
    frm.pack(fill=tk.BOTH, expand=True)

    pulled = int(summary.get("pulled") or 0)
    pushed = int(summary.get("pushed") or 0)
    ttk.Label(
        frm,
        text=f"رفع: {pushed}  •  تنزيل من الخادم: {pulled}  •  تطبيق: {int(summary.get('applied') or 0)}"
        f"  •  تخطي: {int(summary.get('skipped') or 0)}  •  أخطاء تطبيق: {int(summary.get('apply_errors') or 0)}"
        f"  •  seq={summary.get('next_seq')}",
        wraplength=860,
    ).pack(anchor="w", pady=(0, 8))

    nb = ttk.Notebook(frm)
    nb.pack(fill=tk.BOTH, expand=True)

    cols = ("seq", "etype", "src", "uuid", "summary", "note")

    def _tab(title: str, rows: List[Dict[str, Any]], note_field: Optional[str]) -> None:
        tab = ttk.Frame(nb, padding=4)
        nb.add(tab, text=title)
        tv = ttk.Treeview(tab, columns=cols, show="headings", height=14)
        heads = {
            "seq": "تسلسل",
            "etype": "نوع الحدث",
            "src": "من جهاز",
            "uuid": "المعرّف",
            "summary": "ملخص",
            "note": "ملاحظة / خطأ",
        }
        widths = (70, 130, 110, 120, 320, 220)
        for c, w in zip(cols, widths):
            tv.heading(c, text=heads.get(c, c))
            tv.column(c, width=w, anchor="center")
        ys = ttk.Scrollbar(tab, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=ys.set)
        tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ys.pack(side=tk.RIGHT, fill=tk.Y)
        for r in rows:
            u = str(r.get("event_uuid") or "")
            udisp = (u[:12] + "…") if len(u) > 12 else u
            note = ""
            if note_field == "reason":
                note = str(r.get("reason") or "")
            elif note_field == "error":
                note = str(r.get("error") or "")
            else:
                note = ""
            tv.insert("", tk.END, values=(
                r.get("server_seq"),
                r.get("event_type"),
                r.get("source_device") or "",
                udisp,
                r.get("summary") or "",
                note,
            ))

    applied = summary.get("applied_events") or []
    skipped = summary.get("skipped_events") or []
    errors = summary.get("error_events") or []

    _tab("تم التطبيق", applied, None)
    _tab("تم التخطي", skipped, "reason")
    _tab("أخطاء التطبيق", errors, "error")

    bf = ttk.Frame(frm)
    bf.pack(fill=tk.X, pady=(8, 0))
    ttk.Button(bf, text="إغلاق", command=dlg.destroy).pack(side=tk.RIGHT)


def present_sync_cycle_summary(master: tk.Misc, summary: Dict[str, Any]) -> None:
    """UI feedback after a successful sync cycle (periodic or manual)."""
    try:
        _notify_host_synced(master)
    except Exception:
        pass

    pulled = int(summary.get("pulled") or 0)
    applied = int(summary.get("applied") or 0)
    skipped = int(summary.get("skipped") or 0)
    errc = int(summary.get("apply_errors") or 0)

    if errc:
        errs = summary.get("error_events") or []
        first = (errs[0].get("error") if errs else "") or ""
        messagebox.showwarning(
            "مزامنة — أخطاء تطبيق",
            f"فشل تطبيق {errc} حدث/أحداث.\n"
            f"أول خطأ: {first}\n\n"
            "افتح «تفاصيل ما وصل» من الإشعار أو نافذة المزامنة لمراجعة القائمة.",
            parent=master.winfo_toplevel(),
        )


def present_sync_cycle_failure(master: tk.Misc, err_text: str) -> None:
    messagebox.showerror("فشل المزامنة", str(err_text), parent=master.winfo_toplevel())


# ------------------------------- setup ---------------------------------- #

class SyncSetupDialog(tk.Toplevel):
    """One-time setup: server URL, device name, optional API key.

    Saves into device_identity. Offers a "Test connection" button that
    calls /v1/health + /v1/auth/token so the user can confirm their
    values before closing.
    """

    def __init__(self, master: tk.Misc, db_conn: sqlite3.Connection) -> None:
        super().__init__(master)
        self.db_conn = db_conn
        self.title("إعدادات المزامنة")
        self.geometry("520x360")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        cfg = sync_client.load_sync_config(db_conn)
        self._var_url = tk.StringVar(value=cfg.get("server_url") or "http://127.0.0.1:8000")
        self._var_name = tk.StringVar(value=cfg.get("device_name") or "")
        self._var_key = tk.StringVar(value=cfg.get("api_token") or "")
        self._role = cfg.get("device_role") or "—"
        self._build()

    def _build(self) -> None:
        frm = ttk.Frame(self, padding=14)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="إعدادات مزامنة الجهاز",
                  font=("Segoe UI", 13, "bold")).pack(anchor="e", pady=(0, 10))

        info = ttk.Label(
            frm,
            text=(
                "Enter the sync server URL and this device name (for example POS-01). "
                "API key is optional with the new server, and only needed for old compatibility flows."
            ),
            wraplength=460, justify="right",
        )
        info.pack(anchor="e", pady=(0, 10))

        grid = ttk.Frame(frm)
        grid.pack(fill=tk.X, pady=2)

        ttk.Label(grid, text="عنوان الخادم:").grid(row=0, column=1, sticky="e", pady=4)
        ttk.Entry(grid, textvariable=self._var_url, width=46, justify="left") \
            .grid(row=0, column=0, sticky="w", padx=(6, 0), pady=4)

        ttk.Label(grid, text="اسم الجهاز:").grid(row=1, column=1, sticky="e", pady=4)
        ttk.Entry(grid, textvariable=self._var_name, width=46, justify="left") \
            .grid(row=1, column=0, sticky="w", padx=(6, 0), pady=4)


        ttk.Label(grid, text="API key (optional):").grid(row=2, column=1, sticky="e", pady=4)
        ttk.Entry(grid, textvariable=self._var_key, width=46, justify="left",
                  show="•") \
            .grid(row=2, column=0, sticky="w", padx=(6, 0), pady=4)

        ttk.Label(grid, text="الدور (ثابت):").grid(row=3, column=1, sticky="e", pady=4)
        ttk.Label(grid, text=self._role, foreground="#64748b") \
            .grid(row=3, column=0, sticky="w", padx=(6, 0), pady=4)

        self._status_var = tk.StringVar(value="")
        ttk.Label(frm, textvariable=self._status_var, foreground="#0ea5e9") \
            .pack(anchor="e", pady=(12, 0))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(14, 0))
        ttk.Button(btns, text="إلغاء", command=self.destroy).pack(side=tk.LEFT)
        ttk.Button(btns, text="اختبار الاتصال", command=self._test).pack(side=tk.RIGHT, padx=(0, 6))
        ttk.Button(btns, text="حفظ", command=self._save).pack(side=tk.RIGHT)

    # ---- actions ----

    def _collect(self) -> Optional[Dict[str, str]]:
        url = self._var_url.get().strip().rstrip("/")
        name = self._var_name.get().strip()
        key = self._var_key.get().strip()
        if not url or not url.lower().startswith(("http://", "https://")):
            messagebox.showwarning("إعدادات المزامنة", "عنوان الخادم غير صالح.", parent=self)
            return None
        if not name:
            messagebox.showwarning("إعدادات المزامنة", "اسم الجهاز مطلوب.", parent=self)
            return None
        return {"url": url, "name": name, "key": key}

    def _save(self) -> None:
        d = self._collect()
        if not d:
            return
        sync_client.save_setup(
            self.db_conn,
            server_url=d["url"],
            device_name=d["name"],
            api_token=d["key"],
        )
        self._status_var.set("تم الحفظ.")
        messagebox.showinfo("إعدادات المزامنة", "تم حفظ إعدادات المزامنة.", parent=self)
        self.destroy()

    def _test(self) -> None:
        d = self._collect()
        if not d:
            return
        # Save first so SyncClient reads the fresh values.
        sync_client.save_setup(
            self.db_conn,
            server_url=d["url"],
            device_name=d["name"],
            api_token=d["key"],
        )
        self._status_var.set("جاري الاختبار...")
        self.update_idletasks()
        client = sync_client.SyncClient(self.db_conn)
        try:
            result = client.test_connection()
        except sync_client.SyncError as e:
            self._status_var.set("فشل الاتصال")
            messagebox.showerror("إعدادات المزامنة", f"فشل الاتصال: {e}", parent=self)
            return
        self._status_var.set("اتصال ناجح ✔")
        messagebox.showinfo(
            "إعدادات المزامنة",
            "تم الاتصال بالخادم بنجاح.\n\n"
            f"حالة الخادم: {result.get('health', {}).get('status', 'ok')}\n"
            f"الرمز: {result.get('token_prefix', '—')}",
            parent=self,
        )


# ------------------------------- sync main ------------------------------ #

class SyncDialog(tk.Toplevel):
    """Main sync dashboard. Shows config summary + runs manual sync.

    Layout:
        +------------------------------------------------+
        | العنوان: http://...                            |
        | الجهاز:  POS-01   الدور: pos                  |
        | آخر رفع: ...  آخر تنزيل: ...                   |
        | قيد الانتظار: 12    تم تنزيل حتى: seq 37       |
        | --------------------------------------------- |
        | [ مزامنة الآن ]  [ اختبار ]  [ الإعدادات ]     |
        | --------------------------------------------- |
        | <scrollable log>                               |
        +------------------------------------------------+
    """

    POLL_MS = 100

    def __init__(self, master: tk.Misc, db_conn: sqlite3.Connection) -> None:
        super().__init__(master)
        self.db_conn = db_conn
        self.title("مزامنة البيانات")
        self.geometry("640x480")
        self.minsize(560, 420)
        self.transient(master)
        self.grab_set()

        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None

        self._build()
        self._refresh_state()
        self.after(self.POLL_MS, self._pump_queue)

    # ---- layout ----

    def _build(self) -> None:
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="مزامنة البيانات مع الخادم",
                  font=("Segoe UI", 13, "bold")).pack(anchor="e", pady=(0, 8))

        card = ttk.LabelFrame(frm, text="حالة الجهاز", padding=10)
        card.pack(fill=tk.X)

        self._lbl_server = tk.StringVar()
        self._lbl_device = tk.StringVar()
        self._lbl_counts = tk.StringVar()
        self._lbl_last   = tk.StringVar()

        row = ttk.Frame(card)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="الخادم:", width=12, anchor="e").pack(side=tk.RIGHT)
        ttk.Label(row, textvariable=self._lbl_server, foreground="#0f172a").pack(side=tk.RIGHT, padx=(0, 8))

        row = ttk.Frame(card)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="الجهاز:", width=12, anchor="e").pack(side=tk.RIGHT)
        ttk.Label(row, textvariable=self._lbl_device, foreground="#0f172a").pack(side=tk.RIGHT, padx=(0, 8))

        row = ttk.Frame(card)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="قيد الانتظار:", width=12, anchor="e").pack(side=tk.RIGHT)
        ttk.Label(row, textvariable=self._lbl_counts, foreground="#0f172a").pack(side=tk.RIGHT, padx=(0, 8))

        row = ttk.Frame(card)
        row.pack(fill=tk.X, pady=2)
        ttk.Label(row, text="آخر مزامنة:", width=12, anchor="e").pack(side=tk.RIGHT)
        ttk.Label(row, textvariable=self._lbl_last, foreground="#475569").pack(side=tk.RIGHT, padx=(0, 8))

        btns = ttk.Frame(frm)
        btns.pack(fill=tk.X, pady=(12, 6))
        self._btn_setup = ttk.Button(btns, text="الإعدادات", command=self._open_setup)
        self._btn_setup.pack(side=tk.LEFT)
        self._btn_test = ttk.Button(btns, text="اختبار الاتصال", command=self._test_connection)
        self._btn_test.pack(side=tk.LEFT, padx=6)
        self._btn_sync = ttk.Button(btns, text="مزامنة الآن", command=self._start_sync)
        self._btn_sync.pack(side=tk.RIGHT)

        log_frame = ttk.LabelFrame(frm, text="سجل المزامنة", padding=6)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self._log = tk.Text(
            log_frame, height=10, wrap="word",
            font=("Consolas", 9),
            background="#0f172a", foreground="#e2e8f0",
            insertbackground="#e2e8f0", state="disabled",
        )
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self._log.yview)
        self._log.configure(yscrollcommand=scroll.set)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # ---- state ----

    def _refresh_state(self) -> None:
        cfg = sync_client.load_sync_config(self.db_conn)
        self._lbl_server.set(_fmt(cfg.get("server_url")))
        self._lbl_device.set(
            f"{_fmt(cfg.get('device_name'))}  ({_fmt(cfg.get('device_role'))})"
        )
        try:
            pending = sync_core.outbox_pending_count(self.db_conn)
        except Exception:
            pending = 0
        try:
            dead = sync_core.dead_letter_count(self.db_conn)
        except Exception:
            dead = 0
        extra = f"  •  DLQ={dead}" if dead else ""
        self._lbl_counts.set(f"{pending} حدث غير مُرسَل{extra}")
        last_push = _fmt(cfg.get("last_push_at"))
        last_pull = _fmt(cfg.get("last_pull_at"))
        seq = cfg.get("last_pulled_seq") or 0
        self._lbl_last.set(f"رفع: {last_push}  |  تنزيل: {last_pull}  |  seq={seq}")

    def _append_log(self, msg: str, kind: str = "info") -> None:
        color = {
            "info":  "#e2e8f0",
            "ok":    "#4ade80",
            "err":   "#f87171",
            "warn":  "#facc15",
        }.get(kind, "#e2e8f0")
        self._log.configure(state="normal")
        self._log.insert("end", msg + "\n", kind)
        self._log.tag_configure(kind, foreground=color)
        self._log.see("end")
        self._log.configure(state="disabled")

    # ---- actions ----

    def _open_setup(self) -> None:
        SyncSetupDialog(self, self.db_conn).wait_window()
        self._refresh_state()

    def _test_connection(self) -> None:
        client = sync_client.SyncClient(self.db_conn)
        try:
            result = client.test_connection()
        except sync_client.SyncError as e:
            self._append_log(f"اختبار الاتصال: فشل ({e})", "err")
            return
        self._append_log(
            f"اختبار الاتصال: ناجح ({result.get('token_prefix', '—')})", "ok"
        )

    def _start_sync(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            return
        self._btn_sync.configure(state="disabled")
        self._btn_setup.configure(state="disabled")
        self._btn_test.configure(state="disabled")
        self._append_log("── بدء دورة مزامنة ──", "info")

        db_path = self.db_conn.execute("PRAGMA database_list").fetchone()[2]

        def worker() -> None:
            # Background thread MUST open its own connection — sqlite3
            # connections are not thread-safe when isolation is auto.
            conn = sqlite3.connect(db_path, timeout=30.0, isolation_level=None, check_same_thread=False)
            conn.row_factory = sqlite3.Row
            conn.execute("PRAGMA busy_timeout=30000;")
            try:
                client = sync_client.SyncClient(conn)
                summary = client.run_cycle(progress=lambda m: self._q.put(("log", m, "info")))
                self._q.put(("done", summary))
            except sync_client.SyncError as e:
                self._q.put(("error", str(e)))
            except Exception as e:
                self._q.put(("error", f"unexpected: {e}"))
            finally:
                try:
                    conn.close()
                except Exception:
                    pass

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _pump_queue(self) -> None:
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._append_log(msg[1], msg[2])
                elif kind == "done":
                    summary = msg[1]
                    self._append_log(
                        f"تمت المزامنة • رفع {summary['pushed']} "
                        f"• تنزيل {summary['pulled']} "
                        f"• تخطي {summary.get('skipped', 0)} "
                        f"• DLQ={summary.get('dead_lettered', 0)} "
                        f"• seq={summary['next_seq']} "
                        f"• cycle={summary.get('cycle_id', '-')}",
                        "ok",
                    )
                    self._enable_buttons()
                    self._refresh_state()
                    try:
                        present_sync_cycle_summary(self.master, summary)
                    except Exception:
                        _notify_host_synced(self.master)
                elif kind == "error":
                    self._append_log(f"فشل المزامنة: {msg[1]}", "err")
                    self._enable_buttons()
                    self._refresh_state()
                    try:
                        present_sync_cycle_failure(self.master, str(msg[1]))
                    except Exception:
                        pass
        except queue.Empty:
            pass
        self.after(self.POLL_MS, self._pump_queue)

    def _enable_buttons(self) -> None:
        self._btn_sync.configure(state="normal")
        self._btn_setup.configure(state="normal")
        self._btn_test.configure(state="normal")


# --------------------- module-level open helpers ---------------------- #

def open_sync_dialog(master: tk.Misc, db_conn: sqlite3.Connection) -> None:
    """Open the main sync dialog. Safe to call from a menu/button."""
    SyncDialog(master, db_conn)


def open_sync_setup(master: tk.Misc, db_conn: sqlite3.Connection) -> None:
    """Open the setup dialog directly."""
    SyncSetupDialog(master, db_conn)


def run_sync_now(master: tk.Misc, db_conn: sqlite3.Connection, *, reason: str = "") -> None:
    """Run one sync cycle immediately in background.

    Useful for "send + sync now" workflows after enqueuing business events.
    """
    host = master.winfo_toplevel()
    note = f" ({reason})" if str(reason or "").strip() else ""
    try:
        db_path = db_conn.execute("PRAGMA database_list").fetchone()[2]
    except Exception as e:
        messagebox.showerror("المزامنة", f"تعذر بدء المزامنة{note}:\n{e}", parent=host)
        return

    q: "queue.Queue[tuple]" = queue.Queue()

    def worker() -> None:
        conn = sqlite3.connect(db_path, timeout=30.0, isolation_level=None, check_same_thread=False)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA busy_timeout=30000;")
        try:
            client = sync_client.SyncClient(conn)
            summary = client.run_cycle(progress=None)
            q.put(("done", summary))
        except Exception as ex:
            q.put(("error", str(ex)))
        finally:
            try:
                conn.close()
            except Exception:
                pass

    threading.Thread(target=worker, daemon=True).start()

    def pump() -> None:
        try:
            kind, payload = q.get_nowait()
        except queue.Empty:
            host.after(120, pump)
            return
        if kind == "done":
            try:
                present_sync_cycle_summary(host, payload)
            except Exception:
                _notify_host_synced(host)
        else:
            present_sync_cycle_failure(host, str(payload))

    host.after(120, pump)
