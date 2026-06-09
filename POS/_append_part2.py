
# Helper script to append Part 2 (UI helpers + IncomeFrame) to HosnyWarehouse.py
import os

TARGET = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\HosnyWarehouse.py"

PART2 = """

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

        self.btn = ttk.Button(row, text="\u25bc", width=2, takefocus=0, command=self._on_caret_click)
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
        self.btn = ttk.Button(row, text="\U0001f4c5", width=3, command=self._open)
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
            tp.focus_force()
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
        ttk.Button(hdr, text="\u25c4", width=2, command=self._prev_month).pack(side=tk.LEFT)
        self._title = ttk.Label(hdr, font=("", 10, "bold")); self._title.pack(side=tk.LEFT, padx=6)
        ttk.Button(hdr, text="\u25ba", width=2, command=self._next_month).pack(side=tk.RIGHT)

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

            ttk.Label(cell, text="\u0627\u0644\u0643\u0645\u064a\u0629", font=("", 8)).pack()
            v_qty = tk.StringVar()
            ttk.Entry(cell, textvariable=v_qty, width=6).pack()

            ttk.Label(cell, text="\u0627\u0644\u0633\u0639\u0631", font=("", 8)).pack()
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
"""

with open(TARGET, "a", encoding="utf-8") as f:
    f.write(PART2)

print("Part2 UI helpers written OK")
