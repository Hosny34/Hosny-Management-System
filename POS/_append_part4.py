
TARGET = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\HosnyWarehouse.py"

PART4 = """

# ------------------- POS Frame (replaces OutcomeFrame) -------------------

class POSFrame(ttk.Frame):
    \"\"\"
    Multi-bill POS with step-by-step item navigation.
    Bills 1-4 are regular sales; bill 5 is a reservation.
    \"\"\"
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=6)
        self.db = db

        # In-memory bill state
        self.bills = {1: [], 2: [], 3: [], 4: [], 5: []}
        self.customers = {1: "", 2: "", 3: "", 4: "", 5: ""}
        self.active_bill = 1

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

        # Quick filter bar
        fbar = ttk.LabelFrame(left, text="تصفية سريعة")
        fbar.pack(fill=tk.X, padx=4, pady=(4, 2))

        self._flt_school = LabeledCombobox(fbar, "المدرسة", self.db, "school")
        self._flt_school.set_supplier(lambda: self.db.get_distinct_filtered("school", {}))
        self._flt_school.grid(row=0, column=0, padx=4, pady=4, sticky="ew")

        self._flt_item = LabeledCombobox(fbar, "النوع", self.db, "item_type")
        self._flt_item.set_supplier(lambda: self.db.get_distinct_filtered("item_type", {}))
        self._flt_item.grid(row=0, column=1, padx=4, pady=4, sticky="ew")

        self._flt_color = LabeledCombobox(fbar, "اللون", self.db, "color")
        self._flt_color.set_supplier(lambda: self.db.get_distinct_filtered("color", {}))
        self._flt_color.grid(row=0, column=2, padx=4, pady=4, sticky="ew")

        ttk.Button(fbar, text="تطبيق", command=self._apply_quick_filter).grid(row=0, column=3, padx=4, pady=4)
        ttk.Button(fbar, text="مسح", command=self._clear_quick_filter).grid(row=0, column=4, padx=4, pady=4)

        for c in range(5):
            fbar.columnconfigure(c, weight=1 if c < 3 else 0)

        # Breadcrumb
        crumb = ttk.Frame(left)
        crumb.pack(fill=tk.X, padx=4, pady=(2, 0))
        self._crumb_var = tk.StringVar(value="اختر النوع")
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

        # ---- RIGHT: bill management ----
        right = ttk.LabelFrame(vsplit, text="الفواتير")
        vsplit.add(right, weight=4)

        # Bill switcher buttons (1-5)
        switcher = ttk.Frame(right)
        switcher.pack(fill=tk.X, padx=4, pady=(4, 2))
        self._bill_btns: Dict[int, ttk.Button] = {}
        for n in range(1, 6):
            label = str(n) if n < 5 else "5 (حجز)"
            btn = ttk.Button(switcher, text=label, width=8,
                             command=lambda b=n: self._switch_bill(b))
            btn.pack(side=tk.LEFT, padx=2)
            self._bill_btns[n] = btn
        self._update_bill_btn_styles()

        # Customer field
        self._customer_var = tk.StringVar()
        cust_row = ttk.Frame(right)
        cust_row.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Label(cust_row, text="العميل:").pack(side=tk.LEFT)
        self._cust_entry = ttk.Entry(cust_row, textvariable=self._customer_var)
        self._cust_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 0))
        self._cust_entry.bind("<FocusOut>", lambda e: self._save_customer(), add="+")

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
        self.total_var = tk.StringVar(value="0.00")
        ttk.Label(tot_row, textvariable=self.total_var, font=("Segoe UI", 12, "bold")).pack(side=tk.RIGHT, padx=(0, 6))

        # Reservation extras (only visible for bill 5)
        self._res_frame = ttk.LabelFrame(right, text="بيانات الحجز")
        self._res_frame.pack(fill=tk.X, padx=4, pady=(2, 2))
        ttk.Label(self._res_frame, text="المبلغ المدفوع:").grid(row=0, column=0, padx=4, pady=4, sticky="e")
        self._paid_var = tk.StringVar(value="0")
        ttk.Entry(self._res_frame, textvariable=self._paid_var, width=12).grid(row=0, column=1, padx=4, pady=4, sticky="w")
        ttk.Label(self._res_frame, text="ملاحظة:").grid(row=1, column=0, padx=4, pady=4, sticky="e")
        self._res_note_var = tk.StringVar()
        ttk.Entry(self._res_frame, textvariable=self._res_note_var, width=24).grid(row=1, column=1, padx=4, pady=4, sticky="ew")
        self._res_frame.columnconfigure(1, weight=1)

        # Finalize button
        ttk.Button(right, text="تأكيد الفاتورة / الحجز", command=self._finalize).pack(fill=tk.X, padx=4, pady=(4, 4))

        self._update_res_frame_visibility()

        # Initial render
        self._render_items()

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
    def _render_items(self):
        self._sel_item = None
        self._sel_color = None
        self._sel_size = None
        self._crumb_var.set("اختر النوع")
        self._clear_grid()

        school_f = self._flt_school.get() or None
        try:
            items = sorted({
                r["item_type"]
                for r in self.db.current_inventory({"school": school_f} if school_f else {})
                if r.get("item_type")
            })
        except Exception:
            items = []

        self._mk_grid_buttons(items, self._select_item, cols=4)

    def _select_item(self, item_type: str):
        self._sel_item = item_type
        self._price_user_edited = False
        self._render_schools()

    def _render_schools(self):
        self._sel_school = None
        self._sel_color = None
        self._sel_size = None

        if self._sel_item:
            self._crumb_var.set(f"النوع: {self._sel_item}  \\u27f6  اختر المدرسة")
        else:
            self._crumb_var.set("اختر مدرسة")

        self._clear_grid()
        color_f = self._flt_color.get() or None
        try:
            constraints = {"item_type": self._sel_item}
            if color_f:
                constraints["color"] = color_f
            schools = self.db.get_distinct_filtered("school", constraints)
        except Exception:
            schools = []

        self._mk_grid_buttons(schools, self._select_school, cols=4)
        ttk.Button(self._grid_host, text="\\u25c4 رجوع إلى الأنواع", command=self._render_items)\
            .pack(anchor="w", padx=4, pady=4)

    def _select_school(self, school: str):
        self._sel_school = school
        self._price_user_edited = False
        self._render_colors()

    def _render_colors(self):
        self._sel_color = None
        self._sel_size = None
        self._crumb_var.set(f"المدرسة: {self._sel_school}  \\u27f6  النوع: {self._sel_item}  \\u27f6  اختر اللون")
        self._clear_grid()

        item_f = self._flt_item.get() or self._sel_item
        try:
            pairs = self.db.list_items_for_school(self._sel_school or "")
            colors = sorted({cl for (it, cl) in pairs if it == item_f})
        except Exception:
            colors = []

        self._mk_grid_buttons(colors, self._select_color, cols=4)
        ttk.Button(self._grid_host, text="\\u25c4 رجوع إلى المدارس", command=self._render_schools)\
            .pack(anchor="w", padx=4, pady=4)

    def _select_color(self, color: str):
        self._sel_color = color
        self._price_user_edited = False
        self._render_sizes()

    def _render_sizes(self):
        self._sel_size = None
        self._crumb_var.set(
            f"المدرسة: {self._sel_school}  \\u27f6  النوع: {self._sel_item}  \\u27f6  اللون: {self._sel_color}  \\u27f6  اختر المقاس"
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
                sz = str(r.get("size") or "").strip()
                if not sz:
                    continue
                stock_rows.setdefault(sz, {"count": 0, "last_price": r.get("unit_price")})
                stock_rows[sz]["count"] += int(r.get("count") or 0)
        except Exception:
            pass

        self._sizes_cache = []
        for sz in raw_sizes:
            r = stock_rows.get(sz, {})
            self._sizes_cache.append({
                "size": sz,
                "count": r.get("count", 0),
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
            style = "SizeZero.TButton" if cnt == 0 else "Size.TButton"

            btn = ttk.Button(row, text=label, style=style)
            # Double-click adds to active bill directly
            btn.configure(command=lambda v=label, b=btn: self._on_size_click(v, b))
            btn.bind("<Double-Button-1>", lambda e, v=label, b=btn: self._on_size_double_click(v, b))
            btn.pack(side=tk.LEFT, padx=3, pady=3)

            self._size_btns[label] = (btn, cnt == 0)

        ttk.Button(self._grid_host, text="\\u25c4 رجوع إلى الألوان", command=self._render_colors)\
            .pack(anchor="w", padx=4, pady=4)

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

        # Find price
        price = self._compute_price_for_size(size)
        if price is None:
            price = 0.0

        bill_lines = self.bills[self.active_bill]
        # Merge if same specs and price
        for existing in bill_lines:
            if (existing.get("item_type") == self._sel_item
                    and existing.get("school") == self._sel_school
                    and existing.get("color") == self._sel_color
                    and existing.get("size") == size
                    and abs(float(existing.get("unit_price", 0)) - price) < 0.001):
                existing["qty"] = int(existing["qty"]) + qty
                self._sync_bill_table()
                return

        bill_lines.append({
            "item_type": self._sel_item,
            "school": self._sel_school,
            "color": self._sel_color,
            "size": size,
            "unit_price": price,
            "qty": qty,
            "allow_factory_fill": False,
            "has_badge": 0,
        })
        self._sync_bill_table()

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

        if profile:
            r1s, r1e, r2s, r2e, has_alpha = profile
            if r1s is not None and r1e is not None:
                labels = ALLOWED_NUMERIC_RANGES.get((r1s, r1e))
                if labels:
                    sizes.extend(labels)
            if r2s is not None and r2e is not None:
                labels = ALLOWED_NUMERIC_RANGES.get((r2s, r2e))
                if labels:
                    sizes.extend(labels)
            if has_alpha:
                sizes.extend(ALPHA_SIZES)
            return sizes

        try:
            rows = self.db.current_inventory({"school": school, "item_type": item, "color": color})
            seen: set = set()
            for r in rows:
                sz = str(r.get("size") or "").strip()
                if sz and sz not in seen:
                    seen.add(sz)
                    sizes.append(sz)
            return sizes
        except Exception:
            return []

    # ------------------------------------------------------------------ quick filter
    def _apply_quick_filter(self):
        sc = self._flt_school.get()
        it = self._flt_item.get()
        cl = self._flt_color.get()

        if sc and it and cl:
            self._sel_item = it
            self._sel_school = sc
            self._sel_color = cl
            self._render_sizes()
        elif sc and it:
            self._sel_item = it
            self._sel_school = sc
            self._render_colors()
        elif sc:
            self._sel_item = None
            self._render_schools()
        else:
            self._render_items()

    def _clear_quick_filter(self):
        self._flt_school.set("")
        self._flt_item.set("")
        self._flt_color.set("")
        self._render_items()

    # ------------------------------------------------------------------ bill management
    def _switch_bill(self, n: int):
        # Save current customer
        self.customers[self.active_bill] = self._customer_var.get()
        self.active_bill = n
        # Restore customer for this bill
        self._customer_var.set(self.customers[n])
        self._update_bill_btn_styles()
        self._sync_bill_table()
        self._update_res_frame_visibility()

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

    def _update_res_frame_visibility(self):
        if self.active_bill == 5:
            self._res_frame.pack(fill=tk.X, padx=4, pady=(2, 2))
        else:
            self._res_frame.pack_forget()

    def _sync_bill_table(self):
        self.bill_table.delete(*self.bill_table.get_children())
        total = 0.0
        for idx, ln in enumerate(self.bills[self.active_bill]):
            line_total = float(ln["unit_price"]) * int(ln["qty"])
            total += line_total
            self.bill_table.insert(
                "", tk.END, iid=str(idx),
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"],
                        f"{float(ln['unit_price']):.2f}", ln["qty"], f"{line_total:.2f}")
            )
        self.total_var.set(f"{total:.2f}")

    def _inc_qty(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        lines = self.bills[self.active_bill]
        if 0 <= idx < len(lines):
            lines[idx]["qty"] = int(lines[idx]["qty"]) + 1
            self._sync_bill_table()
            self.bill_table.selection_set(str(idx))

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

    def _remove_line(self):
        sel = self.bill_table.selection()
        if not sel:
            return
        idx = int(sel[0])
        lines = self.bills[self.active_bill]
        if 0 <= idx < len(lines):
            lines.pop(idx)
            self._sync_bill_table()

    def _clear_bill(self):
        self.bills[self.active_bill].clear()
        self._sync_bill_table()

    # ------------------------------------------------------------------ finalize
    def _finalize(self):
        lines = self.bills[self.active_bill]
        if not lines:
            messagebox.showwarning("فارغ", "لا توجد أصناف في الفاتورة.")
            return

        self._save_customer()
        customer = (self.customers[self.active_bill] or "").strip()

        if self.active_bill == 5:
            # Reservation
            try:
                paid = float((self._paid_var.get() or "0").strip())
            except Exception:
                paid = 0.0
            note = self._res_note_var.get().strip()
            # Add note to each line
            for ln in lines:
                ln["note"] = note
            try:
                ids = self.db.create_reservation(customer, lines, paid_amount=paid)
                messagebox.showinfo("تم الحجز", f"تم إنشاء {len(ids)} حجز/حجوزات بنجاح.", parent=self)
                self.bills[5].clear()
                self._paid_var.set("0")
                self._res_note_var.set("")
                self._sync_bill_table()
            except Exception as ex:
                messagebox.showerror("فشل الحجز", str(ex), parent=self)
        else:
            # Regular sale
            try:
                bill_id = self.db.create_bill(customer, lines)
            except Exception as ex:
                messagebox.showerror("فشل الحفظ", str(ex), parent=self)
                return

            total = self.total_var.get()
            self.bills[self.active_bill].clear()
            self._sync_bill_table()

            if messagebox.askyesno("طباعة الفاتورة",
                                   f"تم إنشاء الفاتورة #{bill_id} (الإجمالي: {total}).\\nطباعة الفاتورة؟"):
                self._direct_print_bill(bill_id)

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


# ------------------- Factory Item Dialog -------------------

class FactoryItemDialog(tk.Toplevel):
    \"\"\"
    Create an ad-hoc bill line that ships directly from the factory.
    \"\"\"
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

        self.whv = tk.StringVar(value=preset.get("warehouse_no",""))
        self.pkv = tk.StringVar(value=preset.get("package_no",""))

        row_wp = ttk.Frame(frm); row_wp.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(0,6))
        ttk.Label(row_wp, text="المخزن:").pack(side=tk.LEFT)
        ttk.Entry(row_wp, textvariable=self.whv, width=8).pack(side=tk.LEFT, padx=(4,10))
        ttk.Label(row_wp, text="العبوة:").pack(side=tk.LEFT)
        ttk.Entry(row_wp, textvariable=self.pkv, width=10).pack(side=tk.LEFT, padx=(4,0))

        self.qv = tk.StringVar(value=preset.get("qty","1"))
        self.pv = tk.StringVar(value=preset.get("unit_price",""))
        self.badge = tk.BooleanVar(value=bool(int(preset.get("has_badge","0") or 0)))

        grid2 = ttk.Frame(frm); grid2.grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        ttk.Label(grid2, text="الكمية:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        ttk.Entry(grid2, textvariable=self.qv, width=12).grid(row=0, column=1, sticky="w")
        ttk.Label(grid2, text="سعر الوحدة:").grid(row=0, column=2, sticky="e", padx=12, pady=4)
        ttk.Entry(grid2, textvariable=self.pv, width=12).grid(row=0, column=3, sticky="w")
        ttk.Checkbutton(grid2, text="بادج", variable=self.badge).grid(row=0, column=4, sticky="w", padx=(12,0))

        btns = ttk.Frame(frm); btns.grid(row=4, column=0, columnspan=2, sticky="e", padx=6, pady=(10,0))
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
                self.pv.set(f"{p:.2f}")
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
            "allow_factory_fill": True,
            "has_badge": 1 if self.badge.get() else 0,
        }
        self.destroy()

    @staticmethod
    def _err(msg: str) -> None:
        raise RuntimeError(msg)

    def run(self) -> Optional[Dict[str, str]]:
        self.wait_window()
        return self._result
"""

with open(TARGET, "a", encoding="utf-8") as f:
    f.write(PART4)

print("Part4 POSFrame+FactoryItemDialog written OK")
