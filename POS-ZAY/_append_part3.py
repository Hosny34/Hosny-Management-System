
TARGET = r"c:\Users\youssef.sherif\Downloads\ادارة المخازن\ادارة المخازن\HosnyWarehouse.py"

PART3 = """

# ------------------- Income Frame -------------------
class IncomeFrame(ttk.Frame):

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=10)
        self.db = db
        self._income_r1 = tk.StringVar()
        self._income_r2 = tk.StringVar()
        self._income_has_alpha = tk.BooleanVar(value=False)

        self._build()

    def _build(self):
        ttk.Label(self, text="وارد (إضافة أصناف جديدة)", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 8))

        pkg_frame = ttk.LabelFrame(self, text="حاوية (حسب المخزن/العبوة)")
        pkg_frame.pack(fill=tk.X, pady=(0, 8))

        self.wh  = LabeledStaticCombo(pkg_frame, "رقم المخزن", values=["", "1", "2", "3", "4"])
        self.pkg = LabeledEntry(pkg_frame, "رقم العبوة")
        self.pkg_hint_var  = tk.StringVar(value="")
        self.pkg_status_var= tk.StringVar(value="")
        self.pkg_count_var = tk.StringVar(value="0")

        self.wh.grid(row=0, column=0, padx=6, pady=6, sticky="ew")
        self.pkg.grid(row=0, column=1, padx=6, pady=6, sticky="ew")
        self._income_r1.trace_add("write", lambda *_: self._rebuild_sizes_grid())
        self._income_r2.trace_add("write", lambda *_: self._rebuild_sizes_grid())
        self._income_has_alpha.trace_add("write", lambda *_: self._rebuild_sizes_grid())

        hint_box = ttk.Frame(pkg_frame)
        hint_box.grid(row=1, column=0, columnspan=2, sticky="ew", padx=6, pady=(12, 6))
        ttk.Label(hint_box, textvariable=self.pkg_hint_var).pack(anchor="w")

        stat_box = ttk.Frame(pkg_frame)
        stat_box.grid(row=2, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 6))
        ttk.Label(stat_box, text="الحالة:").pack(side=tk.LEFT)
        ttk.Label(stat_box, textvariable=self.pkg_status_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(stat_box, text="عدد صفوف المخزون داخل العبوة:").pack(side=tk.LEFT)
        ttk.Label(stat_box, textvariable=self.pkg_count_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(4, 0))

        pkg_frame.columnconfigure(0, weight=1)
        pkg_frame.columnconfigure(1, weight=1)

        grid = ttk.LabelFrame(self, text="بيانات الصنف")
        grid.pack(fill=tk.BOTH, expand=True)

        self.item_type = LabeledCombobox(grid, "النوع", self.db, "item_type")
        self.school    = LabeledCombobox(grid, "المدرسة", self.db, "school")
        self.color     = LabeledCombobox(grid, "اللون", self.db, "color")

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
            w.bind("<<ComboboxSelected>>", lambda e: (self._rebuild_sizes_grid(), self._auto_fill_price_for_grid()), add="+")
            w.bind("<FocusOut>",           lambda e: (self._rebuild_sizes_grid(), self._auto_fill_price_for_grid()), add="+")

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

        def _on_mousewheel(event):
            try:
                if hasattr(event, "delta") and event.delta:
                    step = int(-1 * (event.delta / 120))
                    sizes_canvas.yview_scroll(step, "units")
                else:
                    if event.num == 4:
                        sizes_canvas.yview_scroll(-1, "units")
                    elif event.num == 5:
                        sizes_canvas.yview_scroll(1, "units")
            except Exception:
                pass

        def _bind_wheel(_e=None):
            sizes_canvas.bind_all("<MouseWheel>", _on_mousewheel)
            sizes_canvas.bind_all("<Button-4>", _on_mousewheel)
            sizes_canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_wheel(_e=None):
            sizes_canvas.unbind_all("<MouseWheel>")
            sizes_canvas.unbind_all("<Button-4>")
            sizes_canvas.unbind_all("<Button-5>")

        sizes_canvas.bind("<Enter>", _bind_wheel)
        sizes_canvas.bind("<Leave>", _unbind_wheel)

        self.has_badge = tk.BooleanVar(value=False)
        badge_row = ttk.Frame(grid)
        badge_row.grid(row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 6))
        ttk.Checkbutton(badge_row, text="بادج", variable=self.has_badge).pack(side=tk.LEFT)

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

        pkg_btns = ttk.Frame(grid)
        pkg_btns.grid(row=5, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 6))

        ttk.Button(pkg_btns, text="إغلاق العبوة", command=self._close_current_package).pack(side=tk.RIGHT)
        ttk.Button(pkg_btns, text="إضافة", command=self._on_add).pack(side=tk.LEFT)

        btns = ttk.Frame(self)
        btns.pack(fill=tk.X, pady=(12, 0))
        ttk.Button(btns, text="تفريغ (يبقي المخزن/العبوة)", command=self._on_reset_keep_pkg).pack(side=tk.LEFT, padx=8)

        self.wh.cb.bind("<<ComboboxSelected>>", lambda *_: (self._refresh_pkg_hints(), self._refresh_pkg_status()))
        self.pkg.var.trace_add("write", lambda *_: self._refresh_pkg_status())

        self._refresh_pkg_hints()
        self._refresh_pkg_status()

    def _rebuild_sizes_grid(self):
        try:
            self.sizes_grid.destroy()
        except Exception:
            pass

        numeric_ranges = []

        def _parse_range(label: str):
            if not label:
                return None
            a, b = label.split("\\u2192")
            return int(a.strip()), int(b.strip())

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

    def _parse_int_or_none(self, s: str) -> Optional[int]:
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
            more = "" if len(info["free"]) <= 20 else " ..."
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
        if not messagebox.askyesno("تأكيد", f"إغلاق العبوة {p} في المخزن {w} ؟\\nلن يُسمح بأي إضافات لاحقاً."):
            return
        try:
            self.db.close_package(w, p)
            messagebox.showinfo("تم", "تم إغلاق العبوة.")
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

        if self.sizes_grid is None:
            messagebox.showwarning("فارغ", "لم تُدخل أي كميات في شبكة المقاسات.", parent=self)
            return

        rows = self.sizes_grid.get_rows()
        rows = [r for r in rows if int(r.get("qty") or 0) > 0]
        if not rows:
            messagebox.showwarning("فارغ", "لم تُدخل أي كميات في شبكة المقاسات.", parent=self)
            return

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

                self.db.ensure_default_price(item_type, price)

                added += 1
            except Exception as ex:
                errors.append(f"{size}: {ex}")

        msg_parts = []
        if added:
            msg_parts.append(f"تمت إضافة {added} صفًا للمخزون.")
        if errors:
            msg_parts.append("أخطاء:\\n" + "\\n".join(errors))
        if not msg_parts:
            messagebox.showinfo("معلومة", "لم تُجر أي تغييرات.", parent=self)
        else:
            messagebox.showinfo("النتيجة", "\\n\\n".join(msg_parts), parent=self)

        if added:
            self.item_type.set("")
            self.school.set("")
            self.color.set("")

            self._income_r1.set("")
            self._income_r2.set("")
            self._income_has_alpha.set(False)

            if self.sizes_grid:
                try:
                    self.sizes_grid.destroy()
                except Exception:
                    pass
                self.sizes_grid = None

            self.has_badge.set(False)

            self._refresh_pkg_hints()
            self._refresh_pkg_status()

            try:
                self.item_type.cb.focus_set()
            except Exception:
                pass

    @staticmethod
    def _err(msg: str) -> None:
        raise RuntimeError(msg)
"""

with open(TARGET, "a", encoding="utf-8") as f:
    f.write(PART3)

print("Part3 IncomeFrame written OK")
