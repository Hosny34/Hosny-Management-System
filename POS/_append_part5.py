
import os

TARGET = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Warehouse", "HosnyWarehouse.py"))

PART5 = """

# ------------------- Statistics Frame -------------------

class StatisticsFrame(ttk.Frame):
    \"\"\"Statistics and reporting tab.\"\"\"

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master, padding=6)
        self.db = db
        self._build()

    def _build(self):
        ttk.Label(self, text="الإحصائيات والتقارير", font=("Segoe UI", 12, "bold")).pack(anchor="w", pady=(0, 6))

        # Date range filter
        filter_row = ttk.Frame(self)
        filter_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(filter_row, text="من:").pack(side=tk.LEFT)
        self._df = DateField(filter_row, "")
        self._df.pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(filter_row, text="إلى:").pack(side=tk.LEFT)
        self._dt = DateField(filter_row, "")
        self._dt.pack(side=tk.LEFT, padx=(4, 10))
        ttk.Button(filter_row, text="تحديث", command=self._refresh_all).pack(side=tk.LEFT, padx=4)

        # ---- Section 1: Money flow ----
        mf = ttk.LabelFrame(self, text="التدفق المالي")
        mf.pack(fill=tk.X, padx=4, pady=(0, 8))

        self._sales_count_var = tk.StringVar(value="0")
        self._sales_total_var = tk.StringVar(value="0.00")
        self._res_count_var   = tk.StringVar(value="0")
        self._res_total_var   = tk.StringVar(value="0.00")
        self._res_paid_var    = tk.StringVar(value="0.00")

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

        ttk.Button(mv, text="طباعة", command=self._print_movements).pack(anchor="e", padx=8, pady=4)

        # ---- Section 3: Reservations ----
        rv = ttk.LabelFrame(self, text="الحجوزات")
        rv.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 8))

        rv_wrap = ttk.Frame(rv)
        rv_wrap.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._res_table = ttk.Treeview(
            rv_wrap,
            columns=("id", "created", "customer", "item", "school", "color", "size", "qty", "total", "paid", "status"),
            show="headings",
            height=5,
        )
        for col, txt, w in [
            ("id", "رقم", 40), ("created", "التاريخ", 140), ("customer", "العميل", 100),
            ("item", "النوع", 80), ("school", "المدرسة", 100), ("color", "اللون", 60),
            ("size", "المقاس", 55), ("qty", "الكمية", 55), ("total", "الإجمالي", 70),
            ("paid", "المدفوع", 70), ("status", "الحالة", 70),
        ]:
            self._res_table.heading(col, text=txt)
            self._res_table.column(col, width=w, anchor="center")
        rv_ysb = ttk.Scrollbar(rv_wrap, orient="vertical", command=self._res_table.yview)
        self._res_table.configure(yscrollcommand=rv_ysb.set)
        self._res_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        rv_ysb.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Button(rv, text="طباعة", command=self._print_reservations).pack(anchor="e", padx=8, pady=4)

        self._refresh_all()

    def _refresh_all(self):
        df = self._df.get() or None
        dt = self._dt.get() or None

        # Money flow
        try:
            stats = self.db.get_sales_stats(df, dt)
            self._sales_count_var.set(str(stats["sales_count"]))
            self._sales_total_var.set(f"{stats['sales_total']:.2f}")
            self._res_count_var.set(str(stats["res_count"]))
            self._res_total_var.set(f"{stats['res_total']:.2f}")
            self._res_paid_var.set(f"{stats['res_paid']:.2f}")
        except Exception:
            pass

        # Item movements
        try:
            self._mv_table.delete(*self._mv_table.get_children())
            for r in self.db.get_item_movement_stats():
                received = int(r.get("received") or 0)
                sold = int(r.get("sold") or 0)
                reserved = int(r.get("reserved") or 0)
                adjusted = int(r.get("adjusted") or 0)
                remaining = received - sold - reserved - adjusted
                self._mv_table.insert("", tk.END, values=(
                    r.get("item_type",""), r.get("school",""), r.get("color",""), r.get("size",""),
                    received, sold, reserved, remaining
                ))
        except Exception:
            pass

        # Reservations
        try:
            self._res_table.delete(*self._res_table.get_children())
            for r in self.db.list_reservations():
                self._res_table.insert("", tk.END, values=(
                    r.get("id",""), r.get("created_at",""), r.get("customer",""),
                    r.get("item_type",""), r.get("school",""), r.get("color",""), r.get("size",""),
                    r.get("qty",""), f"{float(r.get('total_amount',0)):.2f}",
                    f"{float(r.get('paid_amount',0)):.2f}", r.get("status","")
                ))
        except Exception:
            pass

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
<tr><td>إجمالي المبيعات</td><td>{stats['sales_total']:.2f}</td></tr>
<tr><td>عدد الحجوزات</td><td>{stats['res_count']}</td></tr>
<tr><td>إجمالي الحجوزات</td><td>{stats['res_total']:.2f}</td></tr>
<tr><td>المبلغ المدفوع من الحجوزات</td><td>{stats['res_paid']:.2f}</td></tr>
</tbody></table>
</body></html>"""
            import tempfile, os
            path = os.path.join(tempfile.gettempdir(), "money_flow.html")
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)

    def _print_movements(self):
        try:
            rows = self.db.get_item_movement_stats()
            rows_html = ""
            for r in rows:
                received = int(r.get("received") or 0)
                sold = int(r.get("sold") or 0)
                reserved = int(r.get("reserved") or 0)
                adjusted = int(r.get("adjusted") or 0)
                remaining = received - sold - reserved - adjusted
                rows_html += f"<tr><td>{_html(r.get('item_type',''))}</td><td>{_html(r.get('school',''))}</td><td>{_html(r.get('color',''))}</td><td>{_html(r.get('size',''))}</td><td>{received}</td><td>{sold}</td><td>{reserved}</td><td>{remaining}</td></tr>"
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
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)

    def _print_reservations(self):
        try:
            rows = self.db.list_reservations()
            rows_html = ""
            for r in rows:
                rows_html += (f"<tr><td>{r.get('id','')}</td><td>{_html(r.get('created_at',''))}</td>"
                              f"<td>{_html(r.get('customer',''))}</td><td>{_html(r.get('item_type',''))}</td>"
                              f"<td>{_html(r.get('school',''))}</td><td>{_html(r.get('color',''))}</td>"
                              f"<td>{_html(r.get('size',''))}</td><td>{r.get('qty','')}</td>"
                              f"<td>{float(r.get('total_amount',0)):.2f}</td><td>{float(r.get('paid_amount',0)):.2f}</td>"
                              f"<td>{_html(r.get('status',''))}</td></tr>")
            html = f"""<!DOCTYPE html>
<html lang='ar' dir='rtl'>
<head><meta charset='utf-8'><title>الحجوزات</title>
<style>body{{font-family:Tahoma,Arial;margin:20px}} table{{border-collapse:collapse;width:100%}} th,td{{border:1px solid #777;padding:6px;text-align:center}}</style>
</head><body>
<h2>الحجوزات</h2>
<table><thead><tr><th>رقم</th><th>التاريخ</th><th>العميل</th><th>النوع</th><th>المدرسة</th><th>اللون</th><th>المقاس</th><th>الكمية</th><th>الإجمالي</th><th>المدفوع</th><th>الحالة</th></tr></thead>
<tbody>{rows_html}</tbody></table>
</body></html>"""
            import tempfile, os
            path = os.path.join(tempfile.gettempdir(), "reservations.html")
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            _print_html_auto(path, copies=1, parent=self)
        except Exception as ex:
            messagebox.showerror("خطأ", str(ex), parent=self)


# ------------------- Inventory Window -------------------

class InventoryWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("المخزون")
        self.geometry("1000x560")

        self.multi: Dict[str, List[Any]] = {
            "item_type": [], "school": [], "color": [], "size": [],
            "warehouse_no": [], "package_no": [],
        }

        self._multi_btns: Dict[str, ttk.Button] = {}
        self._field_widgets: Dict[str, tk.Widget] = {}

        self._build()

    def _build(self):
        filters = ttk.LabelFrame(self, text="تصنيف")
        filters.pack(fill=tk.X, padx=8, pady=8)

        self.f_type   = LabeledCombobox(filters, "النوع",    self.db, "item_type")
        self.f_school = LabeledCombobox(filters, "المدرسة", self.db, "school")
        self.f_color  = LabeledCombobox(filters, "اللون",    self.db, "color")
        self.f_size   = LabeledCombobox(filters, "المقاس",   self.db, "size")
        self.f_wh     = LabeledStaticCombo(filters, "رقم المخزن", values=["", "1", "2", "3", "4"])
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
        self.f_wh.cb.bind("<<ComboboxSelected>>", lambda e: self._schedule_refresh(), add="+")
        self.f_pkg.var.trace_add("write", lambda *_: self._schedule_refresh())

        btns = ttk.Frame(filters)
        btns.grid(row=4, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 6))
        ttk.Button(btns, text="بحث", command=self._refresh).pack(side=tk.LEFT)
        ttk.Button(btns, text="مسح", command=self._clear_all).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="تصدير إلى إكسل", command=self._export_excel).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="طباعة جداول المقاسات", command=self._print_size_sheets).pack(side=tk.LEFT, padx=8)
        ttk.Button(btns, text="تعديل نطاقات المقاسات...", command=self._edit_size_ranges_dialog).pack(side=tk.LEFT, padx=8)

        table_wrap = ttk.Frame(self)
        table_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))
        self.table = ttk.Treeview(
            table_wrap,
            columns=("id", "type", "school", "color", "size", "wh", "pkg", "price", "count", "value"),
            show="headings",
            selectmode="extended",
        )
        for col, txt, w in [
            ("id","المعرّف",60), ("type","النوع",140), ("school","المدرسة",160),
            ("color","اللون",80), ("size","المقاس",70), ("wh","المخزن",50),
            ("pkg","العبوة",60), ("price","السعر",80), ("count","الكمية",70), ("value","القيمة",90),
        ]:
            self.table.heading(col, text=txt)
            self.table.column(col, width=w, anchor="center")
        ysb = ttk.Scrollbar(table_wrap, orient="vertical", command=self.table.yview)
        xsb = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        self.table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)
        xsb.pack(side=tk.BOTTOM, fill=tk.X)

        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=8, pady=(0, 8))
        self.sum_qty = tk.StringVar(value="0")
        self.sum_val = tk.StringVar(value="0.00")
        ttk.Label(bar, text="إجمالي الكمية:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_qty, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(bar, text="إجمالي القيمة:").pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.sum_val, font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Button(bar, text="تعديل السعر...", command=self._edit_price_dialog).pack(side=tk.RIGHT)
        ttk.Button(bar, text="تعديل المواصفات...", command=self._edit_specs_dialog).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(bar, text="حذف المحدد...", command=self._remove_selected_dialog).pack(side=tk.RIGHT, padx=(8, 0))

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
            v = (v or "").strip()
            if not v:
                return None, None
            a, b = v.split("\\u2192")
            return int(a.strip()), int(b.strip())

        def on_save():
            try:
                r1s, r1e = _parse_range_label(r1_var.get())
                r2s, r2e = _parse_range_label(r2_var.get())
                self.db.upsert_size_profile(item_type, school, color,
                    r1_start=r1s, r1_end=r1e, r2_start=r2s, r2_end=r2e,
                    has_alpha=has_alpha_var.get())
                messagebox.showinfo("تم", "تم حفظ نطاقات المقاسات.", parent=dlg)
                dlg.destroy()
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=on_save).pack(side=tk.RIGHT, padx=6)

    def _open_multi_dialog(self, field: str):
        if field in ("item_type", "school", "color", "size"):
            try:
                values = self.db.get_distinct(field)
            except Exception:
                values = []
        elif field == "warehouse_no":
            values = ["1", "2", "3", "4"]
        else:
            try:
                cur = self.db.conn.execute("SELECT DISTINCT package_no FROM stocks ORDER BY package_no ASC")
                values = [str(r[0]) for r in cur.fetchall()]
            except Exception:
                values = []

        dlg = MultiSelectDialog(self, title="اختيار متعدد", values=values,
                                preselected=[str(x) for x in self.multi[field]])
        picked = dlg.run()
        if picked is None:
            return

        if field in ("warehouse_no", "package_no"):
            self.multi[field] = [int(x) for x in picked]
        else:
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
        elif field == "warehouse_no": self.f_wh.set("")
        elif field == "package_no": self.f_pkg.set("")

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
            school_groups[school].setdefault((item, color), []).append(r)

        def build_size_ranges_from_profile(profile):
            numeric_tables = []
            alpha_labels = []
            if profile is None:
                return numeric_tables, alpha_labels
            r1_start, r1_end, r2_start, r2_end, has_alpha = profile
            if r1_start is not None and r1_end is not None:
                key = (r1_start, r1_end)
                labels = ALLOWED_NUMERIC_RANGES.get(key)
                if labels:
                    numeric_tables.append(labels[:])
            if r2_start is not None and r2_end is not None:
                key = (r2_start, r2_end)
                labels = ALLOWED_NUMERIC_RANGES.get(key)
                if labels:
                    numeric_tables.append(labels[:])
            if has_alpha:
                alpha_labels = ALPHA_SIZES[:]
            return numeric_tables, alpha_labels

        tables_html = []

        for sch, item_groups in school_groups.items():
            for (t, clr), items in item_groups.items():
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

                head = f"<div class='hdr'><span>النوع: {_html(t)}</span><span>المدرسة: {_html(sch)}</span><span>اللون: {_html(clr)}</span></div>"

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
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)
        _print_html_auto(path, copies=1, parent=self)

    def _schedule_refresh(self, delay_ms: int = 250):
        if hasattr(self, "_inv_job") and self._inv_job:
            self.after_cancel(self._inv_job)
        self._inv_job = self.after(delay_ms, self._refresh)

    def _filters(self) -> Dict[str, Any]:
        active_multi = [k for k, v in self.multi.items() if v]
        if active_multi:
            fld = active_multi[0]
            f: Dict[str, Any] = {
                "item_type": None, "school": None, "color": None, "size": None,
                "warehouse_no": None, "package_no": None,
            }
            f[fld] = self.multi[fld][:]
            return f

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
                        f"{float(r['unit_price']):.2f}",
                        r["count"], f"{float(r['value']):.2f}")
            )
            total_qty += int(r["count"])
            total_value += float(r["value"])
        self.sum_qty.set(str(total_qty))
        self.sum_val.set(f"{total_value:.2f}")

    def _clear_all(self):
        for w in (self.f_type, self.f_school, self.f_color, self.f_size):
            w.set("")
        self.f_wh.set("")
        self.f_pkg.set("")
        for k in self.multi.keys():
            self.multi[k] = []
        for b in self._multi_btns.values():
            b.configure(text="اختيار متعدد...")
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
            messagebox.showinfo("تم الحفظ", f"تم حفظ المخزون إلى:\\n{path}", parent=self)
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)

    def _remove_selected_dialog(self):
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("لم يتم التحديد", "اختر صفاً واحداً أو أكثر من المخزون أولاً.", parent=self)
            return

        rows = [self.table.item(i, "values") for i in sel]
        ids = [int(r[0]) for r in rows]

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
                messagebox.showinfo("تم", f"تم حذف {total_removed} وحدة من {len(ids)} صف(وف).", parent=dlg)
                dlg.destroy()
                self._refresh()
            except Exception as ex:
                messagebox.showerror("خطأ", str(ex), parent=dlg)

        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حذف", command=on_ok).pack(side=tk.RIGHT, padx=6)

    def _edit_price_dialog(self):
        row = 0
        sel = self.table.selection()
        if not sel:
            messagebox.showwarning("لم يتم التحديد", "اختر صفاً من المخزون أولاً.", parent=self)
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

        if multi:
            ttk.Label(frm, text=f"عدد الصفوف المحددة: {len(rows)}", font=("Segoe UI", 10, "bold"))\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
        else:
            ttk.Label(frm, text=f"الصنف: {first[1]} / {first[2]} / {first[3]} / {first[4]}")\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
            ttk.Label(frm, text=f"المخزن/العبوة: {first[5]} / {first[6]}")\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 6))
            row += 1
            ttk.Label(frm, text=f"السعر الحالي: {current_price:.2f}")\
                .grid(row=row, column=0, columnspan=2, sticky="w", pady=(0, 10))
            row += 1

        ttk.Label(frm, text="السعر الجديد:").grid(row=row, column=0, sticky="e", padx=4, pady=6)
        price_var = tk.StringVar(value=f"{current_price:.2f}")
        ttk.Entry(frm, textvariable=price_var, width=16).grid(row=row, column=1, sticky="w", padx=4, pady=6)
        row += 1

        scope_var = tk.StringVar(value="row")

        if not multi:
            scope_box = ttk.LabelFrame(frm, text="النطاق")
            scope_box.grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 10))
            ttk.Radiobutton(scope_box, text="الصف المحدد فقط", variable=scope_var, value="row").pack(anchor="w", padx=8, pady=4)
            ttk.Radiobutton(scope_box, text="كل الصفوف بنفس (النوع/المدرسة/المقاس) داخل نفس المخزن/العبوة",
                            variable=scope_var, value="same_pkg").pack(anchor="w", padx=8, pady=4)
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
                        updated_total += self.db.update_prices({"id": int(vals[0])}, new_price, note="Price update (multi-selection)")
                else:
                    if scope_var.get() == "row":
                        updated_total = self.db.update_prices({"id": int(first[0])}, new_price, note="Price update (single row)")
                    else:
                        updated_total = self.db.update_prices({
                            "item_type": first[1], "school": first[2], "size": first[4],
                            "warehouse_no": first[5], "package_no": first[6],
                        }, new_price, note="Price update (same type/school/size)")

                if updated_total == 0:
                    messagebox.showinfo("لا تغييرات", "لم يتم العثور على صفوف مطابقة.", parent=dlg)
                else:
                    messagebox.showinfo("تم", f"تم تحديث السعر في {updated_total} صف(وف).", parent=dlg)
                dlg.destroy()
                self._refresh()
            except Exception as ex:
                messagebox.showerror("فشل التحديث", str(ex), parent=dlg)

        btns = ttk.Frame(frm)
        btns.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="إلغاء", command=dlg.destroy).pack(side=tk.RIGHT)
        ttk.Button(btns, text="حفظ", command=on_ok).pack(side=tk.RIGHT, padx=6)
        dlg.update_idletasks()

    def _edit_specs_dialog(self):
        sel = self.table.selection()
        ids: List[int] = []
        if sel:
            for iid in sel:
                vals = self.table.item(iid, "values")
                ids.append(int(vals[0]))
            scope_text = f"عدد الصفوف المحددة: {len(ids)}"
            scope_mode = "ids"
        else:
            wh_txt = (self.f_wh.get() or "").strip()
            pkg_txt = (self.f_pkg.get() or "").strip()
            if not (wh_txt and pkg_txt and wh_txt.isdigit() and pkg_txt.isdigit()):
                messagebox.showwarning("حدد النطاق",
                    "اختر صفوفاً من الجدول أو أدخل (المخزن/العبوة) في أعلى الشاشة لتطبيق التعديل على العبوة كلها.",
                    parent=self)
                return
            w = int(wh_txt); p = int(pkg_txt)
            scope_text = f"النطاق: المخزن {w} / العبوة {p}"
            scope_mode = "pkg"
            ids = [w, p]

        dlg = tk.Toplevel(self)
        dlg.title("تعديل المواصفات")
        dlg.transient(self)
        dlg.grab_set()

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=scope_text, font=("", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,8))
        ttk.Label(frm, text="اترك الحقل فارغاً إذا كنت لا تريد تغييره.").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0,10))

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

                if updated == 0:
                    messagebox.showinfo("لا تغييرات", "لم يتم العثور على صفوف مطابقة.", parent=dlg)
                else:
                    messagebox.showinfo("تم", f"تم تعديل المواصفات في {updated} صف(وف).", parent=dlg)
                dlg.destroy()
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
        self.geometry("1000x560")
        self._build()

    def _build(self):
        top = ttk.Frame(self)
        top.pack(fill=tk.X, padx=8, pady=8)
        ttk.Label(top, text="سجل الفواتير", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT)
        ttk.Button(top, text="تحديث", command=self._refresh).pack(side=tk.RIGHT)
        ttk.Button(top, text="فتح", command=self._print_selected).pack(side=tk.RIGHT, padx=8)
        ttk.Button(top, text="تصدير المحدد إلى إكسل", command=self._export_selected).pack(side=tk.RIGHT)

        bills_wrap = ttk.Frame(self)
        bills_wrap.pack(fill=tk.BOTH, expand=False, padx=8, pady=(0, 6))
        self.bills_table = ttk.Treeview(
            bills_wrap, columns=("id", "created_at", "customer", "total"), show="headings", height=10
        )
        for col, txt, w in [("id","المعرّف",80), ("created_at","التاريخ",180), ("customer","العميل",250), ("total","الإجمالي",120)]:
            self.bills_table.heading(col, text=txt)
            self.bills_table.column(col, width=w, anchor="center")
        bills_ysb = ttk.Scrollbar(bills_wrap, orient="vertical", command=self.bills_table.yview)
        bills_xsb = ttk.Scrollbar(bills_wrap, orient="horizontal", command=self.bills_table.xview)
        self.bills_table.configure(yscrollcommand=bills_ysb.set, xscrollcommand=bills_xsb.set)
        self.bills_table.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bills_ysb.pack(side=tk.RIGHT, fill=tk.Y)
        bills_xsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.bills_table.bind("<<TreeviewSelect>>", lambda e: self._load_items())

        ttk.Label(self, text="بنود الفاتورة").pack(anchor="w", padx=8)

        items_wrap = ttk.Frame(self)
        items_wrap.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
        self.items_table = ttk.Treeview(
            items_wrap,
            columns=("type", "school", "color", "size", "origin", "wh", "pkg", "price", "qty", "total"),
            show="headings",
            height=12,
        )
        for col, txt, w in [
            ("type","النوع",140), ("school","المدرسة",160), ("color","اللون",80), ("size","المقاس",70),
            ("origin","المصدر",90), ("wh","المخزن",50), ("pkg","العبوة",60), ("price","السعر",80),
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

        self._refresh()

    def _refresh(self):
        self.bills_table.delete(*self.bills_table.get_children())
        for b in self.db.list_bills():
            self.bills_table.insert(
                "", tk.END, iid=str(b["id"]),
                values=(b["id"], b["created_at"], b.get("customer") or "", f"{float(b['total']):.2f}")
            )
        self.items_table.delete(*self.items_table.get_children())

    def _get_selected_bill_id(self) -> Optional[int]:
        sel = self.bills_table.selection()
        if not sel:
            return None
        return int(sel[0])

    def _load_items(self):
        bill_id = self._get_selected_bill_id()
        self.items_table.delete(*self.items_table.get_children())
        if bill_id is None:
            return
        items = self.db.list_bill_items(bill_id)
        for ln in items:
            origin_txt = "من المخزون" if ln.get("origin") == "STOCK" else ("من المصنع" if ln.get("origin") == "FACTORY" else "")
            wh_txt  = "" if (ln.get("warehouse_no") in (None, "", 0, "0")) else ln.get("warehouse_no")
            pkg_txt = "" if (ln.get("package_no")  in (None, "", 0, "0")) else ln.get("package_no")
            self.items_table.insert(
                "", tk.END,
                values=(ln["item_type"], ln["school"], ln["color"], ln["size"], origin_txt,
                        wh_txt, pkg_txt, f"{float(ln['unit_price']):.2f}",
                        ln["qty"], f"{float(ln['line_total']):.2f}")
            )

    def _export_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.")
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
        headers = ["type", "school", "color", "size", "origin", "warehouse_no", "package_no", "unit_price", "qty", "line_total"]
        def _origin_txt(o: Optional[str]) -> str:
            return "من المخزون" if o == "STOCK" else ("من المصنع" if o == "FACTORY" else "")
        rows = [
            [
                ln["item_type"], ln["school"], ln["color"], ln["size"],
                _origin_txt(ln.get("origin")),
                ln["warehouse_no"], ln["package_no"],
                float(ln["unit_price"]), int(ln["qty"]), float(ln["line_total"]),
            ]
            for ln in items
        ]
        try:
            export_to_excel(path, headers, rows)
            messagebox.showinfo("تم الحفظ", f"تم تصدير الفاتورة إلى:\\n{path}")
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex))

    def _print_selected(self):
        bill_id = self._get_selected_bill_id()
        if bill_id is None:
            messagebox.showwarning("لم يتم التحديد", "اختر فاتورة أولاً.")
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


# ------------------- Movements Window -------------------

class MovementsWindow(tk.Toplevel):
    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("سجل الحركات")
        self.geometry("1180x560")
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
            height=16,
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
        headers = ["id","ts","direction","item_type","school","color","size","warehouse_no","package_no","unit_price","qty","note","bill_id","stock_id"]
        table = [[
            m.get("id"), m.get("ts"), m.get("direction"),
            m.get("item_type",""), m.get("school",""), m.get("color",""), m.get("size",""),
            m.get("warehouse_no",""), m.get("package_no",""), m.get("unit_price",""),
            m.get("qty"), m.get("note",""), m.get("bill_id"), m.get("stock_id")
        ] for m in rows]
        try:
            export_to_excel(path, headers, table)
            messagebox.showinfo("تم الحفظ", f"تم حفظ الحركات إلى:\\n{path}", parent=self)
        except Exception as ex:
            messagebox.showerror("فشل التصدير", str(ex), parent=self)
"""

with open(TARGET, "a", encoding="utf-8") as f:
    f.write(PART5)

print("Part5 written OK")
