

# ------------------- Admin Window -------------------

class AdminWindow(tk.Toplevel):
    """Admin configuration window."""

    def __init__(self, master, db: SqliteDatabase):
        super().__init__(master)
        self.db = db
        self.title("الإعدادات والإدارة")
        self.geometry("520x520")
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

        # ---- Tab 2: Import ----
        imp_tab = ttk.Frame(nb, padding=12)
        nb.add(imp_tab, text="استيراد")

        ttk.Label(imp_tab, text="استيراد من ملف إكسل", font=("Segoe UI", 11, "bold")).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(imp_tab, text="الملف:").grid(row=1, column=0, sticky="e", padx=6, pady=4)
        self._imp_path = tk.StringVar()
        ttk.Entry(imp_tab, textvariable=self._imp_path, width=40).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Button(imp_tab, text="استعراض...", command=self._browse_import).grid(row=2, column=1, sticky="w", padx=6, pady=2)

        ttk.Label(imp_tab, text="رقم المخزن:").grid(row=3, column=0, sticky="e", padx=6, pady=4)
        self._imp_wh = ttk.Combobox(imp_tab, values=["1", "2", "3", "4"], state="readonly", width=8)
        self._imp_wh.set("1")
        self._imp_wh.grid(row=3, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(imp_tab, text="رقم العبوة:").grid(row=4, column=0, sticky="e", padx=6, pady=4)
        self._imp_pkg = ttk.Entry(imp_tab, width=12)
        self._imp_pkg.grid(row=4, column=1, sticky="w", padx=6, pady=4)

        imp_tab.columnconfigure(1, weight=1)

        imp_btns = ttk.Frame(imp_tab)
        imp_btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(imp_btns, text="استيراد", command=self._do_import).pack(side=tk.RIGHT)

        # ---- Tab 3: Adjustment ----
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

        ttk.Label(adj_tab, text="المخزن:").grid(row=5, column=0, sticky="e", padx=6, pady=4)
        self._adj_wh = ttk.Combobox(adj_tab, values=["1", "2", "3", "4"], state="readonly", width=8)
        self._adj_wh.set("1")
        self._adj_wh.grid(row=5, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(adj_tab, text="رقم العبوة:").grid(row=6, column=0, sticky="e", padx=6, pady=4)
        self._adj_pkg = ttk.Entry(adj_tab, width=12)
        self._adj_pkg.grid(row=6, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(adj_tab, text="الكمية (موجب=إضافة، سالب=حذف):").grid(row=7, column=0, sticky="e", padx=6, pady=4)
        self._adj_qty = ttk.Entry(adj_tab, width=12)
        self._adj_qty.grid(row=7, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(adj_tab, text="السعر:").grid(row=8, column=0, sticky="e", padx=6, pady=4)
        self._adj_price = ttk.Entry(adj_tab, width=12)
        self._adj_price.grid(row=8, column=1, sticky="w", padx=6, pady=4)

        ttk.Label(adj_tab, text="ملاحظة:").grid(row=9, column=0, sticky="e", padx=6, pady=4)
        self._adj_note = ttk.Entry(adj_tab, width=28)
        self._adj_note.grid(row=9, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(adj_tab, text="كلمة المرور:").grid(row=10, column=0, sticky="e", padx=6, pady=4)
        self._adj_pw = ttk.Entry(adj_tab, show="*", width=16)
        self._adj_pw.grid(row=10, column=1, sticky="w", padx=6, pady=4)

        adj_tab.columnconfigure(1, weight=1)

        adj_btns = ttk.Frame(adj_tab)
        adj_btns.grid(row=11, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(adj_btns, text="تطبيق التعديل", command=self._do_adjustment).pack(side=tk.RIGHT)

        # ---- Tab 4: Reset Counts ----
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
        ttk.Button(reset_btns, text="إعادة التعيين", command=self._reset_counts).pack(side=tk.RIGHT)

        # ---- Close button ----
        ttk.Button(self, text="إغلاق", command=self.destroy).pack(side=tk.BOTTOM, pady=8)

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
        messagebox.showinfo("تم", "تم تغيير كلمة المرور بنجاح.", parent=self)
        self._pw_cur.delete(0, tk.END)
        self._pw_new.delete(0, tk.END)
        self._pw_confirm.delete(0, tk.END)

    def _browse_import(self):
        path = filedialog.askopenfilename(
            title="اختر ملف إكسل",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            parent=self,
        )
        if path:
            self._imp_path.set(path)

    def _do_import(self):
        path = self._imp_path.get().strip()
        wh_txt = self._imp_wh.get().strip()
        pkg_txt = self._imp_pkg.get().strip()

        if not path:
            messagebox.showwarning("مسار ناقص", "اختر ملف إكسل أولاً.", parent=self)
            return
        if not wh_txt.isdigit():
            messagebox.showwarning("مخزن غير صالح", "اختر رقم المخزن.", parent=self)
            return
        if not pkg_txt.isdigit():
            messagebox.showwarning("عبوة غير صالحة", "أدخل رقم العبوة.", parent=self)
            return

        try:
            count = self.db.import_from_excel(path, int(wh_txt), int(pkg_txt))
            messagebox.showinfo("اكتمل الاستيراد", f"تم استيراد {count} صف بنجاح.", parent=self)
        except Exception as ex:
            messagebox.showerror("فشل الاستيراد", str(ex), parent=self)

    def _do_adjustment(self):
        pw = self._adj_pw.get()
        if not self.db.verify_admin_password(pw):
            messagebox.showerror("مرفوض", "كلمة المرور غير صحيحة.", parent=self)
            return

        item_type = self._adj_type.get().strip()
        school    = self._adj_school.get().strip()
        color     = self._adj_color.get().strip()
        size      = self._adj_size.get().strip()
        wh_txt    = self._adj_wh.get().strip()
        pkg_txt   = self._adj_pkg.get().strip()
        qty_txt   = self._adj_qty.get().strip()
        price_txt = self._adj_price.get().strip()
        note      = self._adj_note.get().strip() or "Manual adjustment"

        if not (item_type and school and color and size):
            messagebox.showwarning("بيانات ناقصة", "أدخل النوع والمدرسة واللون والمقاس.", parent=self)
            return
        if not wh_txt.isdigit() or not pkg_txt.isdigit():
            messagebox.showwarning("بيانات غير صالحة", "أدخل رقم المخزن ورقم العبوة.", parent=self)
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
                warehouse_no=int(wh_txt), package_no=int(pkg_txt),
                qty=qty, unit_price=price, note=note
            )
            messagebox.showinfo("تم", f"تم تطبيق التعديل ({qty:+d}) على المخزون.", parent=self)
        except Exception as ex:
            messagebox.showerror("فشل التعديل", str(ex), parent=self)

    def _reset_counts(self):
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
            messagebox.showinfo("تم", "تم إعادة تعيين جميع العدادات بنجاح.", parent=self)
            self._reset_pw.delete(0, tk.END)
        except Exception as ex:
            messagebox.showerror("فشل", str(ex), parent=self)


# ------------------- Main Application -------------------

class WarehouseApp:
    """Main application controller."""

    def __init__(self, root: tk.Tk, db: SqliteDatabase):
        self.root = root
        self.db = db
        root.title("إدارة المخازن والمبيعات")
        root.geometry("1280x760")
        root.minsize(900, 600)
        self._build()

    def _build(self):
        # ---- Top toolbar ----
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, side=tk.TOP, padx=6, pady=(6, 0))

        ttk.Button(toolbar, text="المخزون", command=self._open_inventory).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="سجل الفواتير", command=self._open_bills_history).pack(side=tk.LEFT, padx=4)
        ttk.Button(toolbar, text="سجل الحركات", command=self._open_movements).pack(side=tk.LEFT)
        ttk.Button(toolbar, text="الإعدادات", command=self._open_admin).pack(side=tk.RIGHT)

        # ---- Main notebook ----
        self._nb = ttk.Notebook(self.root)
        self._nb.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # Tab 1: Income (warehouse receiving)
        income_tab = ttk.Frame(self._nb)
        self._nb.add(income_tab, text="الوارد")
        self._income_frame = IncomeFrame(income_tab, self.db)
        self._income_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 2: POS (point of sale)
        pos_tab = ttk.Frame(self._nb)
        self._nb.add(pos_tab, text="نقطة البيع")
        self._pos_frame = POSFrame(pos_tab, self.db)
        self._pos_frame.pack(fill=tk.BOTH, expand=True)

        # Tab 3: Statistics
        stats_tab = ttk.Frame(self._nb)
        self._nb.add(stats_tab, text="الإحصائيات")
        self._stats_frame = StatisticsFrame(stats_tab, self.db)
        self._stats_frame.pack(fill=tk.BOTH, expand=True)

        # Refresh statistics when tab is selected
        self._nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _on_tab_changed(self, event=None):
        try:
            current = self._nb.select()
            tab_text = self._nb.tab(current, "text")
            if tab_text == "الإحصائيات":
                self._stats_frame._refresh_all()
        except Exception:
            pass

    def _open_inventory(self):
        try:
            InventoryWindow(self.root, self.db)
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


# ------------------- Entry Point -------------------

def main():
    import os
    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "warehouse.db")
    db = SqliteDatabase(db_path)
    root = tk.Tk()
    try:
        root.tk.call("source", os.path.join(os.path.dirname(os.path.abspath(__file__)), "azure.tcl"))
        root.tk.call("set_theme", "light")
    except Exception:
        pass
    try:
        root.option_add("*Font", "TkDefaultFont 10")
    except Exception:
        pass
    _app = WarehouseApp(root, db)
    root.mainloop()


if __name__ == "__main__":
    main()
