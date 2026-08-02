"""HosnyFactory UI module."""

from __future__ import annotations

from typing import *
import re

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from factory_core import *  # noqa: F401,F403

_DIGIT_TRANSLATION = str.maketrans({
    "\u0660": "0", "\u0661": "1", "\u0662": "2", "\u0663": "3", "\u0664": "4",
    "\u0665": "5", "\u0666": "6", "\u0667": "7", "\u0668": "8", "\u0669": "9",
    "\u06f0": "0", "\u06f1": "1", "\u06f2": "2", "\u06f3": "3", "\u06f4": "4",
    "\u06f5": "5", "\u06f6": "6", "\u06f7": "7", "\u06f8": "8", "\u06f9": "9",
})
_LTR_MARK = "\u200e"
_RTL_MARK = "\u200f"
_NUMBER_RUN_RE = re.compile(r"(?<!\u200e)([0-9](?:[0-9.,:/\\\- ]*[0-9])?)(?!\u200e)")


def western_digits(value: Any) -> str:
    return ("" if value is None else str(value)).translate(_DIGIT_TRANSLATION)


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

    for entry_cls in (tk.Entry, ttk.Entry, ttk.Combobox):
        original_get = entry_cls.get

        def get(self, _original_get=original_get):
            return _strip_digit_marks(_original_get(self))

        entry_cls.get = get

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
            string = western_digits_for_display(string)
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

def setup_style(root: tk.Tk) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    root.configure(bg=UI["BG"])
    root.option_add("*Font", FONTS["body"])
    style.configure(".", font=FONTS["body"], background=UI["BG"], foreground=UI["TEXT"])
    style.configure("TFrame", background=UI["BG"])
    style.configure("Card.TFrame", background=UI["SURFACE"])
    style.configure("TLabel", background=UI["BG"], foreground=UI["TEXT"])
    style.configure("Card.TLabel", background=UI["SURFACE"], foreground=UI["TEXT"])
    style.configure("H1.TLabel", background=UI["BG"], foreground=UI["TEXT"], font=FONTS["h1"])
    style.configure("H2.TLabel", background=UI["BG"], foreground=UI["TEXT"], font=FONTS["h2"])
    style.configure("H3.TLabel", background=UI["BG"], foreground=UI["TEXT"], font=FONTS["h3"])
    style.configure("Dim.TLabel", background=UI["BG"], foreground=UI["TEXT_SEC"], font=FONTS["small"])

    style.configure(
        "TButton",
        padding=(12, 6),
        font=FONTS["btn"],
        background=UI["ACCENT"],
        foreground="white",
        borderwidth=0,
    )
    style.map(
        "TButton",
        background=[("active", UI["ACCENT_H"]), ("disabled", UI["BORDER"])],
        foreground=[("disabled", UI["TEXT_DIM"])],
    )
    style.configure("Ok.TButton",     background=UI["OK"],     foreground="white")
    style.map("Ok.TButton",     background=[("active", "#15803d")])
    style.configure("Warn.TButton",   background=UI["WARN"],   foreground="white")
    style.map("Warn.TButton",   background=[("active", "#b45309")])
    style.configure("Danger.TButton", background=UI["DANGER"], foreground="white")
    style.map("Danger.TButton", background=[("active", "#991b1b")])
    style.configure("Ghost.TButton", background=UI["SURFACE"], foreground=UI["TEXT"])
    style.map("Ghost.TButton", background=[("active", UI["SURFACE2"])])

    style.configure(
        "Treeview",
        background=UI["SURFACE"],
        fieldbackground=UI["SURFACE"],
        foreground=UI["TEXT"],
        rowheight=26,
        font=FONTS["body"],
        borderwidth=0,
    )
    style.configure(
        "Treeview.Heading",
        background=UI["SURFACE2"],
        foreground=UI["TEXT"],
        font=FONTS["body_b"],
        relief="flat",
    )
    style.map("Treeview",
              background=[("selected", UI["SEL_BG"])],
              foreground=[("selected", UI["SEL_FG"])])

    style.configure("TNotebook", background=UI["BG"], borderwidth=0)
    style.configure(
        "TNotebook.Tab",
        padding=(16, 8),
        font=FONTS["btn"],
        background=UI["SURFACE2"],
        foreground=UI["TEXT_SEC"],
    )
    style.map(
        "TNotebook.Tab",
        background=[("selected", UI["SURFACE"])],
        foreground=[("selected", UI["ACCENT"])],
    )

    style.configure("TEntry", fieldbackground=UI["SURFACE"], foreground=UI["TEXT"], padding=6)
    style.configure("TCombobox", fieldbackground=UI["SURFACE"], foreground=UI["TEXT"], padding=4)
    style.configure("Vertical.TScrollbar", background=UI["SURFACE2"], troughcolor=UI["BG"], borderwidth=0)
    style.configure("Horizontal.TScrollbar", background=UI["SURFACE2"], troughcolor=UI["BG"], borderwidth=0)


def show_error(parent, title: str, msg: str) -> None:
    log.error("%s: %s\n%s", title, msg, traceback.format_exc())
    messagebox.showerror(title, msg, parent=parent)


class Toast(tk.Toplevel):
    """Simple auto-dismissing toast."""

    def __init__(
        self,
        parent: tk.Misc,
        message: str,
        *,
        kind: str = "info",
        duration: int = 2500,
    ) -> None:
        super().__init__(parent)
        self.overrideredirect(True)
        try:
            self.attributes("-topmost", True)
        except Exception:
            pass
        bg, fg = STATUS_COLOR.get(
            {"info": "ready", "success": "done", "warn": "paused", "error": "blocked"}.get(kind, "ready"),
            (UI["INFO_L"], UI["INFO"]),
        )
        self.configure(bg=bg)
        frame = tk.Frame(self, bg=bg, padx=16, pady=10)
        frame.pack()
        tk.Label(
            frame,
            text=message,
            bg=bg,
            fg=fg,
            font=FONTS["body_b"],
            justify="right",
        ).pack()
        self.update_idletasks()
        try:
            rx = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            ry = parent.winfo_rooty() + 20
            self.geometry(f"+{rx}+{ry}")
        except Exception:
            pass
        self.after(duration, self.destroy)


def toast(parent: tk.Misc, message: str, kind: str = "info") -> None:
    try:
        Toast(parent, message, kind=kind)
    except Exception:
        log.exception("toast failed")


def require_manager(parent: tk.Misc, db: DB, action_label: str) -> bool:  # noqa: ARG001
    """Password gating is disabled — sensitive actions run immediately.

    Call sites still invoke this for future re-enablement; we keep the
    signature stable but always grant access.
    """
    return True


def hint_bar(parent: tk.Misc, *lines: str) -> tk.Frame:
    """Build a visually distinct instruction banner for admin screens."""
    bar = tk.Frame(
        parent, bg=UI["INFO_L"], highlightbackground=UI["INFO"],
        highlightthickness=1, padx=14, pady=8,
    )
    tk.Label(
        bar, text="💡  إرشادات", bg=UI["INFO_L"], fg=UI["INFO"],
        font=FONTS["body_b"],
    ).pack(anchor="e")
    for line in lines:
        tk.Label(
            bar, text=line, bg=UI["INFO_L"], fg=UI["TEXT"],
            font=FONTS["small"], justify="right", anchor="e", wraplength=1100,
        ).pack(anchor="e", pady=1)
    return bar


# =============================================================================
# Reusable form helpers
# =============================================================================

def labeled_entry(
    parent: tk.Misc, label: str, *, width: int = 24
) -> Tuple[tk.Frame, ttk.Entry, tk.StringVar]:
    frame = tk.Frame(parent, bg=UI["SURFACE"])
    tk.Label(frame, text=label, bg=UI["SURFACE"], fg=UI["TEXT"], font=FONTS["body"]).pack(anchor="e")
    var = tk.StringVar()
    ent = ttk.Entry(frame, textvariable=var, width=width, justify="right")
    ent.pack(fill=tk.X, pady=(2, 0))
    return frame, ent, var


def labeled_combo(
    parent: tk.Misc, label: str, values: List[str], *, width: int = 24
) -> Tuple[tk.Frame, ttk.Combobox, tk.StringVar]:
    frame = tk.Frame(parent, bg=UI["SURFACE"])
    tk.Label(frame, text=label, bg=UI["SURFACE"], fg=UI["TEXT"], font=FONTS["body"]).pack(anchor="e")
    var = tk.StringVar()
    combo = ttk.Combobox(frame, textvariable=var, values=values, state="readonly", width=width, justify="right")
    combo.pack(fill=tk.X, pady=(2, 0))
    return frame, combo, var


def tree_with_scrollbars(
    parent: tk.Misc,
    *,
    columns: Tuple[str, ...],
    height: int,
) -> ttk.Treeview:
    """Create a treeview hosted inside a frame with both scrollbars."""
    wrap = tk.Frame(parent, bg=UI["SURFACE"])
    wrap.pack(fill=tk.BOTH, expand=True)
    tree = ttk.Treeview(wrap, columns=columns, show="headings", height=height)
    vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side=tk.RIGHT, fill=tk.Y)
    hsb.pack(side=tk.BOTTOM, fill=tk.X)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    return tree


# =============================================================================
# Admin tabs: Stages, Products, Workers
# =============================================================================


class _BaseAdminFrame(ttk.Frame):
    """Common scaffolding for list-edit admin screens."""

    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.configure(style="TFrame")


class SkillsAdmin(_BaseAdminFrame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent, app)
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="المهارات", style="H2.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="➕ إضافة مهارة جديدة", style="Ok.TButton",
                   command=self._add).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="✏️ تعديل المختارة",
                   command=self._edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="🗑️ حذف المختارة", style="Danger.TButton",
                   command=self._delete).pack(side=tk.LEFT, padx=4)

        hint_bar(
            self,
            "المهارات هي القدرات التي يمتلكها العمال (مثل: قص، حياكة، كوي).",
            "كل مرحلة في الإنتاج تتطلب مهارة معينة، ويتم توجيه العمل فقط للعمال الذين يمتلكونها.",
            "لإضافة مهارة: اضغط زر « ➕ إضافة مهارة جديدة » أعلاه.",
            "للتعديل أو الحذف: اختر سطرًا من الجدول ثم اضغط الزر المناسب.",
        ).pack(fill=tk.X, padx=12, pady=(0, 8))

        cols = ("id", "name", "notes")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        self.tree.heading("id", text="#")
        self.tree.heading("name", text="المهارة")
        self.tree.heading("notes", text="ملاحظات")
        self.tree.column("id", width=60, anchor="center")
        self.tree.column("name", width=220, anchor="e")
        self.tree.column("notes", width=400, anchor="e")
        self.tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for r in self.db.query("SELECT id,name,notes FROM skills ORDER BY name"):
            self.tree.insert("", "end", values=(r["id"], r["name"], r["notes"] or ""))

    def _selected(self) -> Optional[int]:
        sel = self.tree.focus()
        if not sel:
            return None
        return int(self.tree.item(sel)["values"][0])

    def _add(self) -> None:
        name = simpledialog.askstring("إضافة مهارة", "اسم المهارة:", parent=self)
        if not name:
            return
        notes = simpledialog.askstring("إضافة مهارة", "ملاحظات (اختياري):", parent=self) or ""
        try:
            self.db.execute("INSERT INTO skills(name,notes) VALUES(?,?)", (name.strip(), notes.strip()))
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "المهارة موجودة بالفعل")
            return
        self.refresh()
        self.app.invalidate_caches()

    def _edit(self) -> None:
        sid = self._selected()
        if not sid:
            return
        row = self.db.query_one("SELECT * FROM skills WHERE id=?", (sid,))
        if not row:
            return
        name = simpledialog.askstring("تعديل مهارة", "اسم المهارة:", initialvalue=row["name"], parent=self)
        if not name:
            return
        notes = simpledialog.askstring("تعديل مهارة", "ملاحظات:", initialvalue=row["notes"] or "", parent=self) or ""
        self.db.execute("UPDATE skills SET name=?, notes=? WHERE id=?", (name.strip(), notes.strip(), sid))
        self.refresh()
        self.app.invalidate_caches()

    def _delete(self) -> None:
        sid = self._selected()
        if not sid:
            return
        if not messagebox.askyesno("تأكيد", "حذف هذه المهارة؟", parent=self):
            return
        try:
            self.db.execute("DELETE FROM skills WHERE id=?", (sid,))
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "لا يمكن الحذف — المهارة مستخدمة")
            return
        self.refresh()
        self.app.invalidate_caches()


class StagesAdmin(_BaseAdminFrame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent, app)
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="قوالب المراحل", style="H2.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="➕ إضافة مرحلة جديدة", style="Ok.TButton",
                   command=self._add).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="✏️ تعديل المختارة",
                   command=self._edit).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="⏼ تفعيل/تعطيل", style="Warn.TButton",
                   command=self._toggle_active).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="🗑️ حذف", style="Danger.TButton",
                   command=self._delete).pack(side=tk.LEFT, padx=4)

        hint_bar(
            self,
            "قوالب المراحل تمثل خطوات الإنتاج القابلة لإعادة الاستخدام (مثل: قص، حياكة، كوي).",
            "لكل مرحلة: وقت إعداد ثابت + وقت لكل قطعة × الكمية = إجمالي وقت المرحلة.",
            "المرحلة نفسها هي معيار الإسناد للعمال (لا يوجد تبويب مهارات منفصل).",
            "لإضافة مرحلة: اضغط « ➕ إضافة مرحلة جديدة » ثم املأ الاسم والأوقات.",
            "بعد إنشاء المراحل، اذهب إلى تبويب « المنتجات » لربطها بمنتج في مسار إنتاج.",
        ).pack(fill=tk.X, padx=12, pady=(0, 8))

        cols = ("id", "code", "name", "setup", "per_unit", "active")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        for c, t, w, a in [
            ("id", "#", 50, "center"),
            ("code", "الكود", 100, "center"),
            ("name", "الاسم", 180, "e"),
            ("setup", "إعداد (د)", 90, "center"),
            ("per_unit", "لكل قطعة (د)", 110, "center"),
            ("active", "فعّال", 70, "center"),
        ]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        rows = self.db.query(
            """SELECT st.*
               FROM stage_templates st
               ORDER BY st.is_active DESC, st.name"""
        )
        for r in rows:
            self.tree.insert(
                "", "end",
                values=(
                    r["id"],
                    r["code"],
                    r["name"],
                    f"{r['default_setup_minutes']:.1f}",
                    f"{r['default_per_unit_minutes']:.2f}",
                    "نعم" if r["is_active"] else "لا",
                ),
            )

    def _selected(self) -> Optional[int]:
        sel = self.tree.focus()
        return int(self.tree.item(sel)["values"][0]) if sel else None

    def _prompt_stage_dialog(
        self, initial: Optional[sqlite3.Row] = None
    ) -> Optional[Dict[str, Any]]:
        dlg = tk.Toplevel(self)
        dlg.title("قالب مرحلة")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        code_var = tk.StringVar(value=initial["code"] if initial else "")
        name_var = tk.StringVar(value=initial["name"] if initial else "")
        setup_var = tk.StringVar(
            value=f"{initial['default_setup_minutes']:.1f}" if initial else "0"
        )
        per_unit_var = tk.StringVar(
            value=f"{initial['default_per_unit_minutes']:.2f}" if initial else "0"
        )
        notes_var = tk.StringVar(value=(initial["notes"] if initial else "") or "")

        pad = {"padx": 10, "pady": 5}
        for row, (lbl, var) in enumerate([
            ("الكود", code_var),
            ("الاسم", name_var),
            ("وقت الإعداد (دقيقة)", setup_var),
            ("الوقت لكل قطعة (دقيقة)", per_unit_var),
            ("ملاحظات", notes_var),
        ]):
            tk.Label(dlg, text=lbl, bg=UI["SURFACE"], fg=UI["TEXT"]).grid(row=row, column=1, sticky="e", **pad)
            ttk.Entry(dlg, textvariable=var, width=28, justify="right").grid(row=row, column=0, sticky="we", **pad)
        result: Dict[str, Any] = {}

        def ok() -> None:
            try:
                setup = float(setup_var.get() or 0)
                per_unit = float(per_unit_var.get() or 0)
            except ValueError:
                messagebox.showerror("خطأ", "أدخل قيمًا رقمية صحيحة للأوقات", parent=dlg)
                return
            if not code_var.get().strip() or not name_var.get().strip():
                messagebox.showerror("خطأ", "الكود والاسم مطلوبان", parent=dlg)
                return
            result.update({
                "code": code_var.get().strip().upper(),
                "name": name_var.get().strip(),
                "setup": setup,
                "per_unit": per_unit,
                "notes": notes_var.get().strip(),
            })
            dlg.destroy()

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.grid(row=5, column=0, columnspan=2, pady=12)
        ttk.Button(btns, text="حفظ", style="Ok.TButton", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton", command=dlg.destroy).pack(side=tk.LEFT, padx=6)

        dlg.wait_window()
        return result or None

    def _add(self) -> None:
        data = self._prompt_stage_dialog()
        if not data:
            return
        try:
            self.db.execute(
                """INSERT INTO stage_templates
                   (code,name,default_setup_minutes,default_per_unit_minutes,notes)
                   VALUES(?,?,?,?,?)""",
                (data["code"], data["name"], data["setup"], data["per_unit"], data["notes"]),
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "كود المرحلة موجود مسبقًا")
            return
        self.refresh()
        self.app.invalidate_caches()

    def _edit(self) -> None:
        sid = self._selected()
        if not sid:
            return
        row = self.db.query_one("SELECT * FROM stage_templates WHERE id=?", (sid,))
        if not row:
            return
        data = self._prompt_stage_dialog(row)
        if not data:
            return
        try:
            self.db.execute(
                """UPDATE stage_templates SET
                    code=?, name=?, default_setup_minutes=?, default_per_unit_minutes=?,
                    notes=?
                   WHERE id=?""",
                (data["code"], data["name"], data["setup"], data["per_unit"], data["notes"], sid),
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "كود المرحلة موجود مسبقًا")
            return
        self.refresh()
        self.app.invalidate_caches()

    def _toggle_active(self) -> None:
        sid = self._selected()
        if not sid:
            return
        self.db.execute(
            "UPDATE stage_templates SET is_active = 1 - is_active WHERE id=?", (sid,)
        )
        self.refresh()

    def _delete(self) -> None:
        sid = self._selected()
        if not sid:
            return
        if not messagebox.askyesno("تأكيد", "حذف هذا القالب؟", parent=self):
            return
        # Show explicit "where used" details before trying delete.
        refs = self.db.query(
            """SELECT p.id, p.code, COALESCE(p.product_type, p.name) AS product_name
               FROM product_stages ps
               JOIN products p ON p.id = ps.product_id
               WHERE ps.stage_template_id=? AND p.is_active=1
               ORDER BY p.product_type, p.name""",
            (sid,),
        )
        if refs:
            preview = refs[:8]
            lines = [f"- {r['product_name']} ({r['code']})" for r in preview]
            if len(refs) > len(preview):
                lines.append(f"... +{len(refs) - len(preview)} منتج إضافي")
            msg = (
                "لا يمكن حذف هذه المرحلة لأنها مستخدمة في المنتجات التالية:\n\n"
                + "\n".join(lines)
                + "\n\nاحذف المرحلة أولًا من مسار هذه المنتجات (تبويب «المنتجات»)."
            )
            if messagebox.askyesno(
                "المرحلة مستخدمة",
                msg + "\n\nهل تريد الانتقال الآن إلى أول منتج؟",
                parent=self,
            ):
                self._jump_to_product(int(refs[0]["id"]))
            return
        try:
            self.db.execute("DELETE FROM stage_templates WHERE id=?", (sid,))
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "لا يمكن حذف المرحلة الآن بسبب ارتباطات موجودة.")
            return
        self.refresh()
        self.app.invalidate_caches()

    def _jump_to_product(self, product_id: int) -> None:
        """Open products tab and focus a given product id."""
        try:
            self.app.nb.select(self.app.products)
            self.app.products.refresh_products()
            tree = self.app.products.products_tree
            for iid in tree.get_children():
                vals = tree.item(iid).get("values", [])
                if vals and int(vals[0]) == int(product_id):
                    tree.selection_set(iid)
                    tree.focus(iid)
                    self.app.products._on_product_select()
                    tree.see(iid)
                    break
        except Exception:
            log.exception("failed to jump to product %s", product_id)


class ProductsAdmin(_BaseAdminFrame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent, app)
        self.current_product_id: Optional[int] = None
        self._build()
        self.refresh_products()

    def _build(self) -> None:
        header = tk.Frame(self, bg=UI["BG"])
        header.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(header, text="أنواع المنتجات ومسارات الإنتاج", style="H2.TLabel").pack(side=tk.RIGHT)

        hint_bar(
            self,
            "هذه الشاشة تحدد « أنواع المنتجات » فقط (قميص، بنطلون، جاكيت …) ومسار إنتاج كل نوع.",
            "تفاصيل الطلب نفسها (المدرسة، المقاس، اللون، الكمية) تُدخَل عند إنشاء أمر إنتاج جديد.",
            "الخطوة 1 — على اليمين: أضف نوعًا أو اختر نوعًا موجودًا.",
            "الخطوة 2 — على اليسار: أضف المراحل للمسار بالترتيب (قص ← حياكة ← كوي ...).",
            "يمكنك تحديد المرحلة « افتراضية » (مختارة تلقائيًا في الأوامر الجديدة) أو « اختيارية » (يمكن تبديلها في الأمر).",
        ).pack(fill=tk.X, padx=12, pady=(0, 8))

        body = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        # Left: product list
        left = tk.Frame(body, bg=UI["BG"], padx=2, pady=2)
        ttk.Label(left, text="1️⃣ أنواع المنتجات", style="H3.TLabel").pack(anchor="e")
        btns = tk.Frame(left, bg=UI["BG"])
        btns.pack(fill=tk.X, pady=(2, 6), anchor="e")
        ttk.Button(btns, text="➕ نوع منتج جديد", style="Ok.TButton",
                   command=self._add_product).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="✏️ تعديل",
                   command=self._edit_product).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="🗑️ حذف", style="Danger.TButton",
                   command=self._delete_product).pack(side=tk.LEFT, padx=3)

        cols = ("id", "code", "type")
        self.products_tree = tree_with_scrollbars(left, columns=cols, height=14)
        for c, t, w, a in [
            ("id",   "#",       40,  "center"),
            ("code", "الكود",   100, "center"),
            ("type", "النوع",   220, "e"),
        ]:
            self.products_tree.heading(c, text=t)
            self.products_tree.column(c, width=w, anchor=a)
        self.products_tree.bind("<<TreeviewSelect>>", self._on_product_select)

        # Right: routing editor
        right = tk.Frame(body, bg=UI["BG"], padx=2, pady=2)
        self.routing_title = ttk.Label(right, text="2️⃣ مسار الإنتاج — اختر منتجًا من اليمين",
                                       style="H3.TLabel")
        self.routing_title.pack(anchor="e")
        rbtns = tk.Frame(right, bg=UI["BG"])
        rbtns.pack(fill=tk.X, pady=(2, 6), anchor="e")
        ttk.Button(rbtns, text="➕ إضافة مرحلة للمسار", style="Ok.TButton",
                   command=self._add_routing).pack(side=tk.LEFT, padx=3)
        ttk.Button(rbtns, text="✏️ تعديل",
                   command=self._edit_routing).pack(side=tk.LEFT, padx=3)
        ttk.Button(rbtns, text="⬆ لأعلى",
                   command=lambda: self._move_routing(-1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(rbtns, text="⬇ لأسفل",
                   command=lambda: self._move_routing(+1)).pack(side=tk.LEFT, padx=3)
        ttk.Button(rbtns, text="🗑️ حذف", style="Danger.TButton",
                   command=self._delete_routing).pack(side=tk.LEFT, padx=3)

        cols = ("id", "seq", "stage", "default", "optional", "setup", "per_unit")
        self.routing_tree = tree_with_scrollbars(right, columns=cols, height=14)
        for c, t, w, a in [
            ("id", "#", 50, "center"),
            ("seq", "الترتيب", 80, "center"),
            ("stage", "المرحلة", 180, "e"),
            ("default", "افتراضي", 90, "center"),
            ("optional", "اختياري", 90, "center"),
            ("setup", "إعداد (د)", 90, "center"),
            ("per_unit", "لكل قطعة (د)", 110, "center"),
        ]:
            self.routing_tree.heading(c, text=t)
            self.routing_tree.column(c, width=w, anchor=a)
        body.add(left, weight=2)
        body.add(right, weight=3)

    # --- products list --- #

    def refresh_products(self) -> None:
        self.products_tree.delete(*self.products_tree.get_children())
        rows = self.db.query(
            """SELECT id,code,name,product_type
               FROM products
               WHERE is_active=1
               ORDER BY product_type"""
        )
        for r in rows:
            self.products_tree.insert(
                "", "end",
                values=(r["id"], r["code"], r["product_type"] or r["name"] or "—"),
            )
        if rows:
            first = self.products_tree.get_children()[0]
            self.products_tree.selection_set(first)
            self.products_tree.focus(first)
            self._on_product_select()
        else:
            self.current_product_id = None
            self.routing_tree.delete(*self.routing_tree.get_children())
            self.routing_title.config(text="لا يوجد أنواع — اضغط « ➕ نوع منتج جديد »")

    def _on_product_select(self, _evt=None) -> None:
        sel = self.products_tree.focus()
        if not sel:
            return
        pid = int(self.products_tree.item(sel)["values"][0])
        self.current_product_id = pid
        row = self.db.query_one("SELECT code,name FROM products WHERE id=?", (pid,))
        if row:
            self.routing_title.config(
                text=f"2️⃣ مسار الإنتاج — {row['name']} ({row['code']})"
            )
        self._refresh_routing()

    def _prompt_product_dialog(
        self, initial: Optional[sqlite3.Row] = None
    ) -> Optional[Dict[str, str]]:
        """Modal dialog for creating/editing a product type.

        Fields: product_type (the canonical name), code (auto), notes.
        The `name` stored in DB mirrors `product_type` for back-compat.
        """
        is_edit = initial is not None
        dlg = tk.Toplevel(self)
        dlg.title("تعديل نوع منتج" if is_edit else "إضافة نوع منتج")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()
        dlg.resizable(False, False)

        type_values = sorted(set(self.db.distinct_product_values("product_type")) |
                             set(COMMON_PRODUCT_TYPES))

        type_var  = tk.StringVar(value=(initial["product_type"] if is_edit else "") or "")
        code_var  = tk.StringVar(value=(initial["code"]         if is_edit else
                                         self.db.next_product_code()))
        notes_var = tk.StringVar(value=(initial["notes"]        if is_edit else "") or "")

        pad = {"padx": 10, "pady": 6}

        tk.Label(dlg, text="نوع المنتج", bg=UI["SURFACE"], fg=UI["ACCENT"],
                 font=FONTS["h3"]).grid(row=0, column=0, columnspan=2, sticky="e", **pad)

        tk.Label(dlg, text="النوع (Type)", bg=UI["SURFACE"], font=FONTS["body_b"]).grid(
            row=1, column=1, sticky="e", **pad
        )
        ttk.Combobox(dlg, textvariable=type_var, values=type_values, width=34,
                     justify="right").grid(row=1, column=0, sticky="we", **pad)

        tk.Label(dlg, text="الكود", bg=UI["SURFACE"], font=FONTS["body_b"]).grid(
            row=2, column=1, sticky="e", **pad
        )
        ttk.Entry(dlg, textvariable=code_var, width=36, justify="right").grid(
            row=2, column=0, sticky="we", **pad
        )

        tk.Label(dlg, text="ملاحظات", bg=UI["SURFACE"], font=FONTS["body_b"]).grid(
            row=3, column=1, sticky="e", **pad
        )
        ttk.Entry(dlg, textvariable=notes_var, width=36, justify="right").grid(
            row=3, column=0, sticky="we", **pad
        )

        tk.Label(
            dlg,
            text="المدرسة، المقاس، اللون، والكمية تُدخل عند إنشاء أمر الإنتاج.",
            bg=UI["SURFACE"], fg=UI["TEXT_SEC"], font=FONTS["small"],
            justify="right", wraplength=400,
        ).grid(row=4, column=0, columnspan=2, sticky="e", **pad)

        result: Dict[str, str] = {}

        def _ok() -> None:
            t = type_var.get().strip()
            if not t:
                messagebox.showerror(
                    "حقل مطلوب", "الرجاء إدخال نوع المنتج.", parent=dlg
                )
                return
            code = code_var.get().strip() or self.db.next_product_code()
            result.update({
                "code": code,
                "name": t,
                "product_type": t,
                "notes": notes_var.get().strip(),
            })
            dlg.destroy()

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.grid(row=5, column=0, columnspan=2, pady=(12, 14))
        ttk.Button(btns, text="💾 حفظ", style="Ok.TButton", command=_ok).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton",
                   command=dlg.destroy).pack(side=tk.LEFT, padx=6)

        dlg.wait_window()
        return result or None

    def _add_product(self) -> None:
        data = self._prompt_product_dialog()
        if not data:
            return
        try:
            self.db.execute(
                """INSERT INTO products(code, name, product_type, notes)
                   VALUES(?,?,?,?)""",
                (data["code"], data["name"], data["product_type"], data["notes"]),
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "كود المنتج موجود مسبقًا — غيّر الكود يدويًا")
            return
        self.refresh_products()
        self.app.invalidate_caches()

    def _edit_product(self) -> None:
        if not self.current_product_id:
            return
        row = self.db.query_one(
            "SELECT * FROM products WHERE id=?", (self.current_product_id,)
        )
        if not row:
            return
        data = self._prompt_product_dialog(row)
        if not data:
            return
        try:
            self.db.execute(
                """UPDATE products SET code=?, name=?, product_type=?, notes=?
                   WHERE id=?""",
                (data["code"], data["name"], data["product_type"],
                 data["notes"], self.current_product_id),
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "كود المنتج موجود مسبقًا")
            return
        self.refresh_products()
        self.app.invalidate_caches()

    def _delete_product(self) -> None:
        if not self.current_product_id:
            return
        if not messagebox.askyesno(
            "تأكيد",
            "حذف هذا المنتج (سيتم الاحتفاظ بأوامر الإنتاج السابقة)؟",
            parent=self,
        ):
            return
        try:
            self.db.execute("UPDATE products SET is_active=0 WHERE id=?", (self.current_product_id,))
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "لا يمكن إخفاء المنتج")
            return
        self.refresh_products()
        self.app.invalidate_caches()

    # --- routing editor --- #

    def _refresh_routing(self) -> None:
        self.routing_tree.delete(*self.routing_tree.get_children())
        if not self.current_product_id:
            return
        rows = self.db.query(
            """SELECT ps.*, st.code, st.name, st.default_setup_minutes, st.default_per_unit_minutes
               FROM product_stages ps
               JOIN stage_templates st ON st.id = ps.stage_template_id
               WHERE ps.product_id=?
               ORDER BY ps.sequence""",
            (self.current_product_id,),
        )
        for r in rows:
            setup = r["override_setup_minutes"] if r["override_setup_minutes"] is not None else r["default_setup_minutes"]
            per_unit = r["override_per_unit_minutes"] if r["override_per_unit_minutes"] is not None else r["default_per_unit_minutes"]
            self.routing_tree.insert(
                "", "end",
                values=(
                    r["id"],
                    r["sequence"],
                    f"{r['name']} ({r['code']})",
                    "نعم" if r["is_default"] else "لا",
                    "نعم" if r["is_optional"] else "لا",
                    f"{setup:.1f}",
                    f"{per_unit:.2f}",
                ),
            )

    def _routing_selected(self) -> Optional[int]:
        sel = self.routing_tree.focus()
        return int(self.routing_tree.item(sel)["values"][0]) if sel else None

    def _prompt_routing_dialog(
        self, initial: Optional[sqlite3.Row] = None
    ) -> Optional[Dict[str, Any]]:
        dlg = tk.Toplevel(self)
        dlg.title("مرحلة في المسار")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        stages = self.db.query(
            "SELECT id,code,name FROM stage_templates WHERE is_active=1 ORDER BY name"
        )
        stage_labels = [f"{s['name']} ({s['code']})" for s in stages]
        stage_ids = [s["id"] for s in stages]
        stage_var = tk.StringVar(value=stage_labels[0] if stage_labels else "")
        if initial:
            for i, sid in enumerate(stage_ids):
                if sid == initial["stage_template_id"]:
                    stage_var.set(stage_labels[i])
                    break

        default_var = tk.BooleanVar(value=bool(initial["is_default"]) if initial else True)
        optional_var = tk.BooleanVar(value=bool(initial["is_optional"]) if initial else False)
        setup_var = tk.StringVar(
            value="" if (initial is None or initial["override_setup_minutes"] is None)
            else f"{initial['override_setup_minutes']:.2f}"
        )
        per_unit_var = tk.StringVar(
            value="" if (initial is None or initial["override_per_unit_minutes"] is None)
            else f"{initial['override_per_unit_minutes']:.2f}"
        )

        pad = {"padx": 10, "pady": 5}
        tk.Label(dlg, text="المرحلة", bg=UI["SURFACE"]).grid(row=0, column=1, sticky="e", **pad)
        ttk.Combobox(dlg, textvariable=stage_var, values=stage_labels, state="readonly", width=30, justify="right").grid(row=0, column=0, sticky="we", **pad)

        tk.Checkbutton(dlg, text="افتراضية (مختارة تلقائيًا)", variable=default_var, bg=UI["SURFACE"], anchor="e").grid(row=1, column=0, columnspan=2, sticky="e", **pad)
        tk.Checkbutton(dlg, text="اختيارية (يمكن إلغاؤها)", variable=optional_var, bg=UI["SURFACE"], anchor="e").grid(row=2, column=0, columnspan=2, sticky="e", **pad)

        tk.Label(dlg, text="إعداد مخصص (د) — فارغ = افتراضي", bg=UI["SURFACE"]).grid(row=3, column=1, sticky="e", **pad)
        ttk.Entry(dlg, textvariable=setup_var, width=30, justify="right").grid(row=3, column=0, sticky="we", **pad)
        tk.Label(dlg, text="لكل قطعة مخصص (د) — فارغ = افتراضي", bg=UI["SURFACE"]).grid(row=4, column=1, sticky="e", **pad)
        ttk.Entry(dlg, textvariable=per_unit_var, width=30, justify="right").grid(row=4, column=0, sticky="we", **pad)

        result: Dict[str, Any] = {}

        def ok() -> None:
            if not stage_var.get():
                return
            try:
                override_setup = float(setup_var.get()) if setup_var.get().strip() else None
                override_pu = float(per_unit_var.get()) if per_unit_var.get().strip() else None
            except ValueError:
                messagebox.showerror("خطأ", "قيم رقمية غير صالحة", parent=dlg)
                return
            idx = stage_labels.index(stage_var.get())
            result.update({
                "stage_template_id": stage_ids[idx],
                "is_default": 1 if default_var.get() else 0,
                "is_optional": 1 if optional_var.get() else 0,
                "override_setup_minutes": override_setup,
                "override_per_unit_minutes": override_pu,
            })
            dlg.destroy()

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.grid(row=5, column=0, columnspan=2, pady=12)
        ttk.Button(btns, text="حفظ", style="Ok.TButton", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton", command=dlg.destroy).pack(side=tk.LEFT, padx=6)
        dlg.wait_window()
        return result or None

    def _add_routing(self) -> None:
        if not self.current_product_id:
            messagebox.showinfo("تنبيه", "اختر منتجًا أولًا", parent=self)
            return
        data = self._prompt_routing_dialog()
        if not data:
            return
        row = self.db.query_one(
            "SELECT COALESCE(MAX(sequence),0) AS m FROM product_stages WHERE product_id=?",
            (self.current_product_id,),
        )
        nxt_seq = int(row["m"]) + 1
        try:
            self.db.execute(
                """INSERT INTO product_stages
                   (product_id, stage_template_id, sequence, is_default, is_optional,
                    override_setup_minutes, override_per_unit_minutes)
                   VALUES(?,?,?,?,?,?,?)""",
                (
                    self.current_product_id,
                    data["stage_template_id"],
                    nxt_seq,
                    data["is_default"],
                    data["is_optional"],
                    data["override_setup_minutes"],
                    data["override_per_unit_minutes"],
                ),
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "لا يمكن إضافة هذه المرحلة")
            return
        self._refresh_routing()

    def _edit_routing(self) -> None:
        rid = self._routing_selected()
        if not rid:
            return
        row = self.db.query_one("SELECT * FROM product_stages WHERE id=?", (rid,))
        if not row:
            return
        data = self._prompt_routing_dialog(row)
        if not data:
            return
        self.db.execute(
            """UPDATE product_stages SET
                stage_template_id=?, is_default=?, is_optional=?,
                override_setup_minutes=?, override_per_unit_minutes=?
               WHERE id=?""",
            (
                data["stage_template_id"],
                data["is_default"],
                data["is_optional"],
                data["override_setup_minutes"],
                data["override_per_unit_minutes"],
                rid,
            ),
        )
        self._refresh_routing()

    def _move_routing(self, delta: int) -> None:
        rid = self._routing_selected()
        if not rid or not self.current_product_id:
            return
        rows = self.db.query(
            "SELECT id, sequence FROM product_stages WHERE product_id=? ORDER BY sequence",
            (self.current_product_id,),
        )
        ids = [r["id"] for r in rows]
        if rid not in ids:
            return
        idx = ids.index(rid)
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(ids):
            return
        a, b = ids[idx], ids[new_idx]
        sa = rows[idx]["sequence"]
        sb = rows[new_idx]["sequence"]
        # swap sequences (use large temp to avoid unique violation)
        tmp = max(r["sequence"] for r in rows) + 100
        self.db.execute("UPDATE product_stages SET sequence=? WHERE id=?", (tmp, a))
        self.db.execute("UPDATE product_stages SET sequence=? WHERE id=?", (sa, b))
        self.db.execute("UPDATE product_stages SET sequence=? WHERE id=?", (sb, a))
        self._refresh_routing()
        for child in self.routing_tree.get_children():
            if int(self.routing_tree.item(child)["values"][0]) == a:
                self.routing_tree.selection_set(child)
                self.routing_tree.focus(child)
                break

    def _delete_routing(self) -> None:
        rid = self._routing_selected()
        if not rid:
            return
        if not messagebox.askyesno("تأكيد", "حذف هذه المرحلة من المسار؟", parent=self):
            return
        self.db.execute("DELETE FROM product_stages WHERE id=?", (rid,))
        self._refresh_routing()


class WorkersAdmin(_BaseAdminFrame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent, app)
        self.current_worker_id: Optional[int] = None
        self._build()
        self.refresh_workers()

    def _build(self) -> None:
        header = tk.Frame(self, bg=UI["BG"])
        header.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(header, text="العمال والمراحل المسموح بها", style="H2.TLabel").pack(side=tk.RIGHT)

        hint_bar(
            self,
            "هذه الشاشة لإدارة العمال وتحديد المراحل التي يمكن لكل عامل تنفيذها.",
            "الخطوة 1 — على اليمين: أضف العامل (كود + اسم).",
            "الخطوة 2 — على اليسار: اختر العامل، ثم فعّل المراحل التي يُسمح له بها.",
            "عند تعيين عامل لمرحلة في أمر إنتاج، يظهر فقط العمال المصرّح لهم بهذه المرحلة.",
        ).pack(fill=tk.X, padx=12, pady=(0, 8))

        body = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        left = tk.Frame(body, bg=UI["BG"], padx=2, pady=2)
        ttk.Label(left, text="1️⃣ العمال", style="H3.TLabel").pack(anchor="e")
        btns = tk.Frame(left, bg=UI["BG"])
        btns.pack(fill=tk.X, pady=(2, 6), anchor="e")
        ttk.Button(btns, text="➕ عامل جديد", style="Ok.TButton",
                   command=self._add_worker).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="✏️ تعديل",
                   command=self._edit_worker).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="⏼ تفعيل/تعطيل", style="Warn.TButton",
                   command=self._toggle_worker).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="🗑️ حذف", style="Danger.TButton",
                   command=self._delete_worker).pack(side=tk.LEFT, padx=3)

        self.workers_tree = tree_with_scrollbars(
            left, columns=("id", "code", "name", "active"), height=14
        )
        for c, t, w, a in [("id", "#", 50, "center"), ("code", "الكود", 90, "center"),
                           ("name", "الاسم", 200, "e"), ("active", "فعّال", 70, "center")]:
            self.workers_tree.heading(c, text=t)
            self.workers_tree.column(c, width=w, anchor=a)
        self.workers_tree.bind("<<TreeviewSelect>>", self._on_worker_select)

        right = tk.Frame(body, bg=UI["BG"], padx=2, pady=2)
        self.skills_title = ttk.Label(right, text="2️⃣ مراحل العامل — اختر عاملًا من اليمين",
                                      style="H3.TLabel")
        self.skills_title.pack(anchor="e")
        sbtns = tk.Frame(right, bg=UI["BG"])
        sbtns.pack(fill=tk.X, pady=(2, 6), anchor="e")
        ttk.Button(sbtns, text="✓ السماح بالمرحلة المختارة",   style="Ok.TButton",
                   command=lambda: self._toggle_skill(enable=True)).pack(side=tk.LEFT, padx=3)
        ttk.Button(sbtns, text="✗ منع المرحلة المختارة",   style="Danger.TButton",
                   command=lambda: self._toggle_skill(enable=False)).pack(side=tk.LEFT, padx=3)

        self.skills_list = tree_with_scrollbars(
            right, columns=("id", "name", "has"), height=14
        )
        for c, t, w, a in [("id", "#", 50, "center"),
                           ("name", "المرحلة", 240, "e"),
                           ("has", "مسموح", 90, "center")]:
            self.skills_list.heading(c, text=t)
            self.skills_list.column(c, width=w, anchor=a)
        body.add(left, weight=3)
        body.add(right, weight=2)

    def refresh_workers(self) -> None:
        self.workers_tree.delete(*self.workers_tree.get_children())
        rows = self.db.query("SELECT id,code,name,is_active FROM workers ORDER BY is_active DESC, name")
        for r in rows:
            self.workers_tree.insert("", "end",
                                     values=(r["id"], r["code"], r["name"], "نعم" if r["is_active"] else "لا"))
        if rows:
            first = self.workers_tree.get_children()[0]
            self.workers_tree.selection_set(first)
            self.workers_tree.focus(first)
            self._on_worker_select()
        else:
            self.current_worker_id = None
            self.skills_list.delete(*self.skills_list.get_children())
            self.skills_title.config(text="لا يوجد عمال")

    def _on_worker_select(self, _evt=None) -> None:
        sel = self.workers_tree.focus()
        if not sel:
            return
        wid = int(self.workers_tree.item(sel)["values"][0])
        self.current_worker_id = wid
        row = self.db.query_one("SELECT name FROM workers WHERE id=?", (wid,))
        self.skills_title.config(text=f"مراحل العامل — {row['name']}" if row else "مراحل العامل")
        self._refresh_skills()

    def _refresh_skills(self) -> None:
        self.skills_list.delete(*self.skills_list.get_children())
        if not self.current_worker_id:
            return
        rows = self.db.query(
            """SELECT st.id, st.name, ws.id AS ws_id
               FROM stage_templates st
               LEFT JOIN worker_stages ws
                    ON ws.stage_template_id = st.id AND ws.worker_id = ?
               WHERE st.is_active=1
               ORDER BY st.name""",
            (self.current_worker_id,),
        )
        for r in rows:
            self.skills_list.insert(
                "", "end",
                values=(
                    r["id"], r["name"],
                    "نعم" if r["ws_id"] else "لا",
                ),
            )

    def _skill_selected(self) -> Optional[int]:
        sel = self.skills_list.focus()
        return int(self.skills_list.item(sel)["values"][0]) if sel else None

    # --- worker CRUD --- #

    def _prompt_worker(
        self, initial: Optional[sqlite3.Row] = None
    ) -> Optional[Tuple[str, str, str]]:
        dlg = tk.Toplevel(self)
        dlg.title("عامل")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self.winfo_toplevel())
        dlg.grab_set()

        code_var = tk.StringVar(value=initial["code"] if initial else "")
        name_var = tk.StringVar(value=initial["name"] if initial else "")
        notes_var = tk.StringVar(value=(initial["notes"] if initial else "") or "")

        pad = {"padx": 10, "pady": 5}
        for row, (lbl, var) in enumerate([("الكود", code_var), ("الاسم", name_var), ("ملاحظات", notes_var)]):
            tk.Label(dlg, text=lbl, bg=UI["SURFACE"]).grid(row=row, column=1, sticky="e", **pad)
            ttk.Entry(dlg, textvariable=var, width=28, justify="right").grid(row=row, column=0, sticky="we", **pad)

        result: Dict[str, str] = {}

        def ok() -> None:
            if not code_var.get().strip() or not name_var.get().strip():
                messagebox.showerror("خطأ", "الكود والاسم مطلوبان", parent=dlg)
                return
            result["code"] = code_var.get().strip()
            result["name"] = name_var.get().strip()
            result["notes"] = notes_var.get().strip()
            dlg.destroy()

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.grid(row=3, column=0, columnspan=2, pady=12)
        ttk.Button(btns, text="حفظ", style="Ok.TButton", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton", command=dlg.destroy).pack(side=tk.LEFT, padx=6)
        dlg.wait_window()
        if not result:
            return None
        return result["code"], result["name"], result["notes"]

    def _add_worker(self) -> None:
        data = self._prompt_worker()
        if not data:
            return
        code, name, notes = data
        try:
            self.db.execute(
                "INSERT INTO workers(code,name,notes) VALUES(?,?,?)", (code, name, notes)
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "كود العامل موجود مسبقًا")
            return
        self.refresh_workers()
        self.app.invalidate_caches()

    def _edit_worker(self) -> None:
        if not self.current_worker_id:
            return
        row = self.db.query_one("SELECT * FROM workers WHERE id=?", (self.current_worker_id,))
        if not row:
            return
        data = self._prompt_worker(row)
        if not data:
            return
        code, name, notes = data
        try:
            self.db.execute(
                "UPDATE workers SET code=?, name=?, notes=? WHERE id=?",
                (code, name, notes, self.current_worker_id),
            )
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "كود العامل موجود مسبقًا")
            return
        self.refresh_workers()
        self.app.invalidate_caches()

    def _toggle_worker(self) -> None:
        if not self.current_worker_id:
            return
        self.db.execute(
            "UPDATE workers SET is_active = 1 - is_active WHERE id=?",
            (self.current_worker_id,),
        )
        self.refresh_workers()

    def _delete_worker(self) -> None:
        if not self.current_worker_id:
            return
        if not messagebox.askyesno("تأكيد", "حذف هذا العامل؟ (لن يتم حذف تاريخه في الأوامر)", parent=self):
            return
        try:
            self.db.execute("UPDATE workers SET is_active=0 WHERE id=?", (self.current_worker_id,))
        except sqlite3.IntegrityError:
            show_error(self, "خطأ", "لا يمكن إخفاء العامل")
            return
        self.refresh_workers()
        self.app.invalidate_caches()

    # --- skills matrix --- #

    def _toggle_skill(self, *, enable: bool) -> None:
        sid = self._skill_selected()
        if not sid or not self.current_worker_id:
            return
        if enable:
            try:
                self.db.execute(
                    "INSERT INTO worker_stages(worker_id,stage_template_id) VALUES(?,?)",
                    (self.current_worker_id, sid),
                )
            except sqlite3.IntegrityError:
                pass
        else:
            self.db.execute(
                "DELETE FROM worker_stages WHERE worker_id=? AND stage_template_id=?",
                (self.current_worker_id, sid),
            )
        self._refresh_skills()


# =============================================================================
# New Order Wizard
# =============================================================================


class NewOrderWizard(tk.Toplevel):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.title("أمر إنتاج جديد")
        self.configure(bg=UI["BG"])
        self.transient(parent.winfo_toplevel())
        self.grab_set()
        self.geometry("840x720")

        self._product_rows = self.db.query(
            """SELECT id,code,name,product_type
               FROM products WHERE is_active=1
               ORDER BY product_type"""
        )
        # Past values pulled live from the order history for autocomplete.
        self._schools_hist = self.db.distinct_order_values("school_name")
        self._sizes_hist   = self.db.distinct_order_values("size")
        self._colors_hist  = self.db.distinct_order_values("color")
        self._qty_hist     = self.db.distinct_order_quantities()

        self._stage_rows: List[sqlite3.Row] = []
        self._checks: Dict[int, tk.BooleanVar] = {}
        self._start_idx_var = tk.IntVar(value=0)
        self._line_items: List[Dict[str, Any]] = []
        self._editing_line_index: Optional[int] = None
        self._build()

    @staticmethod
    def _product_label(p: sqlite3.Row) -> str:
        return f"{p['product_type'] or p['name']} ({p['code']})"

    def _build(self) -> None:
        header = tk.Frame(self, bg=UI["BG"])
        header.pack(fill=tk.X, padx=16, pady=12)
        ttk.Label(header, text="أمر إنتاج جديد", style="H1.TLabel").pack(side=tk.RIGHT)

        # Scrollable body: keeps all fields/stages reachable on smaller screens.
        body_wrap = tk.Frame(self, bg=UI["BG"])
        body_wrap.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 8))

        body_canvas = tk.Canvas(body_wrap, bg=UI["BG"], highlightthickness=0, bd=0)
        body_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        body_scroll = ttk.Scrollbar(body_wrap, orient="vertical", command=body_canvas.yview)
        body_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        body_canvas.configure(yscrollcommand=body_scroll.set)

        content = tk.Frame(body_canvas, bg=UI["BG"])
        win_id = body_canvas.create_window((0, 0), window=content, anchor="nw")

        def _sync_scrollregion(_e=None):
            body_canvas.configure(scrollregion=body_canvas.bbox("all"))

        def _sync_width(event):
            body_canvas.itemconfigure(win_id, width=event.width)

        content.bind("<Configure>", _sync_scrollregion)
        body_canvas.bind("<Configure>", _sync_width)

        def _wheel(event):
            body_canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")

        # Local mouse-wheel binding for this dialog only.
        self.bind("<MouseWheel>", _wheel)

        form = tk.Frame(content, bg=UI["SURFACE"], padx=16, pady=12)
        form.pack(fill=tk.X)

        # --- Row 0: Product type (dropdown — must pick from the catalog) ---
        tk.Label(form, text="نوع المنتج *", bg=UI["SURFACE"],
                 font=FONTS["body_b"]).grid(row=0, column=1, sticky="e", padx=8, pady=6)
        product_labels = [self._product_label(p) for p in self._product_rows]
        self.product_var = tk.StringVar(value=product_labels[0] if product_labels else "")
        self.product_combo = ttk.Combobox(
            form, values=product_labels, textvariable=self.product_var,
            state="readonly", width=36, justify="right",
        )
        self.product_combo.grid(row=0, column=0, sticky="we", padx=8, pady=6)
        self.product_combo.bind("<<ComboboxSelected>>", self._on_product_changed)

        # --- Row 1: School name (free text + autocomplete) ---
        tk.Label(form, text="اسم المدرسة *", bg=UI["SURFACE"],
                 font=FONTS["body_b"]).grid(row=1, column=1, sticky="e", padx=8, pady=6)
        self.school_var = tk.StringVar()
        ttk.Combobox(
            form, textvariable=self.school_var, values=self._schools_hist,
            width=36, justify="right",
        ).grid(row=1, column=0, sticky="we", padx=8, pady=6)

        # --- Row 2: Size (used by the line editor below) ---
        tk.Label(form, text="المقاس *", bg=UI["SURFACE"],
                 font=FONTS["body_b"]).grid(row=2, column=1, sticky="e", padx=8, pady=6)
        self.size_var = tk.StringVar()
        size_values = list(dict.fromkeys(self._sizes_hist + COMMON_SIZES))
        ttk.Combobox(
            form, textvariable=self.size_var, values=size_values,
            width=36, justify="right",
        ).grid(row=2, column=0, sticky="we", padx=8, pady=6)

        # --- Row 3: Color (used by the line editor below) ---
        tk.Label(form, text="اللون *", bg=UI["SURFACE"],
                 font=FONTS["body_b"]).grid(row=3, column=1, sticky="e", padx=8, pady=6)
        self.color_var = tk.StringVar()
        color_values = list(dict.fromkeys(self._colors_hist + COMMON_COLORS))
        ttk.Combobox(
            form, textvariable=self.color_var, values=color_values,
            width=36, justify="right",
        ).grid(row=3, column=0, sticky="we", padx=8, pady=6)

        # --- Row 4: Quantity (used by the line editor below) ---
        tk.Label(form, text="الكمية *", bg=UI["SURFACE"],
                 font=FONTS["body_b"]).grid(row=4, column=1, sticky="e", padx=8, pady=6)
        self.qty_var = tk.StringVar(value="1")
        qty_values = list(dict.fromkeys(self._qty_hist + COMMON_QUANTITIES))
        ttk.Combobox(
            form, textvariable=self.qty_var, values=qty_values,
            width=36, justify="right",
        ).grid(row=4, column=0, sticky="we", padx=8, pady=6)
        self.qty_var.trace_add("write", lambda *_: self._refresh_eta_preview())

        lines_card = tk.Frame(form, bg=UI["SURFACE2"], padx=10, pady=10)
        lines_card.grid(row=5, column=0, columnspan=2, sticky="we", padx=8, pady=(8, 6))
        lines_card.grid_columnconfigure(0, weight=1)
        tk.Label(
            lines_card,
            text="مقاسات/ألوان/كميات هذا الإدخال: كل سطر منها سيُنشئ أمر إنتاج مستقل في البايبلاين",
            bg=UI["SURFACE2"],
            fg=UI["TEXT_SEC"],
            font=FONTS["small"],
            anchor="e",
            justify="right",
        ).grid(row=0, column=0, columnspan=4, sticky="we", pady=(0, 8))
        self.line_status_lbl = tk.Label(
            lines_card,
            text="أضف سطرًا واحدًا على الأقل",
            bg=UI["SURFACE2"],
            fg=UI["TEXT_SEC"],
            font=FONTS["small"],
            anchor="e",
            justify="right",
        )
        self.line_status_lbl.grid(row=1, column=0, columnspan=4, sticky="we", pady=(0, 6))
        btns = tk.Frame(lines_card, bg=UI["SURFACE2"])
        btns.grid(row=2, column=0, columnspan=4, sticky="w")
        ttk.Button(btns, text="إضافة السطر", style="Ok.TButton", command=self._add_line_item).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="تعديل المحدد", command=self._edit_line_item).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="حذف المحدد", style="Danger.TButton", command=self._remove_line_item).pack(side=tk.LEFT, padx=3)
        ttk.Button(btns, text="تفريغ الحقول", style="Ghost.TButton", command=self._clear_line_editor).pack(side=tk.LEFT, padx=3)

        self.lines_tree = ttk.Treeview(
            lines_card,
            columns=("size", "color", "qty"),
            show="headings",
            height=4,
        )
        for col, title, width, anchor in [
            ("size", "المقاس", 180, "e"),
            ("color", "اللون", 180, "e"),
            ("qty", "الكمية", 100, "center"),
        ]:
            self.lines_tree.heading(col, text=title)
            self.lines_tree.column(col, width=width, anchor=anchor, stretch=True)
        self.lines_tree.grid(row=3, column=0, columnspan=4, sticky="nsew", pady=(8, 0))
        self.lines_tree.bind("<<TreeviewSelect>>", self._load_selected_line_item)

        # --- Row 6: Priority ---
        tk.Label(form, text="الأولوية (1=الأعلى)", bg=UI["SURFACE"],
                 font=FONTS["body_b"]).grid(row=6, column=1, sticky="e", padx=8, pady=6)
        self.priority_var = tk.StringVar(value="3")
        ttk.Combobox(form, textvariable=self.priority_var,
                     values=["1", "2", "3", "4", "5"],
                     state="readonly", width=34, justify="right"
                     ).grid(row=6, column=0, sticky="we", padx=8, pady=6)

        form.grid_columnconfigure(0, weight=1)

        # Stages section
        stages_frame = tk.Frame(content, bg=UI["SURFACE"], padx=16, pady=12)
        stages_frame.pack(fill=tk.BOTH, expand=True, pady=12)
        ttk.Label(stages_frame, text="مراحل الإنتاج (الافتراضية مُفعّلة، والاختيارية يمكن تبديلها)",
                  style="H3.TLabel").pack(anchor="e")

        self.stages_container = tk.Frame(stages_frame, bg=UI["SURFACE"])
        self.stages_container.pack(fill=tk.BOTH, expand=True, pady=8)

        # ETA preview
        self.eta_label = ttk.Label(stages_frame, text="", style="Dim.TLabel")
        self.eta_label.pack(anchor="e")

        # Footer (inside scrollable content so it's always reachable)
        footer = tk.Frame(content, bg=UI["BG"])
        footer.pack(fill=tk.X, pady=(0, 12))
        ttk.Button(footer, text="إضافة وبدء تلقائي", style="Ok.TButton",
                   command=self._save).pack(side=tk.LEFT, padx=4)
        ttk.Button(footer, text="إلغاء", style="Danger.TButton",
                   command=self.destroy).pack(side=tk.LEFT, padx=4)

        if self._product_rows:
            self._on_product_changed()

    def _on_product_changed(self, _evt=None) -> None:
        for w in self.stages_container.winfo_children():
            w.destroy()
        self._checks.clear()
        if not self._product_rows or not self.product_var.get():
            return
        try:
            idx = [self._product_label(p) for p in self._product_rows].index(self.product_var.get())
        except ValueError:
            return
        pid = self._product_rows[idx]["id"]
        self._stage_rows = self.db.query(
            """SELECT ps.*, st.code, st.name, st.default_setup_minutes, st.default_per_unit_minutes
               FROM product_stages ps
               JOIN stage_templates st ON st.id=ps.stage_template_id
               WHERE ps.product_id=?
               ORDER BY ps.sequence""",
            (pid,),
        )
        if not self._stage_rows:
            tk.Label(self.stages_container, text="هذا المنتج لا يحتوي على مراحل — أضِفها من شاشة المنتجات",
                     bg=UI["SURFACE"], fg=UI["DANGER"]).pack(anchor="e", pady=8)
            return

        header_row = tk.Frame(self.stages_container, bg=UI["SURFACE2"])
        header_row.pack(fill=tk.X, pady=(0, 4))
        for t, w in [("إدراج", 70), ("البدء", 70), ("#", 50),
                     ("المرحلة", 260), ("افتراضية", 90), ("اختيارية", 90), ("الوقت المقدر", 160)]:
            tk.Label(header_row, text=t, bg=UI["SURFACE2"], fg=UI["TEXT_SEC"],
                     font=FONTS["small"], width=max(6, w // 10), anchor="center").pack(side=tk.RIGHT, padx=2)

        qty = self._qty()
        for i, r in enumerate(self._stage_rows):
            row = tk.Frame(self.stages_container, bg=UI["SURFACE"])
            row.pack(fill=tk.X, pady=2)

            check_var = tk.BooleanVar(value=bool(r["is_default"]))
            cb = tk.Checkbutton(row, variable=check_var, bg=UI["SURFACE"],
                                command=self._refresh_eta_preview)
            cb.pack(side=tk.RIGHT, padx=4)
            if not r["is_optional"]:
                cb.configure(state="disabled")
            self._checks[r["id"]] = check_var

            start_rb = tk.Radiobutton(row, variable=self._start_idx_var, value=i,
                                      bg=UI["SURFACE"], command=self._refresh_eta_preview)
            start_rb.pack(side=tk.RIGHT, padx=4)

            tk.Label(row, text=str(r["sequence"]), bg=UI["SURFACE"], fg=UI["TEXT_SEC"],
                     font=FONTS["small"], width=4).pack(side=tk.RIGHT, padx=2)
            tk.Label(row, text=f"{r['name']} ({r['code']})", bg=UI["SURFACE"], fg=UI["TEXT"],
                     font=FONTS["body"], anchor="e", width=26).pack(side=tk.RIGHT, padx=6)
            tk.Label(row, text="نعم" if r["is_default"] else "لا", bg=UI["SURFACE"],
                     fg=UI["TEXT_SEC"], font=FONTS["small"], width=6).pack(side=tk.RIGHT, padx=2)
            tk.Label(row, text="نعم" if r["is_optional"] else "لا", bg=UI["SURFACE"],
                     fg=UI["TEXT_SEC"], font=FONTS["small"], width=6).pack(side=tk.RIGHT, padx=2)

            setup = r["override_setup_minutes"] if r["override_setup_minutes"] is not None else r["default_setup_minutes"]
            per_unit = r["override_per_unit_minutes"] if r["override_per_unit_minutes"] is not None else r["default_per_unit_minutes"]
            dur = float(setup) + float(per_unit) * max(1, qty)
            tk.Label(row, text=fmt_minutes(dur), bg=UI["SURFACE"], fg=UI["TEXT_SEC"],
                     font=FONTS["small"], width=14).pack(side=tk.RIGHT, padx=2)

        self._refresh_eta_preview()

    def _qty(self) -> int:
        try:
            return max(1, int(self.qty_var.get() or 1))
        except ValueError:
            return 1

    def _line_item_from_editor(self) -> Optional[Dict[str, Any]]:
        size = self.size_var.get().strip()
        color = self.color_var.get().strip()
        qty_text = self.qty_var.get().strip()
        if not (size and color and qty_text):
            return None
        try:
            qty = max(1, int(qty_text))
        except ValueError:
            return None
        return {"size": size, "color": color, "qty": qty}

    def _clear_line_editor(self) -> None:
        self.size_var.set("")
        self.color_var.set("")
        self.qty_var.set("1")
        self._editing_line_index = None
        try:
            self.lines_tree.selection_remove(self.lines_tree.selection())
        except Exception:
            pass
        self._refresh_line_items_view()

    def _refresh_line_items_view(self) -> None:
        for item in self.lines_tree.get_children():
            self.lines_tree.delete(item)
        for idx, line in enumerate(self._line_items):
            self.lines_tree.insert("", "end", iid=str(idx), values=(line["size"], line["color"], line["qty"]))
        if self._editing_line_index is None:
            self.line_status_lbl.config(
                text=f"عدد السطور المضافة: {len(self._line_items)}",
                fg=UI["TEXT_SEC"],
            )
        else:
            self.line_status_lbl.config(
                text=f"تعديل السطر رقم {self._editing_line_index + 1}",
                fg=UI["ACCENT"],
            )
        self._refresh_eta_preview()

    def _add_line_item(self) -> None:
        line = self._line_item_from_editor()
        if not line:
            messagebox.showerror(
                "بيانات السطر ناقصة",
                "أدخل المقاس واللون والكمية الرقمية قبل إضافة السطر.",
                parent=self,
            )
            return
        if self._editing_line_index is None:
            self._line_items.append(line)
        else:
            self._line_items[self._editing_line_index] = line
        self._clear_line_editor()

    def _selected_line_index(self) -> Optional[int]:
        sel = self.lines_tree.selection()
        if not sel:
            return None
        try:
            idx = int(sel[0])
        except Exception:
            return None
        if 0 <= idx < len(self._line_items):
            return idx
        return None

    def _load_selected_line_item(self, _evt=None) -> None:
        idx = self._selected_line_index()
        if idx is None:
            return
        line = self._line_items[idx]
        self.size_var.set(str(line["size"]))
        self.color_var.set(str(line["color"]))
        self.qty_var.set(str(line["qty"]))
        self._editing_line_index = idx
        self._refresh_line_items_view()

    def _edit_line_item(self) -> None:
        idx = self._selected_line_index()
        if idx is None:
            messagebox.showinfo("اختيار مطلوب", "اختر سطرًا من القائمة أولًا.", parent=self)
            return
        line = self._line_items[idx]
        self.size_var.set(str(line["size"]))
        self.color_var.set(str(line["color"]))
        self.qty_var.set(str(line["qty"]))
        self._editing_line_index = idx
        self._refresh_line_items_view()

    def _remove_line_item(self) -> None:
        idx = self._selected_line_index()
        if idx is None:
            messagebox.showinfo("اختيار مطلوب", "اختر سطرًا من القائمة أولًا.", parent=self)
            return
        del self._line_items[idx]
        self._editing_line_index = None
        self._refresh_line_items_view()

    def _collected_stages(self) -> List[Tuple[int, sqlite3.Row]]:
        """Return [(sequence_number, stage_row), ...] with selection respected."""
        qty = self._qty()  # noqa: F841 (computed for preview only)
        start = self._start_idx_var.get()
        out: List[Tuple[int, sqlite3.Row]] = []
        for i, r in enumerate(self._stage_rows):
            if i < start:
                continue
            selected = self._checks.get(r["id"])
            is_sel = True if selected is None else bool(selected.get())
            if not is_sel:
                continue
            out.append((len(out) + 1, r))
        return out

    def _refresh_eta_preview(self) -> None:
        lines = list(self._line_items)
        if not lines:
            current = self._line_item_from_editor()
            if current:
                lines = [current]
        total = 0.0
        for line in lines:
            qty = int(line["qty"])
            for _seq, r in self._collected_stages():
                setup = r["override_setup_minutes"] if r["override_setup_minutes"] is not None else r["default_setup_minutes"]
                per_unit = r["override_per_unit_minutes"] if r["override_per_unit_minutes"] is not None else r["default_per_unit_minutes"]
                total += float(setup) + float(per_unit) * max(1, qty)
        hours = ShopHours.from_db(self.db)
        eta = add_working_minutes(datetime.now(), total, hours)
        self.eta_label.config(
            text=f"عدد الأوامر الناتجة: {len(lines)}  ·  إجمالي وقت العمل: {fmt_minutes(total)}  ·  "
                 f"ETA تقديري: {eta.strftime('%Y-%m-%d %H:%M')}"
        )

    def _save(self) -> None:
        if not self._product_rows:
            messagebox.showerror("خطأ", "لا يوجد منتجات. أضِف منتجًا أولًا.", parent=self)
            return

        school = self.school_var.get().strip()
        if not school:
            messagebox.showerror(
                "حقول ناقصة",
                "الرجاء إدخال اسم المدرسة.",
                parent=self,
            )
            return
        line_items = list(self._line_items)
        if not line_items:
            current_line = self._line_item_from_editor()
            if current_line:
                line_items = [current_line]
        if not line_items:
            messagebox.showerror(
                "لا توجد مقاسات/كميات",
                "أضف سطرًا واحدًا على الأقل للمقاس/اللون/الكمية.",
                parent=self,
            )
            return
        stages = self._collected_stages()
        if not stages:
            messagebox.showerror("خطأ", "اختر مرحلة واحدة على الأقل", parent=self)
            return
        try:
            idx = [self._product_label(p) for p in self._product_rows].index(self.product_var.get())
        except ValueError:
            messagebox.showerror("خطأ", "اختر نوع المنتج", parent=self)
            return
        product = self._product_rows[idx]

        try:
            priority = int(self.priority_var.get())
        except ValueError:
            priority = 3
        now_iso = datetime.now().isoformat(timespec="minutes")
        status = "released"
        due_iso = None
        notes = ""
        created_numbers: List[str] = []
        for line in line_items:
            qty = int(line["qty"])
            order_number = self.db.next_order_number()
            self.db.execute(
                """INSERT INTO production_orders
                   (order_number, product_id, quantity, priority, due_at, status,
                    school_name, size, color, notes, created_at, released_at)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                (order_number, product["id"], qty, priority, due_iso, status,
                 school, line["size"], line["color"], notes,
                 now_iso, now_iso),
            )
            order_id = self.db.query_one("SELECT id FROM production_orders WHERE order_number=?", (order_number,))["id"]
            self.db.execute(
                "UPDATE production_orders SET root_order_id=?, original_quantity=? WHERE id=?",
                (order_id, qty, order_id),
            )

            for seq, r in stages:
                setup = r["override_setup_minutes"] if r["override_setup_minutes"] is not None else r["default_setup_minutes"]
                per_unit = r["override_per_unit_minutes"] if r["override_per_unit_minutes"] is not None else r["default_per_unit_minutes"]
                stage_tpl = self.db.query_one("SELECT name FROM stage_templates WHERE id=?",
                                              (r["stage_template_id"],))
                self.db.execute(
                    """INSERT INTO order_stages
                       (order_id, stage_template_id, sequence, name_snapshot,
                        setup_minutes, per_unit_minutes,
                        is_optional_selected, status)
                       VALUES(?,?,?,?,?,?,1,'planned')""",
                    (order_id, r["stage_template_id"], seq, stage_tpl["name"] if stage_tpl else r["name"],
                     setup, per_unit),
                )

            self.db.log_event(
                order_id, "ORDER_CREATED",
                actor="user",
                payload={
                    "order_number": order_number, "qty": qty,
                    "school": school, "size": line["size"], "color": line["color"],
                    "batch_lines": len(line_items),
                },
            )
            started_now = sync_order_state(self.db, order_id, ensure_started=True)
            if started_now:
                self.db.log_event(order_id, "ORDER_STARTED", actor="system")
            else:
                log.info("order %s created and queued; no worker available immediately", order_id)
            created_numbers.append(order_number)

        # Persist combobox history even after this order is later deleted or completed.
        self.db.record_wizard_suggestion("school_name", school)
        for line in line_items:
            self.db.record_wizard_suggestion("size", str(line["size"]))
            self.db.record_wizard_suggestion("color", str(line["color"]))
            self.db.record_wizard_suggestion("quantity", str(line["qty"]))

        # Close the modal first so the app does not feel stuck while refreshing.
        self.destroy()
        if len(created_numbers) == 1:
            toast(self.app, f"تم إنشاء الأمر {created_numbers[0]}", kind="success")
        else:
            toast(self.app, f"تم إنشاء {len(created_numbers)} أوامر إنتاج", kind="success")
        self.app.after_idle(self.app.notify_order_changed)


# =============================================================================
# Order Detail
# =============================================================================


class OrderDetail(tk.Toplevel):
    def __init__(self, parent: tk.Misc, app: "FactoryApp", order_id: int) -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.order_id = order_id
        self.title("تفاصيل الأمر")
        self.configure(bg=UI["BG"])
        self.transient(parent.winfo_toplevel())
        self.geometry("1020x700")
        self._build()
        self.refresh()

    def _build(self) -> None:
        head = tk.Frame(self, bg=UI["BG"])
        head.pack(fill=tk.X, padx=12, pady=8)
        self.title_lbl = ttk.Label(head, text="", style="H1.TLabel")
        self.title_lbl.pack(side=tk.RIGHT)
        self.meta_lbl = ttk.Label(head, text="", style="Dim.TLabel")
        self.meta_lbl.pack(side=tk.RIGHT, padx=12)

        info = tk.Frame(self, bg=UI["SURFACE"])
        info.pack(fill=tk.X, padx=12, pady=6)
        self.status_lbl = tk.Label(info, text="", bg=UI["SURFACE"], font=FONTS["h2"])
        self.status_lbl.pack(side=tk.RIGHT, padx=12, pady=8)
        self.eta_lbl = tk.Label(info, text="", bg=UI["SURFACE"], font=FONTS["body"])
        self.eta_lbl.pack(side=tk.RIGHT, padx=12, pady=8)
        self.lots_lbl = tk.Label(
            info,
            text="",
            bg=UI["SURFACE"],
            fg=UI["TEXT_SEC"],
            font=FONTS["small"],
            justify="right",
        )
        self.lots_lbl.pack(side=tk.RIGHT, padx=12, pady=8)

        body = tk.Frame(self, bg=UI["BG"])
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        # Stages tree
        left = tk.Frame(body, bg=UI["BG"])
        left.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 8))
        ttk.Label(left, text="المراحل", style="H3.TLabel").pack(anchor="e")
        cols = ("id", "seq", "name", "status", "worker", "planned", "actual")
        self.stages_tree = ttk.Treeview(left, columns=cols, show="headings", height=14)
        for c, t, w, a in [
            ("id", "#", 50, "center"),
            ("seq", "الترتيب", 70, "center"),
            ("name", "المرحلة", 180, "e"),
            ("status", "الحالة", 100, "center"),
            ("worker", "العامل", 160, "e"),
            ("planned", "المخطط (بدء → نهاية)", 230, "e"),
            ("actual", "الفعلي (د)", 100, "center"),
        ]:
            self.stages_tree.heading(c, text=t)
            self.stages_tree.column(c, width=w, anchor=a)
        self.stages_tree.pack(fill=tk.BOTH, expand=True, pady=(4, 6))
        self.stages_tree.tag_configure("running",  background=UI["OK_L"])
        self.stages_tree.tag_configure("paused",   background=UI["WARN_L"])
        self.stages_tree.tag_configure("blocked",  background=UI["DANGER_L"])
        self.stages_tree.tag_configure("done",     background=UI["SURFACE2"])
        self.stages_tree.tag_configure("skipped",  background=UI["SURFACE2"])

        # Actions
        actions = tk.Frame(left, bg=UI["BG"])
        actions.pack(fill=tk.X, pady=4, anchor="e")
        ttk.Button(actions, text="بدء", style="Ok.TButton", command=self._start_stage).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="إنهاء", style="Ok.TButton", command=self._complete_stage).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="إيقاف مؤقت", style="Warn.TButton", command=self._pause_stage).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="استئناف", command=self._resume_stage).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="تعيين/استبدال عامل", command=self._assign_worker).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="تخطي (مدير)", style="Danger.TButton", command=self._skip_stage).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="إدراج مرحلة (مدير)", style="Warn.TButton", command=self._insert_stage).pack(side=tk.LEFT, padx=3)
        ttk.Button(actions, text="تشغيل التالي بالتوازي (مدير)", command=self._early_release).pack(side=tk.LEFT, padx=3)

        order_actions = tk.Frame(left, bg=UI["BG"])
        order_actions.pack(fill=tk.X, pady=2, anchor="e")
        ttk.Button(order_actions, text="إلغاء الأمر", style="Danger.TButton", command=self._cancel_order).pack(side=tk.LEFT, padx=3)
        ttk.Button(order_actions, text="إرجاع كمية/تقسيم", style="Warn.TButton", command=self._split_rework_lot).pack(side=tk.LEFT, padx=3)

        # Events log
        right = tk.Frame(body, bg=UI["BG"])
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        ttk.Label(right, text="سجل الأحداث", style="H3.TLabel").pack(anchor="e")
        cols2 = ("ts", "type", "stage", "actor", "reason")
        self.events_tree = ttk.Treeview(right, columns=cols2, show="headings", height=14)
        for c, t, w, a in [("ts", "الوقت", 140, "center"), ("type", "الحدث", 180, "e"),
                           ("stage", "المرحلة", 140, "e"),
                           ("actor", "بواسطة", 90, "center"), ("reason", "التفاصيل", 260, "e")]:
            self.events_tree.heading(c, text=t)
            self.events_tree.column(c, width=w, anchor=a)
        self.events_tree.pack(fill=tk.BOTH, expand=True, pady=(4, 6))

    # ------ refresh ------ #

    def refresh(self) -> None:
        order = self.db.query_one(
            """SELECT po.*, p.code AS p_code, p.name AS p_name,
                      root.order_number AS root_order_number,
                      parent.order_number AS parent_order_number
               FROM production_orders po
               JOIN products p ON p.id = po.product_id
               LEFT JOIN production_orders root ON root.id = po.root_order_id
               LEFT JOIN production_orders parent ON parent.id = po.parent_order_id
               WHERE po.id=?""",
            (self.order_id,),
        )
        if not order:
            self.destroy()
            return

        # Build a descriptive title from the order's own fields
        desc_parts = [
            order["p_name"],
            order["school_name"] or "",
            order["size"] or "",
            order["color"] or "",
        ]
        desc = " · ".join(p for p in desc_parts if p)
        lot_hint = ""
        if order["parent_order_id"]:
            lot_hint = f" (إعادة من {order['parent_order_number'] or 'دفعة سابقة'})"
        self.title_lbl.config(text=f"{order['order_number']}{lot_hint} — {desc}")
        qty_text = f"الكمية: {order['quantity']}"
        if order["original_quantity"] and int(order["original_quantity"]) != int(order["quantity"]):
            qty_text += f" من أصل {order['original_quantity']}"
        self.meta_lbl.config(
            text=(
                f"الكود: {order['p_code']}  ·  {qty_text}  ·  "
                f"الأولوية: {order['priority']}  ·  استحقاق: {fmt_ts(order['due_at'])}"
            )
        )
        bg, fg = STATUS_COLOR.get(order["status"], (UI["SURFACE2"], UI["TEXT_SEC"]))
        self.status_lbl.config(text=f"الحالة: {status_label(order['status'])}", bg=bg, fg=fg)
        eta = find_order_eta(self.db, order["id"])
        self.eta_lbl.config(text=f"ETA: {eta.strftime('%Y-%m-%d %H:%M') if eta else '—'}")
        root_id = int(order["root_order_id"] or order["id"])
        family = self.db.query(
            """SELECT order_number, quantity, status
               FROM production_orders
               WHERE COALESCE(root_order_id, id)=?
               ORDER BY created_at, id""",
            (root_id,),
        )
        if len(family) > 1:
            family_bits = [
                f"{r['order_number']}: {r['quantity']} ({status_label(r['status'])})"
                for r in family[:4]
            ]
            if len(family) > 4:
                family_bits.append(f"+{len(family) - 4}")
            self.lots_lbl.config(text="الدفعات: " + " | ".join(family_bits))
        else:
            self.lots_lbl.config(text="")

        self.stages_tree.delete(*self.stages_tree.get_children())
        stages = self.db.query(
            """SELECT os.*, w.name AS worker_name
               FROM order_stages os
               LEFT JOIN workers w ON w.id = os.assigned_worker_id
               WHERE os.order_id=? ORDER BY os.sequence""",
            (self.order_id,),
        )
        for s in stages:
            status = s["status"]
            if not s["is_optional_selected"]:
                status = "skipped"
            planned = (
                f"{fmt_ts(s['planned_start'])} ← {fmt_ts(s['planned_end'])}"
                if s["planned_start"] else "—"
            )
            actual = fmt_minutes(s["actual_minutes"]) if s["actual_minutes"] else "—"
            tag = status if status in ("running", "paused", "blocked", "done", "skipped") else ""
            self.stages_tree.insert(
                "", "end",
                values=(
                    s["id"], s["sequence"], s["name_snapshot"],
                    status_label(status), s["worker_name"] or "—", planned, actual,
                ),
                tags=(tag,) if tag else (),
            )

        self.events_tree.delete(*self.events_tree.get_children())
        events = self.db.query(
            """SELECT oe.*, os.name_snapshot AS stage_name
               FROM order_events oe
               LEFT JOIN order_stages os ON os.id = oe.order_stage_id
               WHERE oe.order_id=?
               ORDER BY oe.created_at DESC, oe.id DESC""",
            (self.order_id,),
        )
        for e in events:
            self.events_tree.insert(
                "", "end",
                values=(fmt_ts(e["created_at"]), e["event_type"], e["stage_name"] or "—",
                        e["actor"] or "—", e["reason"] or ""),
            )

    def _selected_stage_id(self) -> Optional[int]:
        sel = self.stages_tree.focus()
        return int(self.stages_tree.item(sel)["values"][0]) if sel else None

    def _require_selected_stage(self, action_label: str) -> Optional[sqlite3.Row]:
        sid = self._selected_stage_id()
        if not sid:
            toast(self, f"اختر مرحلة أولًا لتنفيذ: {action_label}", kind="warn")
            return None
        s = self._stage(sid)
        if not s:
            toast(self, "المرحلة المختارة غير متاحة", kind="warn")
            return None
        return s

    def _apply_and_refresh(self) -> None:
        sync_order_state(self.db, self.order_id)
        self._maybe_auto_complete_order()
        self.refresh()
        self.app.notify_order_changed()

    def _maybe_auto_complete_order(self) -> None:
        rows = self.db.query(
            """SELECT COUNT(*) AS total,
                      SUM(CASE WHEN status IN ('done','skipped') OR is_optional_selected=0 THEN 1 ELSE 0 END) AS finished
               FROM order_stages WHERE order_id=?""",
            (self.order_id,),
        )
        r = rows[0] if rows else None
        if not r:
            return
        if int(r["total"] or 0) > 0 and int(r["finished"] or 0) == int(r["total"]):
            order = self.db.query_one("SELECT status FROM production_orders WHERE id=?", (self.order_id,))
            if order and order["status"] not in ("done", "cancelled"):
                now = datetime.now().isoformat(timespec="minutes")
                self.db.execute(
                    "UPDATE production_orders SET status='done', completed_at=? WHERE id=?",
                    (now, self.order_id),
                )
                self.db.log_event(self.order_id, "PRODUCTION_COMPLETED", actor="system")

    # ------ stage actions ------ #

    def _stage(self, sid: int) -> Optional[sqlite3.Row]:
        return self.db.query_one("SELECT * FROM order_stages WHERE id=?", (sid,))

    def _available_workers_for_stage(self, stage_template_id: int, *, exclude_stage_id: Optional[int] = None) -> List[sqlite3.Row]:
        """Active workers allowed for this stage and currently available."""
        return _available_workers_for_stage_db(
            self.db, stage_template_id, exclude_stage_id=exclude_stage_id
        )

    def _auto_assign_if_needed(self, s: sqlite3.Row) -> Optional[int]:
        """Assign a best-fit available worker to a stage if it's unassigned."""
        assigned_id: Optional[int] = None
        if s["assigned_worker_id"]:
            assigned_id = int(s["assigned_worker_id"])
            if _worker_is_available_now(self.db, assigned_id, exclude_stage_id=int(s["id"])):
                return assigned_id
        workers = self._available_workers_for_stage(int(s["stage_template_id"]), exclude_stage_id=int(s["id"]))
        workers = [w for w in workers if int(w["id"]) != (assigned_id or 0)]
        if assigned_id is not None and not workers:
            return None
        if not workers:
            return None
        wid = int(workers[0]["id"])
        self.db.execute("UPDATE order_stages SET assigned_worker_id=? WHERE id=?", (wid, int(s["id"])))
        self.db.log_event(
            self.order_id,
            "WORKER_ASSIGNED",
            order_stage_id=int(s["id"]),
            actor="system",
            payload={"worker_id": wid, "worker_name": workers[0]["name"], "auto": True},
        )
        return wid

    def _worker_block_reason(self, s: sqlite3.Row) -> str:
        if s["assigned_worker_id"]:
            busy = self.db.query_one(
                """SELECT w.name AS worker_name, po.order_number, os.name_snapshot
                   FROM order_stages os
                   JOIN production_orders po ON po.id = os.order_id
                   LEFT JOIN workers w ON w.id = os.assigned_worker_id
                   WHERE os.assigned_worker_id=?
                     AND os.id<>?
                     AND po.status NOT IN ('done','cancelled')
                     AND os.status='running'
                   LIMIT 1""",
                (int(s["assigned_worker_id"]), int(s["id"])),
            )
            if busy:
                return (
                    f"العامل {busy['worker_name'] or 'المعين'} مشغول الآن في "
                    f"{busy['order_number']} - {busy['name_snapshot']}"
                )
        return "لا يوجد عامل متاح ومخوّل لهذه المرحلة الآن"

    def _auto_start_next_stage(self) -> None:
        """After completion, start next eligible planned stage automatically."""
        auto_dispatch_order(self.db, self.order_id)

    def _set_stage_status(self, sid: int, new_status: str, event_type: str,
                          *, actor: str = "user", reason: Optional[str] = None,
                          set_actual_start: bool = False,
                          set_actual_end: bool = False) -> None:
        now = datetime.now().isoformat(timespec="minutes")
        fields: List[str] = ["status=?"]
        values: List[Any] = [new_status]
        if set_actual_start:
            fields.append("actual_start=?")
            values.append(now)
        if set_actual_end:
            fields.append("actual_end=?")
            values.append(now)
            # compute actual_minutes
            row = self._stage(sid)
            if row and row["actual_start"]:
                try:
                    start = datetime.fromisoformat(row["actual_start"])
                    mins = (datetime.now() - start).total_seconds() / 60.0
                    fields.append("actual_minutes=?")
                    values.append(max(0.0, mins))
                except Exception:
                    pass
        values.append(sid)
        self.db.execute(
            f"UPDATE order_stages SET {', '.join(fields)} WHERE id=?",
            tuple(values),
        )
        # keep production_orders.status synced
        row = self._stage(sid)
        if row:
            order_status_map = {
                "running": "running",
                "paused": "paused",
                "blocked": "blocked",
            }
            if new_status in order_status_map:
                self.db.execute(
                    "UPDATE production_orders SET status=? WHERE id=? AND status NOT IN ('done','cancelled')",
                    (order_status_map[new_status], self.order_id),
                )
            self.db.log_event(self.order_id, event_type,
                              order_stage_id=sid, actor=actor, reason=reason)

    def _start_stage(self) -> None:
        s = self._require_selected_stage("بدء")
        if not s:
            return
        sid = int(s["id"])
        if s["status"] not in ("planned", "paused"):
            toast(self, "لا يمكن بدء هذه المرحلة من حالتها الحالية", kind="warn")
            return
        # require predecessor done or early_release
        pred = self.db.query_one(
            """SELECT * FROM order_stages
               WHERE order_id=? AND sequence < ? AND is_optional_selected=1
               ORDER BY sequence DESC LIMIT 1""",
            (self.order_id, s["sequence"]),
        )
        if pred and pred["status"] != "done" and not pred["early_release"]:
            toast(self, "المرحلة السابقة لم تنتهِ بعد", kind="warn")
            return
        if not self._auto_assign_if_needed(s):
            toast(self, self._worker_block_reason(s), kind="warn")
            return
        needs_start = s["status"] == "planned" and not s["actual_start"]
        self._set_stage_status(
            sid, "running", "STAGE_STARTED" if needs_start else "STAGE_RESUMED",
            set_actual_start=needs_start,
        )
        self._apply_and_refresh()

    def _complete_stage(self) -> None:
        s = self._require_selected_stage("إنهاء")
        if not s:
            return
        sid = int(s["id"])
        if s["status"] == "done":
            toast(self, "هذه المرحلة منتهية بالفعل", kind="warn")
            return
        if s["status"] != "running":
            toast(self, "يمكن إنهاء المرحلة فقط عندما تكون قيد التنفيذ", kind="warn")
            return
        self._set_stage_status(sid, "done", "STAGE_COMPLETED", set_actual_end=True)
        self._apply_and_refresh()

    def _split_rework_lot(self) -> None:
        s = self._require_selected_stage("إرجاع كمية")
        if not s:
            return
        if not s["is_optional_selected"] or s["status"] == "skipped":
            toast(self, "لا يمكن التقسيم من مرحلة متخطاة", kind="warn")
            return
        order = self.db.query_one("SELECT * FROM production_orders WHERE id=?", (self.order_id,))
        if not order:
            return
        if order["status"] in ("done", "cancelled"):
            toast(self, "لا يمكن تقسيم أمر منتهي أو ملغي", kind="warn")
            return
        current_qty = int(order["quantity"])
        if current_qty <= 1:
            toast(self, "لا توجد كمية كافية للتقسيم", kind="warn")
            return

        return_stages = self.db.query(
            """SELECT id, sequence, name_snapshot, status
               FROM order_stages
               WHERE order_id=?
                 AND is_optional_selected=1
                 AND status<>'skipped'
                 AND sequence<=?
               ORDER BY sequence""",
            (self.order_id, int(s["sequence"])),
        )
        if not return_stages:
            toast(self, "لا توجد مرحلة صالحة للرجوع", kind="warn")
            return

        dlg = tk.Toplevel(self)
        dlg.title("إرجاع كمية/تقسيم دفعة")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        try:
            dlg.geometry(f"+{self.winfo_rootx() + 90}+{self.winfo_rooty() + 90}")
        except Exception:
            pass

        tk.Label(
            dlg,
            text=f"مرحلة المشكلة: {s['name_snapshot']} - الكمية الحالية {current_qty}",
            bg=UI["SURFACE"],
            fg=UI["TEXT"],
            font=FONTS["body_b"],
        ).grid(row=0, column=0, columnspan=2, sticky="e", padx=12, pady=(12, 8))

        tk.Label(dlg, text="كمية الإرجاع", bg=UI["SURFACE"]).grid(row=1, column=1, sticky="e", padx=8, pady=6)
        qty_var = tk.StringVar(value="1")
        qty_spin = tk.Spinbox(
            dlg,
            from_=1,
            to=current_qty - 1,
            textvariable=qty_var,
            width=12,
            justify="center",
        )
        qty_spin.grid(row=1, column=0, sticky="we", padx=8, pady=6)

        labels = [
            f"#{int(r['sequence'])} - {r['name_snapshot']} ({status_label(r['status'])})"
            for r in return_stages
        ]
        default_idx = max(0, len(labels) - 2)
        stage_var = tk.StringVar(value=labels[default_idx])
        tk.Label(dlg, text="يرجع إلى مرحلة", bg=UI["SURFACE"]).grid(row=2, column=1, sticky="e", padx=8, pady=6)
        ttk.Combobox(
            dlg,
            textvariable=stage_var,
            values=labels,
            state="readonly",
            width=34,
            justify="right",
        ).grid(row=2, column=0, sticky="we", padx=8, pady=6)

        tk.Label(dlg, text="السبب", bg=UI["SURFACE"]).grid(row=3, column=1, sticky="e", padx=8, pady=6)
        reason_var = tk.StringVar()
        ttk.Entry(dlg, textvariable=reason_var, width=36, justify="right").grid(
            row=3, column=0, sticky="we", padx=8, pady=6
        )

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.grid(row=4, column=0, columnspan=2, pady=12)

        def ok() -> None:
            try:
                qty = int(qty_var.get())
            except ValueError:
                messagebox.showerror("خطأ", "أدخل كمية صحيحة.", parent=dlg)
                return
            if qty <= 0 or qty >= current_qty:
                messagebox.showerror(
                    "خطأ",
                    f"الكمية يجب أن تكون من 1 إلى {current_qty - 1}.",
                    parent=dlg,
                )
                return
            try:
                return_idx = labels.index(stage_var.get())
            except ValueError:
                messagebox.showerror("خطأ", "اختر مرحلة الرجوع.", parent=dlg)
                return
            reason = reason_var.get().strip() or None
            try:
                child_id = split_order_lot(
                    self.db,
                    self.order_id,
                    qty,
                    int(s["id"]),
                    int(return_stages[return_idx]["id"]),
                    actor="user",
                    reason=reason,
                )
            except Exception as exc:
                messagebox.showerror("تعذر التقسيم", str(exc), parent=dlg)
                return
            child = self.db.query_one(
                "SELECT order_number FROM production_orders WHERE id=?",
                (child_id,),
            )
            dlg.destroy()
            toast(
                self,
                f"تم إنشاء دفعة إعادة: {child['order_number'] if child else child_id}",
                kind="success",
            )
            self.refresh()
            self.app.notify_order_changed()

        ttk.Button(btns, text="تقسيم", style="Ok.TButton", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton", command=dlg.destroy).pack(side=tk.LEFT, padx=6)
        qty_spin.focus_set()

    def _pause_stage(self) -> None:
        s = self._require_selected_stage("إيقاف مؤقت")
        if not s:
            return
        sid = int(s["id"])
        if s["status"] != "running":
            toast(self, "يمكن إيقاف مرحلة قيد التنفيذ فقط", kind="warn")
            return
        self._set_stage_status(sid, "paused", "STAGE_PAUSED")
        self._apply_and_refresh()

    def _resume_stage(self) -> None:
        s = self._require_selected_stage("استئناف")
        if not s:
            return
        sid = int(s["id"])
        if s["status"] != "paused":
            toast(self, "يمكن استئناف مرحلة متوقفة مؤقتًا فقط", kind="warn")
            return
        if not self._auto_assign_if_needed(s):
            toast(self, "لا يوجد عامل متاح ومخوّل لهذه المرحلة الآن", kind="warn")
            return
        self._set_stage_status(sid, "running", "STAGE_RESUMED")
        self._apply_and_refresh()

    def _assign_worker(self) -> None:
        s = self._require_selected_stage("تعيين عامل")
        if not s:
            return
        sync_order_state(self.db, self.order_id)
        self.refresh()
        s = self._stage(int(s["id"]))
        if not s:
            toast(self, "تعذر تحديث بيانات المرحلة", kind="warn")
            return
        sid = int(s["id"])
        stage_start = _parse_iso(s["actual_start"]) or _parse_iso(s["planned_start"])
        stage_end = _parse_iso(s["actual_end"]) or _parse_iso(s["planned_end"])
        if not stage_start or not stage_end or stage_end <= stage_start:
            toast(self, "لا توجد نافذة زمنية صالحة لهذه المرحلة بعد", kind="warn")
            return

        eligible_workers = self.db.query(
            """SELECT w.id, w.name
               FROM workers w
               JOIN worker_stages ws ON ws.worker_id = w.id
               WHERE ws.stage_template_id = ? AND w.is_active = 1
               ORDER BY CASE WHEN w.id = ? THEN 0 ELSE 1 END, w.name""",
            (int(s["stage_template_id"]), int(s["assigned_worker_id"] or 0)),
        )
        if not eligible_workers:
            messagebox.showerror(
                "لا يوجد عمال مؤهلون",
                "لا يوجد أي عامل مسموح له بهذه المرحلة حاليًا.",
                parent=self,
            )
            return

        dlg = tk.Toplevel(self)
        dlg.title("تعيين/استبدال عامل")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        try:
            dlg.geometry(f"+{self.winfo_rootx() + 80}+{self.winfo_rooty() + 80}")
        except Exception:
            pass

        tk.Label(
            dlg,
            text=f"العاملون المتاحون للمرحلة: {s['name_snapshot']}",
            bg=UI["SURFACE"],
            fg=UI["TEXT"],
            font=FONTS["body_b"],
        ).pack(padx=12, pady=(12, 4), anchor="e")
        tk.Label(
            dlg,
            text=f"{stage_start.strftime('%Y-%m-%d %H:%M')} ← {stage_end.strftime('%Y-%m-%d %H:%M')}",
            bg=UI["SURFACE"],
            fg=UI["TEXT_SEC"],
            font=FONTS["small"],
        ).pack(padx=12, pady=(0, 8), anchor="e")

        cols = ("worker_id", "worker", "state")
        tree = ttk.Treeview(dlg, columns=cols, show="headings", height=min(8, max(3, len(eligible_workers))))
        tree.heading("worker_id", text="#")
        tree.column("worker_id", width=50, anchor="center")
        tree.heading("worker", text="العامل")
        tree.column("worker", width=220, anchor="e")
        tree.heading("state", text="الحالة")
        tree.column("state", width=140, anchor="center")
        tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))

        available_ids: List[int] = []
        current_id = int(s["assigned_worker_id"] or 0)
        for worker in eligible_workers:
            wid = int(worker["id"])
            conflicts = self.db.query(
                """SELECT os.id, os.status, os.planned_start, os.planned_end, os.actual_start, os.actual_end
                   FROM order_stages os
                   JOIN production_orders po ON po.id = os.order_id
                   WHERE os.assigned_worker_id=?
                     AND os.id<>?
                     AND po.status NOT IN ('done','cancelled')
                     AND os.status NOT IN ('done','skipped')""",
                (wid, sid),
            )
            is_available = True
            has_running_overlap = False
            for conflict in conflicts:
                c_start = _parse_iso(conflict["actual_start"]) or _parse_iso(conflict["planned_start"])
                c_end = _parse_iso(conflict["actual_end"]) or _parse_iso(conflict["planned_end"])
                if not c_start or not c_end or c_end <= c_start:
                    continue
                if stage_start < c_end and c_start < stage_end:
                    is_available = False
                    if conflict["status"] == "running":
                        has_running_overlap = True
                    break
            if is_available or wid == current_id:
                state = "متاح"
            elif has_running_overlap:
                state = "مشغول الآن"
            else:
                state = "مشغول بخطة أخرى"
            if is_available or wid == current_id:
                available_ids.append(wid)
            item = tree.insert("", "end", values=(wid, worker["name"], state))
            if wid == current_id:
                tree.selection_set(item)
                tree.focus(item)

        if not tree.selection():
            for item in tree.get_children():
                vals = tree.item(item).get("values", [])
                if vals and vals[2] == "متاح":
                    tree.selection_set(item)
                    tree.focus(item)
                    break

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.pack(pady=12)

        def ok() -> None:
            selected = tree.focus()
            if not selected:
                messagebox.showwarning("تنبيه", "اختر عاملًا أولًا.", parent=dlg)
                return
            vals = tree.item(selected).get("values", [])
            if not vals:
                return
            selected_id = int(vals[0])
            selected_name = str(vals[1])

            running_conflicts = self.db.query(
                """SELECT os.id
                   FROM order_stages os
                   JOIN production_orders po ON po.id = os.order_id
                   WHERE os.assigned_worker_id=?
                     AND os.id<>?
                     AND po.status NOT IN ('done','cancelled')
                     AND os.status='running'""",
                (selected_id, sid),
            )
            if s["status"] == "running" and running_conflicts and selected_id != current_id:
                messagebox.showerror(
                    "العامل مشغول الآن",
                    "لا يمكن نقل عامل يعمل الآن في مرحلة جارية أخرى إلى مرحلتين قيد التنفيذ في نفس الوقت.",
                    parent=dlg,
                )
                return

            bumped_stage_ids: List[int] = []
            planned_conflicts = self.db.query(
                """SELECT os.id, os.order_id, os.name_snapshot
                   FROM order_stages os
                   JOIN production_orders po ON po.id = os.order_id
                   WHERE os.assigned_worker_id=?
                     AND os.id<>?
                     AND po.status NOT IN ('done','cancelled')
                     AND os.status='planned'""",
                (selected_id, sid),
            )
            for conflict in planned_conflicts:
                conflict_stage = self.db.query_one(
                    "SELECT planned_start, planned_end, actual_start, actual_end FROM order_stages WHERE id=?",
                    (int(conflict["id"]),),
                )
                if not conflict_stage:
                    continue
                c_start = _parse_iso(conflict_stage["actual_start"]) or _parse_iso(conflict_stage["planned_start"])
                c_end = _parse_iso(conflict_stage["actual_end"]) or _parse_iso(conflict_stage["planned_end"])
                if not c_start or not c_end or c_end <= c_start:
                    continue
                if stage_start < c_end and c_start < stage_end:
                    self.db.execute(
                        "UPDATE order_stages SET assigned_worker_id=NULL WHERE id=?",
                        (int(conflict["id"]),),
                    )
                    self.db.log_event(
                        int(conflict["order_id"]),
                        "WORKER_UNASSIGNED",
                        order_stage_id=int(conflict["id"]),
                        actor="system",
                        payload={
                            "worker_id": selected_id,
                            "worker_name": selected_name,
                            "reason": "manual_priority_reassignment",
                            "taken_by_order_id": self.order_id,
                            "taken_by_stage_id": sid,
                        },
                    )
                    bumped_stage_ids.append(int(conflict["id"]))

            self.db.execute(
                "UPDATE order_stages SET assigned_worker_id=? WHERE id=?",
                (selected_id, sid),
            )
            self.db.log_event(self.order_id, "WORKER_ASSIGNED",
                              order_stage_id=sid, actor="user",
                              payload={
                                  "worker_id": selected_id,
                                  "worker_name": selected_name,
                                  "reassigned": True,
                                  "bumped_stage_ids": bumped_stage_ids,
                              })
            dlg.destroy()
            self._apply_and_refresh()

        ttk.Button(btns, text="حفظ", style="Ok.TButton", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton", command=dlg.destroy).pack(side=tk.LEFT, padx=6)
        tree.bind("<Double-1>", lambda _e: ok())
        dlg.wait_visibility()
        tree.focus_force()

    def _skip_stage(self) -> None:
        s = self._require_selected_stage("تخطي")
        if not s:
            return
        sid = int(s["id"])
        if not require_manager(self, self.db, "تخطي مرحلة"):
            return
        self.db.execute(
            "UPDATE order_stages SET status='skipped', is_optional_selected=0 WHERE id=?",
            (sid,),
        )
        self.db.log_event(self.order_id, "STAGE_SKIPPED",
                          order_stage_id=sid, actor="manager")
        self._apply_and_refresh()

    def _insert_stage(self) -> None:
        if not require_manager(self, self.db, "إدراج مرحلة جديدة"):
            return
        stage_rows = self.db.query(
            "SELECT id,code,name,default_setup_minutes,default_per_unit_minutes "
            "FROM stage_templates WHERE is_active=1 ORDER BY name"
        )
        if not stage_rows:
            messagebox.showerror("خطأ", "لا يوجد قوالب مراحل", parent=self)
            return
        labels = [f"{s['name']} ({s['code']})" for s in stage_rows]

        dlg = tk.Toplevel(self)
        dlg.title("إدراج مرحلة")
        dlg.configure(bg=UI["SURFACE"])
        dlg.transient(self)
        dlg.grab_set()

        var = tk.StringVar(value=labels[0])
        tk.Label(dlg, text="المرحلة:", bg=UI["SURFACE"]).grid(row=0, column=1, sticky="e", padx=8, pady=6)
        ttk.Combobox(dlg, textvariable=var, values=labels, state="readonly",
                     width=30, justify="right").grid(row=0, column=0, sticky="we", padx=8, pady=6)

        # pick insertion point: before which remaining stage?
        remaining = self.db.query(
            """SELECT id, sequence, name_snapshot FROM order_stages
               WHERE order_id=? AND status NOT IN ('done','skipped') AND is_optional_selected=1
               ORDER BY sequence""",
            (self.order_id,),
        )
        rem_labels = [f"قبل #{r['sequence']} — {r['name_snapshot']}" for r in remaining] + ["في النهاية"]
        pos_var = tk.StringVar(value=rem_labels[0])
        tk.Label(dlg, text="موضع الإدراج:", bg=UI["SURFACE"]).grid(row=1, column=1, sticky="e", padx=8, pady=6)
        ttk.Combobox(dlg, textvariable=pos_var, values=rem_labels, state="readonly",
                     width=30, justify="right").grid(row=1, column=0, sticky="we", padx=8, pady=6)

        def ok() -> None:
            idx = labels.index(var.get())
            stage = stage_rows[idx]
            pos_idx = rem_labels.index(pos_var.get())
            if pos_idx < len(remaining):
                target_seq = int(remaining[pos_idx]["sequence"])
            else:
                mx = self.db.query_one(
                    "SELECT COALESCE(MAX(sequence),0) AS m FROM order_stages WHERE order_id=?",
                    (self.order_id,),
                )
                target_seq = int(mx["m"]) + 1
            # shift later sequences up by 1 (use big offset to avoid unique violation)
            self.db.execute(
                "UPDATE order_stages SET sequence = sequence + 1000 WHERE order_id=? AND sequence >= ?",
                (self.order_id, target_seq),
            )
            self.db.execute(
                "UPDATE order_stages SET sequence = sequence - 999 WHERE order_id=? AND sequence >= ?",
                (self.order_id, target_seq + 1000),
            )
            self.db.execute(
                """INSERT INTO order_stages
                   (order_id, stage_template_id, sequence, name_snapshot,
                    setup_minutes, per_unit_minutes,
                    is_optional_selected, status)
                   VALUES(?,?,?,?,?,?,1,'planned')""",
                (self.order_id, stage["id"], target_seq, stage["name"],
                 stage["default_setup_minutes"],
                 stage["default_per_unit_minutes"]),
            )
            new_id = self.db.query_one(
                "SELECT id FROM order_stages WHERE order_id=? AND sequence=?",
                (self.order_id, target_seq),
            )["id"]
            self.db.log_event(self.order_id, "STAGE_INSERTED",
                              order_stage_id=new_id, actor="manager",
                              payload={"stage": stage["name"], "sequence": target_seq})
            dlg.destroy()
            self._apply_and_refresh()

        btns = tk.Frame(dlg, bg=UI["SURFACE"])
        btns.grid(row=2, column=0, columnspan=2, pady=12)
        ttk.Button(btns, text="إدراج", style="Ok.TButton", command=ok).pack(side=tk.LEFT, padx=6)
        ttk.Button(btns, text="إلغاء", style="Ghost.TButton", command=dlg.destroy).pack(side=tk.LEFT, padx=6)

    def _early_release(self) -> None:
        s = self._require_selected_stage("إطلاق مبكر")
        if not s:
            return
        sid = int(s["id"])
        target_stage = s
        if s["status"] not in ("running", "paused"):
            pred = self.db.query_one(
                """SELECT * FROM order_stages
                   WHERE order_id=? AND sequence < ? AND is_optional_selected=1
                   ORDER BY sequence DESC LIMIT 1""",
                (self.order_id, int(s["sequence"])),
            )
            if pred and pred["status"] in ("running", "paused"):
                target_stage = pred
                sid = int(pred["id"])
            else:
                toast(self, "اختر المرحلة الجارية أو المرحلة التالية لها مباشرة", kind="warn")
                return
        if not require_manager(self, self.db, "تشغيل المرحلة التالية بالتوازي"):
            return
        if not start_next_stage_in_parallel(self.db, self.order_id, sid, actor="manager"):
            toast(self, "تعذر تشغيل المرحلة التالية بالتوازي الآن", kind="warn")
            return
        self._apply_and_refresh()

    def _cancel_order(self) -> None:
        if not require_manager(self, self.db, "إلغاء أمر الإنتاج"):
            return
        self.db.execute(
            "UPDATE production_orders SET status='cancelled', completed_at=datetime('now') WHERE id=?",
            (self.order_id,),
        )
        self.db.log_event(self.order_id, "ORDER_CANCELLED", actor="manager")
        self.refresh()
        self.app.notify_order_changed()


# =============================================================================
# Pipeline Board
# =============================================================================

PIPELINE_COLUMNS = [
    ("planned",  "مخطط",         UI["SURFACE2"], UI["TEXT_SEC"]),
    ("ready",    "جاهز",         UI["INFO_L"],   UI["INFO"]),
    ("running",  "قيد التنفيذ",  UI["OK_L"],     UI["OK"]),
    ("paused",   "متوقف",        UI["WARN_L"],   UI["WARN"]),
    ("blocked",  "محجوب",        UI["DANGER_L"], UI["DANGER"]),
    ("done",     "منتهي (اليوم)", UI["SURFACE2"], UI["TEXT_SEC"]),
]


class PipelineBoard(ttk.Frame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="لوحة الإنتاج", style="H1.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="أمر إنتاج جديد", style="Ok.TButton",
                   command=self.app.open_new_order).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="تحديث",
                   command=self._recalc_and_refresh).pack(side=tk.LEFT, padx=4)

        body = tk.Frame(self, bg=UI["BG"])
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        self.columns: Dict[str, tk.Frame] = {}
        self.counts: Dict[str, tk.Label] = {}
        for key, title, bg, fg in PIPELINE_COLUMNS:
            col = tk.Frame(body, bg=UI["SURFACE"], highlightbackground=UI["BORDER"],
                           highlightthickness=1)
            col.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=4)
            header = tk.Frame(col, bg=bg, padx=10, pady=6)
            header.pack(fill=tk.X)
            tk.Label(header, text=title, bg=bg, fg=fg, font=FONTS["h3"]).pack(side=tk.RIGHT)
            count_lbl = tk.Label(header, text="0", bg=bg, fg=fg, font=FONTS["body_b"])
            count_lbl.pack(side=tk.LEFT)
            self.counts[key] = count_lbl

            # Scrollable content
            wrap = tk.Frame(col, bg=UI["SURFACE"])
            wrap.pack(fill=tk.BOTH, expand=True)
            canvas = tk.Canvas(wrap, bg=UI["SURFACE"], highlightthickness=0, bd=0)
            canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scroll = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
            scroll.pack(side=tk.RIGHT, fill=tk.Y)
            canvas.configure(yscrollcommand=scroll.set)
            content = tk.Frame(canvas, bg=UI["SURFACE"])
            window_id = canvas.create_window((0, 0), window=content, anchor="nw")

            def _on_resize(event, canvas=canvas, window_id=window_id):
                canvas.itemconfigure(window_id, width=event.width)
            canvas.bind("<Configure>", _on_resize)
            content.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))
            self.columns[key] = content

    def _recalc_and_refresh(self) -> None:
        recompute_all(self.db)
        self.refresh()

    def refresh(self) -> None:
        for key, frame in self.columns.items():
            for w in frame.winfo_children():
                w.destroy()

        orders = self.db.query(
            """SELECT po.*, p.name AS product_name, p.code AS product_code,
                      root.order_number AS root_order_number,
                      parent.order_number AS parent_order_number
               FROM production_orders po
               JOIN products p ON p.id = po.product_id
               LEFT JOIN production_orders root ON root.id = po.root_order_id
               LEFT JOIN production_orders parent ON parent.id = po.parent_order_id
               WHERE po.status NOT IN ('cancelled')
               ORDER BY po.priority ASC, COALESCE(po.due_at,'9999'), po.created_at"""
        )
        buckets: Dict[str, List[sqlite3.Row]] = {k: [] for k, *_ in PIPELINE_COLUMNS}
        today = date.today().isoformat()
        for o in orders:
            bucket = self._bucket_for(o, today)
            buckets[bucket].append(o)

        total_orders = len(orders)
        for key, rows in buckets.items():
            self.counts[key].config(text=str(len(rows)))
            if not rows:
                tk.Label(
                    self.columns[key],
                    text="لا يوجد",
                    bg=UI["SURFACE"],
                    fg=UI["TEXT_DIM"],
                    font=FONTS["small"],
                ).pack(pady=18)
                continue
            for o in rows:
                self._render_card(self.columns[key], o)
        if total_orders == 0:
            # Show a single prominent hint across the board
            hint = tk.Frame(self.columns["planned"], bg=UI["SURFACE"])
            hint.pack(fill=tk.X, pady=8)
            tk.Label(
                hint,
                text="لا يوجد أوامر إنتاج بعد",
                bg=UI["SURFACE"],
                fg=UI["TEXT"],
                font=FONTS["h3"],
            ).pack(pady=(6, 2))
            tk.Label(
                hint,
                text="اضغط \"أمر إنتاج جديد\" للبدء",
                bg=UI["SURFACE"],
                fg=UI["TEXT_SEC"],
                font=FONTS["small"],
            ).pack()

    def _bucket_for(self, order: sqlite3.Row, today_iso: str) -> str:
        s = order["status"]
        if s == "running":
            return "running"
        if s == "paused":
            return "paused"
        if s == "blocked":
            return "blocked"
        if s == "done":
            completed = (order["completed_at"] or "")[:10]
            if completed == today_iso:
                return "done"
            return "done"  # still show for simplicity in V1
        next_stage = self.db.query_one(
            """SELECT * FROM order_stages
               WHERE order_id=? AND status='planned' AND is_optional_selected=1
               ORDER BY sequence LIMIT 1""",
            (order["id"],),
        )
        if not next_stage:
            return "planned"
        pred = self.db.query_one(
            """SELECT * FROM order_stages
               WHERE order_id=? AND sequence < ? AND is_optional_selected=1
               ORDER BY sequence DESC LIMIT 1""",
            (order["id"], next_stage["sequence"]),
        )
        if not pred or pred["status"] == "done" or pred["early_release"]:
            return "ready"
        return "planned"

    def _stage_progress(self, order_id: int) -> Tuple[int, int]:
        row = self.db.query_one(
            """SELECT COUNT(*) AS total,
                      SUM(CASE WHEN status='done' THEN 1 ELSE 0 END) AS done_cnt
               FROM order_stages
               WHERE order_id=? AND is_optional_selected=1 AND status<>'skipped'""",
            (order_id,),
        )
        total = int(row["total"] or 0) if row else 0
        done_cnt = int(row["done_cnt"] or 0) if row else 0
        return done_cnt, total

    def _render_card(self, parent: tk.Widget, order: sqlite3.Row) -> None:
        eta = find_order_eta(self.db, order["id"])
        done_cnt, total_cnt = self._stage_progress(int(order["id"]))
        running_stages = self.db.query(
            """SELECT os.*, w.name AS worker_name
               FROM order_stages os
               LEFT JOIN workers w ON w.id = os.assigned_worker_id
               WHERE os.order_id=? AND os.status='running'
               ORDER BY os.sequence""",
            (order["id"],),
        )
        current_stage = running_stages[0] if running_stages else None
        if not current_stage:
            current_stage = self.db.query_one(
                """SELECT os.*, w.name AS worker_name
                   FROM order_stages os
                   LEFT JOIN workers w ON w.id = os.assigned_worker_id
                   WHERE os.order_id=? AND os.status='planned' AND os.is_optional_selected=1
                   ORDER BY os.sequence LIMIT 1""",
                (order["id"],),
            )

        due_iso = order["due_at"]
        overdue = False
        if due_iso and eta:
            try:
                overdue = eta > datetime.fromisoformat(due_iso)
            except Exception:
                overdue = False

        card = tk.Frame(parent, bg=UI["SURFACE2"], highlightbackground=UI["BORDER"],
                        highlightthickness=1, padx=10, pady=8)
        card.pack(fill=tk.X, padx=8, pady=6)

        top = tk.Frame(card, bg=UI["SURFACE2"])
        top.pack(fill=tk.X)
        tk.Label(top, text=order["order_number"], bg=UI["SURFACE2"], fg=UI["TEXT"],
                 font=FONTS["body_b"]).pack(side=tk.RIGHT)
        prio_bg, prio_fg = self._priority_colors(int(order["priority"]))
        tk.Label(top, text=f"P{int(order['priority'])}",
                 bg=prio_bg, fg=prio_fg, font=FONTS["small"], padx=6).pack(side=tk.LEFT)
        if overdue:
            tk.Label(top, text="متأخر", bg=UI["DANGER_L"], fg=UI["DANGER"],
                     font=FONTS["small"], padx=6).pack(side=tk.LEFT, padx=4)
        if order["parent_order_id"]:
            tk.Label(top, text=f"إعادة من {order['parent_order_number'] or 'دفعة'}",
                     bg=UI["WARN_L"], fg=UI["WARN"], font=FONTS["small"],
                     padx=6).pack(side=tk.LEFT, padx=4)

        state_bg, state_fg = STATUS_COLOR.get(order["status"], (UI["SURFACE2"], UI["TEXT_SEC"]))
        tk.Label(top, text=status_label(order["status"]), bg=state_bg, fg=state_fg,
                 font=FONTS["small"], padx=6).pack(side=tk.LEFT, padx=4)

        # Line 1: product type + school
        desc_top_parts = [order["product_name"]]
        if order["school_name"]:
            desc_top_parts.append(order["school_name"])
        tk.Label(card, text=" · ".join(desc_top_parts),
                 bg=UI["SURFACE2"], fg=UI["TEXT"], font=FONTS["body_b"]).pack(
                     anchor="e", pady=(4, 0))

        # Line 2: size · color · qty
        detail_parts = []
        if order["size"]:
            detail_parts.append(f"مقاس {order['size']}")
        if order["color"]:
            detail_parts.append(order["color"])
        detail_parts.append(f"كمية {order['quantity']}")
        tk.Label(card, text=" · ".join(detail_parts),
                 bg=UI["SURFACE2"], fg=UI["TEXT_SEC"], font=FONTS["small"]).pack(
                     anchor="e", pady=(2, 0))
        if total_cnt > 0:
            progress_row = tk.Frame(card, bg=UI["SURFACE2"])
            progress_row.pack(fill=tk.X, pady=(6, 0))
            tk.Label(progress_row, text=f"{done_cnt}/{total_cnt} مراحل",
                     bg=UI["SURFACE2"], fg=UI["TEXT_SEC"], font=FONTS["small"]).pack(side=tk.LEFT)
            bar = tk.Frame(progress_row, bg=UI["BORDER"], height=8)
            bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(0, 8))
            fill = tk.Frame(bar, bg=UI["OK"] if done_cnt == total_cnt else UI["ACCENT"], height=8)
            fill.place(relx=0, rely=0, relheight=1, relwidth=(done_cnt / total_cnt) if total_cnt else 0.0)
        if running_stages:
            stage_bits = [
                f"{s['name_snapshot']} · {s['worker_name'] or 'بلا عامل'}"
                for s in running_stages
            ]
            label = "المراحل الجارية: " + " | ".join(stage_bits[:2])
            if len(running_stages) > 2:
                label += f" (+{len(running_stages) - 2})"
            tk.Label(card,
                     text=label,
                     bg=UI["SURFACE2"], fg=UI["TEXT"], font=FONTS["small"]).pack(anchor="e", pady=(2, 0))
        elif current_stage:
            tk.Label(card,
                     text=f"المرحلة: {current_stage['name_snapshot']} · {current_stage['worker_name'] or 'بلا عامل'}",
                     bg=UI["SURFACE2"], fg=UI["TEXT"], font=FONTS["small"]).pack(anchor="e", pady=(2, 0))
        tk.Label(card,
                 text=f"ETA: {eta.strftime('%Y-%m-%d %H:%M') if eta else '—'}",
                 bg=UI["SURFACE2"], fg=UI["TEXT_SEC"], font=FONTS["small"]).pack(anchor="e", pady=(2, 0))

        card.bind("<Double-Button-1>", lambda _e, oid=int(order["id"]): self.app.open_order(oid))
        for child in card.winfo_children():
            child.bind("<Double-Button-1>", lambda _e, oid=int(order["id"]): self.app.open_order(oid))

    @staticmethod
    def _priority_colors(p: int) -> Tuple[str, str]:
        return {
            1: (UI["DANGER_L"], UI["DANGER"]),
            2: (UI["WARN_L"],   UI["WARN"]),
            3: (UI["INFO_L"],   UI["INFO"]),
            4: (UI["SURFACE2"], UI["TEXT_SEC"]),
            5: (UI["SURFACE2"], UI["TEXT_DIM"]),
        }.get(p, (UI["SURFACE2"], UI["TEXT_SEC"]))


# =============================================================================
# Gantt chart (timeline view of orders × stages)
# =============================================================================


STAGE_STATUS_COLOR = {
    "planned":  ("#93c5fd", "#1e3a8a"),   # blue
    "running":  (UI["OK"],     "white"),
    "paused":   (UI["WARN"],   "white"),
    "blocked":  (UI["DANGER"], "white"),
    "done":     ("#64748b",    "white"),
    "skipped":  ("#cbd5e1",    "#475569"),
}


class GanttChart(ttk.Frame):
    """Scrollable Canvas-based Gantt: orders on Y, time on X, stages as bars."""

    ROW_HEIGHT = 46
    HEADER_HEIGHT = 44
    LABEL_WIDTH = 220
    BAR_V_PAD = 8
    BAR_LANE_HEIGHT = 30
    BAR_LANE_GAP = 4
    QUICK_DONE_W = 26
    MIN_PX_PER_HOUR = 72
    MAX_PX_PER_HOUR = 160

    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.days_span = 7
        self.px_per_hour = self.MIN_PX_PER_HOUR
        self._order_bars: List[Dict[str, Any]] = []
        self._frozen_offset_x = 0.0
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="مخطط الإنتاج الزمني (Gantt)", style="H1.TLabel").pack(side=tk.RIGHT)

        ttk.Button(top, text="أمر جديد", style="Ok.TButton",
                   command=self.app.open_new_order).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="تحديث", command=self._recalc_and_refresh).pack(side=tk.LEFT, padx=3)
        ttk.Button(top, text="اذهب للآن", command=self._scroll_to_now).pack(side=tk.LEFT, padx=3)

        sep = tk.Frame(top, bg=UI["BORDER"], width=1, height=24)
        sep.pack(side=tk.LEFT, padx=8)

        ttk.Button(top, text="−", width=3, style="Ghost.TButton",
                   command=lambda: self._zoom(0.75)).pack(side=tk.LEFT, padx=2)
        ttk.Button(top, text="+", width=3, style="Ghost.TButton",
                   command=lambda: self._zoom(1.33)).pack(side=tk.LEFT, padx=2)

        sep2 = tk.Frame(top, bg=UI["BORDER"], width=1, height=24)
        sep2.pack(side=tk.LEFT, padx=8)

        self._span_btns: Dict[int, ttk.Button] = {}
        for d in (3, 7, 14, 30):
            b = ttk.Button(top, text=f"{d} يوم", style="Ghost.TButton",
                           command=lambda dd=d: self._set_span(dd))
            b.pack(side=tk.LEFT, padx=2)
            self._span_btns[d] = b

        sep3 = tk.Frame(top, bg=UI["BORDER"], width=1, height=24)
        sep3.pack(side=tk.LEFT, padx=8)

        self._density_btns: Dict[str, ttk.Button] = {}
        for key, label, px in (
            ("half", "30 دقيقة", self.MIN_PX_PER_HOUR),
            ("detail", "تفصيلي", 120),
            ("max", "أقصى", self.MAX_PX_PER_HOUR),
        ):
            b = ttk.Button(top, text=label, style="Ghost.TButton",
                           command=lambda pp=px: self._set_density(pp))
            b.pack(side=tk.LEFT, padx=2)
            self._density_btns[key] = b

        # Legend
        legend = tk.Frame(self, bg=UI["BG"])
        legend.pack(fill=tk.X, padx=12)
        for key, label in [
            ("planned", "مخطط"),
            ("running", "قيد التنفيذ"),
            ("paused",  "متوقف"),
            ("blocked", "محجوب"),
            ("done",    "منتهي"),
        ]:
            bg, fg = STAGE_STATUS_COLOR[key]
            chip = tk.Frame(legend, bg=UI["BG"])
            chip.pack(side=tk.RIGHT, padx=6, pady=4)
            tk.Label(chip, text="  ", bg=bg, width=3).pack(side=tk.RIGHT)
            tk.Label(chip, text=label, bg=UI["BG"], fg=UI["TEXT_SEC"],
                     font=FONTS["small"]).pack(side=tk.RIGHT, padx=4)

        # Canvas with scrollbars
        body = tk.Frame(self, bg=UI["SURFACE"], highlightthickness=1,
                        highlightbackground=UI["BORDER"])
        body.pack(fill=tk.BOTH, expand=True, padx=12, pady=(4, 12))

        self.canvas = tk.Canvas(body, bg=UI["SURFACE"], highlightthickness=0, bd=0)
        self.hsb = ttk.Scrollbar(body, orient="horizontal", command=self._xview)
        self.vsb = ttk.Scrollbar(body, orient="vertical",   command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self._on_xscroll, yscrollcommand=self.vsb.set)

        self.vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas.bind("<Button-1>", self._on_click)
        self.canvas.bind("<Double-Button-1>", self._on_click)
        self.canvas.bind("<MouseWheel>",
                         lambda e: self.canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        self.canvas.bind("<Shift-MouseWheel>",
                         lambda e: self._xview("scroll", -2 if e.delta > 0 else 2, "units"))

    # ---- controls ---- #

    def _zoom(self, factor: float) -> None:
        self.px_per_hour = self._clamp_px_per_hour(self.px_per_hour * factor)
        self.refresh()

    def _set_span(self, days: int) -> None:
        self.days_span = days
        self.refresh()

    def _set_density(self, px_per_hour: int) -> None:
        self.px_per_hour = self._clamp_px_per_hour(px_per_hour)
        self.refresh()

    def _clamp_px_per_hour(self, px_per_hour: float) -> int:
        return max(self.MIN_PX_PER_HOUR, min(self.MAX_PX_PER_HOUR, int(px_per_hour)))

    def _recalc_and_refresh(self) -> None:
        recompute_all(self.db)
        self.refresh()

    def _scroll_to_now(self) -> None:
        self.canvas.update_idletasks()
        scroll = self.canvas.cget("scrollregion")
        if not scroll:
            return
        try:
            _, _, total_w, _ = map(float, scroll.split())
        except Exception:
            return
        if total_w <= 0:
            return
        # position "now" line roughly 1/4 from the left
        self.canvas.xview_moveto(max(0.0, (self._now_x - self.LABEL_WIDTH - 100) / total_w))
        self._sync_frozen_column()

    def _xview(self, *args: Any) -> None:
        self.canvas.xview(*args)
        self._sync_frozen_column()

    def _on_xscroll(self, first: str, last: str) -> None:
        self.hsb.set(first, last)
        self._sync_frozen_column()

    def _sync_frozen_column(self) -> None:
        view_x = float(self.canvas.canvasx(0))
        delta = view_x - self._frozen_offset_x
        if abs(delta) > 0.01:
            self.canvas.move("frozen", delta, 0)
            self._frozen_offset_x = view_x
        self.canvas.tag_raise("frozen")

    # ---- drawing ---- #

    def refresh(self) -> None:
        cv = self.canvas
        cv.delete("all")
        self._order_bars = []
        self._frozen_offset_x = 0.0
        for d, btn in self._span_btns.items():
            btn.state(["disabled"] if d == self.days_span else ["!disabled"])
        density_key = (
            "half" if self.px_per_hour <= self.MIN_PX_PER_HOUR else
            "max" if self.px_per_hour >= self.MAX_PX_PER_HOUR else
            "detail"
        )
        for key, btn in self._density_btns.items():
            btn.state(["disabled"] if key == density_key else ["!disabled"])

        orders = self.db.query(
            """SELECT po.*, p.name AS product_name
               FROM production_orders po
               JOIN products p ON p.id = po.product_id
               WHERE po.status NOT IN ('done','cancelled')
                 AND EXISTS (
                     SELECT 1
                       FROM order_stages os
                      WHERE os.order_id = po.id
                        AND os.is_optional_selected = 1
                        AND os.status NOT IN ('done','skipped')
                 )
               ORDER BY po.priority ASC, COALESCE(po.due_at,'9999'), po.created_at"""
        )

        now = datetime.now()
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=self.days_span)
        px_per_min = self.px_per_hour / 60.0
        total_minutes = int((end - start).total_seconds() / 60)
        timeline_width = max(200, int(total_minutes * px_per_min))
        canvas_width = self.LABEL_WIDTH + timeline_width
        order_layouts: List[Dict[str, Any]] = []
        for order in orders:
            stages = self.db.query(
                """SELECT os.*, w.name AS worker_name
                   FROM order_stages os
                   LEFT JOIN workers w ON w.id = os.assigned_worker_id
                   WHERE order_id=? AND is_optional_selected=1
                   ORDER BY sequence""",
                (order["id"],),
            )
            visible_stages: List[Dict[str, Any]] = []
            for stage in stages:
                window = self._stage_time_window(stage, start, end)
                if not window:
                    continue
                dt1, dt2, dt1c, dt2c = window
                visible_stages.append({
                    "stage": stage,
                    "dt1": dt1,
                    "dt2": dt2,
                    "dt1c": dt1c,
                    "dt2c": dt2c,
                    "lane": 0,
                })

            lane_ends: List[datetime] = []
            for item in sorted(
                visible_stages,
                key=lambda it: (it["dt1"], it["dt2"], int(it["stage"]["sequence"])),
            ):
                lane_idx = None
                for idx, lane_end in enumerate(lane_ends):
                    if item["dt1"] >= lane_end:
                        lane_idx = idx
                        break
                if lane_idx is None:
                    lane_idx = len(lane_ends)
                    lane_ends.append(item["dt2"])
                else:
                    lane_ends[lane_idx] = item["dt2"]
                item["lane"] = lane_idx

            lane_count = max(1, len(lane_ends))
            row_height = max(
                self.ROW_HEIGHT,
                (self.BAR_V_PAD * 2)
                + (lane_count * self.BAR_LANE_HEIGHT)
                + ((lane_count - 1) * self.BAR_LANE_GAP),
            )
            order_layouts.append({
                "order": order,
                "stages": visible_stages,
                "lane_count": lane_count,
                "height": row_height,
            })

        canvas_height = self.HEADER_HEIGHT + sum(int(l["height"]) for l in order_layouts) + 20
        if not order_layouts:
            canvas_height = self.HEADER_HEIGHT + self.ROW_HEIGHT + 20

        cv.configure(scrollregion=(0, 0, canvas_width, canvas_height))

        # ---- background bands: weekends and non-shop hours ---- #
        hours = ShopHours.from_db(self.db)
        for d in range(self.days_span):
            day = (start + timedelta(days=d)).date()
            x0 = self.LABEL_WIDTH + int(d * 24 * 60 * px_per_min)
            x1 = x0 + int(24 * 60 * px_per_min)
            if not hours.is_working_day(day):
                cv.create_rectangle(x0, self.HEADER_HEIGHT, x1, canvas_height,
                                    fill="#fef2f2", outline="")
            else:
                # shade non-working window
                shop_start_x = x0 + int((hours.start_h * 60 + hours.start_m) * px_per_min)
                shop_end_x = x0 + int((hours.end_h * 60 + hours.end_m) * px_per_min)
                cv.create_rectangle(x0, self.HEADER_HEIGHT, shop_start_x, canvas_height,
                                    fill="#f8fafc", outline="")
                cv.create_rectangle(shop_end_x, self.HEADER_HEIGHT, x1, canvas_height,
                                    fill="#f8fafc", outline="")

        # ---- day columns + half-hour/hour ticks ---- #
        show_time_labels = self.px_per_hour >= self.MIN_PX_PER_HOUR

        for d in range(self.days_span + 1):
            day = start + timedelta(days=d)
            x = self.LABEL_WIDTH + int(d * 24 * 60 * px_per_min)
            cv.create_line(x, self.HEADER_HEIGHT, x, canvas_height,
                           fill=UI["BORDER"], width=1)
            if d < self.days_span:
                day_w = int(24 * 60 * px_per_min)
                is_today = day.date() == now.date()
                bg = "#e0f2fe" if is_today else UI["SURFACE2"]
                cv.create_rectangle(x, 0, x + day_w, self.HEADER_HEIGHT,
                                    fill=bg, outline="")
                cv.create_text(x + 8, 8,
                               text=day.strftime("%a %Y-%m-%d"),
                               anchor="nw", fill=UI["TEXT"], font=FONTS["body_b"])
                # Vertical timeline: every 30 minutes (minor), every hour (major)
                for minute in range(0, 24 * 60 + 1, 30):
                    hx = x + int(minute * px_per_min)
                    is_hour = (minute % 60 == 0)
                    cv.create_line(
                        hx, self.HEADER_HEIGHT, hx, canvas_height,
                        fill=UI["BORDER"] if is_hour else "#edf2f7",
                        width=1,
                        dash=() if is_hour else (2, 4),
                    )
                    if show_time_labels and minute < 24 * 60:
                        h = minute // 60
                        m = minute % 60
                        cv.create_text(
                            hx + 2, self.HEADER_HEIGHT - 14,
                            text=f"{h:02d}:{m:02d}", anchor="nw",
                            fill=UI["TEXT_DIM"], font=FONTS["small"],
                        )

        # ---- now line ---- #
        self._now_x = self.LABEL_WIDTH + int((now - start).total_seconds() / 60 * px_per_min)
        if self.LABEL_WIDTH <= self._now_x <= canvas_width:
            cv.create_line(self._now_x, self.HEADER_HEIGHT, self._now_x, canvas_height,
                           fill=UI["DANGER"], width=2, dash=(4, 3))
            cv.create_text(self._now_x + 4, self.HEADER_HEIGHT + 2,
                           text="الآن", anchor="nw",
                           fill=UI["DANGER"], font=FONTS["body_b"])

        # ---- left label column ---- #
        cv.create_rectangle(0, 0, self.LABEL_WIDTH, canvas_height,
                            fill=UI["SURFACE2"], outline="", tags=("frozen",))
        cv.create_line(self.LABEL_WIDTH, 0, self.LABEL_WIDTH, canvas_height,
                       fill=UI["BORDER"], width=1, tags=("frozen",))
        cv.create_text(self.LABEL_WIDTH - 10, 10, text="الأوامر",
                       anchor="ne", fill=UI["TEXT"], font=FONTS["body_b"], tags=("frozen",))

        if not orders:
            cv.create_text(
                self.LABEL_WIDTH + timeline_width / 2,
                self.HEADER_HEIGHT + 60,
                text="لا توجد أوامر نشطة في مخطط جانت — اضغط \"أمر جديد\"",
                fill=UI["TEXT_SEC"], font=FONTS["h3"],
            )
            self._sync_frozen_column()
            return

        # ---- rows ---- #
        y_top = self.HEADER_HEIGHT
        for idx, layout in enumerate(order_layouts):
            order = layout["order"]
            row_height = int(layout["height"])
            y_mid = y_top + row_height // 2
            if order["status"] == "running":
                cv.create_rectangle(0, y_top, canvas_width, y_top + row_height,
                                    fill="#ecfdf5", outline="")
            if idx > 0:
                cv.create_line(0, y_top, canvas_width, y_top, fill=UI["BORDER"])

            # label
            cv.create_text(self.LABEL_WIDTH - 10, y_mid - 8,
                           text=order["order_number"], anchor="e",
                           fill=UI["TEXT"], font=FONTS["body_b"], tags=("frozen",))
            parts = [order["product_name"]]
            if order["school_name"]:
                parts.append(order["school_name"])
            if order["size"]:
                parts.append(f"مقاس {order['size']}")
            if order["color"]:
                parts.append(order["color"])
            parts.append(f"×{order['quantity']}")
            cv.create_text(self.LABEL_WIDTH - 10, y_mid + 8,
                           text=" · ".join(parts),
                           anchor="e", fill=UI["TEXT_SEC"], font=FONTS["small"], tags=("frozen",))

            # due-date marker
            if order["due_at"]:
                try:
                    due_dt = datetime.fromisoformat(order["due_at"])
                    if start <= due_dt <= end:
                        dx = self.LABEL_WIDTH + int((due_dt - start).total_seconds() / 60 * px_per_min)
                        cv.create_line(dx, y_top + 2, dx, y_top + row_height - 2,
                                       fill=UI["DANGER"], width=2)
                        cv.create_text(dx + 3, y_top + 2, text="⏰",
                                       anchor="nw", font=FONTS["small"])
                except Exception:
                    pass

            # stages
            for item in sorted(layout["stages"], key=lambda it: (int(it["lane"]), it["dt1"], int(it["stage"]["sequence"]))):
                lane = int(item["lane"])
                y1 = y_top + self.BAR_V_PAD + lane * (self.BAR_LANE_HEIGHT + self.BAR_LANE_GAP)
                y2 = min(y_top + row_height - self.BAR_V_PAD, y1 + self.BAR_LANE_HEIGHT)
                self._draw_stage_bar(order, item, y1, y2, start, px_per_min)
            y_top += row_height
        self._sync_frozen_column()

    def _stage_time_window(
        self,
        stage: sqlite3.Row,
        start: datetime,
        end: datetime,
    ) -> Optional[Tuple[datetime, datetime, datetime, datetime]]:
        status = stage["status"]
        actual_start = stage["actual_start"]
        actual_end = stage["actual_end"]
        planned_start = stage["planned_start"]
        planned_end = stage["planned_end"]
        planned_start_dt = _parse_iso(planned_start)
        planned_end_dt = _parse_iso(planned_end)
        actual_start_dt = _parse_iso(actual_start)
        actual_end_dt = _parse_iso(actual_end)

        dt1 = dt2 = None
        if status == "done" and actual_start_dt and actual_end_dt:
            dt1, dt2 = actual_start_dt, actual_end_dt
        elif status == "running":
            dt1 = planned_start_dt or actual_start_dt
            live_end = datetime.now()
            if planned_end_dt and planned_end_dt > live_end:
                live_end = planned_end_dt
            dt2 = live_end
        elif status in ("paused", "blocked"):
            dt1 = planned_start_dt or actual_start_dt
            dt2 = planned_end_dt or actual_end_dt or datetime.now()
        elif status == "skipped":
            return None
        elif planned_start_dt and planned_end_dt:
            dt1, dt2 = planned_start_dt, planned_end_dt

        if not dt1 or not dt2:
            return None
        if dt2 <= start or dt1 >= end:
            return None
        return dt1, dt2, max(dt1, start), min(dt2, end)

    def _draw_stage_bar(
        self,
        order: sqlite3.Row,
        item: Dict[str, Any],
        y1: float,
        y2: float,
        start: datetime,
        px_per_min: float,
    ) -> None:
        stage = item["stage"]
        status = stage["status"]
        dt1c = item["dt1c"]
        dt2c = item["dt2c"]
        x1 = self.LABEL_WIDTH + int((dt1c - start).total_seconds() / 60 * px_per_min)
        x2 = self.LABEL_WIDTH + int((dt2c - start).total_seconds() / 60 * px_per_min)
        if x2 - x1 < 3:
            x2 = x1 + 3

        bg, fg = STAGE_STATUS_COLOR.get(status, STAGE_STATUS_COLOR["planned"])
        outline = UI["OK"] if status == "running" else "white"
        width = 2 if status == "running" else 1
        self.canvas.create_rectangle(x1, y1, x2, y2, fill=bg, outline=outline, width=width)

        # label inside bar if there's space
        bar_w = x2 - x1
        if bar_w > 36:
            name = stage["name_snapshot"] or ""
            worker = stage["worker_name"] or "بدون عامل"
            text = f"{name} · {worker}"
            if bar_w < 90:
                text = name[:10]
            self.canvas.create_text(
                x1 + 6, (y1 + y2) // 2, text=text, anchor="w",
                fill=fg, font=FONTS["small"],
            )
        if status == "running":
            self.canvas.create_oval(x1 + 4, y1 + 4, x1 + 10, y1 + 10, fill="white", outline="")

        bar_info: Dict[str, Any] = {
            "x1": x1,
            "y1": y1,
            "x2": x2,
            "y2": y2,
            "order_id": int(order["id"]),
            "stage_id": int(stage["id"]),
            "status": status,
            "action": None,
        }
        if status == "running":
            btn_size = min(self.QUICK_DONE_W, max(16, y2 - y1 - 6))
            bx2 = x2 - 4
            bx1 = max(x1 + 12, bx2 - btn_size)
            by1 = y1 + ((y2 - y1 - btn_size) / 2)
            by2 = by1 + btn_size
            self.canvas.create_rectangle(
                bx1, by1, bx2, by2,
                fill="white",
                outline=UI["OK"],
                width=1,
            )
            self.canvas.create_text(
                (bx1 + bx2) / 2,
                (by1 + by2) / 2,
                text="✓",
                fill=UI["OK"],
                font=FONTS["body_b"],
            )
            bar_info["action"] = {"kind": "complete", "x1": bx1, "y1": by1, "x2": bx2, "y2": by2}
        self._order_bars.append(bar_info)

    def _on_click(self, event: tk.Event) -> None:
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        for bar in self._order_bars:
            if bar["x1"] <= x <= bar["x2"] and bar["y1"] <= y <= bar["y2"]:
                action = bar.get("action")
                if action and action["x1"] <= x <= action["x2"] and action["y1"] <= y <= action["y2"]:
                    self._quick_complete_stage(int(bar["stage_id"]))
                    return
                self.app.open_order(int(bar["order_id"]))
                return

    def _quick_complete_stage(self, stage_id: int) -> None:
        stage = self.db.query_one(
            """SELECT os.id, os.order_id, os.name_snapshot, po.order_number
               FROM order_stages os
               JOIN production_orders po ON po.id = os.order_id
               WHERE os.id=?""",
            (stage_id,),
        )
        if not stage:
            return
        if not messagebox.askyesno(
            "إنهاء المرحلة",
            f"إنهاء المرحلة «{stage['name_snapshot']}» في الأمر {stage['order_number']}؟",
            parent=self,
        ):
            return
        if not complete_stage_and_advance(self.db, stage_id, actor="user"):
            toast(self, "تعذر إنهاء المرحلة من مخطط جانت", kind="warn")
            return
        toast(self, f"تم إنهاء المرحلة في {stage['order_number']}", kind="success")
        self.app.notify_order_changed()


def _parse_iso(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


# =============================================================================
# Worker day plan
# =============================================================================


class WorkerPlanFrame(ttk.Frame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self.selected_day = datetime.now().date()
        self._summary_var = tk.StringVar(value="")
        self._day_var = tk.StringVar(value="")
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="خطة العمال اليومية", style="H1.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="اليوم", style="Ghost.TButton",
                   command=self._go_today).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="اليوم التالي", style="Ghost.TButton",
                   command=lambda: self._shift_day(1)).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="اليوم السابق", style="Ghost.TButton",
                   command=lambda: self._shift_day(-1)).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh).pack(side=tk.LEFT, padx=4)

        meta = tk.Frame(self, bg=UI["BG"])
        meta.pack(fill=tk.X, padx=12, pady=(0, 6))
        ttk.Label(meta, textvariable=self._day_var, style="H2.TLabel").pack(side=tk.RIGHT)
        ttk.Label(meta, textvariable=self._summary_var, style="Dim.TLabel").pack(side=tk.LEFT)

        hint_bar(
            self,
            "يعرض هذا التبويب كل عامل وما المراحل غير المنتهية المسندة له خلال اليوم المحدد.",
            "العامل الذي لا يملك مراحل في هذا اليوم سيظهر على أنه متاح.",
            "انقر مرتين على أي سطر لفتح أمر الإنتاج المرتبط به."
        ).pack(fill=tk.X, padx=12, pady=(0, 8))

        cols = ("order_id", "worker", "from", "to", "stage", "order_no", "product", "school", "size", "color", "status")
        self.tree = tree_with_scrollbars(self, columns=cols, height=20)
        self.tree.heading("order_id", text="#")
        self.tree.column("order_id", width=50, anchor="center")
        self.tree.heading("worker", text="العامل")
        self.tree.column("worker", width=180, anchor="e")
        self.tree.heading("from", text="من")
        self.tree.column("from", width=110, anchor="center")
        self.tree.heading("to", text="إلى")
        self.tree.column("to", width=110, anchor="center")
        self.tree.heading("stage", text="المرحلة")
        self.tree.column("stage", width=180, anchor="e")
        self.tree.heading("order_no", text="الأمر")
        self.tree.column("order_no", width=130, anchor="center")
        self.tree.heading("product", text="المنتج")
        self.tree.column("product", width=120, anchor="e")
        self.tree.heading("school", text="المدرسة")
        self.tree.column("school", width=150, anchor="e")
        self.tree.heading("size", text="المقاس")
        self.tree.column("size", width=80, anchor="center")
        self.tree.heading("color", text="اللون")
        self.tree.column("color", width=90, anchor="e")
        self.tree.heading("status", text="الحالة")
        self.tree.column("status", width=110, anchor="center")
        self.tree.pack_configure(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        self.tree.bind("<Double-1>", self._open_selected)

    def _go_today(self) -> None:
        self.selected_day = datetime.now().date()
        self.refresh()

    def _shift_day(self, delta_days: int) -> None:
        self.selected_day += timedelta(days=delta_days)
        self.refresh()

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        day_start = datetime.combine(self.selected_day, datetime.min.time())
        day_end = day_start + timedelta(days=1)
        self._day_var.set(self.selected_day.strftime("%A %Y-%m-%d"))

        workers = self.db.query(
            "SELECT id, name FROM workers WHERE is_active=1 ORDER BY name"
        )
        stage_rows = self.db.query(
            """SELECT os.order_id,
                      os.assigned_worker_id AS worker_id,
                      os.name_snapshot AS stage_name,
                      os.status AS stage_status,
                      COALESCE(os.actual_start, os.planned_start) AS stage_start,
                      COALESCE(os.actual_end, os.planned_end, os.actual_start, os.planned_start) AS stage_end,
                      po.order_number,
                      po.school_name,
                      po.size,
                      po.color,
                      p.name AS product_name
               FROM order_stages os
               JOIN production_orders po ON po.id = os.order_id
               JOIN products p ON p.id = po.product_id
               WHERE os.assigned_worker_id IS NOT NULL
                 AND po.status <> 'cancelled'
                 AND os.status NOT IN ('done','skipped')
                 AND COALESCE(os.actual_start, os.planned_start) < ?
                 AND COALESCE(os.actual_end, os.planned_end, os.actual_start, os.planned_start) >= ?
               ORDER BY os.assigned_worker_id,
                        COALESCE(os.actual_start, os.planned_start),
                        os.sequence,
                        os.id""",
            (day_end.isoformat(timespec="minutes"), day_start.isoformat(timespec="minutes")),
        )

        by_worker: Dict[int, List[sqlite3.Row]] = {}
        for row in stage_rows:
            by_worker.setdefault(int(row["worker_id"]), []).append(row)

        total_slots = 0
        busy_workers = 0
        for worker in workers:
            wid = int(worker["id"])
            planned = by_worker.get(wid, [])
            if planned:
                busy_workers += 1
                for row in planned:
                    total_slots += 1
                    self.tree.insert(
                        "",
                        "end",
                        values=(
                            row["order_id"],
                            worker["name"],
                            fmt_ts(row["stage_start"]),
                            fmt_ts(row["stage_end"]),
                            row["stage_name"],
                            row["order_number"],
                            row["product_name"],
                            row["school_name"] or "—",
                            row["size"] or "—",
                            row["color"] or "—",
                            status_label(row["stage_status"]),
                        ),
                    )
            else:
                self.tree.insert(
                    "",
                    "end",
                    values=("—", worker["name"], "—", "—", "متاح", "—", "—", "—", "—", "—", "بدون مراحل"),
                )

        self._summary_var.set(
            f"{len(workers)} عامل • {busy_workers} لديهم مراحل • {len(workers) - busy_workers} متاح • {total_slots} مرحلة في اليوم"
        )

    def _open_selected(self, _evt=None) -> None:
        sel = self.tree.focus()
        if not sel:
            return
        values = self.tree.item(sel).get("values", [])
        if not values or values[0] in ("", "—", None):
            return
        self.app.open_order(int(values[0]))


# =============================================================================
# Orders list (simple list view as an alternative to the board)
# =============================================================================


class OrdersList(ttk.Frame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="أوامر الإنتاج", style="H1.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="أمر جديد", style="Ok.TButton",
                   command=self.app.open_new_order).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="حذف الأمر بالكامل", style="Danger.TButton",
                   command=self._hard_delete_order).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="حذف سجل الأمر", style="Warn.TButton",
                   command=self._clear_order_history).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh).pack(side=tk.LEFT, padx=4)

        cols = ("id", "num", "product", "school", "size", "color",
                "qty", "priority", "status", "due", "eta", "created")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=22)
        headings = [
            ("id",       "#",         40,  "center"),
            ("num",      "الرقم",     130, "center"),
            ("product",  "النوع",     110, "e"),
            ("school",   "المدرسة",   150, "e"),
            ("size",     "المقاس",    70,  "center"),
            ("color",    "اللون",     80,  "e"),
            ("qty",      "الكمية",    60,  "center"),
            ("priority", "أولوية",    60,  "center"),
            ("status",   "الحالة",    100, "center"),
            ("due",      "الاستحقاق", 120, "center"),
            ("eta",      "ETA",       120, "center"),
            ("created",  "أُنشئ",     120, "center"),
        ]
        for c, t, w, a in headings:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor=a)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))
        self.tree.bind("<Double-1>", self._open_selected)

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        rows = self.db.query(
            """SELECT po.*, p.name AS product_name,
                      parent.order_number AS parent_order_number
               FROM production_orders po
               JOIN products p ON p.id = po.product_id
               LEFT JOIN production_orders parent ON parent.id = po.parent_order_id
               ORDER BY po.priority ASC, COALESCE(po.due_at,'9999'), po.created_at DESC"""
        )
        for o in rows:
            eta = find_order_eta(self.db, o["id"])
            number_text = o["order_number"]
            if o["parent_order_id"]:
                number_text = f"{number_text} / من {o['parent_order_number'] or 'دفعة'}"
            self.tree.insert(
                "", "end",
                values=(
                    o["id"], number_text, o["product_name"],
                    o["school_name"] or "—", o["size"] or "—", o["color"] or "—",
                    o["quantity"], o["priority"],
                    status_label(o["status"]),
                    fmt_ts(o["due_at"]),
                    eta.strftime("%Y-%m-%d %H:%M") if eta else "—",
                    fmt_ts(o["created_at"]),
                ),
            )

    def _open_selected(self, _evt=None) -> None:
        sel = self.tree.focus()
        if not sel:
            return
        oid = int(self.tree.item(sel)["values"][0])
        self.app.open_order(oid)

    def _clear_order_history(self) -> None:
        sel = self.tree.focus()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر أمرًا من الجدول أولًا.", parent=self)
            return
        vals = self.tree.item(sel).get("values", [])
        if not vals:
            return
        oid = int(vals[0])
        order_num = vals[1]
        cnt_row = self.db.query_one(
            "SELECT COUNT(*) AS c FROM order_events WHERE order_id=?",
            (oid,),
        )
        cnt = int(cnt_row["c"]) if cnt_row else 0
        if cnt == 0:
            messagebox.showinfo("معلومة", "لا يوجد سجل أحداث لهذا الأمر.", parent=self)
            return
        if not messagebox.askyesno(
            "تأكيد حذف السجل",
            f"سيتم حذف {cnt} حدث من سجل الأمر {order_num}.\n"
            "لن يتم حذف الأمر نفسه أو مراحله.\n\nمتأكد؟",
            parent=self,
        ):
            return
        self.db.execute("DELETE FROM order_events WHERE order_id=?", (oid,))
        toast(self.winfo_toplevel(), f"تم حذف سجل الأمر {order_num}", kind="success")
        self.app.notify_order_changed()

    def _hard_delete_order(self) -> None:
        sel = self.tree.focus()
        if not sel:
            messagebox.showwarning("تنبيه", "اختر أمرًا من الجدول أولًا.", parent=self)
            return
        vals = self.tree.item(sel).get("values", [])
        if not vals:
            return
        oid = int(vals[0])
        order_num = str(vals[1])
        status_txt = str(vals[8]) if len(vals) > 8 else ""

        stages_cnt = self.db.query_one(
            "SELECT COUNT(*) AS c FROM order_stages WHERE order_id=?",
            (oid,),
        )
        events_cnt = self.db.query_one(
            "SELECT COUNT(*) AS c FROM order_events WHERE order_id=?",
            (oid,),
        )
        c_stages = int(stages_cnt["c"]) if stages_cnt else 0
        c_events = int(events_cnt["c"]) if events_cnt else 0

        # Manager-style confirmation: clear warning + explicit typed confirmation.
        if not messagebox.askyesno(
            "تحذير حذف نهائي",
            f"سيتم حذف الأمر {order_num} نهائيًا (Hard Delete).\n"
            f"الحالة الحالية: {status_txt}\n\n"
            f"سيتم حذف:\n"
            f"• بيانات الأمر\n"
            f"• {c_stages} مرحلة مرتبطة\n"
            f"• {c_events} سجل أحداث\n\n"
            "هذا الإجراء غير قابل للتراجع.\n\nمتأكد أنك تريد المتابعة؟",
            parent=self,
        ):
            return

        self.db.execute("DELETE FROM production_orders WHERE id=?", (oid,))
        toast(self.winfo_toplevel(), f"تم حذف الأمر {order_num} نهائيًا", kind="success")
        self.app.notify_order_changed()


# =============================================================================
# Reports
# =============================================================================


class ReportsFrame(ttk.Frame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="التقارير", style="H1.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="تحديث", command=self.refresh).pack(side=tk.LEFT, padx=4)

        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        self.bottleneck_tree = self._make_tree(
            nb, ["المرحلة", "عدد المراحل", "متوسط الانتظار (د)", "قيد التنفيذ"]
        )
        nb.add(self.bottleneck_tree.master, text="اختناقات")

        self.planned_vs_tree = self._make_tree(
            nb, ["المرحلة", "متوسط المخطط (د)", "متوسط الفعلي (د)", "عدد المنفذة"]
        )
        nb.add(self.planned_vs_tree.master, text="مخطط مقابل فعلي")

        self.workers_tree = self._make_tree(
            nb, ["العامل", "مراحل منتهية", "إجمالي دقائق فعلية", "متوسط الدقائق"]
        )
        nb.add(self.workers_tree.master, text="إنتاجية العمال")

        self.overdue_tree = self._make_tree(
            nb, ["الأمر", "المنتج", "الاستحقاق", "ETA", "الحالة"]
        )
        nb.add(self.overdue_tree.master, text="متأخرة")

        self.throughput_tree = self._make_tree(
            nb, ["المنتج", "أوامر منتهية", "إجمالي الكميات"]
        )
        nb.add(self.throughput_tree.master, text="الإنتاجية حسب المنتج")

    @staticmethod
    def _make_tree(parent: tk.Misc, columns: List[str]) -> ttk.Treeview:
        wrap = tk.Frame(parent, bg=UI["BG"])
        tree = ttk.Treeview(wrap, columns=list(range(len(columns))), show="headings", height=18)
        for i, c in enumerate(columns):
            tree.heading(i, text=c)
            tree.column(i, width=180, anchor="e" if i == 0 else "center")
        tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        return tree

    def refresh(self) -> None:
        self._fill_bottleneck()
        self._fill_planned_vs_actual()
        self._fill_workers()
        self._fill_overdue()
        self._fill_throughput()

    def _fill_bottleneck(self) -> None:
        t = self.bottleneck_tree
        t.delete(*t.get_children())
        rows = self.db.query(
            """SELECT st.name AS stage,
                      COUNT(os.id) AS cnt,
                      SUM(CASE WHEN os.status='running' THEN 1 ELSE 0 END) AS running,
                      AVG(CASE WHEN os.actual_start IS NOT NULL AND os.planned_start IS NOT NULL
                               THEN (julianday(os.actual_start) - julianday(os.planned_start)) * 24 * 60
                          END) AS wait_min
               FROM order_stages os
               JOIN stage_templates st ON st.id = os.stage_template_id
               GROUP BY st.name
               ORDER BY running DESC, cnt DESC"""
        )
        for r in rows:
            t.insert("", "end", values=(
                r["stage"], r["cnt"], f"{r['wait_min'] or 0:.0f}", r["running"] or 0,
            ))

    def _fill_planned_vs_actual(self) -> None:
        t = self.planned_vs_tree
        t.delete(*t.get_children())
        rows = self.db.query(
            """SELECT st.name AS stage,
                      AVG(CASE WHEN os.planned_start IS NOT NULL AND os.planned_end IS NOT NULL
                               THEN (julianday(os.planned_end) - julianday(os.planned_start)) * 24 * 60
                          END) AS planned_avg,
                      AVG(os.actual_minutes) AS actual_avg,
                      SUM(CASE WHEN os.status='done' THEN 1 ELSE 0 END) AS done_cnt
               FROM order_stages os
               JOIN stage_templates st ON st.id = os.stage_template_id
               GROUP BY st.name
               ORDER BY done_cnt DESC"""
        )
        for r in rows:
            t.insert("", "end", values=(
                r["stage"],
                f"{r['planned_avg'] or 0:.1f}",
                f"{r['actual_avg'] or 0:.1f}",
                r["done_cnt"] or 0,
            ))

    def _fill_workers(self) -> None:
        t = self.workers_tree
        t.delete(*t.get_children())
        rows = self.db.query(
            """SELECT w.name AS worker,
                      COUNT(CASE WHEN os.status='done' THEN 1 END) AS done,
                      COALESCE(SUM(os.actual_minutes),0) AS total_min,
                      COALESCE(AVG(os.actual_minutes),0) AS avg_min
               FROM workers w
               LEFT JOIN order_stages os ON os.assigned_worker_id = w.id
               WHERE w.is_active = 1
               GROUP BY w.id, w.name
               ORDER BY done DESC, w.name"""
        )
        for r in rows:
            t.insert("", "end", values=(
                r["worker"], r["done"], f"{r['total_min']:.0f}", f"{r['avg_min']:.0f}",
            ))

    def _fill_overdue(self) -> None:
        t = self.overdue_tree
        t.delete(*t.get_children())
        orders = self.db.query(
            """SELECT po.id, po.order_number, po.due_at, po.status,
                      p.name AS product_name
               FROM production_orders po
               JOIN products p ON p.id = po.product_id
               WHERE po.status NOT IN ('done','cancelled') AND po.due_at IS NOT NULL
               ORDER BY po.due_at"""
        )
        now = datetime.now()
        for o in orders:
            try:
                due = datetime.fromisoformat(o["due_at"])
            except Exception:
                continue
            eta = find_order_eta(self.db, o["id"])
            is_late = (eta and eta > due) or (due < now)
            if is_late:
                t.insert("", "end", values=(
                    o["order_number"], o["product_name"],
                    due.strftime("%Y-%m-%d %H:%M"),
                    eta.strftime("%Y-%m-%d %H:%M") if eta else "—",
                    status_label(o["status"]),
                ))

    def _fill_throughput(self) -> None:
        t = self.throughput_tree
        t.delete(*t.get_children())
        rows = self.db.query(
            """SELECT p.name AS product,
                      COUNT(CASE WHEN po.status='done' THEN 1 END) AS done_cnt,
                      COALESCE(SUM(CASE WHEN po.status='done' THEN po.quantity ELSE 0 END),0) AS total_qty
               FROM products p
               LEFT JOIN production_orders po ON po.product_id = p.id
               GROUP BY p.id, p.name
               ORDER BY done_cnt DESC, p.name"""
        )
        for r in rows:
            t.insert("", "end", values=(r["product"], r["done_cnt"], r["total_qty"]))


# =============================================================================
# Dashboard
# =============================================================================


class DashboardFrame(ttk.Frame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self._labels: Dict[str, tk.Label] = {}
        self._build()
        self.refresh()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="اللوحة الرئيسية", style="H1.TLabel").pack(side=tk.RIGHT)
        ttk.Button(top, text="أمر جديد", style="Ok.TButton",
                   command=self.app.open_new_order).pack(side=tk.LEFT, padx=4)
        ttk.Button(top, text="تحديث", command=self.refresh).pack(side=tk.LEFT, padx=4)

        cards = tk.Frame(self, bg=UI["BG"])
        cards.pack(fill=tk.X, padx=12, pady=8)

        self._kpi(cards, "planned",  "مخطط",        UI["SURFACE2"], UI["TEXT_SEC"])
        self._kpi(cards, "running",  "قيد التنفيذ", UI["OK_L"],     UI["OK"])
        self._kpi(cards, "paused",   "متوقف",       UI["WARN_L"],   UI["WARN"])
        self._kpi(cards, "blocked",  "محجوب",       UI["DANGER_L"], UI["DANGER"])
        self._kpi(cards, "done_tod", "منتهي اليوم", UI["INFO_L"],   UI["INFO"])
        self._kpi(cards, "overdue",  "متأخر",       UI["DANGER_L"], UI["DANGER"])

        info = tk.Frame(self, bg=UI["BG"])
        info.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        left = tk.Frame(info, bg=UI["SURFACE"], padx=12, pady=10)
        left.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(0, 6))
        ttk.Label(left, text="قادم اليوم", style="H3.TLabel").pack(anchor="e")
        self.upcoming_tree = ttk.Treeview(
            left, columns=("num", "product", "stage", "worker", "eta"), show="headings", height=10
        )
        for c, t, w, a in [("num", "الأمر", 130, "center"), ("product", "المنتج", 200, "e"),
                           ("stage", "المرحلة", 160, "e"), ("worker", "العامل", 140, "e"),
                           ("eta", "ETA", 140, "center")]:
            self.upcoming_tree.heading(c, text=t)
            self.upcoming_tree.column(c, width=w, anchor=a)
        self.upcoming_tree.pack(fill=tk.BOTH, expand=True, pady=6)

        right = tk.Frame(info, bg=UI["SURFACE"], padx=12, pady=10)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(6, 0))
        ttk.Label(right, text="العمال — مشغول / متاح", style="H3.TLabel").pack(anchor="e")
        self.workers_tree = ttk.Treeview(
            right, columns=("name", "busy", "stage"), show="headings", height=10
        )
        for c, t, w, a in [("name", "العامل", 180, "e"),
                           ("busy", "الحالة", 100, "center"),
                           ("stage", "المرحلة الحالية", 260, "e")]:
            self.workers_tree.heading(c, text=t)
            self.workers_tree.column(c, width=w, anchor=a)
        self.workers_tree.pack(fill=tk.BOTH, expand=True, pady=6)

    def _kpi(self, parent: tk.Widget, key: str, title: str, bg: str, fg: str) -> None:
        card = tk.Frame(parent, bg=bg, padx=16, pady=12)
        card.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=4)
        tk.Label(card, text=title, bg=bg, fg=fg, font=FONTS["h3"]).pack(anchor="e")
        big = tk.Label(card, text="0", bg=bg, fg=fg, font=FONTS["big"])
        big.pack(anchor="e")
        self._labels[key] = big

    def refresh(self) -> None:
        # KPI counts
        rows = self.db.query(
            """SELECT status, COUNT(*) AS c
               FROM production_orders
               GROUP BY status"""
        )
        counts = {r["status"]: r["c"] for r in rows}
        running = counts.get("running", 0)
        paused = counts.get("paused", 0)
        blocked = counts.get("blocked", 0)
        planned_row = self.db.query_one(
            """SELECT COUNT(*) AS c
               FROM production_orders po
               WHERE po.status NOT IN ('running','paused','blocked','done','cancelled')
                  OR EXISTS (
                      SELECT 1
                      FROM order_stages os
                      WHERE os.order_id = po.id
                        AND os.is_optional_selected=1
                        AND os.status='planned'
                  )"""
        )
        planned = int(planned_row["c"]) if planned_row else 0
        today = date.today().isoformat()
        done_today_row = self.db.query_one(
            """SELECT COUNT(*) AS c FROM production_orders
               WHERE status='done' AND substr(completed_at,1,10)=?""",
            (today,),
        )
        done_today = int(done_today_row["c"]) if done_today_row else 0

        # Overdue
        orders = self.db.query(
            """SELECT id, due_at FROM production_orders
               WHERE status NOT IN ('done','cancelled') AND due_at IS NOT NULL"""
        )
        now = datetime.now()
        overdue = 0
        for o in orders:
            try:
                due = datetime.fromisoformat(o["due_at"])
            except Exception:
                continue
            eta = find_order_eta(self.db, o["id"])
            if (eta and eta > due) or due < now:
                overdue += 1

        self._labels["planned"].config(text=str(planned))
        self._labels["running"].config(text=str(running))
        self._labels["paused"].config(text=str(paused))
        self._labels["blocked"].config(text=str(blocked))
        self._labels["done_tod"].config(text=str(done_today))
        self._labels["overdue"].config(text=str(overdue))

        # upcoming today
        self.upcoming_tree.delete(*self.upcoming_tree.get_children())
        upcoming = self.db.query(
            """SELECT po.order_number AS num, p.name AS product,
                      os.name_snapshot AS stage_name, w.name AS worker_name,
                      COALESCE(os.planned_start, os.actual_start) AS start_at
               FROM order_stages os
               JOIN production_orders po ON po.id = os.order_id
               JOIN products p ON p.id = po.product_id
               LEFT JOIN workers w ON w.id = os.assigned_worker_id
               WHERE os.is_optional_selected=1
                 AND os.status IN ('planned','running')
                 AND po.status NOT IN ('done','cancelled')
               ORDER BY start_at
               LIMIT 20"""
        )
        for u in upcoming:
            self.upcoming_tree.insert(
                "", "end",
                values=(u["num"], u["product"], u["stage_name"],
                        u["worker_name"] or "—", fmt_ts(u["start_at"])),
            )

        self.workers_tree.delete(*self.workers_tree.get_children())
        workers = self.db.query(
            """SELECT w.id, w.name,
                      (SELECT COUNT(*) FROM order_stages os
                       WHERE os.assigned_worker_id = w.id AND os.status='running') AS busy,
                      (SELECT os.name_snapshot FROM order_stages os
                       WHERE os.assigned_worker_id = w.id AND os.status='running'
                       ORDER BY os.actual_start DESC LIMIT 1) AS cur_stage
               FROM workers w WHERE w.is_active=1 ORDER BY w.name"""
        )
        for w in workers:
            self.workers_tree.insert(
                "", "end",
                values=(w["name"], "مشغول" if w["busy"] else "متاح",
                        w["cur_stage"] or "—"),
            )


# =============================================================================
# Settings
# =============================================================================


class SettingsFrame(ttk.Frame):
    def __init__(self, parent: tk.Misc, app: "FactoryApp") -> None:
        super().__init__(parent)
        self.app = app
        self.db = app.db
        self._build()

    def _build(self) -> None:
        top = tk.Frame(self, bg=UI["BG"])
        top.pack(fill=tk.X, padx=12, pady=8)
        ttk.Label(top, text="الإعدادات", style="H1.TLabel").pack(side=tk.RIGHT)

        card = tk.Frame(self, bg=UI["SURFACE"], padx=16, pady=16)
        card.pack(fill=tk.BOTH, expand=False, padx=12, pady=6)

        # Shop hours
        self.start_var = tk.StringVar(value=self.db.get_setting("shop_start") or "08:00")
        self.end_var = tk.StringVar(value=self.db.get_setting("shop_end") or "18:00")
        self.weekend_var = tk.StringVar(value=self.db.get_setting("weekend_days") or "fri")

        tk.Label(card, text="بداية يوم العمل (HH:MM)", bg=UI["SURFACE"]).grid(row=0, column=1, sticky="e", padx=8, pady=6)
        ttk.Entry(card, textvariable=self.start_var, width=20, justify="right").grid(row=0, column=0, sticky="we", padx=8, pady=6)

        tk.Label(card, text="نهاية يوم العمل (HH:MM)", bg=UI["SURFACE"]).grid(row=1, column=1, sticky="e", padx=8, pady=6)
        ttk.Entry(card, textvariable=self.end_var, width=20, justify="right").grid(row=1, column=0, sticky="we", padx=8, pady=6)

        tk.Label(card, text="أيام العطلة (مثل: fri أو fri,sat)", bg=UI["SURFACE"]).grid(row=2, column=1, sticky="e", padx=8, pady=6)
        ttk.Entry(card, textvariable=self.weekend_var, width=20, justify="right").grid(row=2, column=0, sticky="we", padx=8, pady=6)

        ttk.Button(card, text="حفظ ساعات العمل", style="Ok.TButton",
                   command=self._save_hours).grid(row=3, column=0, columnspan=2, pady=10)

        # Backup / restore
        bk_card = tk.Frame(self, bg=UI["SURFACE"], padx=16, pady=16)
        bk_card.pack(fill=tk.X, padx=12, pady=6)
        ttk.Label(bk_card, text="نسخ احتياطي", style="H3.TLabel").pack(anchor="e")
        row = tk.Frame(bk_card, bg=UI["SURFACE"])
        row.pack(fill=tk.X, pady=6, anchor="e")
        ttk.Button(row, text="تصدير قاعدة البيانات...", style="Ghost.TButton",
                   command=self._backup).pack(side=tk.LEFT, padx=4)
        ttk.Button(row, text="فتح أرشيف التحليل الشهري", style="Ghost.TButton",
                   command=self._open_analytics_folder).pack(side=tk.LEFT, padx=4)

    def _save_hours(self) -> None:
        try:
            parse_hhmm(self.start_var.get(), (8, 0))
            parse_hhmm(self.end_var.get(), (18, 0))
        except Exception:
            messagebox.showerror("خطأ", "صيغة الوقت غير صحيحة", parent=self)
            return
        self.db.set_setting("shop_start", self.start_var.get())
        self.db.set_setting("shop_end", self.end_var.get())
        self.db.set_setting("weekend_days", self.weekend_var.get() or "fri")
        recompute_all(self.db)
        toast(self.winfo_toplevel(), "تم حفظ ساعات العمل", kind="success")
        self.app.notify_reference_data_changed()

    def _backup(self) -> None:
        target = filedialog.asksaveasfilename(
            parent=self,
            title="حفظ نسخة احتياطية",
            defaultextension=".sqlite3",
            initialfile=f"factory_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.sqlite3",
        )
        if not target:
            return
        try:
            import shutil
            shutil.copy2(DB_PATH, target)
        except Exception as e:
            show_error(self, "خطأ", f"فشل التصدير: {e}")
            return
        toast(self.winfo_toplevel(), "تم حفظ النسخة الاحتياطية", kind="success")

    def _open_analytics_folder(self) -> None:
        try:
            export_monthly_analytics(self.db)
            os.startfile(ANALYTICS_DIR)
        except Exception as e:
            show_error(self, "خطأ", f"تعذر فتح أرشيف التحليل: {e}")


# =============================================================================
# Main application
# =============================================================================


HELP_SECTIONS = [
    ("البداية السريعة", [
        "1) أنشئ قوالب المراحل (تبويب F4) — مع وقت الإعداد ووقت كل قطعة لكل مرحلة.",
        "2) أنشئ المنتجات ومسار الإنتاج (تبويب F3) — أضف منتجًا ثم أضف له المراحل بالترتيب.",
        "3) أضف العمال (تبويب F6) — وفعّل لكل عامل المراحل التي يُسمح له بها.",
        "4) اضغط « أمر إنتاج جديد » (Ctrl+N) — اختر المنتج والكمية واضغط « إطلاق ».",
        "5) راقب التنفيذ في « لوحة الإنتاج » (F1) و « مخطط جانت » (F10).",
    ]),
    ("إدارة المراحل (F4)", [
        "اضغط « ➕ إضافة مرحلة جديدة ».",
        "املأ: كود فريد (مثل CUT) + اسم + وقت الإعداد (دقائق ثابتة لبداية الدفعة) + وقت لكل قطعة.",
        "لا يوجد اختيار «مهارة» منفصل؛ المرحلة نفسها تُستخدم لاحقًا في صلاحيات العمال.",
        "إجمالي وقت المرحلة = إعداد + (وقت القطعة × الكمية).",
        "لتعطيل مرحلة مؤقتًا بدون حذف: استخدم « ⏼ تفعيل/تعطيل ».",
    ]),
    ("إدارة المنتجات (F3)", [
        "هنا تُسجَّل « أنواع المنتجات » فقط (قميص، بنطلون، جاكيت …) ومسار الإنتاج الخاص بكل نوع.",
        "تفاصيل الطلب (اسم المدرسة، المقاس، اللون، الكمية) تُدخَل عند إنشاء أمر الإنتاج نفسه.",
        "على اليمين: اضغط « ➕ نوع منتج جديد » لإضافة نوع (مع كود اختياري).",
        "على اليسار: بعد اختيار نوع، اضغط « ➕ إضافة مرحلة للمسار » لبناء مسار الإنتاج.",
        "كل مرحلة لها خيارين:",
        "   • افتراضية = يتم اختيارها تلقائيًا عند إنشاء أمر جديد.",
        "   • اختيارية = يمكن إلغاؤها في معالج الأمر (مثل: التشطيب للزبائن المميزين فقط).",
        "استخدم « ⬆ لأعلى / ⬇ لأسفل » لترتيب المراحل (قص → حياكة → كوي → ...).",
        "لتخصيص أوقات لمنتج معين: أدخل قيمًا في « إعداد مخصص » و « لكل قطعة مخصص » (اترك فارغًا لاستخدام الافتراضي).",
    ]),
    ("إدارة العمال (F6)", [
        "على اليمين: اضغط « ➕ عامل جديد » وأدخل الكود والاسم.",
        "على اليسار: اختر العامل، ثم اختر مرحلة من القائمة واضغط « ✓ السماح بالمرحلة المختارة ».",
        "عند تعيين عامل لمرحلة في أمر إنتاج، يظهر فقط العمال المسموح لهم بهذه المرحلة.",
    ]),
    ("إنشاء أمر إنتاج جديد (Ctrl+N)", [
        "« نوع المنتج » يُختار من قائمة منسدلة (قائمة الأنواع تُدار من شاشة المنتجات).",
        "« اسم المدرسة » و « المقاس » و « اللون » و « الكمية »: حقول نصية حرة — تكتب فيها ما تريد،",
        "  وتظهر في قائمتها القيم التي استخدمتها في أوامر سابقة لإعادة الاختيار بسرعة.",
        "أدخل الأولوية (1 = الأعلى) وتاريخ الاستحقاق (اختياري).",
        "المراحل الافتراضية للنوع تظهر مُفعّلة، والاختيارية يمكنك تبديلها.",
        "يمكنك تحديد « البدء من » لأمر يبدأ من منتصف المسار (مثلًا: قطع جاهزة بالفعل).",
        "بمجرد الضغط على زر الإنشاء، يبدأ أمر الإنتاج فورًا وتتحول أول مرحلة إلى قيد التنفيذ تلقائيًا.",
    ]),
    ("تنفيذ الأوامر (من لوحة الإنتاج أو الأوامر)", [
        "انقر نقرًا مزدوجًا على أي بطاقة/أمر لفتح شاشة التفاصيل.",
        "اختر مرحلة من الجدول ثم استخدم الأزرار:",
        "   • بدء → إنهاء → للمرحلة التالية.",
        "   • إيقاف مؤقت/استئناف عند توقف العمل.",
        "   • « تخطي » و « إدراج مرحلة » و « تشغيل التالي بالتوازي » و « إلغاء » متاحة بدون كلمة مرور.",
    ]),
    ("لوحة الإنتاج ومخطط جانت", [
        "لوحة الإنتاج (F1): عرض كانبان — الأوامر موزعة على أعمدة حسب الحالة.",
        "مخطط جانت (F10): عرض زمني — كل أمر سطر، والمراحل مستطيلات ملوّنة على محور الزمن.",
        "الخط الأحمر المتقطع في جانت = الآن. علامة ⏰ الحمراء = تاريخ الاستحقاق.",
        "اضغط على أي مستطيل في جانت لفتح الأمر مباشرة.",
    ]),
    ("الإعدادات (F8)", [
        "ساعات العمل وأيام العطلة تُستخدم لحساب ETA (وقت الانتهاء المتوقع).",
        "زر « تصدير قاعدة البيانات » يحفظ نسخة احتياطية في مكان تختاره.",
    ]),
]


class HelpDialog(tk.Toplevel):
    """Step-by-step user guide shown on first launch or via the Help button."""

    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.title("دليل المستخدم — HosnyFactory")
        self.configure(bg=UI["BG"])
        self.transient(parent)
        self.geometry("920x680")
        try:
            self.grab_set()
        except Exception:
            pass

        header = tk.Frame(self, bg=UI["SURFACE"], padx=16, pady=12)
        header.pack(fill=tk.X)
        tk.Label(header, text="دليل استخدام HosnyFactory", bg=UI["SURFACE"],
                 fg=UI["ACCENT"], font=FONTS["h1"]).pack(side=tk.RIGHT)
        tk.Label(header, text="اتبع الخطوات بالترتيب لإعداد المصنع وبدء الإنتاج",
                 bg=UI["SURFACE"], fg=UI["TEXT_SEC"], font=FONTS["small"]).pack(side=tk.RIGHT, padx=12)

        # Scrollable content
        body = tk.Frame(self, bg=UI["BG"])
        body.pack(fill=tk.BOTH, expand=True, padx=16, pady=8)

        canvas = tk.Canvas(body, bg=UI["BG"], highlightthickness=0)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=vsb.set)
        content = tk.Frame(canvas, bg=UI["BG"])
        win_id = canvas.create_window((0, 0), window=content, anchor="nw")

        def _on_resize(event, c=canvas, wid=win_id):
            c.itemconfigure(wid, width=event.width)
        canvas.bind("<Configure>", _on_resize)
        content.bind("<Configure>", lambda e, c=canvas: c.configure(scrollregion=c.bbox("all")))

        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"),
        )

        for title, items in HELP_SECTIONS:
            card = tk.Frame(content, bg=UI["SURFACE"], padx=14, pady=10,
                            highlightbackground=UI["BORDER"], highlightthickness=1)
            card.pack(fill=tk.X, pady=6)
            tk.Label(card, text=title, bg=UI["SURFACE"], fg=UI["ACCENT"],
                     font=FONTS["h3"], anchor="e").pack(fill=tk.X)
            for line in items:
                tk.Label(card, text="• " + line, bg=UI["SURFACE"], fg=UI["TEXT"],
                         font=FONTS["body"], justify="right", anchor="e",
                         wraplength=820).pack(fill=tk.X, pady=2)

        footer = tk.Frame(self, bg=UI["BG"])
        footer.pack(fill=tk.X, padx=16, pady=(4, 12))
        ttk.Button(footer, text="إغلاق", style="Ok.TButton",
                   command=self._close).pack(side=tk.LEFT)
        tk.Label(footer, text="يمكنك فتح هذا الدليل لاحقًا من زر « مساعدة » في الأعلى.",
                 bg=UI["BG"], fg=UI["TEXT_SEC"], font=FONTS["small"]).pack(side=tk.RIGHT)

    def _close(self) -> None:
        try:
            self.unbind_all("<MouseWheel>")
        except Exception:
            pass
        self.destroy()


class FactoryApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.db = DB()
        normalize_order_statuses(self.db)
        export_monthly_analytics(self.db)
        self.title("HosnyFactory — إدارة الإنتاج")
        self.geometry("1320x820")
        try:
            self.state("zoomed")
        except Exception:
            pass
        setup_style(self)

        self._build_header()
        self._build_tabs()
        self._bind_shortcuts()
        self._tick_auto_refresh()
        self.after(400, self._maybe_show_first_run_help)

    # ------ layout ------ #

    def _build_header(self) -> None:
        header = tk.Frame(self, bg=UI["SURFACE"], padx=12, pady=8)
        header.pack(fill=tk.X)
        tk.Label(header, text="HosnyFactory", bg=UI["SURFACE"], fg=UI["ACCENT"],
                 font=FONTS["h1"]).pack(side=tk.RIGHT)
        tk.Label(header, text="إدارة الإنتاج المحلية — V1",
                 bg=UI["SURFACE"], fg=UI["TEXT_SEC"], font=FONTS["small"]).pack(side=tk.RIGHT, padx=12)
        ttk.Button(header, text="➕ أمر إنتاج جديد (Ctrl+N)", style="Ok.TButton",
                   command=self.open_new_order).pack(side=tk.LEFT, padx=4)
        ttk.Button(header, text="❓ مساعدة (F1 = لوحة، F12 = دليل)",
                   style="Ghost.TButton",
                   command=self.open_help).pack(side=tk.LEFT, padx=4)

    def _build_tabs(self) -> None:
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)

        self.dashboard = DashboardFrame(self.nb, self)
        self.nb.add(self.dashboard, text="اللوحة الرئيسية (F9)")

        self.pipeline = PipelineBoard(self.nb, self)
        self.nb.add(self.pipeline, text="لوحة الإنتاج (F1)")

        self.gantt = GanttChart(self.nb, self)
        self.nb.add(self.gantt, text="مخطط جانت (F10)")

        self.worker_plan = WorkerPlanFrame(self.nb, self)
        self.nb.add(self.worker_plan, text="خطة العمال (F11)")

        self.orders = OrdersList(self.nb, self)
        self.nb.add(self.orders, text="الأوامر (F2)")

        self.products = ProductsAdmin(self.nb, self)
        self.nb.add(self.products, text="المنتجات (F3)")

        self.stages = StagesAdmin(self.nb, self)
        self.nb.add(self.stages, text="المراحل (F4)")

        self.workers = WorkersAdmin(self.nb, self)
        self.nb.add(self.workers, text="العمال (F6)")

        self.reports = ReportsFrame(self.nb, self)
        self.nb.add(self.reports, text="التقارير (F7)")

        self.settings = SettingsFrame(self.nb, self)
        self.nb.add(self.settings, text="الإعدادات (F8)")

        self.nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def _bind_shortcuts(self) -> None:
        # Tab order: 0=Dashboard 1=Pipeline 2=Gantt 3=WorkerPlan 4=Orders
        #            5=Products 6=Stages 7=Workers 8=Reports 9=Settings
        binds = {
            "<F9>":  0,
            "<F1>":  1,
            "<F10>": 2,
            "<F11>": 3,
            "<F2>":  4,
            "<F3>":  5,
            "<F4>":  6,
            "<F6>":  7,
            "<F7>":  8,
            "<F8>":  9,
        }
        for k, idx in binds.items():
            self.bind_all(k, lambda _e, i=idx: self.nb.select(i))
        self.bind_all("<Control-n>", lambda _e: self.open_new_order())
        self.bind_all("<F12>", lambda _e: self.open_help())

    # ------ helpers ------ #

    def open_new_order(self) -> None:
        NewOrderWizard(self, self)

    def open_order(self, order_id: int) -> None:
        OrderDetail(self, self, order_id)

    def open_help(self) -> None:
        HelpDialog(self)

    def _maybe_show_first_run_help(self) -> None:
        """Show the help guide the first time the app is opened."""
        seen = self.db.get_setting("help_seen")
        if seen == "1":
            return
        try:
            HelpDialog(self)
        except Exception:
            log.exception("first-run help failed")
        self.db.set_setting("help_seen", "1")

    def notify_data_changed(self) -> None:
        """Backward-compatible alias for order-centric refreshes."""
        self.notify_order_changed()

    def notify_order_changed(self) -> None:
        """Refresh active operational views without repainting the whole app."""
        try:
            export_monthly_analytics(self.db)
        except Exception:
            log.exception("monthly analytics export failed")
        frames: List[tk.Misc] = [self.dashboard, self.pipeline, self.orders]
        try:
            cur = self.nb.select()
            if cur:
                tab = self.nb.nametowidget(cur)
                if tab not in frames:
                    frames.append(tab)
        except Exception:
            log.exception("failed selecting active tab for order refresh")
        self._refresh_frames(frames)
        try:
            self.update_idletasks()
        except Exception:
            log.exception("UI update after data change")

    def notify_reference_data_changed(self) -> None:
        """Refresh all tabs after master-data/settings changes."""
        try:
            export_monthly_analytics(self.db)
        except Exception:
            log.exception("monthly analytics export failed")
        self._refresh_all_tabs()

    def _refresh_visible_tabs(self) -> None:
        """Refresh only active tab + key high-level views for responsiveness."""
        frames: List[tk.Misc] = [self.dashboard, self.pipeline, self.orders]
        try:
            cur = self.nb.select()
            if cur:
                tab = self.nb.nametowidget(cur)
                if tab not in frames:
                    frames.append(tab)
        except Exception:
            log.exception("failed selecting active tab for refresh")
        for frame in frames:
            try:
                refresh = getattr(frame, "refresh", None)
                if callable(refresh):
                    refresh()
            except Exception:
                log.exception("refresh failed for %s", type(frame).__name__)

    def invalidate_caches(self) -> None:
        # no caches yet, but keep the hook for future
        self.notify_reference_data_changed()

    def _refresh_frames(self, frames: Iterable[tk.Misc]) -> None:
        seen: set[int] = set()
        for frame in frames:
            ident = id(frame)
            if ident in seen:
                continue
            seen.add(ident)
            try:
                refresh = getattr(frame, "refresh", None)
                if callable(refresh):
                    refresh()
            except Exception:
                log.exception("refresh failed for %s", type(frame).__name__)

    def _refresh_all_tabs(self) -> None:
        self._refresh_frames(
            (
                self.dashboard,
                self.pipeline,
                self.gantt,
                self.worker_plan,
                self.orders,
                self.products,
                self.stages,
                self.workers,
                self.reports,
            )
        )

    def _on_tab_changed(self, _evt=None) -> None:
        try:
            idx = self.nb.index(self.nb.select())
            tab = self.nb.nametowidget(self.nb.tabs()[idx])
            refresh = getattr(tab, "refresh", None)
            if callable(refresh):
                refresh()
        except Exception:
            log.exception("tab change refresh failed")

    def _tick_auto_refresh(self) -> None:
        """Light periodic refresh so running-stage cards stay fresh."""
        try:
            # Keep timer lightweight to avoid UI freezes.
            for r in self.db.query(
                """SELECT DISTINCT po.id
                   FROM production_orders po
                   LEFT JOIN order_stages os ON os.order_id = po.id
                   WHERE po.status IN ('running','paused','blocked')
                      OR (
                          po.status NOT IN ('done','cancelled')
                          AND os.is_optional_selected=1
                          AND os.status='planned'
                      )"""
            ):
                oid = int(r["id"])
                # Re-plan and try auto-dispatch so queued stages can start
                # as soon as a qualified worker is free.
                sync_order_state(self.db, oid, dispatch=True, ensure_started=True)
            cur = self.nb.select()
            if cur:
                tab = self.nb.nametowidget(cur)
                if isinstance(tab, (PipelineBoard, DashboardFrame, GanttChart, WorkerPlanFrame)):
                    tab.refresh()
        except Exception:
            log.exception("auto refresh failed")
        # Refresh more frequently so newly free workers pick up queued stages faster.
        self.after(10_000, self._tick_auto_refresh)


# =============================================================================
# Entrypoint
# =============================================================================


def main() -> None:
    try:
        app = FactoryApp()
    except Exception as e:
        log.exception("startup failed")
        try:
            messagebox.showerror("خطأ", f"فشل بدء التطبيق:\n{e}")
        except Exception:
            pass
        return
    app.mainloop()


if __name__ == "__main__":
    main()
