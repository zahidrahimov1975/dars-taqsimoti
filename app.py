#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dars taqsimoti / O'qituvchilar yuklamasi — mustaqil dastur.
Standalone teaching-load (workload) planner. Stdlib only: tkinter + sqlite3.

Build into a Windows .exe with:
    pip install pyinstaller
    pyinstaller --onefile --windowed --name DarsTaqsimoti app.py
The database file `dars_taqsimoti.db` lives next to the .exe.
"""

import os
import sys
import csv
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_TITLE = "Dars taqsimoti — O'qituvchilar yuklamasi"
DB_NAME = "dars_taqsimoti.db"

# ----- Lookups (matching the original Access value lists) -----
UNVON = ["DSc", "PhD", "Darajasiz"]
KATEGORIYA = ["1", "2"]
YONALISH = ["Buxgalteriya Hisobi", "Iqtisodiyot", "Biznesni Boshqarish"]
TALIM_TURI = ["Kunduzgi", "Sirtqi", "Masofaviy", "Kechki"]
TUR_SOAT = ["Maruza", "Amaliyot"]                # + "Reyting" only when subject is Masofaviy
SEMESTR = [str(i) for i in range(1, 13)]
ALL = "— hammasi —"

SCHEMA = """
CREATE TABLE IF NOT EXISTS Domlalar(
  DomlaID INTEGER PRIMARY KEY AUTOINCREMENT,
  FIO TEXT NOT NULL, IlmiyUnvon TEXT, Kategoriya INTEGER,
  Stavka REAL DEFAULT 1, Meyor1St REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS Fanlar(
  FanID INTEGER PRIMARY KEY AUTOINCREMENT,
  FanNomi TEXT NOT NULL, Yonalish TEXT, TalimTuri TEXT, Kategoriya INTEGER,
  Semestr INTEGER, Maruza REAL DEFAULT 0, Amaliyot REAL DEFAULT 0,
  Potok INTEGER DEFAULT 1, Guruh INTEGER DEFAULT 1, Reyting REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS Taqsimot(
  TaqsimotID INTEGER PRIMARY KEY AUTOINCREMENT,
  DomlaID INTEGER, FanID INTEGER, TurSoat TEXT, Soat REAL DEFAULT 0,
  FOREIGN KEY(DomlaID) REFERENCES Domlalar(DomlaID),
  FOREIGN KEY(FanID)   REFERENCES Fanlar(FanID));
"""


def db_path():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, DB_NAME)


def connect(path=None):
    con = sqlite3.connect(path or db_path())
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA foreign_keys = ON")
    con.executescript(SCHEMA)
    return con


# ----- Computed values (the empty 'Jami' columns in Access) -----
def fan_totals(maruza, amaliyot, potok, guruh, reyting):
    mj = (maruza or 0) * (potok or 0)
    aj = (amaliyot or 0) * (guruh or 0)
    return mj, aj, mj + aj + (reyting or 0)


def meyor_jami(meyor1st, stavka):
    return (meyor1st or 0) * (stavka or 0)


def workload_rows(con):
    sql = """
    SELECT d.DomlaID, d.FIO, d.Stavka, d.Meyor1St,
           COALESCE(SUM(CASE WHEN t.TurSoat='Maruza'   THEN t.Soat END),0) AS maruza,
           COALESCE(SUM(CASE WHEN t.TurSoat='Amaliyot' THEN t.Soat END),0) AS amaliyot,
           COALESCE(SUM(CASE WHEN t.TurSoat='Reyting'  THEN t.Soat END),0) AS reyting,
           COALESCE(SUM(t.Soat),0) AS jami
    FROM Domlalar d LEFT JOIN Taqsimot t ON t.DomlaID = d.DomlaID
    GROUP BY d.DomlaID ORDER BY d.FIO COLLATE NOCASE"""
    out = []
    for r in con.execute(sql):
        norm = meyor_jami(r["Meyor1St"], r["Stavka"])
        diff = r["jami"] - norm
        pct = (r["jami"] / norm * 100) if norm else 0
        out.append({"id": r["DomlaID"], "fio": r["FIO"], "stavka": r["Stavka"],
                    "norm": norm, "maruza": r["maruza"], "amaliyot": r["amaliyot"],
                    "reyting": r["reyting"], "jami": r["jami"], "diff": diff, "pct": pct})
    return out


def g(v):
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else f"{f:g}"
    except (TypeError, ValueError):
        return "" if v is None else str(v)


def short_name(fio):
    """Familiya Ism only (first two name parts)."""
    parts = (fio or "").split()
    return " ".join(parts[:2]) if parts else (fio or "")


def fan_rules(dlg):
    """Reyting only applies to Masofaviy subjects; otherwise disable the field and force 0."""
    w = dlg.widgets.get("Reyting")
    if w is None:
        return
    if dlg.vars["TalimTuri"].get().strip() == "Masofaviy":
        w.config(state="normal")
    else:
        if dlg.vars["Reyting"].get() not in ("", "0", "0.0"):
            dlg.vars["Reyting"].set("0")
        w.config(state="disabled")


# ===================== Generic Add/Edit dialog (Domlalar, Fanlar) =====================
class FormDialog(tk.Toplevel):
    FIXED = {"IlmiyUnvon", "Kategoriya", "Semestr"}

    def __init__(self, master, title, spec, values=None, computed=None, rules=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.spec = spec
        self.vars = {}
        self.widgets = {}
        self.computed = computed
        self.rules = rules          # optional callable(self) to enable/disable dependent fields
        self._guard = False
        values = values or {}

        frm = ttk.Frame(self, padding=16)
        frm.grid(sticky="nsew")
        for i, (key, label, kind, opts) in enumerate(spec):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=4, padx=(0, 12))
            var = tk.StringVar(value="" if values.get(key) is None else str(values.get(key)))
            self.vars[key] = var
            if kind == "combo":
                state = "readonly" if key in self.FIXED else "normal"
                w = ttk.Combobox(frm, textvariable=var, values=opts, state=state, width=28)
            else:
                w = ttk.Entry(frm, textvariable=var, width=30)
            w.grid(row=i, column=1, sticky="ew", pady=4)
            self.widgets[key] = w
            var.trace_add("write", lambda *_: self._on_change())

        self.preview = None
        if computed:
            self.preview = ttk.Label(frm, text="", foreground="#0a58ca")
            self.preview.grid(row=len(spec), column=0, columnspan=2, sticky="w", pady=(8, 0))
            self._refresh_preview()

        btns = ttk.Frame(frm)
        btns.grid(row=len(spec) + 1, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(btns, text="Saqlash", command=self._save).pack(side="right")
        ttk.Button(btns, text="Bekor qilish", command=self.destroy).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda e: self._save())
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(master)
        self.grab_set()
        self.update_idletasks()
        x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x,0)}+{max(y,0)}")
        self._apply_rules()
        self.wait_window(self)

    def _on_change(self):
        if self._guard:
            return
        self._guard = True
        try:
            self._refresh_preview()
            self._apply_rules()
        finally:
            self._guard = False

    def _apply_rules(self):
        if self.rules:
            self.rules(self)

    def _collect(self):
        out = {}
        for key, label, kind, opts in self.spec:
            raw = self.vars[key].get().strip()
            if kind == "int":
                out[key] = int(float(raw)) if raw else 0
            elif kind == "float":
                out[key] = float(raw.replace(",", ".")) if raw else 0.0
            else:
                out[key] = raw
        return out

    def _refresh_preview(self):
        if not (self.computed and self.preview):
            return
        try:
            self.preview.config(text=self.computed(self._collect()))
        except (ValueError, KeyError):
            self.preview.config(text="")

    def _save(self):
        try:
            data = self._collect()
        except ValueError:
            messagebox.showerror("Xato", "Raqamli maydonlarni to'g'ri kiriting.", parent=self)
            return
        first = self.spec[0]
        if first[2] == "text" and not data.get(first[0]):
            messagebox.showerror("Xato", f"\"{first[1]}\" maydoni bo'sh bo'lmasligi kerak.", parent=self)
            return
        self.result = data
        self.destroy()


# ===================== Taqsimot assignment dialog (like the Access 'Domla yuklamasi' form) =====================
class TaqsimotDialog(tk.Toplevel):
    def __init__(self, master, con, values=None, editing_id=None):
        super().__init__(master)
        self.title("Taqsimot yozuvi — Domla yuklamasi")
        self.resizable(False, False)
        self.con = con
        self.result = None
        self.editing_id = editing_id      # TaqsimotID being edited (excluded from assigned-hours)

        self.domlalar = con.execute(
            "SELECT DomlaID, FIO FROM Domlalar ORDER BY FIO COLLATE NOCASE").fetchall()
        self.fanlar = con.execute(
            "SELECT FanID, FanNomi, Yonalish, TalimTuri, Semestr, Maruza, Amaliyot, Reyting, Potok, Guruh "
            "FROM Fanlar ORDER BY FanNomi COLLATE NOCASE").fetchall()
        self.domla_ids = [r["DomlaID"] for r in self.domlalar]
        self.assigned_hours = self._load_assigned_hours()
        self.components = []              # parallel to cb_fan values

        pad = dict(padx=6, pady=5)
        frm = ttk.Frame(self, padding=16)
        frm.grid()
        ttk.Label(frm, text="DOMLA YUKLAMASI — DARS TAQSIMOTI",
                  style="Title.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        ttk.Label(frm, text="Domla:").grid(row=1, column=0, sticky="w", **pad)
        self.cb_domla = ttk.Combobox(frm, state="readonly", width=54,
                                     values=[short_name(r["FIO"]) for r in self.domlalar])
        self.cb_domla.grid(row=1, column=1, columnspan=3, sticky="we", **pad)

        ttk.Label(frm, text="— Fan tanlash (filtrlar ixtiyoriy) —",
                  foreground="#666").grid(row=2, column=0, columnspan=4, sticky="w", pady=(12, 0))

        ttk.Label(frm, text="Yo'nalish:").grid(row=3, column=0, sticky="w", **pad)
        self.cb_yon = ttk.Combobox(frm, state="readonly", width=22, values=[ALL] + self._distinct("Yonalish"))
        self.cb_yon.grid(row=3, column=1, sticky="we", **pad)
        ttk.Label(frm, text="Ta'lim shakli:").grid(row=3, column=2, sticky="w", **pad)
        self.cb_talim = ttk.Combobox(frm, state="readonly", width=16, values=[ALL] + TALIM_TURI)
        self.cb_talim.grid(row=3, column=3, sticky="we", **pad)

        ttk.Label(frm, text="Semestr:").grid(row=4, column=0, sticky="w", **pad)
        self.cb_sem = ttk.Combobox(frm, state="readonly", width=10, values=[ALL] + self._distinct_sem())
        self.cb_sem.grid(row=4, column=1, sticky="w", **pad)

        ttk.Label(frm, text="Fan / yuklama:").grid(row=5, column=0, sticky="w", **pad)
        self.cb_fan = ttk.Combobox(frm, state="readonly", width=54)
        self.cb_fan.grid(row=5, column=1, columnspan=3, sticky="we", **pad)
        ttk.Label(frm, text="(Ma'ruza, Amaliyot va Reyting alohida; biriktirilgani ro'yxatdan chiqadi)",
                  foreground="#888").grid(row=6, column=1, columnspan=3, sticky="w", padx=6)

        ttk.Label(frm, text="Soat:").grid(row=7, column=0, sticky="w", **pad)
        self.var_soat = tk.StringVar()
        self.ent_soat = ttk.Entry(frm, textvariable=self.var_soat, width=12)
        self.ent_soat.grid(row=7, column=1, sticky="w", **pad)
        self.lbl_yuk = ttk.Label(frm, text="Yuklama: —", foreground="#0a58ca")
        self.lbl_yuk.grid(row=7, column=2, columnspan=2, sticky="w", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=8, column=0, columnspan=4, sticky="we", pady=(14, 0))
        ttk.Button(btns, text="Filtrni tozalash", command=self._clear_filters).pack(side="left")
        ttk.Button(btns, text="Saqlash", command=self._save).pack(side="right")
        ttk.Button(btns, text="Bekor qilish", command=self.destroy).pack(side="right", padx=(0, 8))

        for cb in (self.cb_yon, self.cb_talim, self.cb_sem):
            cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_fan())
        self.cb_fan.bind("<<ComboboxSelected>>", lambda e: self._on_fan())

        self.cb_yon.set(ALL)
        self.cb_talim.set(ALL)
        self.cb_sem.set(ALL)
        self._refresh_fan()
        if values:
            self._prefill(values)

        self.transient(master)
        self.grab_set()
        self.bind("<Escape>", lambda e: self.destroy())
        self.update_idletasks()
        self.geometry(f"+{master.winfo_rootx() + 60}+{master.winfo_rooty() + 50}")
        self.wait_window(self)

    def _distinct(self, key):
        seen = []
        for r in self.fanlar:
            v = (r[key] or "").strip()
            if v and v not in seen:
                seen.append(v)
        return sorted(seen)

    def _distinct_sem(self):
        return [str(x) for x in sorted({int(r["Semestr"]) for r in self.fanlar if r["Semestr"]})]

    def _load_assigned_hours(self):
        """Hours already assigned per (FanID, TurSoat) across all teachers, so a component
        can be split among several teachers and only drops out when fully allocated."""
        if self.editing_id:
            cur = self.con.execute("SELECT FanID, TurSoat, COALESCE(SUM(Soat),0) s FROM Taqsimot "
                                   "WHERE TaqsimotID<>? GROUP BY FanID, TurSoat", (self.editing_id,))
        else:
            cur = self.con.execute("SELECT FanID, TurSoat, COALESCE(SUM(Soat),0) s FROM Taqsimot "
                                   "GROUP BY FanID, TurSoat")
        return {(r["FanID"], r["TurSoat"]): r["s"] for r in cur}

    def _match(self, r):
        y, t, s = self.cb_yon.get(), self.cb_talim.get(), self.cb_sem.get()
        if y != ALL and (r["Yonalish"] or "") != y:
            return False
        if t != ALL and (r["TalimTuri"] or "") != t:
            return False
        if s != ALL:
            rsem = str(int(r["Semestr"])) if r["Semestr"] else ""
            if rsem != s:
                return False
        return True

    def _build_components(self):
        """One entry per course-component, with type-specific rules:
          - Maruza  : one professor for the whole course (total = Maruza), NOT splittable.
          - Amaliyot: total = Amaliyot × Guruh (groups), splittable across teachers (per group).
          - Reyting : total = Reyting, splittable; only for Masofaviy, eligible teachers only.
        A splittable component stays (showing remaining hours) until fully assigned."""
        comps = []
        for r in self.fanlar:
            if not self._match(r):
                continue
            specs = [("Maruza", (r["Maruza"] or 0), False, (r["Maruza"] or 0)),
                     ("Amaliyot", (r["Amaliyot"] or 0) * (r["Guruh"] or 1), True, (r["Amaliyot"] or 0))]
            if (r["TalimTuri"] or "") == "Masofaviy":        # Reyting only for distance education
                specs.append(("Reyting", (r["Reyting"] or 0), True, (r["Reyting"] or 0)))
            for turi, total, splittable, unit in specs:
                if total <= 0:
                    continue
                done = self.assigned_hours.get((r["FanID"], turi), 0)
                if not splittable:
                    if done > 0:                             # lecture already taken -> hide
                        continue
                    remaining, default = total, total
                else:
                    remaining = total - done
                    if remaining <= 0:                       # fully assigned -> hide
                        continue
                    default = min(unit, remaining) if unit else remaining
                comps.append({"FanID": r["FanID"], "FanNomi": r["FanNomi"], "TurSoat": turi,
                              "total": total, "remaining": remaining, "default": default,
                              "splittable": splittable, "row": r})
        return comps

    def _comp_label(self, c):
        r = c["row"]
        extra = []
        if r["TalimTuri"]:
            extra.append(r["TalimTuri"])
        if r["Semestr"]:
            extra.append(f'{int(r["Semestr"])}-sem')
        tail = (" · " + ", ".join(extra)) if extra else ""
        if c["splittable"] and abs(c["remaining"] - c["total"]) > 1e-9:
            hrs = f'qoldi {g(c["remaining"])}/{g(c["total"])} soat'
        else:
            hrs = f'{g(c["total"])} soat'
        return f'{c["FanNomi"]} — {c["TurSoat"]} ({hrs}){tail}'

    def _refresh_fan(self):
        self.components = self._build_components()
        self.cb_fan["values"] = [self._comp_label(c) for c in self.components]
        self.cb_fan.set("")
        self.var_soat.set("")
        self.ent_soat.config(state="normal")
        self.lbl_yuk.config(text="Yuklama: —")

    def _current(self):
        i = self.cb_fan.current()
        if i < 0 or i >= len(self.components):
            return None
        return self.components[i]

    def _on_fan(self):
        c = self._current()
        if not c:
            return
        self.var_soat.set(g(c["default"]))
        if c["splittable"]:
            self.ent_soat.config(state="normal")
            self.lbl_yuk.config(text=f"{c['TurSoat']}: jami {g(c['total'])} soat, qoldi {g(c['remaining'])} soat")
        else:
            self.ent_soat.config(state="disabled")
            self.lbl_yuk.config(text=f"{c['TurSoat']}: {g(c['total'])} soat — bitta professor")

    def _clear_filters(self):
        self.cb_yon.set(ALL)
        self.cb_talim.set(ALL)
        self.cb_sem.set(ALL)
        self._refresh_fan()

    def _prefill(self, v):
        if v.get("DomlaID") in self.domla_ids:
            self.cb_domla.current(self.domla_ids.index(v["DomlaID"]))
        fan = next((r for r in self.fanlar if r["FanID"] == v.get("FanID")), None)
        if fan:
            self.cb_yon.set(fan["Yonalish"] or ALL)
            self.cb_talim.set(fan["TalimTuri"] or ALL)
            self.cb_sem.set(str(int(fan["Semestr"])) if fan["Semestr"] else ALL)
            self._refresh_fan()
            for idx, c in enumerate(self.components):
                if c["FanID"] == v.get("FanID") and c["TurSoat"] == v.get("TurSoat"):
                    self.cb_fan.current(idx)
                    self._on_fan()
                    break
        if v.get("Soat") is not None:
            self.var_soat.set(g(v["Soat"]))

    def _save(self):
        di = self.cb_domla.current()
        if di < 0:
            messagebox.showerror("Xato", "Domla tanlanishi shart.", parent=self)
            return
        c = self._current()
        if not c:
            messagebox.showerror("Xato", "Fan / yuklama tanlanishi shart.", parent=self)
            return
        try:
            txt = self.var_soat.get().strip().replace(",", ".")
            soat = float(txt) if txt else 0.0
        except ValueError:
            messagebox.showerror("Xato", "Soat raqam bo'lishi kerak.", parent=self)
            return
        if soat <= 0:
            messagebox.showerror("Xato", "Soat 0 dan katta bo'lishi kerak.", parent=self)
            return
        if soat > c["remaining"] + 1e-9:
            if not messagebox.askyesno("Diqqat",
                    f"Bu komponent uchun qolgan soat: {g(c['remaining'])}.\n"
                    f"Siz {g(soat)} soat kiritdingiz. Baribir saqlansinmi?", parent=self):
                return
        if c["TurSoat"] == "Reyting":
            eligible = self.con.execute(
                "SELECT 1 FROM Taqsimot WHERE DomlaID=? AND FanID=? AND TurSoat IN ('Maruza','Amaliyot') "
                "LIMIT 1", (self.domla_ids[di], c["FanID"])).fetchone()
            if not eligible:
                messagebox.showerror("Xato",
                    "Reyting faqat shu fanning Ma'ruza yoki Amaliyotini o'qitadigan domlaga "
                    "biriktiriladi.\nAvval o'sha domlaga shu fandan Ma'ruza yoki Amaliyot biriktiring.",
                    parent=self)
                return
        self.result = {"DomlaID": self.domla_ids[di], "FanID": c["FanID"],
                       "TurSoat": c["TurSoat"], "Soat": soat}
        self.destroy()


# ===================== Main application =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1160x680")
        self.minsize(940, 540)
        try:
            ttk.Style().theme_use("clam")
        except tk.TclError:
            pass
        self._style()
        self.con = connect()

        header = ttk.Frame(self, padding=(14, 10))
        header.pack(fill="x")
        ttk.Label(header, text="Dars taqsimoti va o'qituvchilar yuklamasi",
                  style="Title.TLabel").pack(side="left")

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self._build_domlalar()
        self._build_fanlar()
        self._build_taqsimot()
        self._build_yuklama()
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._on_tab())

        self.status = tk.StringVar(value=f"Ma'lumotlar bazasi: {db_path()}")
        ttk.Label(self, textvariable=self.status, anchor="w",
                  relief="sunken", padding=(8, 3)).pack(fill="x", side="bottom")
        self.refresh_all()

    def _style(self):
        st = ttk.Style()
        st.configure("Title.TLabel", font=("Segoe UI", 15, "bold"))
        st.configure("Treeview", rowheight=24)
        st.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))
        st.configure("TButton", padding=(9, 4))

    # ---------- generic helpers ----------
    def _make_tab(self, title):
        f = ttk.Frame(self.nb, padding=8)
        self.nb.add(f, text=title)
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(0, 6))
        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)
        return f, bar, body

    def _add_search(self, bar, var):
        ttk.Label(bar, text="🔍 Qidirish:").pack(side="left", padx=(16, 4))
        e = ttk.Entry(bar, textvariable=var, width=24)
        e.pack(side="left")

    def _make_tree(self, parent, columns, widths, anchors=None):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="browse")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        anchors = anchors or {}
        for c, w in zip(columns, widths):
            tree.heading(c, text=c)
            tree.column(c, width=w, anchor=anchors.get(c, "w"),
                        stretch=(c in ("F.I.Sh.", "Fan nomi", "Domla (F.I.Sh.)", "Fan")))
        tree.tag_configure("odd", background="#f5f7fa")
        return tree

    @staticmethod
    def _fill(tree, rows):
        sel = tree.selection()
        tree.delete(*tree.get_children())
        for i, (iid, vals) in enumerate(rows):
            tree.insert("", "end", iid=str(iid), values=vals, tags=("odd",) if i % 2 else ())
        if sel and tree.exists(sel[0]):
            tree.selection_set(sel[0])

    @staticmethod
    def _selected_id(tree):
        s = tree.selection()
        return int(s[0]) if s else None

    @staticmethod
    def _i(v, default=0):
        try:
            return int(float(str(v).strip()))
        except (ValueError, TypeError):
            return default

    @staticmethod
    def _f(v, default=0.0):
        try:
            return float(str(v).strip().replace(",", "."))
        except (ValueError, TypeError):
            return default

    # ---------- CSV template + import ----------
    def _save_template(self, headers, example, default_name):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")], initialfile=default_name)
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(headers)
                w.writerow(example)
            messagebox.showinfo("Tayyor", f"Shablon saqlandi:\n{path}\n\n"
                                "Faylni Excel'da to'ldiring, namuna qatorini o'chiring va "
                                "\"CSV import\" orqali yuklang.")
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    def _csv_import(self, table, expected, insert_sql, row_to_tuple):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Hamma fayllar", "*.*")])
        if not path:
            return
        try:
            with open(path, newline="", encoding="utf-8-sig") as fh:
                reader = csv.DictReader(fh)
                rows = [row_to_tuple({k.strip(): v for k, v in r.items()}) for r in reader]
        except (OSError, csv.Error) as e:
            messagebox.showerror("Xato", f"CSV o'qishda xato:\n{e}")
            return
        rows = [r for r in rows if r and r[0]]
        if not rows:
            messagebox.showwarning("Bo'sh", f"Mos qatorlar topilmadi.\nKutilgan ustunlar: {', '.join(expected)}")
            return
        if not messagebox.askyesno("Import", f"{len(rows)} ta yozuv \"{table}\" jadvaliga qo'shilsinmi?"):
            return
        self.con.executemany(insert_sql, rows)
        self.con.commit()
        self.refresh_all()
        messagebox.showinfo("Tayyor", f"{len(rows)} ta yozuv qo'shildi.")

    # ================= DOMLALAR =================
    def _build_domlalar(self):
        _, bar, body = self._make_tab("  Domlalar  ")
        ttk.Button(bar, text="+ Qo'shish", command=self.dom_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", command=self.dom_edit).pack(side="left", padx=4)
        ttk.Button(bar, text="O'chirish", command=self.dom_del).pack(side="left")
        ttk.Button(bar, text="CSV import", command=self.dom_import).pack(side="left", padx=(12, 0))
        ttk.Button(bar, text="Shablon", command=self.dom_template).pack(side="left", padx=4)
        self.dom_q = tk.StringVar()
        self._add_search(bar, self.dom_q)
        self.dom_q.trace_add("write", lambda *_: self.load_domlalar())
        cols = ["ID", "F.I.Sh.", "Ilmiy unvon", "Kat.", "Stavka", "Meyor (1 st.)", "Meyor (jami)"]
        w = [50, 280, 110, 50, 70, 100, 100]
        an = {"ID": "center", "Kat.": "center", "Stavka": "e", "Meyor (1 st.)": "e", "Meyor (jami)": "e"}
        self.t_dom = self._make_tree(body, cols, w, an)
        self.t_dom.bind("<Double-1>", lambda e: self.dom_edit())

    def load_domlalar(self):
        q = (self.dom_q.get() if hasattr(self, "dom_q") else "").strip().lower()
        rows = []
        for r in self.con.execute("SELECT * FROM Domlalar ORDER BY FIO COLLATE NOCASE"):
            if q and q not in (r["FIO"] or "").lower() and q not in (r["IlmiyUnvon"] or "").lower():
                continue
            rows.append((r["DomlaID"], (
                r["DomlaID"], r["FIO"], r["IlmiyUnvon"], g(r["Kategoriya"]),
                g(r["Stavka"]), g(r["Meyor1St"]), g(meyor_jami(r["Meyor1St"], r["Stavka"])))))
        self._fill(self.t_dom, rows)

    def _dom_spec(self):
        return [("FIO", "F.I.Sh.", "text", None),
                ("IlmiyUnvon", "Ilmiy unvon", "combo", UNVON),
                ("Kategoriya", "Kategoriya", "combo", KATEGORIYA),
                ("Stavka", "Stavka", "float", None),
                ("Meyor1St", "Meyor (1 stavka)", "float", None)]

    @staticmethod
    def _dom_preview(d):
        return f"Meyor (jami) = {g(meyor_jami(d.get('Meyor1St'), d.get('Stavka')))} soat"

    def dom_add(self):
        d = FormDialog(self, "Yangi domla", self._dom_spec(),
                       {"Stavka": 1, "IlmiyUnvon": "PhD", "Kategoriya": "1"},
                       computed=self._dom_preview).result
        if d:
            self.con.execute("INSERT INTO Domlalar(FIO,IlmiyUnvon,Kategoriya,Stavka,Meyor1St) VALUES(?,?,?,?,?)",
                             (d["FIO"], d["IlmiyUnvon"], int(d["Kategoriya"] or 0), d["Stavka"], d["Meyor1St"]))
            self.con.commit()
            self.refresh_all()

    def dom_edit(self):
        i = self._selected_id(self.t_dom)
        if i is None:
            return
        r = self.con.execute("SELECT * FROM Domlalar WHERE DomlaID=?", (i,)).fetchone()
        d = FormDialog(self, "Domlani tahrirlash", self._dom_spec(), dict(r),
                       computed=self._dom_preview).result
        if d:
            self.con.execute("UPDATE Domlalar SET FIO=?,IlmiyUnvon=?,Kategoriya=?,Stavka=?,Meyor1St=? WHERE DomlaID=?",
                             (d["FIO"], d["IlmiyUnvon"], int(d["Kategoriya"] or 0), d["Stavka"], d["Meyor1St"], i))
            self.con.commit()
            self.refresh_all()

    def dom_del(self):
        i = self._selected_id(self.t_dom)
        if i is None:
            return
        used = self.con.execute("SELECT COUNT(*) c FROM Taqsimot WHERE DomlaID=?", (i,)).fetchone()["c"]
        if used:
            messagebox.showwarning("O'chirib bo'lmaydi",
                                   f"Bu domla taqsimotda {used} marta ishlatilgan. Avval o'sha yozuvlarni o'chiring.")
            return
        if messagebox.askyesno("Tasdiqlang", "Tanlangan domla o'chirilsinmi?"):
            self.con.execute("DELETE FROM Domlalar WHERE DomlaID=?", (i,))
            self.con.commit()
            self.refresh_all()

    def dom_import(self):
        self._csv_import("Domlalar", ["FIO", "IlmiyUnvon", "Kategoriya", "Stavka", "Meyor1St"],
                         "INSERT INTO Domlalar(FIO,IlmiyUnvon,Kategoriya,Stavka,Meyor1St) VALUES(?,?,?,?,?)",
                         lambda r: (r.get("FIO", "").strip(), r.get("IlmiyUnvon", "").strip() or "Darajasiz",
                                    self._i(r.get("Kategoriya")), self._f(r.get("Stavka"), 1), self._f(r.get("Meyor1St"))))

    def dom_template(self):
        self._save_template(["FIO", "IlmiyUnvon", "Kategoriya", "Stavka", "Meyor1St"],
                            ["Familiya Ism Otasi", "PhD", "1", "1.5", "360"], "domlalar_shablon.csv")

    # ================= FANLAR =================
    def _build_fanlar(self):
        _, bar, body = self._make_tab("  Fanlar  ")
        ttk.Button(bar, text="+ Qo'shish", command=self.fan_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", command=self.fan_edit).pack(side="left", padx=4)
        ttk.Button(bar, text="O'chirish", command=self.fan_del).pack(side="left")
        ttk.Button(bar, text="CSV import", command=self.fan_import).pack(side="left", padx=(12, 0))
        ttk.Button(bar, text="Shablon", command=self.fan_template).pack(side="left", padx=4)
        self.fan_q = tk.StringVar()
        self._add_search(bar, self.fan_q)
        self.fan_q.trace_add("write", lambda *_: self.load_fanlar())
        cols = ["ID", "Fan nomi", "Yo'nalish", "Ta'lim turi", "Sem.",
                "Ma'ruza", "Amaliyot", "Potok", "Guruh", "Reyting", "Jami soat"]
        w = [45, 230, 150, 90, 45, 70, 75, 55, 55, 65, 80]
        an = {c: "e" for c in ["Sem.", "Ma'ruza", "Amaliyot", "Potok", "Guruh", "Reyting", "Jami soat"]}
        an["ID"] = "center"
        self.t_fan = self._make_tree(body, cols, w, an)
        self.t_fan.bind("<Double-1>", lambda e: self.fan_edit())

    def load_fanlar(self):
        q = (self.fan_q.get() if hasattr(self, "fan_q") else "").strip().lower()
        rows = []
        for r in self.con.execute("SELECT * FROM Fanlar ORDER BY FanNomi COLLATE NOCASE"):
            if q and q not in (r["FanNomi"] or "").lower() and q not in (r["Yonalish"] or "").lower():
                continue
            _, _, jami = fan_totals(r["Maruza"], r["Amaliyot"], r["Potok"], r["Guruh"], r["Reyting"])
            rows.append((r["FanID"], (
                r["FanID"], r["FanNomi"], r["Yonalish"], r["TalimTuri"], g(r["Semestr"]),
                g(r["Maruza"]), g(r["Amaliyot"]), g(r["Potok"]), g(r["Guruh"]), g(r["Reyting"]), g(jami))))
        self._fill(self.t_fan, rows)

    def _fan_spec(self):
        return [("FanNomi", "Fan nomi", "text", None),
                ("Yonalish", "Yo'nalish", "combo", YONALISH),
                ("TalimTuri", "Ta'lim turi", "combo", TALIM_TURI),
                ("Kategoriya", "Kategoriya", "combo", KATEGORIYA),
                ("Semestr", "Semestr", "combo", SEMESTR),
                ("Maruza", "Ma'ruza (soat)", "float", None),
                ("Amaliyot", "Amaliyot (soat)", "float", None),
                ("Potok", "Potok (oqim soni)", "int", None),
                ("Guruh", "Guruh soni", "int", None),
                ("Reyting", "Reyting (soat)", "float", None)]

    @staticmethod
    def _fan_preview(d):
        mj, aj, js = fan_totals(d.get("Maruza"), d.get("Amaliyot"), d.get("Potok"), d.get("Guruh"), d.get("Reyting"))
        return f"Ma'ruza jami={g(mj)} · Amaliyot jami={g(aj)} · Jami soat={g(js)}"

    def fan_add(self):
        d = FormDialog(self, "Yangi fan", self._fan_spec(),
                       {"Potok": 1, "Guruh": 1, "Kategoriya": "1", "TalimTuri": "Kunduzgi"},
                       computed=self._fan_preview, rules=fan_rules).result
        if d:
            self.con.execute("INSERT INTO Fanlar(FanNomi,Yonalish,TalimTuri,Kategoriya,Semestr,"
                             "Maruza,Amaliyot,Potok,Guruh,Reyting) VALUES(?,?,?,?,?,?,?,?,?,?)",
                             (d["FanNomi"], d["Yonalish"], d["TalimTuri"], int(d["Kategoriya"] or 0),
                              int(d["Semestr"] or 0), d["Maruza"], d["Amaliyot"], d["Potok"], d["Guruh"], d["Reyting"]))
            self.con.commit()
            self.refresh_all()

    def fan_edit(self):
        i = self._selected_id(self.t_fan)
        if i is None:
            return
        r = self.con.execute("SELECT * FROM Fanlar WHERE FanID=?", (i,)).fetchone()
        d = FormDialog(self, "Fanni tahrirlash", self._fan_spec(), dict(r),
                       computed=self._fan_preview, rules=fan_rules).result
        if d:
            self.con.execute("UPDATE Fanlar SET FanNomi=?,Yonalish=?,TalimTuri=?,Kategoriya=?,Semestr=?,"
                             "Maruza=?,Amaliyot=?,Potok=?,Guruh=?,Reyting=? WHERE FanID=?",
                             (d["FanNomi"], d["Yonalish"], d["TalimTuri"], int(d["Kategoriya"] or 0),
                              int(d["Semestr"] or 0), d["Maruza"], d["Amaliyot"], d["Potok"], d["Guruh"], d["Reyting"], i))
            self.con.commit()
            self.refresh_all()

    def fan_del(self):
        i = self._selected_id(self.t_fan)
        if i is None:
            return
        used = self.con.execute("SELECT COUNT(*) c FROM Taqsimot WHERE FanID=?", (i,)).fetchone()["c"]
        if used:
            messagebox.showwarning("O'chirib bo'lmaydi",
                                   f"Bu fan taqsimotda {used} marta ishlatilgan. Avval o'sha yozuvlarni o'chiring.")
            return
        if messagebox.askyesno("Tasdiqlang", "Tanlangan fan o'chirilsinmi?"):
            self.con.execute("DELETE FROM Fanlar WHERE FanID=?", (i,))
            self.con.commit()
            self.refresh_all()

    def fan_import(self):
        cols = ["FanNomi", "Yonalish", "TalimTuri", "Kategoriya", "Semestr",
                "Maruza", "Amaliyot", "Potok", "Guruh", "Reyting"]
        self._csv_import("Fanlar", cols,
                         "INSERT INTO Fanlar(FanNomi,Yonalish,TalimTuri,Kategoriya,Semestr,"
                         "Maruza,Amaliyot,Potok,Guruh,Reyting) VALUES(?,?,?,?,?,?,?,?,?,?)",
                         lambda r: (r.get("FanNomi", "").strip(), r.get("Yonalish", "").strip(),
                                    r.get("TalimTuri", "").strip(), self._i(r.get("Kategoriya")),
                                    self._i(r.get("Semestr")), self._f(r.get("Maruza")), self._f(r.get("Amaliyot")),
                                    self._i(r.get("Potok"), 1), self._i(r.get("Guruh"), 1),
                                    self._f(r.get("Reyting")) if r.get("TalimTuri", "").strip() == "Masofaviy" else 0.0))

    def fan_template(self):
        self._save_template(["FanNomi", "Yonalish", "TalimTuri", "Kategoriya", "Semestr",
                             "Maruza", "Amaliyot", "Potok", "Guruh", "Reyting"],
                            ["Fan nomi", "Iqtisodiyot", "Kunduzgi", "1", "3", "30", "30", "1", "1", "0"],
                            "fanlar_shablon.csv")

    # ================= TAQSIMOT =================
    def _build_taqsimot(self):
        _, bar, body = self._make_tab("  Taqsimot  ")
        ttk.Button(bar, text="+ Qo'shish", command=self.taq_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", command=self.taq_edit).pack(side="left", padx=4)
        ttk.Button(bar, text="O'chirish", command=self.taq_del).pack(side="left")
        self.taq_q = tk.StringVar()
        self._add_search(bar, self.taq_q)
        self.taq_q.trace_add("write", lambda *_: self.load_taqsimot())
        cols = ["ID", "Domla (F.I.Sh.)", "Fan", "Turi", "Soat"]
        w = [50, 260, 320, 110, 80]
        self.t_taq = self._make_tree(body, cols, w, {"ID": "center", "Soat": "e"})
        self.t_taq.bind("<Double-1>", lambda e: self.taq_edit())

    def load_taqsimot(self):
        q = (self.taq_q.get() if hasattr(self, "taq_q") else "").strip().lower()
        sql = """SELECT t.TaqsimotID, d.FIO, f.FanNomi, t.TurSoat, t.Soat
                 FROM Taqsimot t
                 LEFT JOIN Domlalar d ON d.DomlaID=t.DomlaID
                 LEFT JOIN Fanlar f   ON f.FanID=t.FanID
                 ORDER BY d.FIO COLLATE NOCASE, f.FanNomi COLLATE NOCASE"""
        rows = []
        for r in self.con.execute(sql):
            fio, fan = r["FIO"] or "—", r["FanNomi"] or "—"
            if q and q not in fio.lower() and q not in fan.lower():
                continue
            rows.append((r["TaqsimotID"], (r["TaqsimotID"], fio, fan, r["TurSoat"], g(r["Soat"]))))
        self._fill(self.t_taq, rows)

    def _has_base_data(self):
        if not self.con.execute("SELECT 1 FROM Domlalar LIMIT 1").fetchone() or \
           not self.con.execute("SELECT 1 FROM Fanlar LIMIT 1").fetchone():
            messagebox.showwarning("Ma'lumot yetarli emas", "Avval kamida bitta domla va bitta fan kiriting.")
            return False
        return True

    def taq_add(self):
        if not self._has_base_data():
            return
        d = TaqsimotDialog(self, self.con).result
        if d:
            self.con.execute("INSERT INTO Taqsimot(DomlaID,FanID,TurSoat,Soat) VALUES(?,?,?,?)",
                             (d["DomlaID"], d["FanID"], d["TurSoat"], d["Soat"]))
            self.con.commit()
            self.refresh_all()

    def taq_edit(self):
        i = self._selected_id(self.t_taq)
        if i is None:
            return
        r = self.con.execute("SELECT * FROM Taqsimot WHERE TaqsimotID=?", (i,)).fetchone()
        d = TaqsimotDialog(self, self.con, dict(r), editing_id=r["TaqsimotID"]).result
        if d:
            self.con.execute("UPDATE Taqsimot SET DomlaID=?,FanID=?,TurSoat=?,Soat=? WHERE TaqsimotID=?",
                             (d["DomlaID"], d["FanID"], d["TurSoat"], d["Soat"], i))
            self.con.commit()
            self.refresh_all()

    def taq_del(self):
        i = self._selected_id(self.t_taq)
        if i is None:
            return
        if messagebox.askyesno("Tasdiqlang", "Tanlangan yozuv o'chirilsinmi?"):
            self.con.execute("DELETE FROM Taqsimot WHERE TaqsimotID=?", (i,))
            self.con.commit()
            self.refresh_all()

    # ================= YUKLAMA (report) =================
    def _build_yuklama(self):
        _, bar, body = self._make_tab("  Yuklama (hisobot)  ")
        ttk.Button(bar, text="Yangilash", command=self.load_yuklama).pack(side="left")
        ttk.Button(bar, text="CSV ga eksport", command=self.yuk_export).pack(side="left", padx=8)
        self.yuk_q = tk.StringVar()
        self._add_search(bar, self.yuk_q)
        self.yuk_q.trace_add("write", lambda *_: self.load_yuklama())
        self.yuk_summary = tk.StringVar()
        ttk.Label(bar, textvariable=self.yuk_summary, foreground="#444").pack(side="right")
        cols = ["F.I.Sh.", "Stavka", "Meyor (jami)", "Ma'ruza", "Amaliyot", "Reyting",
                "Jami berilgan", "Farq", "Bajarilish %"]
        w = [260, 65, 100, 80, 80, 70, 100, 80, 100]
        an = {c: "e" for c in cols if c != "F.I.Sh."}
        self.t_yuk = self._make_tree(body, cols, w, an)
        self.t_yuk.tag_configure("full", background="#d8f5dd")
        self.t_yuk.tag_configure("partial", background="#fff3cd")
        self.t_yuk.tag_configure("low", background="#f8d7da")

    def load_yuklama(self):
        q = (self.yuk_q.get() if hasattr(self, "yuk_q") else "").strip().lower()
        self.t_yuk.delete(*self.t_yuk.get_children())
        tot_norm = tot_assigned = 0
        shown = 0
        for d in workload_rows(self.con):
            tot_norm += d["norm"]
            tot_assigned += d["jami"]
            if q and q not in d["fio"].lower():
                continue
            shown += 1
            tag = "full" if d["pct"] >= 100 else "partial" if d["pct"] >= 70 else "low"
            self.t_yuk.insert("", "end", values=(
                d["fio"], g(d["stavka"]), g(d["norm"]), g(d["maruza"]), g(d["amaliyot"]),
                g(d["reyting"]), g(d["jami"]), g(d["diff"]), f"{d['pct']:.0f}%"), tags=(tag,))
        self.yuk_summary.set(f"Ko'rsatilgan: {shown}   |   Umumiy meyor: {g(tot_norm)} soat   "
                             f"|   Berilgan: {g(tot_assigned)} soat")

    def yuk_export(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")], initialfile="yuklama_hisobot.csv")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["FIO", "Stavka", "Meyor_jami", "Maruza", "Amaliyot", "Reyting",
                            "Jami_berilgan", "Farq", "Bajarilish_%"])
                for d in workload_rows(self.con):
                    w.writerow([d["fio"], g(d["stavka"]), g(d["norm"]), g(d["maruza"]), g(d["amaliyot"]),
                                g(d["reyting"]), g(d["jami"]), g(d["diff"]), f"{d['pct']:.0f}"])
            messagebox.showinfo("Tayyor", f"Hisobot saqlandi:\n{path}")
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    # ---------- refresh ----------
    def _on_tab(self):
        if self.nb.index(self.nb.select()) == 3:
            self.load_yuklama()

    def refresh_all(self):
        self.load_domlalar()
        self.load_fanlar()
        self.load_taqsimot()
        if hasattr(self, "t_yuk"):
            self.load_yuklama()


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
