#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dars taqsimoti / O'qituvchilar yuklamasi — mustaqil dastur.
Standalone teaching-load (workload) planner. Rebuilt from a Microsoft Access
application into a self-contained Python program (stdlib only: tkinter + sqlite3).

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
TALIM_TURI = ["Kunduzgi", "Kechki", "Masofaviy", "Sirtqi"]
TUR_SOAT = ["Maruza", "Amaliyot"]
SEMESTR = [str(i) for i in range(1, 13)]

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
    """Keep the database next to the executable (or the script when not frozen)."""
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


# ----- Computed values (these were the empty 'Jami' columns in Access) -----
def fan_totals(maruza, amaliyot, potok, guruh, reyting):
    maruza_jami = (maruza or 0) * (potok or 0)
    amaliyot_jami = (amaliyot or 0) * (guruh or 0)
    return maruza_jami, amaliyot_jami, maruza_jami + amaliyot_jami + (reyting or 0)


def meyor_jami(meyor1st, stavka):
    return (meyor1st or 0) * (stavka or 0)


def workload_rows(con):
    """Per-teacher assigned hours vs norm — the core of the 'Yuklama' report."""
    sql = """
    SELECT d.DomlaID, d.FIO, d.Stavka, d.Meyor1St,
           COALESCE(SUM(CASE WHEN t.TurSoat='Maruza'   THEN t.Soat END),0) AS maruza,
           COALESCE(SUM(CASE WHEN t.TurSoat='Amaliyot' THEN t.Soat END),0) AS amaliyot,
           COALESCE(SUM(t.Soat),0) AS jami
    FROM Domlalar d LEFT JOIN Taqsimot t ON t.DomlaID = d.DomlaID
    GROUP BY d.DomlaID ORDER BY d.FIO COLLATE NOCASE"""
    out = []
    for r in con.execute(sql):
        norm = meyor_jami(r["Meyor1St"], r["Stavka"])
        diff = r["jami"] - norm
        pct = (r["jami"] / norm * 100) if norm else 0
        out.append({
            "id": r["DomlaID"], "fio": r["FIO"], "stavka": r["Stavka"],
            "norm": norm, "maruza": r["maruza"], "amaliyot": r["amaliyot"],
            "jami": r["jami"], "diff": diff, "pct": pct,
        })
    return out


def g(v):
    """Format a number without trailing .0 (e.g. 30.0 -> '30')."""
    try:
        f = float(v)
        return str(int(f)) if f == int(f) else f"{f:g}"
    except (TypeError, ValueError):
        return "" if v is None else str(v)


# ===================== Add / Edit dialog =====================
class FormDialog(tk.Toplevel):
    """Generic modal form built from a field spec.
    spec item: (key, label, kind, options)
      kind in {'text','int','float','combo','readonly'}
      options: list for combos; for combo, editable=True unless key in fixed set
    """
    FIXED = {"IlmiyUnvon", "Kategoriya", "TurSoat", "Semestr"}

    def __init__(self, master, title, spec, values=None, computed=None):
        super().__init__(master)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        self.spec = spec
        self.vars = {}
        self.computed = computed  # optional callable(dict)->str for a live preview
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
            elif kind == "readonly":
                w = ttk.Entry(frm, textvariable=var, width=30, state="readonly")
            else:
                w = ttk.Entry(frm, textvariable=var, width=30)
            w.grid(row=i, column=1, sticky="ew", pady=4)
            if kind in ("int", "float"):
                var.trace_add("write", lambda *_: self._refresh_preview())

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
        # centre over parent
        x = master.winfo_rootx() + (master.winfo_width() - self.winfo_width()) // 2
        y = master.winfo_rooty() + (master.winfo_height() - self.winfo_height()) // 3
        self.geometry(f"+{max(x,0)}+{max(y,0)}")
        self.wait_window(self)

    def _collect(self):
        out = {}
        for key, label, kind, opts in self.spec:
            raw = self.vars[key].get().strip()
            if kind == "int":
                out[key] = int(float(raw)) if raw else 0
            elif kind == "float":
                out[key] = float(raw.replace(",", ".")) if raw else 0.0
            elif kind == "readonly":
                continue
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
        # required text field is the first 'text'/'combo' field
        for key, label, kind, opts in self.spec:
            if kind in ("text",) and not data.get(key):
                messagebox.showerror("Xato", f"\"{label}\" maydoni bo'sh bo'lmasligi kerak.", parent=self)
                return
            break
        self.result = data
        self.destroy()


# ===================== Main application =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x660")
        self.minsize(900, 520)
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
        st.configure("TButton", padding=(10, 4))

    # ---------- generic helpers ----------
    def _make_tab(self, title):
        f = ttk.Frame(self.nb, padding=8)
        self.nb.add(f, text=title)
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(0, 6))
        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)
        return f, bar, body

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
            tree.column(c, width=w, anchor=anchors.get(c, "w"), stretch=(c in ("F.I.Sh.", "Fan nomi")))
        tree.tag_configure("odd", background="#f5f7fa")
        return tree

    @staticmethod
    def _fill(tree, rows):
        sel = tree.selection()
        tree.delete(*tree.get_children())
        for i, (iid, vals) in enumerate(rows):
            tree.insert("", "end", iid=str(iid), values=vals,
                        tags=("odd",) if i % 2 else ())
        if sel and tree.exists(sel[0]):
            tree.selection_set(sel[0])

    @staticmethod
    def _selected_id(tree):
        s = tree.selection()
        return int(s[0]) if s else None

    # ================= DOMLALAR (teachers) =================
    def _build_domlalar(self):
        _, bar, body = self._make_tab("  Domlalar  ")
        ttk.Button(bar, text="+ Qo'shish", command=self.dom_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", command=self.dom_edit).pack(side="left", padx=4)
        ttk.Button(bar, text="O'chirish", command=self.dom_del).pack(side="left")
        ttk.Button(bar, text="CSV import", command=self.dom_import).pack(side="left", padx=(12, 0))
        cols = ["ID", "F.I.Sh.", "Ilmiy unvon", "Kat.", "Stavka", "Meyor (1 st.)", "Meyor (jami)"]
        w = [50, 280, 110, 50, 70, 100, 100]
        an = {"ID": "center", "Kat.": "center", "Stavka": "e",
              "Meyor (1 st.)": "e", "Meyor (jami)": "e"}
        self.t_dom = self._make_tree(body, cols, w, an)
        self.t_dom.bind("<Double-1>", lambda e: self.dom_edit())

    def load_domlalar(self):
        rows = []
        for r in self.con.execute("SELECT * FROM Domlalar ORDER BY FIO COLLATE NOCASE"):
            rows.append((r["DomlaID"], (
                r["DomlaID"], r["FIO"], r["IlmiyUnvon"], g(r["Kategoriya"]),
                g(r["Stavka"]), g(r["Meyor1St"]), g(meyor_jami(r["Meyor1St"], r["Stavka"])))))
        self._fill(self.t_dom, rows)

    def _dom_spec(self):
        return [
            ("FIO", "F.I.Sh.", "text", None),
            ("IlmiyUnvon", "Ilmiy unvon", "combo", UNVON),
            ("Kategoriya", "Kategoriya", "combo", KATEGORIYA),
            ("Stavka", "Stavka", "float", None),
            ("Meyor1St", "Meyor (1 stavka)", "float", None),
        ]

    @staticmethod
    def _dom_preview(d):
        return f"Meyor (jami) = {g(meyor_jami(d.get('Meyor1St'), d.get('Stavka')))} soat"

    def dom_add(self):
        d = FormDialog(self, "Yangi domla", self._dom_spec(),
                       {"Stavka": 1, "IlmiyUnvon": "PhD", "Kategoriya": "1"},
                       computed=self._dom_preview).result
        if d:
            self.con.execute(
                "INSERT INTO Domlalar(FIO,IlmiyUnvon,Kategoriya,Stavka,Meyor1St) VALUES(?,?,?,?,?)",
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
            self.con.execute(
                "UPDATE Domlalar SET FIO=?,IlmiyUnvon=?,Kategoriya=?,Stavka=?,Meyor1St=? WHERE DomlaID=?",
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
                                    f"Bu domla taqsimotda {used} marta ishlatilgan. "
                                    "Avval taqsimotdagi yozuvlarni o'chiring.")
            return
        if messagebox.askyesno("Tasdiqlang", "Tanlangan domla o'chirilsinmi?"):
            self.con.execute("DELETE FROM Domlalar WHERE DomlaID=?", (i,))
            self.con.commit()
            self.refresh_all()

    def dom_import(self):
        self._csv_import(
            "Domlalar", ["FIO", "IlmiyUnvon", "Kategoriya", "Stavka", "Meyor1St"],
            "INSERT INTO Domlalar(FIO,IlmiyUnvon,Kategoriya,Stavka,Meyor1St) VALUES(?,?,?,?,?)",
            lambda row: (row.get("FIO", "").strip(), row.get("IlmiyUnvon", "").strip() or "Darajasiz",
                         self._i(row.get("Kategoriya")), self._f(row.get("Stavka"), 1),
                         self._f(row.get("Meyor1St"))))

    # ================= FANLAR (subjects) =================
    def _build_fanlar(self):
        _, bar, body = self._make_tab("  Fanlar  ")
        ttk.Button(bar, text="+ Qo'shish", command=self.fan_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", command=self.fan_edit).pack(side="left", padx=4)
        ttk.Button(bar, text="O'chirish", command=self.fan_del).pack(side="left")
        ttk.Button(bar, text="CSV import", command=self.fan_import).pack(side="left", padx=(12, 0))
        ttk.Label(bar, text="Qidirish:").pack(side="left", padx=(16, 4))
        self.fan_q = tk.StringVar()
        e = ttk.Entry(bar, textvariable=self.fan_q, width=24)
        e.pack(side="left")
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
                g(r["Maruza"]), g(r["Amaliyot"]), g(r["Potok"]), g(r["Guruh"]),
                g(r["Reyting"]), g(jami))))
        self._fill(self.t_fan, rows)

    def _fan_spec(self):
        return [
            ("FanNomi", "Fan nomi", "text", None),
            ("Yonalish", "Yo'nalish", "combo", YONALISH),
            ("TalimTuri", "Ta'lim turi", "combo", TALIM_TURI),
            ("Kategoriya", "Kategoriya", "combo", KATEGORIYA),
            ("Semestr", "Semestr", "combo", SEMESTR),
            ("Maruza", "Ma'ruza (soat)", "float", None),
            ("Amaliyot", "Amaliyot (soat)", "float", None),
            ("Potok", "Potok (oqim soni)", "int", None),
            ("Guruh", "Guruh soni", "int", None),
            ("Reyting", "Reyting (soat)", "float", None),
        ]

    @staticmethod
    def _fan_preview(d):
        mj, aj, js = fan_totals(d.get("Maruza"), d.get("Amaliyot"),
                                d.get("Potok"), d.get("Guruh"), d.get("Reyting"))
        return f"Ma'ruza jami={g(mj)} · Amaliyot jami={g(aj)} · Jami soat={g(js)}"

    def fan_add(self):
        d = FormDialog(self, "Yangi fan", self._fan_spec(),
                       {"Potok": 1, "Guruh": 1, "Kategoriya": "1", "TalimTuri": "Kunduzgi"},
                       computed=self._fan_preview).result
        if d:
            self.con.execute(
                "INSERT INTO Fanlar(FanNomi,Yonalish,TalimTuri,Kategoriya,Semestr,"
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
                       computed=self._fan_preview).result
        if d:
            self.con.execute(
                "UPDATE Fanlar SET FanNomi=?,Yonalish=?,TalimTuri=?,Kategoriya=?,Semestr=?,"
                "Maruza=?,Amaliyot=?,Potok=?,Guruh=?,Reyting=? WHERE FanID=?",
                (d["FanNomi"], d["Yonalish"], d["TalimTuri"], int(d["Kategoriya"] or 0),
                 int(d["Semestr"] or 0), d["Maruza"], d["Amaliyot"], d["Potok"], d["Guruh"],
                 d["Reyting"], i))
            self.con.commit()
            self.refresh_all()

    def fan_del(self):
        i = self._selected_id(self.t_fan)
        if i is None:
            return
        used = self.con.execute("SELECT COUNT(*) c FROM Taqsimot WHERE FanID=?", (i,)).fetchone()["c"]
        if used:
            messagebox.showwarning("O'chirib bo'lmaydi",
                                    f"Bu fan taqsimotda {used} marta ishlatilgan. "
                                    "Avval taqsimotdagi yozuvlarni o'chiring.")
            return
        if messagebox.askyesno("Tasdiqlang", "Tanlangan fan o'chirilsinmi?"):
            self.con.execute("DELETE FROM Fanlar WHERE FanID=?", (i,))
            self.con.commit()
            self.refresh_all()

    def fan_import(self):
        cols = ["FanNomi", "Yonalish", "TalimTuri", "Kategoriya", "Semestr",
                "Maruza", "Amaliyot", "Potok", "Guruh", "Reyting"]
        self._csv_import(
            "Fanlar", cols,
            "INSERT INTO Fanlar(FanNomi,Yonalish,TalimTuri,Kategoriya,Semestr,"
            "Maruza,Amaliyot,Potok,Guruh,Reyting) VALUES(?,?,?,?,?,?,?,?,?,?)",
            lambda row: (row.get("FanNomi", "").strip(), row.get("Yonalish", "").strip(),
                         row.get("TalimTuri", "").strip(), self._i(row.get("Kategoriya")),
                         self._i(row.get("Semestr")), self._f(row.get("Maruza")),
                         self._f(row.get("Amaliyot")), self._i(row.get("Potok"), 1),
                         self._i(row.get("Guruh"), 1), self._f(row.get("Reyting"))))

    # ================= TAQSIMOT (distribution) =================
    def _build_taqsimot(self):
        _, bar, body = self._make_tab("  Taqsimot  ")
        ttk.Button(bar, text="+ Qo'shish", command=self.taq_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", command=self.taq_edit).pack(side="left", padx=4)
        ttk.Button(bar, text="O'chirish", command=self.taq_del).pack(side="left")
        cols = ["ID", "Domla (F.I.Sh.)", "Fan", "Turi", "Soat"]
        w = [50, 280, 320, 110, 80]
        self.t_taq = self._make_tree(body, cols, w, {"ID": "center", "Soat": "e"})
        self.t_taq.bind("<Double-1>", lambda e: self.taq_edit())

    def load_taqsimot(self):
        sql = """SELECT t.TaqsimotID, d.FIO, f.FanNomi, t.TurSoat, t.Soat
                 FROM Taqsimot t
                 LEFT JOIN Domlalar d ON d.DomlaID=t.DomlaID
                 LEFT JOIN Fanlar f   ON f.FanID=t.FanID
                 ORDER BY d.FIO COLLATE NOCASE, f.FanNomi COLLATE NOCASE"""
        rows = [(r["TaqsimotID"], (r["TaqsimotID"], r["FIO"] or "—",
                                   r["FanNomi"] or "—", r["TurSoat"], g(r["Soat"])))
                for r in self.con.execute(sql)]
        self._fill(self.t_taq, rows)

    def _maps(self):
        doms = self.con.execute("SELECT DomlaID,FIO FROM Domlalar ORDER BY FIO COLLATE NOCASE").fetchall()
        fans = self.con.execute("SELECT FanID,FanNomi FROM Fanlar ORDER BY FanNomi COLLATE NOCASE").fetchall()
        dlist = [f"{r['DomlaID']} — {r['FIO']}" for r in doms]
        flist = [f"{r['FanID']} — {r['FanNomi']}" for r in fans]
        return dlist, flist

    @staticmethod
    def _id_from(label):
        try:
            return int(str(label).split(" — ", 1)[0])
        except (ValueError, IndexError):
            return None

    def _taq_dialog(self, values=None):
        if not self.con.execute("SELECT 1 FROM Domlalar LIMIT 1").fetchone() or \
           not self.con.execute("SELECT 1 FROM Fanlar LIMIT 1").fetchone():
            messagebox.showwarning("Ma'lumot yetarli emas",
                                   "Avval kamida bitta domla va bitta fan kiriting.")
            return None
        dlist, flist = self._maps()
        spec = [
            ("Domla", "Domla", "combo", dlist),
            ("Fan", "Fan", "combo", flist),
            ("TurSoat", "Soat turi", "combo", TUR_SOAT),
            ("Soat", "Soat", "float", None),
        ]
        # combos here should not be free-text
        FormDialog.FIXED = FormDialog.FIXED | {"Domla", "Fan"}
        return FormDialog(self, "Taqsimot yozuvi", spec, values or {"TurSoat": "Maruza", "Soat": 30}).result

    def taq_add(self):
        d = self._taq_dialog()
        if not d:
            return
        did, fid = self._id_from(d["Domla"]), self._id_from(d["Fan"])
        if did is None or fid is None:
            messagebox.showerror("Xato", "Domla va fan tanlanishi kerak.")
            return
        self.con.execute("INSERT INTO Taqsimot(DomlaID,FanID,TurSoat,Soat) VALUES(?,?,?,?)",
                         (did, fid, d["TurSoat"], d["Soat"]))
        self.con.commit()
        self.refresh_all()

    def taq_edit(self):
        i = self._selected_id(self.t_taq)
        if i is None:
            return
        r = self.con.execute("""SELECT t.*, d.FIO, f.FanNomi FROM Taqsimot t
                                LEFT JOIN Domlalar d ON d.DomlaID=t.DomlaID
                                LEFT JOIN Fanlar f ON f.FanID=t.FanID
                                WHERE TaqsimotID=?""", (i,)).fetchone()
        pre = {"Domla": f"{r['DomlaID']} — {r['FIO']}", "Fan": f"{r['FanID']} — {r['FanNomi']}",
               "TurSoat": r["TurSoat"], "Soat": r["Soat"]}
        d = self._taq_dialog(pre)
        if not d:
            return
        self.con.execute("UPDATE Taqsimot SET DomlaID=?,FanID=?,TurSoat=?,Soat=? WHERE TaqsimotID=?",
                         (self._id_from(d["Domla"]), self._id_from(d["Fan"]), d["TurSoat"], d["Soat"], i))
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

    # ================= YUKLAMA (workload report) =================
    def _build_yuklama(self):
        _, bar, body = self._make_tab("  Yuklama (hisobot)  ")
        ttk.Button(bar, text="Yangilash", command=self.load_yuklama).pack(side="left")
        ttk.Button(bar, text="CSV ga eksport", command=self.yuk_export).pack(side="left", padx=8)
        self.yuk_summary = tk.StringVar()
        ttk.Label(bar, textvariable=self.yuk_summary, foreground="#444").pack(side="right")
        cols = ["F.I.Sh.", "Stavka", "Meyor (jami)", "Ma'ruza", "Amaliyot",
                "Jami berilgan", "Farq", "Bajarilish %"]
        w = [280, 70, 110, 90, 90, 110, 90, 110]
        an = {c: "e" for c in cols if c != "F.I.Sh."}
        self.t_yuk = self._make_tree(body, cols, w, an)
        self.t_yuk.tag_configure("full", background="#d8f5dd")     # met/exceeded norm
        self.t_yuk.tag_configure("partial", background="#fff3cd")  # 70-99%
        self.t_yuk.tag_configure("low", background="#f8d7da")      # under 70%

    def load_yuklama(self):
        data = workload_rows(self.con)
        self.t_yuk.delete(*self.t_yuk.get_children())
        tot_norm = tot_assigned = 0
        for d in data:
            tot_norm += d["norm"]
            tot_assigned += d["jami"]
            tag = "full" if d["pct"] >= 100 else "partial" if d["pct"] >= 70 else "low"
            self.t_yuk.insert("", "end", values=(
                d["fio"], g(d["stavka"]), g(d["norm"]), g(d["maruza"]), g(d["amaliyot"]),
                g(d["jami"]), g(d["diff"]), f"{d['pct']:.0f}%"), tags=(tag,))
        self.yuk_summary.set(
            f"Domlalar: {len(data)}   |   Umumiy meyor: {g(tot_norm)} soat   "
            f"|   Berilgan: {g(tot_assigned)} soat")

    def yuk_export(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV", "*.csv")],
            initialfile="yuklama_hisobot.csv")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                wtr = csv.writer(fh)
                wtr.writerow(["FIO", "Stavka", "Meyor_jami", "Maruza", "Amaliyot",
                              "Jami_berilgan", "Farq", "Bajarilish_%"])
                for d in workload_rows(self.con):
                    wtr.writerow([d["fio"], g(d["stavka"]), g(d["norm"]), g(d["maruza"]),
                                  g(d["amaliyot"]), g(d["jami"]), g(d["diff"]), f"{d['pct']:.0f}"])
            messagebox.showinfo("Tayyor", f"Hisobot saqlandi:\n{path}")
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    # ---------- shared CSV import ----------
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
        rows = [r for r in rows if r and r[0]]  # require a name in the first column
        if not rows:
            messagebox.showwarning("Bo'sh", f"Mos qatorlar topilmadi.\nKutilgan ustunlar: {', '.join(expected)}")
            return
        if not messagebox.askyesno("Import", f"{len(rows)} ta yozuv \"{table}\" jadvaliga qo'shilsinmi?"):
            return
        self.con.executemany(insert_sql, rows)
        self.con.commit()
        self.refresh_all()
        messagebox.showinfo("Tayyor", f"{len(rows)} ta yozuv qo'shildi.")

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
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
