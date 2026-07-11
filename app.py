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
import time
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
TILLAR = ["O'zbek", "Rus", "Ingliz"]             # taught language
ALL = "— hammasi —"

# ----- UI theme (professional, calm institutional look) -----
FONT = "Segoe UI"
UI = {
    "bg":          "#EDF1F7",   # window canvas (cool light slate)
    "surface":     "#FFFFFF",   # header, tables, status
    "ink":         "#1F2937",   # primary text
    "muted":       "#6B7280",   # secondary text
    "brand":       "#2563EB",   # primary accent
    "brand_dark":  "#1D4ED8",
    "brand_press": "#1A45C0",
    "brand_soft":  "#E8F0FE",   # light brand tint (chips, selection)
    "border":      "#D7DEEA",
    "head_bg":     "#F2F6FC",   # table header row
    "stripe":      "#F7F9FC",   # alternating row
    "btn":         "#E7EBF2",   # secondary button
    "btn_hover":   "#D9E0EC",
    "btn_press":   "#CCD5E5",
    "danger":      "#DC2626",
    "danger_dark": "#B91C1C",
    "danger_soft": "#FBE9E9",
    "danger_hover":"#F6D6D6",
}

SCHEMA = """
CREATE TABLE IF NOT EXISTS Domlalar(
  DomlaID INTEGER PRIMARY KEY AUTOINCREMENT,
  FIO TEXT NOT NULL, IlmiyUnvon TEXT, Kategoriya INTEGER,
  Stavka REAL DEFAULT 1, Meyor1St REAL DEFAULT 0);
CREATE TABLE IF NOT EXISTS Fanlar(
  FanID INTEGER PRIMARY KEY AUTOINCREMENT,
  FanNomi TEXT NOT NULL, Yonalish TEXT, TalimTuri TEXT, Kategoriya INTEGER,
  Semestr INTEGER, Maruza REAL DEFAULT 0, Amaliyot REAL DEFAULT 0,
  Potok INTEGER DEFAULT 1, Guruh INTEGER DEFAULT 1, Reyting REAL DEFAULT 0,
  Til TEXT DEFAULT 'O''zbek');
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
    _migrate(con)
    return con


def _migrate(con):
    """Add columns introduced after a database was first created."""
    cols = {r["name"] for r in con.execute("PRAGMA table_info(Fanlar)")}
    if "Til" not in cols:
        con.execute("ALTER TABLE Fanlar ADD COLUMN Til TEXT DEFAULT 'O''zbek'")
        con.execute("UPDATE Fanlar SET Til='O''zbek' WHERE Til IS NULL OR Til=''")
        con.commit()


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

        frm = ttk.Frame(self, padding=(22, 18))
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
            self.preview = ttk.Label(frm, text="", foreground=UI["brand_dark"], font=(FONT, 9, "bold"))
            self.preview.grid(row=len(spec), column=0, columnspan=2, sticky="w", pady=(8, 0))
            self._refresh_preview()

        btns = ttk.Frame(frm)
        btns.grid(row=len(spec) + 1, column=0, columnspan=2, sticky="e", pady=(16, 0))
        ttk.Button(btns, text="Saqlash", style="Primary.TButton", command=self._save).pack(side="right")
        ttk.Button(btns, text="Bekor qilish", style="Secondary.TButton", command=self.destroy).pack(side="right", padx=(0, 8))

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
        self.title("Taqsimot yozuvi — Professor-o'qituvchi yuklamasi")
        self.resizable(False, False)
        self.con = con
        self.result = None
        self.editing_id = editing_id      # TaqsimotID being edited (excluded from assigned-hours)

        self.domlalar = con.execute(
            "SELECT DomlaID, FIO FROM Domlalar ORDER BY FIO COLLATE NOCASE").fetchall()
        self.fanlar = con.execute(
            "SELECT FanID, FanNomi, Yonalish, TalimTuri, Semestr, Maruza, Amaliyot, Reyting, Potok, Guruh, Til "
            "FROM Fanlar ORDER BY FanNomi COLLATE NOCASE").fetchall()
        self.domla_ids = [r["DomlaID"] for r in self.domlalar]
        self.assigned_hours = self._load_assigned_hours()
        self.components = []              # parallel to cb_fan values

        pad = dict(padx=6, pady=5)
        frm = ttk.Frame(self, padding=(22, 18))
        frm.grid()
        ttk.Label(frm, text="PROFESSOR-O'QITUVCHI YUKLAMASI — DARS TAQSIMOTI",
                  style="Heading.TLabel").grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 12))

        ttk.Label(frm, text="Professor-o'qituvchi:").grid(row=1, column=0, sticky="w", **pad)
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
        ttk.Label(frm, text="Til (o'qitish tili):").grid(row=4, column=2, sticky="w", **pad)
        self.cb_til = ttk.Combobox(frm, state="readonly", width=16, values=[ALL] + self._distinct("Til"))
        self.cb_til.grid(row=4, column=3, sticky="we", **pad)

        ttk.Label(frm, text="Fan / yuklama:").grid(row=5, column=0, sticky="w", **pad)
        self.cb_fan = ttk.Combobox(frm, width=54)        # editable: type to filter
        self.cb_fan.grid(row=5, column=1, columnspan=3, sticky="we", **pad)
        ttk.Label(frm, text="(yozib qidiring → tanlang yoki Enter; Ma'ruza/Amaliyot/Reyting alohida; biriktirilgani chiqadi)",
                  foreground="#888").grid(row=6, column=1, columnspan=3, sticky="w", padx=6)

        ttk.Label(frm, text="Soat (avtomatik):").grid(row=7, column=0, sticky="w", **pad)
        self.var_soat = tk.StringVar()
        self.ent_soat = ttk.Entry(frm, textvariable=self.var_soat, width=12)
        self.ent_soat.grid(row=7, column=1, sticky="w", **pad)
        self.var_unitlbl = tk.StringVar(value="Potok soni:")
        ttk.Label(frm, textvariable=self.var_unitlbl).grid(row=7, column=2, sticky="e", **pad)
        self.var_potok = tk.StringVar()
        self.sp_potok = ttk.Spinbox(frm, from_=1, to=1, textvariable=self.var_potok, width=6,
                                    state="disabled", command=self._on_potok)
        self.sp_potok.grid(row=7, column=3, sticky="w", **pad)
        self.sp_potok.bind("<KeyRelease>", lambda e: self._on_potok())
        self.lbl_yuk = ttk.Label(frm, text="Yuklama: —", foreground=UI["brand_dark"], font=(FONT, 10, "bold"))
        self.lbl_yuk.grid(row=8, column=0, columnspan=4, sticky="w", **pad)

        btns = ttk.Frame(frm)
        btns.grid(row=9, column=0, columnspan=4, sticky="we", pady=(14, 0))
        ttk.Button(btns, text="Filtrni tozalash", style="Secondary.TButton", command=self._clear_filters).pack(side="left")
        ttk.Button(btns, text="Saqlash", style="Primary.TButton", command=self._save).pack(side="right")
        ttk.Button(btns, text="Bekor qilish", style="Secondary.TButton", command=self.destroy).pack(side="right", padx=(0, 8))

        for cb in (self.cb_yon, self.cb_talim, self.cb_sem, self.cb_til):
            cb.bind("<<ComboboxSelected>>", lambda e: self._refresh_fan())
        self.cb_fan.bind("<<ComboboxSelected>>", lambda e: self._on_fan())
        self.cb_fan.bind("<KeyRelease>", self._fan_type)
        self.cb_fan.bind("<Return>", self._fan_enter)

        self.cb_yon.set(ALL)
        self.cb_talim.set(ALL)
        self.cb_sem.set(ALL)
        self.cb_til.set(ALL)
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
        y, t, s, l = self.cb_yon.get(), self.cb_talim.get(), self.cb_sem.get(), self.cb_til.get()
        if y != ALL and (r["Yonalish"] or "") != y:
            return False
        if t != ALL and (r["TalimTuri"] or "") != t:
            return False
        if l != ALL and (r["Til"] or "").strip() != l:
            return False
        if s != ALL:
            rsem = str(int(r["Semestr"])) if r["Semestr"] else ""
            if rsem != s:
                return False
        return True

    def _build_components(self):
        """One entry per course-component, with type-specific rules:
          - Maruza  : total = Maruza × Potok. Splittable across teachers, but ONLY in whole
                      potoks: 1 potok (= Maruza hours) is taught fully by one professor and
                      can never be divided. One professor may take one or several potoks.
          - Amaliyot: total = Amaliyot × Guruh (groups), splittable across teachers (per group).
          - Reyting : total = Reyting, splittable; only for Masofaviy, eligible teachers only.
        A splittable component stays (showing remaining hours) until fully assigned."""
        comps = []
        for r in self.fanlar:
            if not self._match(r):
                continue
            specs = [("Maruza", (r["Maruza"] or 0) * (r["Potok"] or 1), True, (r["Maruza"] or 0)),
                     ("Amaliyot", (r["Amaliyot"] or 0) * (r["Guruh"] or 1), True, (r["Amaliyot"] or 0))]
            if (r["TalimTuri"] or "") == "Masofaviy":        # Reyting only for distance education
                specs.append(("Reyting", (r["Reyting"] or 0), True, (r["Reyting"] or 0)))
            for turi, total, splittable, unit in specs:
                if total <= 0:
                    continue
                done = self.assigned_hours.get((r["FanID"], turi), 0)
                remaining = total - done
                if remaining <= 0:                           # fully assigned -> hide
                    continue
                default = min(unit, remaining) if unit else remaining
                comps.append({"FanID": r["FanID"], "FanNomi": r["FanNomi"], "TurSoat": turi,
                              "total": total, "remaining": remaining, "default": default,
                              "unit": unit, "splittable": splittable, "row": r})
        return comps

    def _comp_label(self, c):
        r = c["row"]
        til = (r["Til"] or "").strip()
        name = f'{c["FanNomi"]} ({til.lower()})' if til else c["FanNomi"]
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
        if c["TurSoat"] == "Maruza" and (r["Potok"] or 1) > 1 and c["unit"]:
            hrs += f' · 1 potok = {g(c["unit"])} soat'
        elif c["TurSoat"] == "Amaliyot" and (r["Guruh"] or 1) > 1 and c["unit"]:
            hrs += f' · 1 guruh = {g(c["unit"])} soat'
        return f'{name} — {c["TurSoat"]} ({hrs}){tail}'

    def _refresh_fan(self):
        self.components = self._build_components()
        self._shown = list(self.components)
        self.cb_fan["values"] = [self._comp_label(c) for c in self._shown]
        self.cb_fan.set("")
        self.var_soat.set("")
        self.ent_soat.config(state="normal")
        self.var_potok.set("")
        self.sp_potok.config(state="disabled")
        self.lbl_yuk.config(text="Yuklama: —")

    def _current(self):
        shown = getattr(self, "_shown", self.components)
        i = self.cb_fan.current()
        if i < 0 or i >= len(shown):
            return None
        return shown[i]

    def _fan_type(self, event):
        nav = {"Up", "Down", "Return", "Escape", "Left", "Right", "Tab",
               "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Home", "End"}
        if event.keysym in nav:
            return
        typed = self.cb_fan.get().strip().lower()
        if typed:
            self._shown = [c for c in self.components if typed in self._comp_label(c).lower()]
        else:
            self._shown = list(self.components)
        self.cb_fan["values"] = [self._comp_label(c) for c in self._shown]

    def _fan_enter(self, event):
        if getattr(self, "_shown", None) and self.cb_fan.current() < 0:
            self.cb_fan.current(0)
        self._on_fan()
        return "break"

    def _max_units(self, c):
        """How many whole units (potok for Maruza, guruh for Amaliyot) remain."""
        return max(1, int(round(c["remaining"] / c["unit"]))) if c.get("unit") else 1

    @staticmethod
    def _unit_word(c):
        return "potok" if c["TurSoat"] == "Maruza" else "guruh"

    def _on_fan(self):
        c = self._current()
        if not c:
            return
        self.var_soat.set(g(c["default"]))
        self.ent_soat.config(state="disabled")        # hours are set automatically, not editable
        if c["TurSoat"] in ("Maruza", "Amaliyot") and c.get("unit"):
            word = self._unit_word(c)
            self.var_unitlbl.set(f"{word.capitalize()} soni:")
            mx = self._max_units(c)
            self.sp_potok.config(state="normal", from_=1, to=mx)
            self.var_potok.set("1")
            self.lbl_yuk.config(text=f"{c['TurSoat']}: 1 {word} = {g(c['unit'])} soat (bo'linmaydi), "
                                     f"qoldi {mx} {word} ({g(c['remaining'])} soat)")
        else:
            self.var_potok.set("")
            self.sp_potok.config(state="disabled")
            if c["splittable"]:
                self.lbl_yuk.config(text=f"{c['TurSoat']}: jami {g(c['total'])} soat, qoldi {g(c['remaining'])} soat")
            else:
                self.lbl_yuk.config(text=f"{c['TurSoat']}: {g(c['total'])} soat — bitta professor")

    def _on_potok(self):
        """Unit count changed -> Soat = count × per-unit hours."""
        c = self._current()
        if not c or c["TurSoat"] not in ("Maruza", "Amaliyot") or not c.get("unit"):
            return
        try:
            n = int(float(self.var_potok.get().strip()))
        except (ValueError, AttributeError):
            return
        n = max(1, min(n, self._max_units(c)))
        self.var_soat.set(g(n * c["unit"]))

    def _clear_filters(self):
        self.cb_yon.set(ALL)
        self.cb_talim.set(ALL)
        self.cb_sem.set(ALL)
        self.cb_til.set(ALL)
        self._refresh_fan()

    def _prefill(self, v):
        if v.get("DomlaID") in self.domla_ids:
            self.cb_domla.current(self.domla_ids.index(v["DomlaID"]))
        fan = next((r for r in self.fanlar if r["FanID"] == v.get("FanID")), None)
        if fan:
            self.cb_yon.set(fan["Yonalish"] or ALL)
            self.cb_talim.set(fan["TalimTuri"] or ALL)
            self.cb_sem.set(str(int(fan["Semestr"])) if fan["Semestr"] else ALL)
            self.cb_til.set((fan["Til"] or "").strip() or ALL)
            self._refresh_fan()
            for idx, c in enumerate(self.components):
                if c["FanID"] == v.get("FanID") and c["TurSoat"] == v.get("TurSoat"):
                    self.cb_fan.current(idx)
                    self._on_fan()
                    break
        if v.get("Soat") is not None:
            self.var_soat.set(g(v["Soat"]))
            c = self._current()
            if c and c["TurSoat"] in ("Maruza", "Amaliyot") and c.get("unit"):
                try:
                    self.var_potok.set(str(max(1, int(round(float(v["Soat"]) / c["unit"])))))
                except (ValueError, ZeroDivisionError):
                    pass

    def _save(self):
        di = self.cb_domla.current()
        if di < 0:
            messagebox.showerror("Xato", "Professor-o'qituvchi tanlanishi shart.", parent=self)
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
        if c["TurSoat"] in ("Maruza", "Amaliyot") and c.get("unit"):
            unit, word = c["unit"], self._unit_word(c)
            n = soat / unit
            if abs(n - round(n)) > 1e-6 or round(n) < 1:
                messagebox.showerror("Xato",
                    f"{c['TurSoat']} faqat butun {word} bo'yicha biriktiriladi.\n"
                    f"1 {word} = {g(unit)} soat (bo'linmaydi). Soat {g(unit)} ga karrali bo'lishi kerak "
                    f"(masalan {g(unit)}, {g(2 * unit)}, ...).", parent=self)
                return
            if soat > c["remaining"] + 1e-9:
                messagebox.showerror("Xato",
                    f"Qolgan {word}lar: {self._max_units(c)} ta ({g(c['remaining'])} soat).\n"
                    f"Siz {g(soat)} soat ({int(round(n))} {word}) kiritdingiz — ortiqcha biriktirish mumkin emas.",
                    parent=self)
                return
        elif soat > c["remaining"] + 1e-9:
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
                    "Reyting faqat shu fanning Ma'ruza yoki Amaliyotini o'qitadigan professor-o'qituvchiga "
                    "biriktiriladi.\nAvval o'sha professor-o'qituvchiga shu fandan Ma'ruza yoki Amaliyot biriktiring.",
                    parent=self)
                return
        # Meyor (jami) check: over-assignment is ALLOWED, but the user is notified.
        dom_id = self.domla_ids[di]
        dom = self.con.execute("SELECT FIO, Stavka, Meyor1St FROM Domlalar WHERE DomlaID=?", (dom_id,)).fetchone()
        norm = meyor_jami(dom["Meyor1St"], dom["Stavka"])
        if self.editing_id:
            cur = self.con.execute("SELECT COALESCE(SUM(Soat),0) s FROM Taqsimot "
                                   "WHERE DomlaID=? AND TaqsimotID<>?", (dom_id, self.editing_id)).fetchone()["s"]
        else:
            cur = self.con.execute("SELECT COALESCE(SUM(Soat),0) s FROM Taqsimot "
                                   "WHERE DomlaID=?", (dom_id,)).fetchone()["s"]
        new_total = cur + soat
        if norm and new_total > norm + 1e-9:
            pct = new_total / norm * 100
            messagebox.showwarning("Meyordan ortiq",
                f"Diqqat: {short_name(dom['FIO'])} uchun meyor (jami) — {g(norm)} soat.\n"
                f"Bu biriktirishdan so'ng jami yuklama {g(new_total)} soat bo'ladi "
                f"({g(new_total - norm)} soat meyordan ortiq, bajarilish {pct:.0f}%).\n\n"
                f"Yozuv baribir saqlanadi — hisobotda bu «Meyordan ortiq» deb belgilanadi.",
                parent=self)
        self.result = {"DomlaID": dom_id, "FanID": c["FanID"],
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

        tk.Frame(self, height=3, background=UI["brand"]).pack(fill="x")
        header = ttk.Frame(self, style="Header.TFrame", padding=(20, 13))
        header.pack(fill="x")
        left = ttk.Frame(header, style="Header.TFrame")
        left.pack(side="left")
        ttk.Label(left, text="Dars taqsimoti", style="Title.TLabel").pack(anchor="w")
        ttk.Label(left, text="O'qituvchilar yuklamasini rejalashtirish va hisobot",
                  style="Subtitle.TLabel").pack(anchor="w")
        right = ttk.Frame(header, style="Header.TFrame")
        right.pack(side="right")
        hbtns = ttk.Frame(right, style="Header.TFrame")
        hbtns.pack(anchor="e")
        ttk.Button(hbtns, text="↩ Ortga qaytarish (Ctrl+Z)", style="Ghost.TButton",
                   command=self.show_undo).pack(side="left")
        ttk.Button(hbtns, text="Yordam markazi", style="Ghost.TButton",
                   command=self.show_help).pack(side="left", padx=(8, 0))
        ttk.Label(right, text="Developed by Zaxid Raximov",
                  style="Dev.TLabel").pack(anchor="e", pady=(6, 0))
        tk.Frame(self, height=1, background=UI["border"]).pack(fill="x")

        self.undo_stack = []          # deleted-rows history (session only)
        self.bind("<Control-z>", self.undo_last)
        self.bind("<Control-Z>", self.undo_last)

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        self._build_domlalar()
        self._build_fanlar()
        self._build_taqsimot()
        self._build_yukfan()
        self._build_yuklama()
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self._on_tab())

        status = ttk.Frame(self, style="Status.TFrame", padding=(14, 6))
        status.pack(fill="x", side="bottom")
        self.status = tk.StringVar(value=f"Ma'lumotlar bazasi: {db_path()}")
        ttk.Label(status, textvariable=self.status, style="Status.TLabel", anchor="w").pack(side="left")
        tk.Frame(self, height=1, background=UI["border"]).pack(fill="x", side="bottom")
        self.refresh_all()

    def _style(self):
        st = ttk.Style()
        u = UI
        self.configure(background=u["bg"])
        # dropdown list (combobox popup) colors
        self.option_add("*TCombobox*Listbox.background", u["surface"])
        self.option_add("*TCombobox*Listbox.foreground", u["ink"])
        self.option_add("*TCombobox*Listbox.selectBackground", u["brand"])
        self.option_add("*TCombobox*Listbox.selectForeground", "#FFFFFF")
        self.option_add("*TCombobox*Listbox.font", (FONT, 10))

        # base
        st.configure(".", background=u["bg"], foreground=u["ink"], font=(FONT, 10))
        st.configure("TFrame", background=u["bg"])
        st.configure("TLabel", background=u["bg"], foreground=u["ink"], font=(FONT, 10))

        # surfaces / text
        st.configure("Header.TFrame", background=u["surface"])
        st.configure("Status.TFrame", background=u["surface"])
        st.configure("Card.TFrame", background=u["surface"])
        st.configure("Title.TLabel", background=u["surface"], foreground=u["ink"], font=(FONT, 16, "bold"))
        st.configure("Subtitle.TLabel", background=u["surface"], foreground=u["muted"], font=(FONT, 9))
        st.configure("Status.TLabel", background=u["surface"], foreground=u["muted"], font=(FONT, 9))
        st.configure("Credit.TLabel", background=u["surface"], foreground="#9AA3B2",
                     font=(FONT, 8, "italic"))
        st.configure("Dev.TLabel", background=u["surface"], foreground=u["muted"],
                     font=(FONT, 14, "italic"))
        st.configure("Heading.TLabel", background=u["bg"], foreground=u["ink"], font=(FONT, 14, "bold"))
        st.configure("Muted.TLabel", background=u["bg"], foreground=u["muted"], font=(FONT, 9))
        st.configure("Chip.TLabel", background=u["brand_soft"], foreground=u["brand_dark"],
                     font=(FONT, 10, "bold"), padding=(11, 5))

        # buttons – secondary is the default toolbar look
        def button(name, bg, fg, hover, press, bold=False):
            st.configure(name, font=(FONT, 10, "bold" if bold else "normal"),
                         padding=(13, 6), relief="flat", background=bg, foreground=fg,
                         bordercolor=bg, lightcolor=bg, darkcolor=bg, focuscolor=bg)
            st.map(name, background=[("active", hover), ("pressed", press), ("disabled", "#EDEFF3")],
                   foreground=[("disabled", "#A7AEBA")])
        button("TButton", u["btn"], u["ink"], u["btn_hover"], u["btn_press"])
        button("Secondary.TButton", u["btn"], u["ink"], u["btn_hover"], u["btn_press"])
        button("Primary.TButton", u["brand"], "#FFFFFF", u["brand_dark"], u["brand_press"], bold=True)
        button("Danger.TButton", u["danger_soft"], u["danger_dark"], u["danger_hover"], "#F0C4C4")
        st.configure("Ghost.TButton", font=(FONT, 10), padding=(12, 6), relief="flat",
                     background=u["surface"], foreground=u["brand_dark"],
                     bordercolor=u["surface"], lightcolor=u["surface"], darkcolor=u["surface"],
                     focuscolor=u["surface"])
        st.map("Ghost.TButton", background=[("active", u["brand_soft"]), ("pressed", u["brand_soft"])])

        # notebook
        st.configure("TNotebook", background=u["bg"], borderwidth=0, tabmargins=(4, 6, 4, 0))
        st.configure("TNotebook.Tab", font=(FONT, 10), padding=(18, 9),
                     background="#DFE5EE", foreground=u["muted"], borderwidth=0)
        st.map("TNotebook.Tab",
               background=[("selected", u["surface"]), ("active", "#E9EDF4")],
               foreground=[("selected", u["brand_dark"]), ("active", u["ink"])],
               font=[("selected", (FONT, 10, "bold"))])

        # treeview
        st.configure("Treeview", background=u["surface"], fieldbackground=u["surface"],
                     foreground=u["ink"], rowheight=30, borderwidth=0, font=(FONT, 10))
        st.configure("Treeview.Heading", background=u["head_bg"], foreground=u["ink"],
                     font=(FONT, 10, "bold"), relief="flat", padding=(8, 9), borderwidth=0)
        st.map("Treeview.Heading", background=[("active", "#E7EEF8")])
        st.map("Treeview", background=[("selected", u["brand_soft"])],
               foreground=[("selected", u["ink"])])

        # entry / combobox
        st.configure("TEntry", fieldbackground=u["surface"], foreground=u["ink"], padding=6,
                     bordercolor=u["border"], lightcolor=u["border"], darkcolor=u["border"],
                     insertwidth=1, selectbackground=u["brand_soft"], selectforeground=u["ink"])
        st.map("TEntry", bordercolor=[("focus", u["brand"])],
               lightcolor=[("focus", u["brand"])], darkcolor=[("focus", u["brand"])])
        st.configure("TCombobox", fieldbackground=u["surface"], foreground=u["ink"], padding=6,
                     bordercolor=u["border"], lightcolor=u["border"], darkcolor=u["border"], arrowsize=15)
        st.map("TCombobox",
               fieldbackground=[("readonly", u["surface"]), ("disabled", "#F1F3F8")],
               foreground=[("disabled", "#9AA3B2")],
               bordercolor=[("focus", u["brand"])],
               lightcolor=[("focus", u["brand"])], darkcolor=[("focus", u["brand"])])

        # scrollbars – slim, neutral
        for s in ("Vertical.TScrollbar", "Horizontal.TScrollbar"):
            st.configure(s, background="#C8D0DD", troughcolor=u["bg"], borderwidth=0,
                         arrowcolor=u["muted"], relief="flat")
            st.map(s, background=[("active", "#B4BFD0")])

    # ---------- undo (ortga qaytarish) ----------
    _RESTORE_ORDER = {"Domlalar": 0, "Fanlar": 1, "Taqsimot": 2}

    def _record_delete(self, desc, items):
        """Save one delete operation to the undo stack.
        items: list of (table_name, row_dict) captured BEFORE deletion.
        One bulk delete (e.g. 50 rows at once) = one operation."""
        if not items:
            return
        self.undo_stack.append({"time": time.strftime("%H:%M:%S"), "desc": desc, "items": items})
        if len(self.undo_stack) > 100:
            self.undo_stack.pop(0)

    @staticmethod
    def _snap(con, table, id_col, ids):
        """Snapshot full rows of `table` for given ids as (table, dict) items."""
        if not ids:
            return []
        qm = ",".join("?" * len(ids))
        return [(table, dict(r)) for r in
                con.execute(f"SELECT * FROM {table} WHERE {id_col} IN ({qm})", list(ids))]

    def _restore_op(self, op):
        """Re-insert deleted rows with their original IDs. Parents (Domlalar, Fanlar)
        are restored before Taqsimot so foreign keys stay valid."""
        ok = fail = 0
        for table, row in sorted(op["items"], key=lambda x: self._RESTORE_ORDER.get(x[0], 9)):
            cols = list(row.keys())
            try:
                self.con.execute(
                    f"INSERT INTO {table}({','.join(cols)}) VALUES({','.join('?' * len(cols))})",
                    [row[c] for c in cols])
                ok += 1
            except sqlite3.IntegrityError:
                fail += 1      # e.g. parent row deleted in another operation and not yet restored
        self.con.commit()
        self.refresh_all()
        return ok, fail

    def undo_last(self, event=None):
        """Ctrl+Z — restore the most recent delete operation."""
        if not self.undo_stack:
            messagebox.showinfo("Ortga qaytarish", "Qaytariladigan amal yo'q.")
            return
        op = self.undo_stack.pop()
        ok, fail = self._restore_op(op)
        msg = f"Qaytarildi: {op['desc']}\n{ok} ta yozuv tiklandi."
        if fail:
            msg += (f"\n{fail} ta yozuv tiklanmadi — unga bog'liq o'qituvchi/fan boshqa amalda "
                    "o'chirilgan. Avval o'sha amalni qaytaring.")
        messagebox.showinfo("Ortga qaytarish", msg)

    def show_undo(self):
        """Dialog listing delete operations (latest first); the user picks which to restore."""
        if not self.undo_stack:
            messagebox.showinfo("Ortga qaytarish", "Qaytariladigan amal yo'q.\n\n"
                                "Bu ro'yxatga o'chirilgan yozuvlar tushadi (dastur yopilguncha saqlanadi).")
            return
        win = tk.Toplevel(self)
        win.title("Ortga qaytarish — o'chirilgan amallar")
        win.geometry("640x380")
        win.transient(self)
        win.configure(background=UI["bg"])
        try:
            win.geometry(f"+{self.winfo_rootx() + 120}+{self.winfo_rooty() + 90}")
        except Exception:
            pass
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Qaysi amalni ortga qaytarish kerak? (eng oxirgisi tepada)",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 8))
        cols = ["#", "Vaqt", "Amal", "Yozuvlar"]
        tree = self._make_tree(frm, cols, [40, 70, 380, 80],
                               {"#": "center", "Vaqt": "center", "Yozuvlar": "e"}, stretch=("Amal",))
        # latest first; iid = index in undo_stack
        for pos, idx in enumerate(range(len(self.undo_stack) - 1, -1, -1), 1):
            op = self.undo_stack[idx]
            tree.insert("", "end", iid=str(idx),
                        values=(pos, op["time"], op["desc"], len(op["items"])),
                        tags=("odd",) if pos % 2 == 0 else ())

        def restore(_e=None):
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("Ortga qaytarish", "Avval amalni tanlang.", parent=win)
                return
            idx = int(sel[0])
            op = self.undo_stack.pop(idx)
            ok, fail = self._restore_op(op)
            msg = f"Qaytarildi: {op['desc']}\n{ok} ta yozuv tiklandi."
            if fail:
                msg += (f"\n{fail} ta yozuv tiklanmadi — unga bog'liq o'qituvchi/fan boshqa amalda "
                        "o'chirilgan. Avval o'sha amalni qaytaring.")
            messagebox.showinfo("Ortga qaytarish", msg, parent=win)
            win.destroy()
            if self.undo_stack:
                self.show_undo()

        tree.bind("<Double-1>", restore)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Tanlanganini qaytarish", style="Primary.TButton",
                   command=restore).pack(side="right")
        ttk.Button(btns, text="Yopish", style="Secondary.TButton",
                   command=win.destroy).pack(side="right", padx=(0, 8))
        win.bind("<Escape>", lambda e: win.destroy())
        win.grab_set()

    # ---------- generic helpers ----------
    def _make_tab(self, title):
        f = ttk.Frame(self.nb, padding=(14, 14))
        self.nb.add(f, text=title)
        bar = ttk.Frame(f)
        bar.pack(fill="x", pady=(0, 10))
        body = ttk.Frame(f)
        body.pack(fill="both", expand=True)
        return f, bar, body

    def _add_search(self, bar, var):
        ttk.Label(bar, text="Qidirish:", style="Muted.TLabel").pack(side="left", padx=(16, 6))
        e = ttk.Entry(bar, textvariable=var, width=26)
        e.pack(side="left")

    def _make_tree(self, parent, columns, widths, anchors=None, stretch=None):
        wrap = tk.Frame(parent, background=UI["surface"], highlightthickness=1,
                        highlightbackground=UI["border"], highlightcolor=UI["border"], bd=0)
        wrap.pack(fill="both", expand=True)
        tree = ttk.Treeview(wrap, columns=columns, show="headings", selectmode="extended")
        vs = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        hs = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vs.grid(row=0, column=1, sticky="ns")
        hs.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        anchors = anchors or {}
        if stretch is None:
            stretch = ("F.I.Sh.", "Fan nomi", "Professor-o'qituvchi (F.I.Sh.)", "Fan")
        for c, w in zip(columns, widths):
            tree.heading(c, text=c, command=lambda col=c: self._sort_tree(tree, col))
            tree.column(c, width=w, minwidth=40, anchor=anchors.get(c, "w"),
                        stretch=(c in stretch))
        tree.tag_configure("odd", background=UI["stripe"])
        tree._columns = list(columns)
        tree._sort = {"col": None, "asc": True}
        # Ctrl+A — select all rows (Shift/Ctrl+click also work thanks to selectmode="extended")
        def _select_all(event, t=tree):
            t.selection_set(t.get_children(""))
            return "break"
        tree.bind("<Control-a>", _select_all)
        tree.bind("<Control-A>", _select_all)
        return tree

    def _sort_tree(self, tree, col):
        asc = not (tree._sort.get("col") == col and tree._sort.get("asc"))
        tree._sort = {"col": col, "asc": asc}
        uses_odd = any("odd" in tree.item(i, "tags") for i in tree.get_children(""))

        def key(iid):
            raw = tree.set(iid, col)
            num = (raw or "").strip().replace("%", "").replace(",", ".")
            try:
                return (0, float(num))
            except ValueError:
                return (1, (raw or "").lower())

        for idx, iid in enumerate(sorted(tree.get_children(""), key=key, reverse=not asc)):
            tree.move(iid, "", idx)
        for c in tree._columns:
            arrow = "  ▲" if (c == col and asc) else "  ▼" if c == col else ""
            tree.heading(c, text=c + arrow)
        if uses_odd:
            for idx, iid in enumerate(tree.get_children("")):
                tags = [t for t in tree.item(iid, "tags") if t != "odd"]
                if idx % 2:
                    tags.append("odd")
                tree.item(iid, tags=tags)

    @staticmethod
    def _fill(tree, rows):
        sel = tree.selection()
        tree.delete(*tree.get_children())
        for i, (iid, vals) in enumerate(rows):
            tree.insert("", "end", iid=str(iid), values=vals, tags=("odd",) if i % 2 else ())
        keep = [s for s in sel if tree.exists(s)]
        if keep:
            tree.selection_set(keep)
        if hasattr(tree, "_columns"):          # data reloaded -> back to default order
            tree._sort = {"col": None, "asc": True}
            for c in tree._columns:
                tree.heading(c, text=c)

    @staticmethod
    def _q_match(q, *values):
        """True if every space-separated search word occurs in at least one value
        (search by ID, name, direction, type, semester, language, hours...)."""
        if not q:
            return True
        hay = " | ".join(str(v).lower() for v in values if v is not None)
        return all(w in hay for w in q.split())

    @staticmethod
    def _selected_id(tree):
        s = tree.selection()
        return int(s[0]) if s else None

    @staticmethod
    def _selected_ids(tree):
        return [int(s) for s in tree.selection()]

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

    def _csv_import(self, table, expected, insert_sql, row_to_tuple, preprocess=None):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv"), ("Hamma fayllar", "*.*")])
        if not path:
            return
        try:
            with open(path, newline="", encoding="utf-8-sig") as fh:
                raw_rows = [{k.strip(): v for k, v in r.items()} for r in csv.DictReader(fh)]
        except (OSError, csv.Error) as e:
            messagebox.showerror("Xato", f"CSV o'qishda xato:\n{e}")
            return
        note = ""
        if preprocess:
            raw_rows, note = preprocess(raw_rows)
        rows = [t for t in (row_to_tuple(r) for r in raw_rows) if t and t[0]]
        if not rows:
            messagebox.showwarning("Bo'sh", f"Mos qatorlar topilmadi.\nKutilgan ustunlar: {', '.join(expected)}")
            return
        msg = f"{len(rows)} ta yozuv \"{table}\" jadvaliga qo'shilsinmi?"
        if note:
            msg += f"\n{note}"
        if not messagebox.askyesno("Import", msg):
            return
        self.con.executemany(insert_sql, rows)
        self.con.commit()
        self.refresh_all()
        done = f"{len(rows)} ta yozuv qo'shildi."
        if note:
            done += f"\n{note}"
        messagebox.showinfo("Tayyor", done)

    # ================= DOMLALAR =================
    def _build_domlalar(self):
        _, bar, body = self._make_tab("  Professor-O'qituvchilar  ")
        ttk.Button(bar, text="+ Qo'shish", style="Primary.TButton", command=self.dom_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", style="Secondary.TButton", command=self.dom_edit).pack(side="left", padx=6)
        ttk.Button(bar, text="O'chirish", style="Danger.TButton", command=self.dom_del).pack(side="left")
        ttk.Button(bar, text="CSV import", style="Secondary.TButton", command=self.dom_import).pack(side="left", padx=(16, 0))
        ttk.Button(bar, text="Shablon", style="Secondary.TButton", command=self.dom_template).pack(side="left", padx=6)
        self.dom_total = tk.StringVar(value="Jami meyor: 0")
        ttk.Label(bar, textvariable=self.dom_total, style="Chip.TLabel").pack(side="right", padx=(8, 2))
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
        total = 0
        for r in self.con.execute("SELECT * FROM Domlalar ORDER BY FIO COLLATE NOCASE"):
            if not self._q_match(q, r["DomlaID"], r["FIO"], r["IlmiyUnvon"],
                                 g(r["Kategoriya"]), g(r["Stavka"]), g(r["Meyor1St"])):
                continue
            mj = meyor_jami(r["Meyor1St"], r["Stavka"])
            total += mj or 0
            rows.append((r["DomlaID"], (
                r["DomlaID"], r["FIO"], r["IlmiyUnvon"], g(r["Kategoriya"]),
                g(r["Stavka"]), g(r["Meyor1St"]), g(mj))))
        self._fill(self.t_dom, rows)
        if hasattr(self, "dom_total"):
            self.dom_total.set(f"Jami meyor: {g(total)}")

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
        d = FormDialog(self, "Yangi professor-o'qituvchi", self._dom_spec(),
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
        d = FormDialog(self, "Professor-o'qituvchini tahrirlash", self._dom_spec(), dict(r),
                       computed=self._dom_preview).result
        if d:
            self.con.execute("UPDATE Domlalar SET FIO=?,IlmiyUnvon=?,Kategoriya=?,Stavka=?,Meyor1St=? WHERE DomlaID=?",
                             (d["FIO"], d["IlmiyUnvon"], int(d["Kategoriya"] or 0), d["Stavka"], d["Meyor1St"], i))
            self.con.commit()
            self.refresh_all()

    def dom_del(self):
        ids = self._selected_ids(self.t_dom)
        if not ids:
            return
        used, free = [], []
        for i in ids:
            r = self.con.execute("SELECT FIO, (SELECT COUNT(*) FROM Taqsimot WHERE DomlaID=?) c "
                                 "FROM Domlalar WHERE DomlaID=?", (i, i)).fetchone()
            if not r:
                continue
            (used if r["c"] else free).append((i, r["FIO"]))
        if used and not free:
            names = "\n".join(f"• {short_name(f)}" for _, f in used[:10])
            messagebox.showwarning("O'chirib bo'lmaydi",
                f"Tanlangan professor-o'qituvchi(lar) taqsimotda ishlatilgan:\n{names}\n\n"
                "Avval ularning taqsimot yozuvlarini o'chiring.")
            return
        msg = (f"{len(free)} ta professor-o'qituvchi o'chirilsinmi?"
               if len(free) > 1 else "Tanlangan professor-o'qituvchi o'chirilsinmi?")
        if used:
            names = "\n".join(f"• {short_name(f)}" for _, f in used[:10])
            msg += (f"\n\nQuyidagi {len(used)} tasi taqsimotda ishlatilgani uchun O'CHIRILMAYDI:\n{names}")
        if not messagebox.askyesno("Tasdiqlang", msg):
            return
        free_ids = [i for i, _ in free]
        names = ", ".join(short_name(f) for _, f in free[:3]) + ("..." if len(free) > 3 else "")
        self._record_delete(f"Professor-o'qituvchilar: {len(free)} ta o'chirildi ({names})",
                            self._snap(self.con, "Domlalar", "DomlaID", free_ids))
        self.con.executemany("DELETE FROM Domlalar WHERE DomlaID=?", [(i,) for i in free_ids])
        self.con.commit()
        self.refresh_all()

    def dom_import(self):
        self._csv_import("Professor-o'qituvchilar", ["FIO", "IlmiyUnvon", "Kategoriya", "Stavka", "Meyor1St"],
                         "INSERT INTO Domlalar(FIO,IlmiyUnvon,Kategoriya,Stavka,Meyor1St) VALUES(?,?,?,?,?)",
                         lambda r: (r.get("FIO", "").strip(), r.get("IlmiyUnvon", "").strip() or "Darajasiz",
                                    self._i(r.get("Kategoriya")), self._f(r.get("Stavka"), 1), self._f(r.get("Meyor1St"))))

    def dom_template(self):
        self._save_template(["FIO", "IlmiyUnvon", "Kategoriya", "Stavka", "Meyor1St"],
                            ["Familiya Ism Otasi", "PhD", "1", "1.5", "360"], "professor_oqituvchilar_shablon.csv")

    # ================= FANLAR =================
    def _build_fanlar(self):
        _, bar, body = self._make_tab("  Fanlar yuklamasi  ")
        ttk.Button(bar, text="+ Qo'shish", style="Primary.TButton", command=self.fan_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", style="Secondary.TButton", command=self.fan_edit).pack(side="left", padx=6)
        ttk.Button(bar, text="O'chirish", style="Danger.TButton", command=self.fan_del).pack(side="left")
        ttk.Button(bar, text="CSV import", style="Secondary.TButton", command=self.fan_import).pack(side="left", padx=(16, 0))
        ttk.Button(bar, text="Shablon", style="Secondary.TButton", command=self.fan_template).pack(side="left", padx=6)
        ttk.Button(bar, text="Takror tozalash", style="Secondary.TButton", command=self.fan_dedup).pack(side="left")
        self.fan_total = tk.StringVar(value="Jami soatlar: 0")
        ttk.Label(bar, textvariable=self.fan_total, style="Chip.TLabel").pack(side="right", padx=(8, 2))
        self.fan_q = tk.StringVar()
        self._add_search(bar, self.fan_q)
        self.fan_q.trace_add("write", lambda *_: self.load_fanlar())
        cols = ["ID", "Fan nomi", "Yo'nalish", "Ta'lim turi", "Til", "Sem.",
                "Ma'ruza", "Amaliyot", "Potok", "Guruh", "Reyting", "Jami soat"]
        w = [45, 215, 135, 85, 65, 45, 65, 70, 50, 50, 60, 75]
        an = {c: "e" for c in ["Sem.", "Ma'ruza", "Amaliyot", "Potok", "Guruh", "Reyting", "Jami soat"]}
        an["ID"] = "center"
        self.t_fan = self._make_tree(body, cols, w, an)
        self.t_fan.bind("<Double-1>", lambda e: self.fan_edit())

    def load_fanlar(self):
        q = (self.fan_q.get() if hasattr(self, "fan_q") else "").strip().lower()
        rows = []
        total = 0
        for r in self.con.execute("SELECT * FROM Fanlar ORDER BY FanID"):
            if not self._q_match(q, r["FanID"], r["FanNomi"], r["Yonalish"], r["TalimTuri"],
                                 r["Til"], g(r["Semestr"]), g(r["Kategoriya"])):
                continue
            _, _, jami = fan_totals(r["Maruza"], r["Amaliyot"], r["Potok"], r["Guruh"], r["Reyting"])
            total += jami or 0
            rows.append((r["FanID"], (
                r["FanID"], r["FanNomi"], r["Yonalish"], r["TalimTuri"], r["Til"], g(r["Semestr"]),
                g(r["Maruza"]), g(r["Amaliyot"]), g(r["Potok"]), g(r["Guruh"]), g(r["Reyting"]), g(jami))))
        self._fill(self.t_fan, rows)
        if hasattr(self, "fan_total"):
            self.fan_total.set(f"Jami soatlar: {g(total)}")

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
                ("Reyting", "Reyting (soat)", "float", None),
                ("Til", "Til (o'qitish tili)", "combo", TILLAR)]

    @staticmethod
    def _fan_preview(d):
        mj, aj, js = fan_totals(d.get("Maruza"), d.get("Amaliyot"), d.get("Potok"), d.get("Guruh"), d.get("Reyting"))
        return f"Ma'ruza jami={g(mj)} · Amaliyot jami={g(aj)} · Jami soat={g(js)}"

    def _fan_exists(self, nomi, yon, talim, sem, til, exclude_id=None):
        key = ((nomi or "").strip().lower(), (yon or "").strip().lower(),
               (talim or "").strip().lower(), str(int(sem or 0)), (til or "").strip().lower())
        for r in self.con.execute("SELECT FanID, FanNomi, Yonalish, TalimTuri, Semestr, Til FROM Fanlar"):
            if exclude_id and r["FanID"] == exclude_id:
                continue
            rk = ((r["FanNomi"] or "").strip().lower(), (r["Yonalish"] or "").strip().lower(),
                  (r["TalimTuri"] or "").strip().lower(), str(int(r["Semestr"] or 0)), (r["Til"] or "").strip().lower())
            if rk == key:
                return True
        return False

    def fan_add(self):
        d = FormDialog(self, "Yangi fan", self._fan_spec(),
                       {"Potok": 1, "Guruh": 1, "Kategoriya": "1", "TalimTuri": "Kunduzgi", "Til": "O'zbek"},
                       computed=self._fan_preview, rules=fan_rules).result
        if d:
            if self._fan_exists(d["FanNomi"], d["Yonalish"], d["TalimTuri"], d["Semestr"], d.get("Til")):
                if not messagebox.askyesno("Takror fan",
                        "Nomi, yo'nalishi, ta'lim shakli, semestri va tili bir xil fan allaqachon mavjud.\n"
                        "Baribir qo'shilsinmi?"):
                    return
            self.con.execute("INSERT INTO Fanlar(FanNomi,Yonalish,TalimTuri,Kategoriya,Semestr,"
                             "Maruza,Amaliyot,Potok,Guruh,Reyting,Til) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                             (d["FanNomi"], d["Yonalish"], d["TalimTuri"], int(d["Kategoriya"] or 0),
                              int(d["Semestr"] or 0), d["Maruza"], d["Amaliyot"], d["Potok"], d["Guruh"],
                              d["Reyting"], d.get("Til") or "O'zbek"))
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
                             "Maruza=?,Amaliyot=?,Potok=?,Guruh=?,Reyting=?,Til=? WHERE FanID=?",
                             (d["FanNomi"], d["Yonalish"], d["TalimTuri"], int(d["Kategoriya"] or 0),
                              int(d["Semestr"] or 0), d["Maruza"], d["Amaliyot"], d["Potok"], d["Guruh"],
                              d["Reyting"], d.get("Til") or "O'zbek", i))
            self.con.commit()
            self.refresh_all()

    def fan_del(self):
        ids = self._selected_ids(self.t_fan)
        if not ids:
            return
        used, free = [], []
        for i in ids:
            r = self.con.execute("SELECT FanNomi, (SELECT COUNT(*) FROM Taqsimot WHERE FanID=?) c "
                                 "FROM Fanlar WHERE FanID=?", (i, i)).fetchone()
            if not r:
                continue
            (used if r["c"] else free).append((i, r["FanNomi"]))
        if used and not free:
            names = "\n".join(f"• {f}" for _, f in used[:10])
            messagebox.showwarning("O'chirib bo'lmaydi",
                f"Tanlangan fan(lar) taqsimotda ishlatilgan:\n{names}\n\n"
                "Avval ularning taqsimot yozuvlarini o'chiring.")
            return
        msg = (f"{len(free)} ta fan o'chirilsinmi?" if len(free) > 1 else "Tanlangan fan o'chirilsinmi?")
        if used:
            names = "\n".join(f"• {f}" for _, f in used[:10])
            msg += (f"\n\nQuyidagi {len(used)} tasi taqsimotda ishlatilgani uchun O'CHIRILMAYDI:\n{names}")
        if not messagebox.askyesno("Tasdiqlang", msg):
            return
        free_ids = [i for i, _ in free]
        names = ", ".join(f for _, f in free[:3]) + ("..." if len(free) > 3 else "")
        self._record_delete(f"Fanlar: {len(free)} ta o'chirildi ({names})",
                            self._snap(self.con, "Fanlar", "FanID", free_ids))
        self.con.executemany("DELETE FROM Fanlar WHERE FanID=?", [(i,) for i in free_ids])
        self.con.commit()
        self.refresh_all()

    def _fan_label_dups(self, raw_rows):
        """Keep duplicate courses but append the differing field(s) in brackets to the name,
        e.g. two 'Statistika' rows differing only in Reyting become
        'Statistika (Reyting: 6)' and 'Statistika (Reyting: 7)'."""
        DIFF = [("Kategoriya", "Kat"), ("Maruza", "Ma'ruza"), ("Amaliyot", "Amaliyot"),
                ("Potok", "Potok"), ("Guruh", "Guruh"), ("Reyting", "Reyting")]

        def idk(r):
            return (r.get("FanNomi", "").strip().lower(), r.get("Yonalish", "").strip().lower(),
                    r.get("TalimTuri", "").strip().lower(), str(r.get("Semestr", "")).strip(),
                    (r.get("Til", "") or "").strip().lower())
        groups = {}
        for r in raw_rows:
            groups.setdefault(idk(r), []).append(r)
        labeled = 0
        for members in groups.values():
            if len(members) < 2:
                continue
            differing = [(f, lbl) for f, lbl in DIFF
                         if len({(m.get(f) or "").strip() for m in members}) > 1]
            for i, m in enumerate(members, 1):
                if differing:
                    parts = [f"{lbl}: {(m.get(f) or '').strip() or '0'}" for f, lbl in differing]
                    m["FanNomi"] = m.get("FanNomi", "").strip() + " (" + ", ".join(parts) + ")"
                else:
                    m["FanNomi"] = m.get("FanNomi", "").strip() + f" (nusxa {i})"
                labeled += 1
        note = f"{labeled} ta takror fan nomiga farqi qavs ichida qo'shildi." if labeled else ""
        return raw_rows, note

    def fan_import(self):
        cols = ["FanNomi", "Yonalish", "TalimTuri", "Kategoriya", "Semestr",
                "Maruza", "Amaliyot", "Potok", "Guruh", "Reyting", "Til"]
        self._csv_import("Fanlar yuklamasi", cols,
                         "INSERT INTO Fanlar(FanNomi,Yonalish,TalimTuri,Kategoriya,Semestr,"
                         "Maruza,Amaliyot,Potok,Guruh,Reyting,Til) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                         lambda r: (r.get("FanNomi", "").strip(), r.get("Yonalish", "").strip(),
                                    r.get("TalimTuri", "").strip(), self._i(r.get("Kategoriya")),
                                    self._i(r.get("Semestr")), self._f(r.get("Maruza")), self._f(r.get("Amaliyot")),
                                    self._i(r.get("Potok"), 1), self._i(r.get("Guruh"), 1),
                                    self._f(r.get("Reyting")) if r.get("TalimTuri", "").strip() == "Masofaviy" else 0.0,
                                    (r.get("Til", "") or "").strip() or "O'zbek"),
                         preprocess=self._fan_label_dups)

    def fan_dedup(self):
        """Remove duplicate courses (same Nomi+Yonalish+TalimTuri+Semestr), keeping the first."""
        groups = {}
        for r in self.con.execute("SELECT FanID, FanNomi, Yonalish, TalimTuri, Semestr, Til FROM Fanlar ORDER BY FanID"):
            k = ((r["FanNomi"] or "").strip().lower(), (r["Yonalish"] or "").strip().lower(),
                 (r["TalimTuri"] or "").strip().lower(), str(int(r["Semestr"] or 0)), (r["Til"] or "").strip().lower())
            groups.setdefault(k, []).append(r["FanID"])
        dup_ids = [fid for ids in groups.values() for fid in ids[1:]]
        if not dup_ids:
            messagebox.showinfo("Toza", "Takror fan topilmadi.")
            return
        qm = ",".join("?" * len(dup_ids))
        n_assign = self.con.execute(f"SELECT COUNT(*) c FROM Taqsimot WHERE FanID IN ({qm})", dup_ids).fetchone()["c"]
        warn = (f"{len(dup_ids)} ta takror fan topildi va o'chiriladi "
                "(har bir takror fanning birinchisi qoldiriladi).")
        if n_assign:
            warn += f"\nUlarga bog'langan {n_assign} ta taqsimot yozuvi ham o'chiriladi."
        warn += "\n\nDavom etilsinmi?"
        if not messagebox.askyesno("Takror fanlarni tozalash", warn):
            return
        items = (self._snap(self.con, "Fanlar", "FanID", dup_ids) +
                 [("Taqsimot", dict(r)) for r in
                  self.con.execute(f"SELECT * FROM Taqsimot WHERE FanID IN ({qm})", dup_ids)])
        self._record_delete(f"Takror tozalash: {len(dup_ids)} ta fan"
                            + (f" va {n_assign} ta taqsimot yozuvi" if n_assign else "")
                            + " o'chirildi", items)
        self.con.execute(f"DELETE FROM Taqsimot WHERE FanID IN ({qm})", dup_ids)
        self.con.execute(f"DELETE FROM Fanlar WHERE FanID IN ({qm})", dup_ids)
        self.con.commit()
        self.refresh_all()
        messagebox.showinfo("Tayyor", f"{len(dup_ids)} ta takror fan o'chirildi.")

    def fan_template(self):
        self._save_template(["FanNomi", "Yonalish", "TalimTuri", "Kategoriya", "Semestr",
                             "Maruza", "Amaliyot", "Potok", "Guruh", "Reyting", "Til"],
                            ["Fan nomi", "Iqtisodiyot", "Kunduzgi", "1", "3", "30", "30", "1", "1", "0", "O'zbek"],
                            "fanlar_shablon.csv")

    # ================= TAQSIMOT =================
    def _build_taqsimot(self):
        _, bar, body = self._make_tab("  Taqsimot  ")
        ttk.Button(bar, text="+ Qo'shish", style="Primary.TButton", command=self.taq_add).pack(side="left")
        ttk.Button(bar, text="Tahrirlash", style="Secondary.TButton", command=self.taq_edit).pack(side="left", padx=6)
        ttk.Button(bar, text="O'chirish", style="Danger.TButton", command=self.taq_del).pack(side="left")
        ttk.Button(bar, text="CSV ga eksport", style="Secondary.TButton", command=self.taq_export).pack(side="left", padx=(16, 0))
        self.taq_q = tk.StringVar()
        self._add_search(bar, self.taq_q)
        self.taq_q.trace_add("write", lambda *_: self.load_taqsimot())
        cols = ["ID", "Professor-o'qituvchi (F.I.Sh.)", "Fan", "Turi", "Soat"]
        w = [50, 290, 320, 110, 80]
        self.t_taq = self._make_tree(body, cols, w, {"ID": "center", "Soat": "e"})
        self.t_taq.bind("<Double-1>", lambda e: self.taq_edit())

    def load_taqsimot(self):
        q = (self.taq_q.get() if hasattr(self, "taq_q") else "").strip().lower()
        sql = """SELECT t.TaqsimotID, d.FIO, f.FanNomi, f.Til, f.Yonalish, f.TalimTuri, f.Semestr,
                        t.TurSoat, t.Soat
                 FROM Taqsimot t
                 LEFT JOIN Domlalar d ON d.DomlaID=t.DomlaID
                 LEFT JOIN Fanlar f   ON f.FanID=t.FanID
                 ORDER BY d.FIO COLLATE NOCASE, f.FanNomi COLLATE NOCASE"""
        rows = []
        for r in self.con.execute(sql):
            fio = r["FIO"] or "—"
            fan = r["FanNomi"] or "—"
            til = (r["Til"] or "").strip()
            if til and fan != "—":
                fan = f"{fan} ({til.lower()})"
            if not self._q_match(q, r["TaqsimotID"], fio, fan, r["Yonalish"], r["TalimTuri"],
                                 g(r["Semestr"]), r["TurSoat"], g(r["Soat"])):
                continue
            rows.append((r["TaqsimotID"], (r["TaqsimotID"], fio, fan, r["TurSoat"], g(r["Soat"]))))
        self._fill(self.t_taq, rows)

    def _has_base_data(self):
        if not self.con.execute("SELECT 1 FROM Domlalar LIMIT 1").fetchone() or \
           not self.con.execute("SELECT 1 FROM Fanlar LIMIT 1").fetchone():
            messagebox.showwarning("Ma'lumot yetarli emas", "Avval kamida bitta professor-o'qituvchi va bitta fan kiriting.")
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
        ids = self._selected_ids(self.t_taq)
        if not ids:
            return
        msg = (f"{len(ids)} ta taqsimot yozuvi o'chirilsinmi?" if len(ids) > 1
               else "Tanlangan yozuv o'chirilsinmi?")
        if messagebox.askyesno("Tasdiqlang", msg):
            self._record_delete(f"Taqsimot: {len(ids)} ta yozuv o'chirildi",
                                self._snap(self.con, "Taqsimot", "TaqsimotID", ids))
            self.con.executemany("DELETE FROM Taqsimot WHERE TaqsimotID=?", [(i,) for i in ids])
            self.con.commit()
            self.refresh_all()

    def taq_export(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")], initialfile="taqsimot.csv")
        if not path:
            return
        sql = """SELECT d.FIO, f.FanNomi, f.Til, f.Yonalish, f.TalimTuri, f.Semestr, t.TurSoat, t.Soat
                 FROM Taqsimot t
                 LEFT JOIN Domlalar d ON d.DomlaID=t.DomlaID
                 LEFT JOIN Fanlar f   ON f.FanID=t.FanID
                 ORDER BY d.FIO COLLATE NOCASE, f.FanNomi COLLATE NOCASE"""
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["Professor-o'qituvchi", "Fan", "Til", "Yonalish", "TalimTuri", "Semestr", "Turi", "Soat"])
                for r in self.con.execute(sql):
                    w.writerow([r["FIO"] or "", r["FanNomi"] or "", r["Til"] or "", r["Yonalish"] or "",
                                r["TalimTuri"] or "", g(r["Semestr"]), r["TurSoat"], g(r["Soat"])])
            messagebox.showinfo("Tayyor", f"Taqsimot saqlandi:\n{path}")
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    # ================= YUKLAMA (hisobot - fanlar) =================
    def _build_yukfan(self):
        _, bar, body = self._make_tab("  Yuklama (hisobot - fanlar)  ")
        ttk.Button(bar, text="Yangilash", style="Primary.TButton", command=self.load_yukfan).pack(side="left")
        ttk.Button(bar, text="CSV ga eksport", style="Secondary.TButton", command=self.yukfan_export).pack(side="left", padx=8)
        self.yukfan_q = tk.StringVar()
        self._add_search(bar, self.yukfan_q)
        self.yukfan_q.trace_add("write", lambda *_: self.load_yukfan())
        self.yukfan_summary = tk.StringVar()
        ttk.Label(bar, textvariable=self.yukfan_summary, style="Muted.TLabel").pack(side="right", padx=(8, 2))
        # Excel'dagi "переносить по словам" kabi — sarlavhalar 2 qatorli, ustunlar tor,
        # shunda ekranga ko'proq ustun sig'adi.
        cols = ["ID", "Fan nomi", "Yo'nalish", "Ta'lim\nturi", "Til", "Sem.",
                "Ma'ruza\n(jami)", "Ma'ruza\nberilgan", "Amaliyot\n(jami)", "Amaliyot\nberilgan",
                "Reyting\n(jami)", "Reyting\nberilgan", "Jami", "Berilgan", "Qoldiq", "Bajarilish\n%",
                "O'qituvchilar"]
        w = [40, 180, 110, 72, 55, 42, 68, 68, 72, 68, 62, 62, 56, 66, 58, 72, 210]
        an = {c: "e" for c in cols if c not in ("Fan nomi", "Yo'nalish", "Ta'lim\nturi", "Til", "O'qituvchilar")}
        an["ID"] = "center"
        self.t_yukfan = self._make_tree(body, cols, w, an, stretch=())
        for c in cols:                       # sarlavha matnini markazga tekislash
            self.t_yukfan.heading(c, anchor="center")

    def _fan_report_rows(self):
        """Per-course overview: planned vs assigned hours for each component + teachers."""
        assigned = {(r["FanID"], r["TurSoat"]): r["s"] for r in self.con.execute(
            "SELECT FanID, TurSoat, COALESCE(SUM(Soat),0) s FROM Taqsimot GROUP BY FanID, TurSoat")}
        teachers = {}
        for r in self.con.execute(
                """SELECT t.FanID, d.FIO, t.TurSoat FROM Taqsimot t
                   JOIN Domlalar d ON d.DomlaID=t.DomlaID
                   ORDER BY d.FIO COLLATE NOCASE"""):
            mark = {"Maruza": "M", "Amaliyot": "A", "Reyting": "R"}.get(r["TurSoat"], "?")
            teachers.setdefault(r["FanID"], []).append(f"{short_name(r['FIO'])} ({mark})")
        out = []
        for r in self.con.execute("SELECT * FROM Fanlar ORDER BY FanNomi COLLATE NOCASE"):
            mj, aj, js = fan_totals(r["Maruza"], r["Amaliyot"], r["Potok"], r["Guruh"], r["Reyting"])
            rj = (r["Reyting"] or 0) if (r["TalimTuri"] or "") == "Masofaviy" else 0
            mb = assigned.get((r["FanID"], "Maruza"), 0)
            ab = assigned.get((r["FanID"], "Amaliyot"), 0)
            rb = assigned.get((r["FanID"], "Reyting"), 0)
            jami = mj + aj + rj
            berilgan = mb + ab + rb
            pct = (berilgan / jami * 100) if jami else 0
            out.append({"id": r["FanID"], "nomi": r["FanNomi"], "yon": r["Yonalish"],
                        "talim": r["TalimTuri"], "til": r["Til"], "sem": r["Semestr"],
                        "mj": mj, "mb": mb, "aj": aj, "ab": ab, "rj": rj, "rb": rb,
                        "jami": jami, "berilgan": berilgan, "qoldiq": jami - berilgan,
                        "pct": pct, "oqit": ", ".join(teachers.get(r["FanID"], [])) or "—"})
        return out

    def load_yukfan(self):
        q = (self.yukfan_q.get() if hasattr(self, "yukfan_q") else "").strip().lower()
        self.t_yukfan.delete(*self.t_yukfan.get_children())
        tot_jami = tot_ber = 0
        shown = 0
        for d in sorted(self._fan_report_rows(), key=lambda x: x["pct"], reverse=True):
            tot_jami += d["jami"]
            tot_ber += d["berilgan"]
            if not self._q_match(q, d["id"], d["nomi"], d["yon"], d["talim"],
                                 d["til"], g(d["sem"]), d["oqit"]):
                continue
            shown += 1
            self.t_yukfan.insert("", "end", values=(
                d["id"], d["nomi"], d["yon"], d["talim"], d["til"], g(d["sem"]),
                g(d["mj"]), g(d["mb"]), g(d["aj"]), g(d["ab"]), g(d["rj"]), g(d["rb"]),
                g(d["jami"]), g(d["berilgan"]), g(d["qoldiq"]), f"{d['pct']:.0f}%", d["oqit"]),
                tags=("odd",) if shown % 2 == 0 else ())
        self.yukfan_summary.set(f"Ko'rsatilgan: {shown}   |   Jami reja: {g(tot_jami)} soat   "
                                f"|   Berilgan: {g(tot_ber)} soat   |   Qoldiq: {g(tot_jami - tot_ber)} soat")

    def yukfan_export(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")], initialfile="yuklama_fanlar_hisobot.csv")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["FanID", "FanNomi", "Yonalish", "TalimTuri", "Til", "Semestr",
                            "Maruza_jami", "Maruza_berilgan", "Amaliyot_jami", "Amaliyot_berilgan",
                            "Reyting_jami", "Reyting_berilgan", "Jami", "Berilgan", "Qoldiq",
                            "Bajarilish_%", "Oqituvchilar"])
                for d in self._fan_report_rows():
                    w.writerow([d["id"], d["nomi"], d["yon"], d["talim"], d["til"], g(d["sem"]),
                                g(d["mj"]), g(d["mb"]), g(d["aj"]), g(d["ab"]), g(d["rj"]), g(d["rb"]),
                                g(d["jami"]), g(d["berilgan"]), g(d["qoldiq"]), f"{d['pct']:.0f}", d["oqit"]])
            messagebox.showinfo("Tayyor", f"Hisobot saqlandi:\n{path}")
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    # ================= YUKLAMA (hisobot - o'qituvchi) =================
    def _build_yuklama(self):
        _, bar, body = self._make_tab("  Yuklama (hisobot - o'qituvchi)  ")
        ttk.Button(bar, text="Yangilash", style="Primary.TButton", command=self.load_yuklama).pack(side="left")
        ttk.Button(bar, text="CSV ga eksport", style="Secondary.TButton", command=self.yuk_export).pack(side="left", padx=8)
        self.yuk_q = tk.StringVar()
        self._add_search(bar, self.yuk_q)
        self.yuk_q.trace_add("write", lambda *_: self.load_yuklama())
        self.yuk_summary = tk.StringVar()
        ttk.Label(bar, textvariable=self.yuk_summary, style="Muted.TLabel").pack(side="right", padx=(8, 2))
        cols = ["ID", "F.I.Sh.", "Stavka", "Meyor (jami)", "Ma'ruza", "Amaliyot", "Reyting",
                "Jami berilgan", "Farq", "Bajarilish %", "Izoh"]
        w = [45, 220, 60, 95, 75, 75, 65, 95, 75, 90, 170]
        an = {c: "e" for c in cols if c not in ("F.I.Sh.", "Izoh")}
        an["ID"] = "center"
        self.t_yuk = self._make_tree(body, cols, w, an)
        self.t_yuk.tag_configure("over", background=UI["danger_soft"], foreground=UI["danger_dark"])

    def load_yuklama(self):
        q = (self.yuk_q.get() if hasattr(self, "yuk_q") else "").strip().lower()
        self.t_yuk.delete(*self.t_yuk.get_children())
        tot_norm = tot_assigned = 0
        shown = over = 0
        for d in sorted(workload_rows(self.con), key=lambda x: x["pct"], reverse=True):
            tot_norm += d["norm"]
            tot_assigned += d["jami"]
            if not self._q_match(q, d["id"], d["fio"], g(d["stavka"]), g(d["norm"])):
                continue
            shown += 1
            is_over = d["norm"] and d["jami"] > d["norm"] + 1e-9
            if is_over:
                over += 1
                izoh = f"⚠ Meyordan ortiq (+{g(d['jami'] - d['norm'])} soat)"
            else:
                izoh = ""
            tags = (("over",) if is_over else ()) + (("odd",) if shown % 2 == 0 else ())
            self.t_yuk.insert("", "end", values=(
                d["id"], d["fio"], g(d["stavka"]), g(d["norm"]), g(d["maruza"]), g(d["amaliyot"]),
                g(d["reyting"]), g(d["jami"]), g(d["diff"]), f"{d['pct']:.0f}%", izoh),
                tags=tags)
        summary = (f"Ko'rsatilgan: {shown}   |   Umumiy meyor: {g(tot_norm)} soat   "
                   f"|   Berilgan: {g(tot_assigned)} soat")
        if over:
            summary += f"   |   ⚠ Meyordan ortiq: {over} ta"
        self.yuk_summary.set(summary)

    def yuk_export(self):
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            filetypes=[("CSV", "*.csv")], initialfile="yuklama_hisobot.csv")
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["FIO", "Stavka", "Meyor_jami", "Maruza", "Amaliyot", "Reyting",
                            "Jami_berilgan", "Farq", "Bajarilish_%", "Izoh"])
                for d in workload_rows(self.con):
                    izoh = (f"Meyordan ortiq (+{g(d['jami'] - d['norm'])} soat)"
                            if d["norm"] and d["jami"] > d["norm"] + 1e-9 else "")
                    w.writerow([d["fio"], g(d["stavka"]), g(d["norm"]), g(d["maruza"]), g(d["amaliyot"]),
                                g(d["reyting"]), g(d["jami"]), g(d["diff"]), f"{d['pct']:.0f}", izoh])
            messagebox.showinfo("Tayyor", f"Hisobot saqlandi:\n{path}")
        except OSError as e:
            messagebox.showerror("Xato", str(e))

    # ---------- refresh ----------
    def _on_tab(self):
        idx = self.nb.index(self.nb.select())
        if idx == 3:
            self.load_yukfan()
        elif idx == 4:
            self.load_yuklama()

    def refresh_all(self):
        self.load_domlalar()
        self.load_fanlar()
        self.load_taqsimot()
        if hasattr(self, "t_yukfan"):
            self.load_yukfan()
        if hasattr(self, "t_yuk"):
            self.load_yuklama()

    # ---------- help centre ----------
    def show_help(self):
        win = tk.Toplevel(self)
        win.title("Yordam markazi — Dasturdan foydalanish")
        win.geometry("780x620")
        win.transient(self)
        try:
            win.geometry(f"+{self.winfo_rootx() + 70}+{self.winfo_rooty() + 50}")
        except Exception:
            pass

        win.configure(background=UI["bg"])
        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)
        txt = tk.Text(frm, wrap="word", padx=18, pady=14, background=UI["surface"],
                      relief="flat", cursor="arrow", font=(FONT, 10), foreground=UI["ink"],
                      highlightthickness=1, highlightbackground=UI["border"],
                      highlightcolor=UI["border"], borderwidth=0)
        vs = ttk.Scrollbar(frm, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        txt.pack(side="left", fill="both", expand=True)

        txt.tag_configure("h1", font=(FONT, 14, "bold"), foreground=UI["brand_dark"],
                          spacing1=16, spacing3=7)
        txt.tag_configure("h2", font=(FONT, 11, "bold"), foreground=UI["ink"],
                          spacing1=9, spacing3=3)
        txt.tag_configure("p", font=(FONT, 10), foreground="#374151", spacing3=3, spacing2=2)
        txt.tag_configure("b", font=(FONT, 10), foreground="#374151",
                          lmargin1=18, lmargin2=34, spacing3=3, spacing2=2)

        content = [
            ("h1", "Bu dastur nima uchun?"),
            ("p", "Dastur professor-o'qituvchilarning dars yuklamasini (taqsimotini) rejalashtirish "
                  "uchun mo'ljallangan. Siz o'qituvchilar va fanlar ro'yxatini kiritasiz, so'ngra "
                  "har bir fanni (ma'ruza, amaliyot, reyting) o'qituvchilarga biriktirasiz. Dastur "
                  "har bir o'qituvchining jami yuklamasini va me'yorga nisbatan bajarilishini "
                  "o'zi hisoblab beradi."),

            ("h1", "Asosiy bo'limlar (yuqoridagi 5 ta varaq)"),
            ("b", "Professor-O'qituvchilar — o'qituvchilar ro'yxati (F.I.Sh., ilmiy unvon, stavka, me'yor)."),
            ("b", "Fanlar yuklamasi — fanlar ro'yxati (nomi, yo'nalishi, ta'lim turi, tili, ma'ruza/amaliyot soatlari)."),
            ("b", "Taqsimot — fanlarni o'qituvchilarga biriktirish."),
            ("b", "Yuklama (hisobot - fanlar) — har bir fan bo'yicha reja / berilgan / qoldiq soatlar va o'qituvchilar."),
            ("b", "Yuklama (hisobot - o'qituvchi) — har bir o'qituvchining jami soatlari va bajarilish foizi."),

            ("h1", "1-qadam: O'qituvchilarni kiritish"),
            ("p", "Yuqoridan «Professor-O'qituvchilar» varag'ini tanlang."),
            ("b", "Bittalab kiritish uchun «+ Qo'shish» tugmasini bosing."),
            ("b", "Ko'p bo'lsa: «Shablon» tugmasi bilan tayyor CSV namunasini yuklab oling, uni "
                  "Excel'da to'ldiring (namuna qatorini o'chiring), so'ng «CSV import» orqali yuklang."),
            ("b", "Tahrirlash uchun qatorni ikki marta bosing yoki «Tahrirlash» tugmasini bosing. "
                  "O'chirish uchun qatorni tanlab «O'chirish»."),

            ("h1", "2-qadam: Fanlarni kiritish"),
            ("p", "«Fanlar yuklamasi» varag'iga o'ting va xuddi shu tarzda fanlarni kiriting."),
            ("b", "Har bir fan uchun: nomi, yo'nalishi, ta'lim turi (Kunduzgi/Sirtqi/Masofaviy/Kechki), "
                  "tili, ma'ruza va amaliyot soatlari, potok (oqim soni) va guruh soni."),
            ("b", "Reyting faqat Masofaviy fanlar uchun kiritiladi."),
            ("b", "«Jami soat» ustuni avtomatik hisoblanadi: Ma'ruza×Potok + Amaliyot×Guruh + Reyting."),
            ("b", "O'ng yuqorida «Jami soatlar» — barcha fanlarning umumiy soati ko'rsatiladi."),

            ("h1", "Til (o'qitish tili)"),
            ("p", "Har bir fanning tili bor: O'zbek, Rus yoki Ingliz. Til fan nomi yonida qavs ichida "
                  "ko'rsatiladi, masalan «Ekonometrika (rus)». Agar bitta fan bir nechta tilda o'qitilsa, "
                  "uni har bir til uchun alohida qator qilib kiriting — ular alohida fan sifatida qaraladi."),

            ("h1", "3-qadam: Taqsimot (dars biriktirish)"),
            ("p", "«Taqsimot» varag'ida «+ Qo'shish» tugmasini bosing. Ochilgan oynada:"),
            ("b", "Avval professor-o'qituvchini tanlang."),
            ("b", "Kerak bo'lsa, yo'nalish / ta'lim shakli / semestr / til filtrlari bilan ro'yxatni qisqartiring."),
            ("b", "«Fan / yuklama» maydoniga yozib qidiring — masalan «ekon» deb yozsangiz, mos fanlar "
                  "chiqadi. So'ng ro'yxatdan tanlang yoki Enter bosing."),
            ("b", "«Soat» avtomatik to'ldiriladi; kerak bo'lsa o'zgartiring. So'ng «Saqlash»."),

            ("h2", "Biriktirish qoidalari"),
            ("b", "Ma'ruza — jami soat = Ma'ruza × Potok. Har bir potok bitta o'qituvchi tomonidan to'liq "
                  "o'qitiladi (potok soatini bo'lish mumkin emas). Bitta o'qituvchi bir yoki bir nechta potok "
                  "olishi mumkin — «Potok soni» maydonida tanlang, soat avtomatik hisoblanadi. Barcha potoklar "
                  "taqsimlangach, fan ro'yxatdan chiqadi."),
            ("b", "Amaliyot — jami soat = Amaliyot × Guruh. Har bir guruh bitta o'qituvchi tomonidan to'liq "
                  "o'qitiladi (guruh soatini bo'lish mumkin emas). Bitta o'qituvchi bir yoki bir nechta guruh "
                  "olishi mumkin — «Guruh soni» maydonida tanlang, soat avtomatik hisoblanadi. Masalan 44 guruhli "
                  "fanni 44 ta o'qituvchiga (har biriga 1 guruhdan) yoki kamroq o'qituvchiga ko'proq guruhdan "
                  "berish mumkin. To'liq tarqatilmaguncha ro'yxatda «qoldi X/Y soat» ko'rinishida turadi."),
            ("b", "Bitta o'qituvchi ham ma'ruzani, ham amaliyotni o'qishi mumkin."),
            ("b", "Reyting — faqat shu fanning ma'ruzasi yoki amaliyotini o'qiydigan o'qituvchiga biriktiriladi."),

            ("h1", "4-bo'lim: Yuklama hisobotlari"),
            ("p", "«Yuklama (hisobot - fanlar)» — har bir fan bo'yicha rejadagi soatlar (Ma'ruza×Potok, "
                  "Amaliyot×Guruh, Reyting), berilgan va qolgan soatlar hamda fanni o'qitayotgan "
                  "o'qituvchilar ro'yxati ko'rsatiladi."),
            ("p", "«Yuklama (hisobot - o'qituvchi)» — har bir o'qituvchining me'yori, biriktirilgan soatlari, "
                  "farqi va bajarilish foizi ko'rsatiladi. Yangilash uchun «Yangilash» tugmasini bosing."),
            ("b", "Ikkala hisobot ham avtomatik tarzda «Bajarilish %» bo'yicha kamayish tartibida saralanadi "
                  "(eng yuqori foizdan eng pastiga). Boshqa ustun bo'yicha saralash uchun ustun sarlavhasini bosing."),
            ("b", "Meyordan ortiq yuklama — Taqsimotda biriktirish o'qituvchining meyor (jami)sidan oshsa, "
                  "ogohlantirish chiqadi, lekin yozuv baribir saqlanadi. Hisobotda bunday o'qituvchilar qizil "
                  "rangda va «Izoh» ustunida «⚠ Meyordan ortiq (+X soat)» belgisi bilan ko'rsatiladi."),
            ("p", "«Professor-O'qituvchilar» varag'ida o'ng yuqorida «Jami meyor» — barcha o'qituvchilarning "
                  "umumiy me'yori ko'rsatiladi."),

            ("h1", "Foydali imkoniyatlar"),
            ("b", "Saralash — istalgan ustun sarlavhasini bosing, ma'lumot o'sha ustun bo'yicha tartiblanadi "
                  "(yana bossangiz — teskari tartib). Sonlar son bo'yicha, matn alifbo bo'yicha saralanadi."),
            ("b", "Ko'p qatorni tanlash — Shift+bosish (oraliqni tanlaydi), Ctrl+bosish (bittalab qo'shadi) "
                  "yoki Ctrl+A (barchasini tanlaydi). So'ng «O'chirish» tugmasi bilan tanlangan barcha "
                  "qatorlarni bir vaqtda o'chirish mumkin. Taqsimotda ishlatilgan o'qituvchi/fanlar "
                  "o'chirilmaydi — ular ro'yxatda ko'rsatiladi, qolganlari o'chiriladi."),
            ("b", "Ortga qaytarish — o'chirilgan yozuvlarni tiklash mumkin. Ctrl+Z eng oxirgi o'chirish "
                  "amalini darhol qaytaradi. Yuqoridagi «↩ Ortga qaytarish» tugmasi esa barcha o'chirish "
                  "amallari ro'yxatini ochadi — birdaniga o'chirilgan qatorlar (masalan 50 ta) bitta amal "
                  "sifatida ko'rsatiladi va istalganini tanlab qaytarish mumkin. Ro'yxat dastur yopilguncha saqlanadi."),
            ("b", "Qidirish — har bir varaqdagi «Qidirish» maydoniga yozib, kerakli yozuvni tez toping. "
                  "Qidiruv barcha ustunlar bo'yicha ishlaydi: ID, F.I.Sh., fan nomi, yo'nalish, ta'lim turi, "
                  "til, semestr va h.k. Bir nechta so'z yozsangiz, hammasi mos kelgan qatorlar chiqadi "
                  "(masalan: «ekonometrika kunduzgi»)."),
            ("b", "CSV eksport — «Taqsimot» va «Yuklama» bo'limlarida ma'lumotni Excel uchun CSV faylga saqlash mumkin."),
            ("b", "Takror fanlar — bir xil fan ikki marta kiritilgan bo'lsa, «Takror tozalash» tugmasi ularni "
                  "tozalaydi. CSV import paytida takror fanlarning farqi (masalan reyting) avtomatik qavs ichida belgilanadi."),

            ("h1", "Ma'lumotlar saqlanadimi? (Muhim)"),
            ("p", "Ha. Har bir o'zgarish darhol saqlanadi — alohida «saqlash» tugmasi kerak emas. Dasturni "
                  "yopib qayta ochsangiz, hamma narsa joyida turadi."),
            ("b", "Barcha ma'lumot «dars_taqsimoti.db» faylida saqlanadi — u dastur (exe) bilan bir papkada turadi."),
            ("b", "MUHIM: yuklab olingan ZIP faylni avval haqiqiy papkaga (masalan, Ishchi stol) chiqaring, "
                  "keyin exe'ni o'sha papkadan oching. ZIP ichidan to'g'ridan-to'g'ri ochmang — aks holda "
                  "ma'lumot vaqtinchalik papkaga yoziladi va o'chib ketishi mumkin."),
            ("b", "Zaxira uchun «dars_taqsimoti.db» faylini vaqti-vaqti bilan boshqa joyga (USB yoki Drive'ga) nusxalang."),
        ]
        for tag, text in content:
            txt.insert("end", ("•  " + text if tag == "b" else text) + "\n", tag)
        txt.config(state="disabled")

        ttk.Button(win, text="Yopish", style="Secondary.TButton", command=win.destroy).pack(pady=(2, 12))
        win.bind("<Escape>", lambda e: win.destroy())


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
