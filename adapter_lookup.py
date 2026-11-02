#!/usr/bin/env python3
"""
Adapter Lookup - a search window for an Access database of adapters.

* Finds the .accdb / .mdb next to this script (or asks you to pick one).
* Reads the tables / saved queries and their columns automatically.
* Top search box: part number or keywords (use * as a wildcard). Searches all
  columns, or just the one chosen in the "in" box.
* Left panel: one filter per column - pick-lists for columns with few distinct
  values (connector / interface type...), min-max ranges for numbers (pin
  count, size...), Yes/No for checkboxes, date ranges for dates.
* Click a column header to sort, double-click a row for full details,
  Ctrl+C copies selected rows, File > Export saves results to CSV.

The database is opened read-only; this tool never changes it.
Run with --demo to try it on built-in sample data (no Access driver needed).
"""
import csv
import datetime as dt
import decimal
import glob
import json
import os
import re
import struct
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "Adapter Lookup"
if getattr(sys, "frozen", False):
    # Running as a packaged .exe: look next to the .exe, not in the temp unpack folder.
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "adapter_lookup_settings.json")
MAX_ROWS = 20000         # most connector rows shown at once
DISTINCT_LIMIT = 250     # columns with this many values or fewer get a pick-list

# ---- Settings for the SIGBOARDS database --------------------------------- #
DEFAULT_TABLE = "CONNECTORS"     # the only table used when it exists
GROUP_COLUMN = "BOARD/ADAPTER"   # one adapter = all rows with the same value here
# A row counts as scrapped when one of these columns matches the text (any case).
SCRAP_RULES = [("Port", "SCRAP"), ("Category", "SCRAPPED")]
# Order of the result columns and of the filters (other columns go after).
RESULT_ORDER = ["BOARD/ADAPTER", "CONNECTOR P/N", "Qty", "DESCRIPTION", "Positions",
                "Category", "Contact Type", "Housing Type", "Series", "Arrangement",
                "Shell", "Keying", "Pitch", "Manufacturer", "LOCATION", "Port",
                "CUSTOMER", "Relevant Assemblies", "OBSOLETE", "Voltage Rating"]
FILTER_ORDER = ["BOARD/ADAPTER", "CONNECTOR P/N", "DESCRIPTION", "Positions", "Category",
                "Contact Type", "Housing Type", "Series", "Arrangement", "Shell", "Keying",
                "Pitch", "Manufacturer", "CUSTOMER", "Relevant Assemblies", "LOCATION",
                "Port", "OBSOLETE", "Voltage Rating", "Qty"]
FRIENDLY_LABELS = {"Positions": "Positions (pins)"}
PLACEHOLDERS = {"", "-", "--", "?", "??"}  # left out of pick-lists

# Filters for columns whose names contain these words are listed first.
PRIORITY_HINTS = ["part", "pn", "sku", "model", "connector", "interface", "type",
                  "pin", "size", "gender", "length", "manufacturer", "mfr",
                  "vendor", "brand"]

TEXT_TYPES = {"VARCHAR", "CHAR", "TEXT", "WVARCHAR", "WCHAR", "NVARCHAR", "NCHAR",
              "LONGCHAR", "MEMO", "WLONGVARCHAR", "LONGVARCHAR", "HYPERLINK", "GUID"}
LONG_TEXT_TYPES = {"LONGCHAR", "MEMO", "WLONGVARCHAR", "LONGVARCHAR", "HYPERLINK"}
NUMBER_TYPES = {"INTEGER", "INT", "SMALLINT", "BYTE", "TINYINT", "COUNTER", "BIGINT",
                "DOUBLE", "REAL", "FLOAT", "CURRENCY", "DECIMAL", "NUMERIC", "MONEY"}
DATE_TYPES = {"DATETIME", "DATE", "TIMESTAMP", "TIME"}
BOOL_TYPES = {"BIT", "BOOLEAN", "YESNO"}
BINARY_TYPES = {"LONGBINARY", "VARBINARY", "BINARY", "IMAGE", "OLEOBJECT", "BLOB"}


# --------------------------------------------------------------------------- #
#  Helpers
# --------------------------------------------------------------------------- #
def qi(name):
    """Quote an identifier for Access/SQLite: [name]."""
    return "[" + name.replace("]", "]]") + "]"


def classify(type_name):
    t = (type_name or "").upper().split("(")[0].strip()
    if t in BOOL_TYPES:
        return "bool"
    if t in NUMBER_TYPES:
        return "number"
    if t in DATE_TYPES:
        return "date"
    if t in BINARY_TYPES:
        return "binary"
    return "text"


def is_part_column(name):
    n = name.lower().replace(" ", "").replace("_", "").replace("/", "").replace(".", "")
    return ("part" in n or "sku" in n or n in ("pn", "partno", "model", "modelno")
            or n.startswith("pn") or n.endswith("pn"))


def order_index(name, order):
    lower = [o.lower() for o in order]
    return lower.index(name.lower()) if name.lower() in lower else len(order)


def clean_text(v):
    """Collapse line breaks / repeated spaces that are stored in some cells."""
    return " ".join(str(v).split())


def priority(name):
    n = name.lower()
    i = order_index(name, FILTER_ORDER)
    if i < len(FILTER_ORDER):
        return -100 + i
    if is_part_column(name):
        return 0
    for i, hint in enumerate(PRIORITY_HINTS):
        if hint in n:
            return 1 + i
    return 100


def parse_number(text):
    text = text.strip().replace(",", "")
    if not text:
        return None
    return float(text)


def parse_date(text):
    text = text.strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%Y/%m/%d"):
        try:
            return dt.datetime.strptime(text, fmt)
        except ValueError:
            pass
    raise ValueError(f"'{text}' is not a date (use YYYY-MM-DD or MM/DD/YYYY)")


def fmt_value(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Yes" if v else "No"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:g}"
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if v.time() == dt.time() else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, (bytes, bytearray, memoryview)):
        return "<binary>"
    return clean_text(v)


def sort_key(v):
    if v is None or v == "":
        return (2, "")
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (0, float(v))
    if isinstance(v, (dt.datetime, dt.date)):
        return (0, v.toordinal() if isinstance(v, dt.date) and not isinstance(v, dt.datetime)
                else v.timestamp())
    s = str(v)
    try:
        return (0, float(s))
    except ValueError:
        return (1, s.lower())


class Column:
    def __init__(self, name, type_name):
        self.name = name
        self.type_name = (type_name or "").upper()
        self.kind = classify(type_name)
        self.long_text = self.type_name in LONG_TEXT_TYPES

    def __repr__(self):
        return f"Column({self.name!r}, {self.kind})"


def build_query(table, columns, quick_text="", quick_column=None, filters=(),
                concat="&", date_param=lambda d: d, exclude=(), order_by=None):
    """Return (sql, params).

    filters: iterable of (Column, op, value) where op is one of
      'eq', 'contains', 'min', 'max', 'yes', 'no'.
      'eq' ignores case and spacing (so 'Plastic Crimp Housing' matches a value
      stored with a line break in it).
    exclude: (column name, text) pairs - rows where that column equals the text
      (any case) are left out; used for hiding scrapped adapters.
    order_by: column names to sort by.
    """
    select_cols = [c for c in columns if c.kind != "binary"]
    where, params = [], []

    search_cols = [quick_column] if quick_column else select_cols
    for term in quick_text.split():
        pattern = "%" + term.replace("*", "%").replace("?", "_") + "%"
        ors = [f"({qi(c.name)} {concat} '') LIKE ?" for c in search_cols]
        where.append("(" + " OR ".join(ors) + ")")
        params.extend([pattern] * len(ors))

    for col, op, value in filters:
        c = qi(col.name)
        if op == "eq":
            where.append(f"{c} LIKE ?")
            params.append(re.sub(r"\s+", "%", value.strip()))
        elif op == "contains":
            where.append(f"({c} {concat} '') LIKE ?")
            params.append("%" + value.replace("*", "%").replace("?", "_") + "%")
        elif op in ("min", "max"):
            where.append(f"{c} {'>=' if op == 'min' else '<='} ?")
            params.append(date_param(value) if col.kind == "date" else value)
        elif op == "yes":
            where.append(f"{c} <> 0")
        elif op == "no":
            where.append(f"({c} = 0 OR {c} IS NULL)")
        else:
            raise ValueError(op)

    names = {c.name.lower(): c.name for c in select_cols}
    for col_name, text in exclude:
        if col_name.lower() in names:
            c = qi(names[col_name.lower()])
            where.append(f"({c} IS NULL OR {c} NOT LIKE ?)")
            params.append(text)

    sql = "SELECT " + ", ".join(qi(c.name) for c in select_cols) + " FROM " + qi(table)
    if where:
        sql += " WHERE " + " AND ".join(where)
    if order_by is None:
        first = next((c.name for c in select_cols if is_part_column(c.name)), None)
        order_by = [first] if first else []
    order_by = [names[n.lower()] for n in order_by if n and n.lower() in names]
    if order_by:
        sql += " ORDER BY " + ", ".join(qi(n) for n in order_by)
    return sql, params


# --------------------------------------------------------------------------- #
#  Database back-ends
# --------------------------------------------------------------------------- #
class AccessBackend:
    concat = "&"

    def __init__(self, path):
        try:
            import pyodbc
        except ImportError:
            raise RuntimeError(
                "The 'pyodbc' package isn't installed.\n\n"
                "Double-click 'Install requirements.bat' in the Adapter Lookup "
                "folder (or run:  py -m pip install pyodbc), then start again.")
        drivers = [d for d in pyodbc.drivers() if "Microsoft Access Driver" in d]
        if not drivers:
            bits = struct.calcsize("P") * 8
            raise RuntimeError(
                f"No Microsoft Access ODBC driver was found for {bits}-bit Python.\n\n"
                "Install the free 'Microsoft Access Database Engine 2016 "
                f"Redistributable' ({'x64' if bits == 64 else 'x86'} version) from "
                "Microsoft, or use a Python whose bitness matches your Office "
                "install (64-bit Office -> 64-bit Python).")
        driver = next((d for d in drivers if "accdb" in d), drivers[0])
        self.path = path
        self.conn = pyodbc.connect(
            f"DRIVER={{{driver}}};DBQ={path};ReadOnly=1;", autocommit=True)

    def date_param(self, d):
        return d

    def tables(self):
        cur = self.conn.cursor()
        out = []
        for row in cur.tables():
            if row.table_type in ("TABLE", "VIEW") and not row.table_name.startswith("~"):
                out.append((row.table_name, row.table_type == "VIEW"))
        return sorted(out, key=lambda t: (t[1], t[0].lower()))

    def has_table(self, table):
        try:
            self.execute(f"SELECT * FROM {qi(table)} WHERE 1=0")
            return True
        except Exception:
            return False

    def columns(self, table):
        # Read names/types from an empty query rather than cursor.columns():
        # the Access driver's catalog call sometimes returns corrupt text,
        # which pyodbc fails to decode (UnicodeDecodeError).
        cur = self.execute(f"SELECT * FROM {qi(table)} WHERE 1=0")
        types = {bool: "BIT", int: "INTEGER", float: "DOUBLE", decimal.Decimal: "DECIMAL",
                 dt.datetime: "DATETIME", dt.date: "DATE", bytes: "LONGBINARY",
                 bytearray: "LONGBINARY"}
        out = []
        for d in cur.description:
            name = types.get(d[1])
            if name is None:   # text: over 255 characters means a Memo / Long Text field
                name = "LONGCHAR" if (d[3] or 0) > 255 else "VARCHAR"
            out.append(Column(d[0], name))
        return out

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass


class SqliteBackend:
    """Used for --demo and for testing; mirrors the Access back-end."""
    concat = "||"

    def __init__(self, path):
        import sqlite3
        self.path = path
        self.conn = sqlite3.connect(path)

    def date_param(self, d):
        return d.strftime("%Y-%m-%d")

    def tables(self):
        rows = self.conn.execute(
            "SELECT name, type FROM sqlite_master WHERE type IN ('table','view') "
            "AND name NOT LIKE 'sqlite_%' ORDER BY type, name").fetchall()
        return [(n, t == "view") for n, t in rows]

    def has_table(self, table):
        return self.conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type IN ('table','view') "
            "AND lower(name) = lower(?)", (table,)).fetchone() is not None

    def columns(self, table):
        rows = self.conn.execute(f"PRAGMA table_info({qi(table)})").fetchall()
        return [Column(r[1], r[2]) for r in rows]

    def execute(self, sql, params=()):
        return self.conn.execute(sql, list(params))

    def close(self):
        self.conn.close()


def make_demo_db(path):
    import random
    import sqlite3
    if os.path.exists(path):
        os.remove(path)
    rnd = random.Random(7)
    conn = sqlite3.connect(path)
    conn.execute("""CREATE TABLE Adapters (
        ID INTEGER, PartNumber VARCHAR, Description VARCHAR, ConnectorA VARCHAR,
        ConnectorB VARCHAR, InterfaceType VARCHAR, PinCount INTEGER, Gender VARCHAR,
        Length_mm DOUBLE, Manufacturer VARCHAR, InStock BIT, DateAdded DATETIME,
        Notes LONGCHAR)""")
    conns = [("DB9", 9, "RS-232"), ("DB25", 25, "RS-232"), ("RJ45", 8, "Ethernet"),
             ("USB-A", 4, "USB"), ("USB-C", 24, "USB"), ("M12", 8, "Ethernet"),
             ("BNC", 1, "Coax"), ("SMA", 1, "Coax"), ("D-Sub 15", 15, "VGA"),
             ("MIL-38999", 37, "Mil-Spec"), ("Micro-D 51", 51, "Mil-Spec"),
             ("Terminal block", 6, "Discrete")]
    mfrs = ["Amphenol", "TE Connectivity", "Molex", "Glenair", "L-com", "Startech"]
    rows = []
    for i in range(1, 241):
        a, b = rnd.sample(conns, 2)
        if rnd.random() < .5:
            b = (a[0], a[1], a[2])
        pn = f"ADP-{a[0].split()[0].replace('-', '')[:4].upper()}-{1000 + i}"
        rows.append((i, pn, f"{a[0]} to {b[0]} adapter", a[0], b[0], a[2], a[1],
                     rnd.choice(["M-F", "F-F", "M-M"]),
                     rnd.choice([None, 25.0, 50.0, 100.0, 150.0, 300.0]),
                     rnd.choice(mfrs), rnd.random() < .7,
                     (dt.date(2019, 1, 1) + dt.timedelta(days=rnd.randint(0, 2400))).isoformat(),
                     rnd.choice(["", "", "Shielded", "Strain relief", "Rev B pinout"])))
    conn.executemany("INSERT INTO Adapters VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", rows)
    conn.execute("CREATE TABLE Vendors (VendorName VARCHAR, Phone VARCHAR, Website VARCHAR)")
    conn.executemany("INSERT INTO Vendors VALUES (?,?,?)",
                     [(m, "555-0100", m.lower().replace(" ", "") + ".com") for m in mfrs])
    conn.commit()
    conn.close()


# --------------------------------------------------------------------------- #
#  Settings
# --------------------------------------------------------------------------- #
def load_settings():
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(s):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(s, f, indent=2)
    except Exception:
        pass


# --------------------------------------------------------------------------- #
#  GUI
# --------------------------------------------------------------------------- #
class ScrollFrame(ttk.Frame):
    def __init__(self, master, **kw):
        super().__init__(master, **kw)
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas)
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.bind_all("<MouseWheel>", self._wheel, add="+")
        self.bind_all("<Button-4>", self._wheel, add="+")
        self.bind_all("<Button-5>", self._wheel, add="+")

    def _wheel(self, event):
        w = event.widget
        while w is not None:
            if w is self:
                break
            w = getattr(w, "master", None)
        else:
            return
        if isinstance(event.widget, ttk.Combobox):
            return
        step = -1 if (getattr(event, "num", 0) == 4 or getattr(event, "delta", 0) > 0) else 1
        self.canvas.yview_scroll(step, "units")


class FilterRow:
    """One filter in the left panel."""

    def __init__(self, app, parent, col, distinct, row):
        self.col = col
        self.vars = []
        self.choices = set()
        label = ttk.Label(parent, text=FRIENDLY_LABELS.get(col.name, col.name), anchor="w")
        label.grid(row=row, column=0, sticky="w", padx=(4, 6), pady=2)
        cell = ttk.Frame(parent)
        cell.grid(row=row, column=1, sticky="ew", padx=(0, 4), pady=2)
        if col.kind == "text":
            cell.columnconfigure(0, weight=1)

        if col.kind == "bool":
            v = tk.StringVar(value="Any")
            cb = ttk.Combobox(cell, textvariable=v, values=["Any", "Yes", "No"],
                              state="readonly", width=8)
            cb.grid(row=0, column=0, sticky="w")
            cb.bind("<<ComboboxSelected>>", lambda e: app.run_search())
            self.vars = [v]
        elif col.kind in ("number", "date"):
            lo, hi = tk.StringVar(), tk.StringVar()
            w = 8 if col.kind == "number" else 11
            e1 = ttk.Entry(cell, textvariable=lo, width=w)
            e2 = ttk.Entry(cell, textvariable=hi, width=w)
            e1.grid(row=0, column=0, sticky="w")
            ttk.Label(cell, text=" to ").grid(row=0, column=1)
            e2.grid(row=0, column=2, sticky="w")
            for e in (e1, e2):
                e.bind("<Return>", lambda ev: app.run_search())
            if col.kind == "date":
                ttk.Label(cell, text="YYYY-MM-DD", foreground="gray").grid(
                    row=1, column=0, columnspan=3, sticky="w")
            self.vars = [lo, hi]
        else:
            v = tk.StringVar()
            if distinct is not None:
                self.choices = {d.lower() for d in distinct}
                w = ttk.Combobox(cell, textvariable=v, values=[""] + distinct)
                w.bind("<<ComboboxSelected>>", lambda e: app.run_search())
            else:
                w = ttk.Entry(cell, textvariable=v)
            w.grid(row=0, column=0, sticky="ew")
            w.bind("<Return>", lambda e: app.run_search())
            self.vars = [v]

    def clear(self):
        if self.col.kind == "bool":
            self.vars[0].set("Any")
        else:
            for v in self.vars:
                v.set("")

    def conditions(self):
        """Return list of (Column, op, value); raises ValueError on bad input."""
        c = self.col
        if c.kind == "bool":
            val = self.vars[0].get()
            return [] if val == "Any" else [(c, "yes" if val == "Yes" else "no", None)]
        if c.kind in ("number", "date"):
            parse = parse_number if c.kind == "number" else parse_date
            out = []
            for op, var in (("min", self.vars[0]), ("max", self.vars[1])):
                try:
                    val = parse(var.get())
                except ValueError as e:
                    raise ValueError(f"{c.name}: {e}" if c.kind == "date"
                                     else f"{c.name}: '{var.get()}' is not a number")
                if val is not None:
                    out.append((c, op, val))
            return out
        text = self.vars[0].get().strip()
        if not text:
            return []
        if clean_text(text).lower() in self.choices and "*" not in text:
            return [(c, "eq", text)]
        return [(c, "contains", text)]


class App(tk.Tk):
    def __init__(self, db_path=None, demo=False):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1360x800")
        self.minsize(900, 500)
        self.settings = load_settings()
        self.backend = None
        self.table_map = {}
        self.columns = []
        self.group_col = None
        self.filter_rows = []
        self.result_cols = []
        self.rows = []            # flat mode: matched rows
        self.groups = []          # grouped mode: [(adapter, [(row, matched), ...])]
        self.iid_rows = {}        # tree item id -> list of rows
        self.full_cache = {}      # hide_scrapped flag -> {adapter: [rows]}
        self.sort_state = (None, False)
        self._after_id = None
        self.demo = demo

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("Treeview", rowheight=22)

        self._build_menu()
        self._build_ui()
        self.after(50, lambda: self._startup(db_path))

    # ---------- layout ----------
    def _build_menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=False)
        f.add_command(label="Open database...", command=self.choose_db, accelerator="Ctrl+O")
        f.add_command(label="Reload", command=self.reload, accelerator="F5")
        f.add_separator()
        f.add_command(label="Export results to CSV...", command=self.export_csv,
                      accelerator="Ctrl+E")
        f.add_separator()
        f.add_command(label="Exit", command=self.destroy)
        m.add_cascade(label="File", menu=f)
        v = tk.Menu(m, tearoff=False)
        v.add_command(label="Expand all adapters", command=lambda: self._expand(True))
        v.add_command(label="Collapse all adapters", command=lambda: self._expand(False))
        m.add_cascade(label="View", menu=v)
        self.config(menu=m)
        self.bind("<Control-o>", lambda e: self.choose_db())
        self.bind("<Control-e>", lambda e: self.export_csv())
        self.bind("<F5>", lambda e: self.reload())

    def _build_ui(self):
        top = ttk.Frame(self, padding=(8, 8, 8, 2))
        top.pack(fill="x")
        self.table_label = ttk.Label(top, text="Table:")
        self.table_label.pack(side="left")
        self.table_var = tk.StringVar()
        self.table_cb = ttk.Combobox(top, textvariable=self.table_var, state="readonly",
                                     width=28)
        self.table_cb.pack(side="left", padx=(4, 16))
        self.table_cb.bind("<<ComboboxSelected>>", lambda e: self.load_table())

        self.quick_label = ttk.Label(top, text="Part number / keywords:")
        self.quick_label.pack(side="left")
        self.quick_var = tk.StringVar()
        self.quick_entry = ttk.Entry(top, textvariable=self.quick_var, width=34,
                                     font=("Segoe UI", 11))
        self.quick_entry.pack(side="left", padx=4)
        self.quick_entry.bind("<KeyRelease>", self._quick_typed)
        self.quick_entry.bind("<Return>", lambda e: self.run_search())
        ttk.Label(top, text="in").pack(side="left", padx=(4, 4))
        self.in_var = tk.StringVar(value="All columns")
        self.in_cb = ttk.Combobox(top, textvariable=self.in_var, state="readonly", width=20)
        self.in_cb.pack(side="left")
        self.in_cb.bind("<<ComboboxSelected>>", lambda e: self.run_search())
        ttk.Button(top, text="Search", command=self.run_search).pack(side="left", padx=(10, 4))
        ttk.Button(top, text="Clear all", command=self.clear_all).pack(side="left")
        ttk.Button(top, text="Export CSV", command=self.export_csv).pack(side="right")

        opts = ttk.Frame(self, padding=(8, 0, 8, 2))
        opts.pack(fill="x")
        self.group_var = tk.BooleanVar(value=self.settings.get("group", True))
        self.group_chk = ttk.Checkbutton(
            opts, text="Group by adapter (show every connector on each matching adapter)",
            variable=self.group_var, command=self._options_changed)
        self.group_chk.pack(side="left")
        self.scrap_var = tk.BooleanVar(value=self.settings.get("hide_scrapped", True))
        self.scrap_chk = ttk.Checkbutton(opts, text="Hide scrapped", variable=self.scrap_var,
                                         command=self._options_changed)
        self.scrap_chk.pack(side="left", padx=(16, 0))

        pane = ttk.PanedWindow(self, orient="horizontal")
        pane.pack(fill="both", expand=True, padx=8, pady=4)

        left = ttk.LabelFrame(pane, text="Filters", padding=4)
        self.filter_scroll = ScrollFrame(left)
        self.filter_scroll.pack(fill="both", expand=True)
        self.filter_scroll.inner.columnconfigure(1, weight=1)
        pane.add(left, weight=0)

        right = ttk.Frame(pane)
        self.tree = ttk.Treeview(right, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(right, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(right, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        from tkinter import font as tkfont
        bold = tkfont.nametofont("TkDefaultFont").copy()
        bold.configure(weight="bold")
        self._bold = bold
        self.tree.tag_configure("odd", background="#f3f6fa")
        self.tree.tag_configure("group", background="#dfe8f5", font=bold)
        self.tree.tag_configure("match", font=bold, background="#fff4cc")
        self.tree.tag_configure("other", foreground="#555555")
        self.tree.bind("<Double-1>", self.show_details)
        self.tree.bind("<Return>", self.show_details)
        self.tree.bind("<Control-c>", lambda e: self.copy_rows())
        self.tree.bind("<Button-3>", self._context_menu)
        pane.add(right, weight=1)

        self.status_var = tk.StringVar(value="No database open.")
        ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w",
                  padding=(6, 2)).pack(fill="x", side="bottom")

        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Show details", command=self.show_details)
        self.menu.add_command(label="Copy row(s)", command=self.copy_rows)
        self.menu.add_command(label="Copy adapter number(s)",
                              command=lambda: self.copy_column(self.group_col, unique=True))
        self.menu.add_command(label="Copy connector P/N(s)",
                              command=lambda: self.copy_column(self._part_col()))

    def _options_changed(self):
        self.settings["group"] = self.group_var.get()
        self.settings["hide_scrapped"] = self.scrap_var.get()
        save_settings(self.settings)
        self.run_search()

    # ---------- database ----------
    def _startup(self, db_path):
        if self.demo:
            path = os.path.join(APP_DIR, "demo_adapters.sqlite")
            make_demo_db(path)
            self.open_db(path, sqlite=True)
            return
        if db_path:
            self.open_db(db_path, sqlite=db_path.endswith((".sqlite", ".db")))
            return
        found = sorted(glob.glob(os.path.join(APP_DIR, "*.accdb")) +
                       glob.glob(os.path.join(APP_DIR, "*.mdb")))
        last = self.settings.get("last_db")
        if last and os.path.exists(last):
            self.open_db(last)
        elif len(found) == 1:
            self.open_db(found[0])
        else:
            self.choose_db()

    def choose_db(self):
        path = filedialog.askopenfilename(
            title="Choose the adapter Access database",
            initialdir=APP_DIR,
            filetypes=[("Access databases", "*.accdb *.mdb"), ("All files", "*.*")])
        if path:
            self.open_db(path)

    def open_db(self, path, sqlite=False):
        try:
            backend = SqliteBackend(path) if sqlite else AccessBackend(path)
            # Go straight to the connectors table; only list the tables (a driver
            # catalog call that can fail to decode) when it isn't there.
            tables = ([(DEFAULT_TABLE, False)] if backend.has_table(DEFAULT_TABLE)
                      else backend.tables())
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Couldn't open the database:\n{path}\n\n{e}")
            return
        if self.backend:
            self.backend.close()
        self.backend = backend
        self.full_cache = {}
        self.table_map = {}
        fixed = next((n for n, _ in tables if n.lower() == DEFAULT_TABLE.lower()), None)
        if fixed:
            # Only the connectors table matters; hide the table picker.
            self.table_map[fixed] = fixed
            self.table_label.pack_forget()
            self.table_cb.pack_forget()
        else:
            for name, is_view in tables:
                self.table_map[f"{name}  (query)" if is_view else name] = name
            if not self.table_cb.winfo_ismapped():
                self.table_label.pack(side="left", before=self.quick_label)
                self.table_cb.pack(side="left", padx=(4, 16), before=self.quick_label)
        labels = list(self.table_map)
        self.table_cb["values"] = labels
        if not sqlite:
            self.settings["last_db"] = path
            save_settings(self.settings)
        self.title(f"{APP_TITLE} - {os.path.basename(path)}")
        if not labels:
            messagebox.showwarning(APP_TITLE, "No tables were found in that database.")
            return
        saved = self.settings.get("tables", {}).get(path)
        pick = (saved if saved in labels else
                next((l for l in labels if "adapter" in l.lower()), labels[0]))
        self.table_var.set(pick)
        self.load_table()

    def reload(self):
        if self.backend:
            self.open_db(self.backend.path, sqlite=isinstance(self.backend, SqliteBackend))

    def load_table(self):
        label = self.table_var.get()
        table = self.table_map.get(label)
        if not table:
            return
        self.settings.setdefault("tables", {})[self.backend.path] = label
        save_settings(self.settings)
        self.full_cache = {}
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            cols = self.backend.columns(table)
            self.columns = sorted(cols, key=lambda c: order_index(c.name, RESULT_ORDER))
            self.group_col = next((c.name for c in self.columns
                                   if c.name.lower() == GROUP_COLUMN.lower()), None)
            self._build_filters(table)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Couldn't read table '{table}':\n\n{e}")
            return
        finally:
            self.config(cursor="")
        has_scrap = any(c.name.lower() == r[0].lower() for c in self.columns
                        for r in SCRAP_RULES)
        self.scrap_chk.configure(state="normal" if has_scrap else "disabled")
        self.group_chk.configure(state="normal" if self.group_col else "disabled")
        searchable = [c for c in self.columns if c.kind != "binary"]
        ordered = sorted(searchable, key=lambda c: priority(c.name))
        self.in_cb["values"] = ["All columns"] + [c.name for c in ordered]
        self.in_var.set("All columns")
        self.quick_label.configure(
            text="Adapter / part number / keywords:" if self.group_col
            else "Part number / keywords:")
        self.run_search()

    def _distinct_values(self, table, col):
        if col.kind != "text" or col.long_text:
            return None
        try:
            cur = self.backend.execute(
                f"SELECT DISTINCT {qi(col.name)} FROM {qi(table)} "
                f"WHERE {qi(col.name)} IS NOT NULL")
            vals = cur.fetchmany(DISTINCT_LIMIT * 2 + 1)
        except Exception:
            return None
        if len(vals) > DISTINCT_LIMIT * 2:
            return None
        seen = {}
        for (v,) in vals:
            text = clean_text(v)
            if text in PLACEHOLDERS:
                continue
            seen.setdefault(text.lower(), text)   # merge different capitalisations
        if len(seen) > DISTINCT_LIMIT:
            return None
        return sorted(seen.values(), key=sort_key)

    def _build_filters(self, table):
        inner = self.filter_scroll.inner
        for w in inner.winfo_children():
            w.destroy()
        self.filter_rows = []
        cols = [c for c in self.columns if c.kind != "binary"]
        cols = sorted(cols, key=lambda c: priority(c.name))  # stable: keeps db order otherwise
        ttk.Label(inner, text="Pick from a list, or type and press Enter.\n"
                              "Typing matches anywhere in the value. * = wildcard.",
                  foreground="gray").grid(row=0, column=0, columnspan=2, sticky="w",
                                          padx=4, pady=(0, 6))
        for i, c in enumerate(cols, start=1):
            self.filter_rows.append(
                FilterRow(self, inner, c, self._distinct_values(table, c), i))
        self.filter_scroll.canvas.yview_moveto(0)

    # ---------- searching ----------
    def _quick_typed(self, event):
        if event.keysym in ("Return", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R"):
            return
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(350, self.run_search)

    def _exclude(self):
        return SCRAP_RULES if self.scrap_var.get() else []

    def _grouped(self):
        return bool(self.group_col and self.group_var.get())

    def _part_col(self):
        return next((c for c in self.result_cols if is_part_column(c)), None)

    def _order_by(self):
        order = [self.group_col] if self.group_col else []
        part = next((c.name for c in self.columns if is_part_column(c.name)), None)
        return order + ([part] if part else [])

    def _full_rows_by_group(self, table):
        """All (non-scrapped) rows, grouped by adapter. Cached per table load."""
        key = bool(self.scrap_var.get())
        if key not in self.full_cache:
            sql, params = build_query(table, self.columns, concat=self.backend.concat,
                                      exclude=self._exclude(), order_by=self._order_by())
            cur = self.backend.execute(sql, params)
            names = [d[0] for d in cur.description]
            gi = names.index(self.group_col)
            groups = {}
            for r in self._normalize(names, cur.fetchall()):
                groups.setdefault(r[gi], []).append(r)
            self.full_cache[key] = groups
        return self.full_cache[key]

    def run_search(self):
        self._after_id = None
        if not self.backend or not self.columns:
            return
        table = self.table_map.get(self.table_var.get())
        filters = []
        try:
            for fr in self.filter_rows:
                filters.extend(fr.conditions())
        except ValueError as e:
            self.status_var.set(f"Check filter - {e}")
            self.bell()
            return
        in_name = self.in_var.get()
        quick_col = next((c for c in self.columns if c.name == in_name), None)
        sql, params = build_query(table, self.columns, self.quick_var.get(), quick_col,
                                  filters, self.backend.concat, self.backend.date_param,
                                  exclude=self._exclude(), order_by=self._order_by())
        self.config(cursor="watch")
        self.update_idletasks()
        try:
            cur = self.backend.execute(sql, params)
            names = [d[0] for d in cur.description]
            rows = self._normalize(names, cur.fetchmany(MAX_ROWS + 1))
            full = self._full_rows_by_group(table) if self._grouped() else None
        except Exception as e:
            self.status_var.set(f"Search failed: {e}")
            return
        finally:
            self.config(cursor="")
        truncated = len(rows) > MAX_ROWS
        rows = rows[:MAX_ROWS]
        self.result_cols = names
        self.rows = rows
        self.groups = []
        if full is not None:
            gi = names.index(self.group_col)
            matched = {}
            for r in rows:
                matched.setdefault(r[gi], set()).add(r)
            for adapter, hits in matched.items():
                members = full.get(adapter) or list(hits)
                self.groups.append((adapter, [(r, r in hits) for r in members]))
        self.sort_state = (None, False)
        self._fill_tree()

        active = len(filters) + (1 if self.quick_var.get().strip() else 0)
        if full is not None:
            msg = (f"{len(self.groups)} adapter{'s' if len(self.groups) != 1 else ''} "
                   f"({len(rows)} matching connector{'s' if len(rows) != 1 else ''} "
                   f"in bold)")
        else:
            msg = f"{len(rows)}{'+' if truncated else ''} matching row{'s' if len(rows) != 1 else ''}"
        if truncated:
            msg += f" - only the first {MAX_ROWS} shown"
        if active:
            msg += f"   |   {active} filter{'s' if active != 1 else ''} active"
        if self._exclude():
            msg += "   |   scrapped hidden"
        msg += "   |   Double-click for details, right-click to copy"
        self.status_var.set(msg)

    def _normalize(self, names, rows):
        bools = {i for i, n in enumerate(names)
                 if any(c.name == n and c.kind == "bool" for c in self.columns)}
        return [tuple(bool(v) if i in bools and v is not None else v
                      for i, v in enumerate(r)) for r in rows]

    def _expand(self, open_):
        for iid in self.tree.get_children():
            self.tree.item(iid, open=open_)

    def _fill_tree(self):
        tree = self.tree
        tree.delete(*tree.get_children())
        self.iid_rows = {}
        grouped = self._grouped()
        display = [c for c in self.result_cols if not (grouped and c == self.group_col)]
        if list(tree["columns"]) != display:
            tree["columns"] = display
            sample = self.rows[:300]
            for name in display:
                i = self.result_cols.index(name)
                longest = max([len(name)] + [len(fmt_value(r[i])) for r in sample])
                width = max(50, min(300, longest * 7 + 18))
                tree.heading(name, text=name, command=lambda n=name: self.sort_by(n))
                tree.column(name, width=width, minwidth=40, stretch=False)
        idx = [self.result_cols.index(n) for n in display]
        for name in display:
            arrow = ""
            if self.sort_state[0] == name:
                arrow = " ▼" if self.sort_state[1] else " ▲"
            tree.heading(name, text=name + arrow)

        if grouped:
            tree.configure(show="tree headings")
            head = self.group_col
            if self.sort_state[0] == head:
                head += " ▼" if self.sort_state[1] else " ▲"
            tree.heading("#0", text=head, command=lambda: self.sort_by(self.group_col))
            tree.column("#0", width=170, minwidth=90, stretch=False)
            part = self._part_col()
            pi = self.result_cols.index(part) if part else None
            many = len(self.groups) > 400
            for g, (adapter, members) in enumerate(self.groups):
                gid = f"g{g}"
                summary = [""] * len(display)
                if pi is not None and part in display:
                    pns = [fmt_value(r[pi]) for r, _ in members if fmt_value(r[pi])]
                    summary[display.index(part)] = ", ".join(pns)
                n = len(members)
                tree.insert("", "end", iid=gid, open=not many,
                            text=f"{fmt_value(adapter)}   ({n} conn.)",
                            values=summary, tags=("group",))
                self.iid_rows[gid] = [r for r, _ in members]
                for j, (r, hit) in enumerate(members):
                    cid = f"g{g}_{j}"
                    tree.insert(gid, "end", iid=cid, values=[fmt_value(r[i]) for i in idx],
                                tags=("match",) if hit else ("other",))
                    self.iid_rows[cid] = [r]
        else:
            tree.configure(show="headings")
            for k, r in enumerate(self.rows):
                iid = f"r{k}"
                tree.insert("", "end", iid=iid, values=[fmt_value(r[i]) for i in idx],
                            tags=("odd",) if k % 2 else ())
                self.iid_rows[iid] = [r]

    def sort_by(self, name):
        idx = self.result_cols.index(name)
        desc = self.sort_state[0] == name and not self.sort_state[1]
        key = lambda r: sort_key(r[idx])
        if self._grouped():
            for _, members in self.groups:
                members.sort(key=lambda m: key(m[0]), reverse=desc)
            first = lambda g: next((key(r) for r, hit in g[1] if hit), (3, ""))
            self.groups.sort(key=first, reverse=desc)
        else:
            self.rows.sort(key=key, reverse=desc)
        self.sort_state = (name, desc)
        self._fill_tree()

    def clear_all(self):
        self.quick_var.set("")
        self.in_var.set("All columns")
        for fr in self.filter_rows:
            fr.clear()
        self.run_search()
        self.quick_entry.focus_set()

    # ---------- results actions ----------
    def _selected(self):
        out, seen = [], set()
        for iid in self.tree.selection():
            for r in self.iid_rows.get(iid, []):
                if id(r) not in seen:
                    seen.add(id(r))
                    out.append(r)
        return out

    def _context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.menu.entryconfigure(2, state="normal" if self.group_col else "disabled")
            self.menu.tk_popup(event.x_root, event.y_root)

    def copy_rows(self):
        rows = self._selected()
        if not rows:
            return
        text = "\t".join(self.result_cols) + "\n" + \
            "\n".join("\t".join(fmt_value(v) for v in r) for r in rows)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set(f"Copied {len(rows)} row(s) - paste into Excel.")

    def copy_column(self, name, unique=False):
        if not name:
            return
        i = self.result_cols.index(name)
        vals = [fmt_value(r[i]) for r in self._selected()]
        if unique:
            vals = list(dict.fromkeys(vals))
        if vals:
            self.clipboard_clear()
            self.clipboard_append("\n".join(vals))
            self.status_var.set(f"Copied {len(vals)} value(s) from {name}.")

    def show_details(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        rows = self.iid_rows.get(sel[0], [])
        if not rows:
            return
        title_col = self.group_col or self._part_col() or self.result_cols[0]
        ti = self.result_cols.index(title_col)
        win = tk.Toplevel(self)
        win.title(f"Details - {fmt_value(rows[0][ti])}")
        win.geometry(f"{min(1200, 260 + 230 * len(rows))}x600")
        win.transient(self)
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=8, pady=8)
        cols = ["field"] + [f"c{k}" for k in range(len(rows))]
        t = ttk.Treeview(frame, columns=cols, show="headings")
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=t.xview)
        t.configure(xscrollcommand=hsb.set)
        t.heading("field", text="Field")
        t.column("field", width=150, stretch=False)
        for k in range(len(rows)):
            t.heading(f"c{k}", text=f"Connector {k + 1}" if len(rows) > 1 else "Value")
            t.column(f"c{k}", width=220, stretch=len(rows) == 1)
        for i, name in enumerate(self.result_cols):
            t.insert("", "end", values=[name] + [fmt_value(r[i]) for r in rows])
        t.pack(fill="both", expand=True)
        hsb.pack(fill="x")

        def copy_all():
            win.clipboard_clear()
            win.clipboard_append("\n".join(
                "\t".join([name] + [fmt_value(r[i]) for r in rows])
                for i, name in enumerate(self.result_cols)))
        b = ttk.Frame(win, padding=(8, 0, 8, 8))
        b.pack(fill="x")
        ttk.Button(b, text="Copy all", command=copy_all).pack(side="left")
        ttk.Button(b, text="Close", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())

    def export_csv(self):
        if not self.result_cols:
            return
        table = self.table_map.get(self.table_var.get(), "results")
        path = filedialog.asksaveasfilename(
            title="Export results", defaultextension=".csv",
            initialfile=f"{re.sub(r'[^A-Za-z0-9_-]+', '_', table)}_search.csv",
            filetypes=[("CSV (Excel)", "*.csv")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.writer(f)
                if self._grouped():
                    w.writerow(["Matched"] + self.result_cols)
                    n = 0
                    for _, members in self.groups:
                        for r, hit in members:
                            w.writerow(["Yes" if hit else ""] + [fmt_value(v) for v in r])
                            n += 1
                else:
                    w.writerow(self.result_cols)
                    w.writerows([[fmt_value(v) for v in r] for r in self.rows])
                    n = len(self.rows)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Export failed:\n\n{e}")
            return
        self.status_var.set(f"Exported {n} rows to {path}")


def main():
    args = sys.argv[1:]
    demo = "--demo" in args
    paths = [a for a in args if not a.startswith("--")]
    app = App(db_path=paths[0] if paths else None, demo=demo)
    app.mainloop()


if __name__ == "__main__":
    main()
