#!/usr/bin/env python3
"""
Adapter Lookup - search, and optionally edit, the CONNECTORS table of an
Access database of adapters.

* Finds the .accdb / .mdb (shared Z: drive copy, last used, or next to the app).
* Search box: adapter, part number or keywords (* is a wildcard).
* Filters panel: pick-lists, number ranges, Yes/No, dates - one per column.
  Active filters show as chips above the results; click a chip's x to remove it.
* Details panel on the right shows every field of the selected connector.
* Editing (off by default): click "Enable editing" in the top bar. You can then
  add, edit, duplicate and delete connector rows. Every change is written to a
  change log next to the database, and Ctrl+Z undoes your changes from this session.

Run with --demo to try it on built-in sample data (no Access driver needed).
"""
import csv
import datetime as dt
import decimal
import getpass
import glob
import json
import os
import re
import shutil
import struct
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

APP_TITLE = "Adapter Lookup"
if getattr(sys, "frozen", False):
    # Running as a packaged .exe: look next to the .exe, not in the temp unpack folder.
    APP_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(APP_DIR, "adapter_lookup_settings.json")
MAX_ROWS = 20000         # most connector rows shown at once
ZOOM_DEFAULT, ZOOM_MIN, ZOOM_MAX = 10, 7, 24   # results table font size (points)
DISTINCT_LIMIT = 250     # columns with this many values or fewer get a pick-list
COMMON_FILTERS = 8       # this many filters are always shown; the rest sit under "More filters"

# ---- Settings for the SIGBOARDS database --------------------------------- #
DEFAULT_DB = r"Z:\Engineering\Test Engineering\Test\Test Board DB\SIGBOARDS.mdb"  # used when none is saved
DEFAULT_TABLE = "CONNECTORS"     # the only table used when it exists
GROUP_COLUMN = "BOARD/ADAPTER"   # one adapter = all rows with the same value here
PART_COLUMN = "CONNECTOR P/N"    # the connector part number
REQUIRED_COLUMNS = [GROUP_COLUMN, PART_COLUMN]   # must be filled in when adding / editing
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
FRIENDLY_LABELS = {"Positions": "Positions (pins)", "BOARD/ADAPTER": "Adapter / board",
                   "CONNECTOR P/N": "Connector P/N"}
PLACEHOLDERS = {"", "-", "--", "?", "??"}  # left out of pick-lists

# Filters for columns whose names contain these words are listed first.
PRIORITY_HINTS = ["part", "pn", "sku", "model", "connector", "interface", "type",
                  "pin", "size", "gender", "length", "manufacturer", "mfr",
                  "vendor", "brand"]

TEXT_TYPES = {"VARCHAR", "CHAR", "TEXT", "WVARCHAR", "WCHAR", "NVARCHAR", "NCHAR",
              "LONGCHAR", "MEMO", "WLONGVARCHAR", "LONGVARCHAR", "HYPERLINK", "GUID"}
LONG_TEXT_TYPES = {"LONGCHAR", "MEMO", "WLONGVARCHAR", "LONGVARCHAR", "HYPERLINK"}
INT_TYPES = {"INTEGER", "INT", "SMALLINT", "BYTE", "TINYINT", "COUNTER", "BIGINT"}
NUMBER_TYPES = INT_TYPES | {"DOUBLE", "REAL", "FLOAT", "CURRENCY", "DECIMAL", "NUMERIC",
                            "MONEY"}
DATE_TYPES = {"DATETIME", "DATE", "TIMESTAMP", "TIME"}
BOOL_TYPES = {"BIT", "BOOLEAN", "YESNO"}
BINARY_TYPES = {"LONGBINARY", "VARBINARY", "BINARY", "IMAGE", "OLEOBJECT", "BLOB"}

# ---- Look ------------------------------------------------------------------ #
C = dict(
    bg="#f3f5f8", surface="#ffffff", border="#d8dee6", text="#1d2733", muted="#6a7684",
    faint="#9aa5b1", primary="#2a5ea8", primary_dark="#1e4682", primary_light="#e8eff9",
    header="#1c3557", header_text="#ffffff", header_sub="#b9c8dc", header_btn="#2c4b74",
    heading="#eef1f5", match="#fff3c7", group="#e5edf8", alt="#f8fafc", danger="#b42318",
    danger_light="#fdecea", edit="#b45309", edit_light="#fff4e0", ok="#1f7a4d",
    chip="#e3ecf8", chip_text="#1e4682",
)


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
    if name.lower() == PART_COLUMN.lower():
        return True
    n = name.lower().replace(" ", "").replace("_", "").replace("/", "").replace(".", "")
    return ("part" in n or "sku" in n or n in ("pn", "partno", "model", "modelno")
            or n.startswith("pn") or n.endswith("pn"))


def order_index(name, order):
    lower = [o.lower() for o in order]
    return lower.index(name.lower()) if name.lower() in lower else len(order)


def clean_text(v):
    """Collapse line breaks / repeated spaces that are stored in some cells."""
    return " ".join(str(v).split())


def label_for(name):
    return FRIENDLY_LABELS.get(name, name)


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


def plural(n, word, many=None):
    return f"{n:,} {word if n == 1 else (many or word + 's')}"


class Column:
    def __init__(self, name, type_name, size=None):
        self.name = name
        self.type_name = (type_name or "").upper()
        self.kind = classify(type_name)
        self.long_text = self.type_name in LONG_TEXT_TYPES
        self.is_int = self.type_name.split("(")[0] in INT_TYPES
        self.size = size if (self.kind == "text" and not self.long_text and size
                             and size < 100000) else None

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

    def __init__(self, path, writable=False):
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
        self.writable = writable
        conn_str = f"DRIVER={{{driver}}};DBQ={path};" + ("" if writable else "ReadOnly=1;")
        # Writes run in a transaction so a failed change is rolled back.
        self.conn = pyodbc.connect(conn_str, autocommit=not writable)

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
            out.append(Column(d[0], name, d[3]))
        return out

    def unique_key(self, table):
        """Columns of the table's primary key / first unique index, or None."""
        try:
            idx = {}
            for r in self.conn.cursor().statistics(table, unique=True):
                if r.index_name and r.column_name:
                    idx.setdefault(r.index_name, []).append((r.ordinal_position or 0,
                                                             r.column_name))
            if not idx:
                return None
            name = "PrimaryKey" if "PrimaryKey" in idx else sorted(idx)[0]
            return [c for _, c in sorted(idx[name])]
        except Exception:
            return None

    def execute(self, sql, params=()):
        cur = self.conn.cursor()
        cur.execute(sql, params)
        return cur

    def commit(self):
        if self.writable:
            self.conn.commit()

    def rollback(self):
        if self.writable:
            try:
                self.conn.rollback()
            except Exception:
                pass

    def close(self):
        try:
            self.conn.close()
        except Exception:
            pass


class SqliteBackend:
    """Used for --demo and for testing; mirrors the Access back-end."""
    concat = "||"

    def __init__(self, path, writable=False):
        import sqlite3
        self.path = path
        self.writable = writable
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
        out = []
        for r in rows:
            m = re.search(r"\((\d+)\)", r[2] or "")
            out.append(Column(r[1], r[2], int(m.group(1)) if m else None))
        return out

    def unique_key(self, table):
        return None

    def execute(self, sql, params=()):
        return self.conn.execute(sql, list(params))

    def commit(self):
        self.conn.commit()

    def rollback(self):
        self.conn.rollback()

    def close(self):
        self.conn.close()


def open_backend(path, sqlite=False, writable=False):
    return SqliteBackend(path, writable) if sqlite else AccessBackend(path, writable)


# --------------------------------------------------------------------------- #
#  Writing changes
# --------------------------------------------------------------------------- #
class WriteError(Exception):
    pass


class RecordStore:
    """Adds, changes and deletes single rows.

    The CONNECTORS table has no ID column, so a row is found by matching every
    one of its values (or the table's unique key, if it has one). A change is
    only made when exactly one row matches, so nothing else is touched.
    """

    def __init__(self, backend, table, columns):
        self.backend = backend
        self.table = table
        self.columns = [c for c in columns if c.kind != "binary"]
        names = {c.name for c in self.columns}
        key = backend.unique_key(table)
        self.key_cols = key if key and all(k in names for k in key) else None

    def _where(self, values):
        cols = [c for c in self.columns if c.name in values and not c.long_text]
        if self.key_cols and all(values.get(k) is not None for k in self.key_cols):
            cols = [c for c in cols if c.name in self.key_cols]
        where, params = [], []
        for c in cols:
            v, q = values[c.name], qi(c.name)
            if v is None:
                where.append(f"{q} IS NULL")
            elif c.kind == "bool":
                where.append(f"{q} <> 0" if v else f"{q} = 0")
            elif isinstance(v, float) and not v.is_integer():
                where.append(f"Abs({q} - ?) < 0.000001")
                params.append(v)
            else:
                where.append(f"{q} = ?")
                params.append(v)
        if not where:
            raise WriteError("This row has no values that can be used to find it.")
        return " AND ".join(where), params

    def count(self, values):
        w, p = self._where(values)
        return self.backend.execute(f"SELECT COUNT(*) FROM {qi(self.table)} WHERE {w}",
                                    p).fetchone()[0]

    def _find_one(self, values):
        n = self.count(values)
        if n == 0:
            raise WriteError("This row couldn't be found. It may have just been changed or "
                             "deleted by someone else.\n\nPress F5 to reload, then try again.")
        if n > 1:
            raise WriteError(f"There are {n} identical copies of this row, so it can't be "
                             "changed on its own from here. Change one of them in Access, "
                             "or delete the extra copies there.")

    def _run(self, sql, params):
        try:
            cur = self.backend.execute(sql, params)
            if cur.rowcount not in (1, -1):
                raise WriteError(f"Expected to change 1 row but {cur.rowcount} would change. "
                                 "Nothing was saved.")
            self.backend.commit()
        except WriteError:
            self.backend.rollback()
            raise
        except Exception as e:
            self.backend.rollback()
            raise WriteError(f"The database refused the change:\n\n{e}")

    def insert(self, values):
        values = {c.name: values.get(c.name) for c in self.columns}
        if not self.key_cols and self.count(values):
            raise WriteError("An identical row already exists. Change at least one field "
                             "(for example Qty or LOCATION) so the two rows can be told apart.")
        names = [c.name for c in self.columns]
        sql = (f"INSERT INTO {qi(self.table)} (" + ", ".join(qi(n) for n in names) +
               ") VALUES (" + ", ".join("?" for _ in names) + ")")
        self._run(sql, [values[n] for n in names])
        return values

    def update(self, before, changes):
        changes = {k: v for k, v in changes.items() if before.get(k) != v}
        if not changes:
            return dict(before)
        self._find_one(before)
        after = dict(before)
        after.update(changes)
        if not self.key_cols and self.count(after):
            raise WriteError("Another row already has exactly these values. Change at "
                             "least one more field so the two rows can be told apart.")
        w, p = self._where(before)
        sets = ", ".join(f"{qi(k)} = ?" for k in changes)
        self._run(f"UPDATE {qi(self.table)} SET {sets} WHERE {w}",
                  list(changes.values()) + p)
        return after

    def delete(self, before):
        self._find_one(before)
        w, p = self._where(before)
        self._run(f"DELETE FROM {qi(self.table)} WHERE {w}", p)


class ChangeLog:
    """Appends one line per change to a CSV next to the database."""
    FIELDS = ["When", "User", "Action", "Adapter", "Connector P/N", "Details"]

    def __init__(self, db_path):
        base = os.path.splitext(os.path.basename(db_path))[0]
        self.paths = [os.path.join(os.path.dirname(os.path.abspath(db_path)),
                                   f"{base}_change_log.csv"),
                      os.path.join(APP_DIR, f"{base}_change_log.csv")]
        self.path = self.paths[0]

    @staticmethod
    def user():
        try:
            return os.environ.get("USERNAME") or getpass.getuser()
        except Exception:
            return "unknown"

    def write(self, action, adapter, part, details):
        row = [dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), self.user(), action,
               adapter, part, details]
        for path in self.paths:
            try:
                new = not os.path.exists(path)
                with open(path, "a", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    if new:
                        w.writerow(self.FIELDS)
                    w.writerow(row)
                self.path = path
                return True
            except Exception:
                continue
        return False


def describe_values(values):
    return "; ".join(f"{k}={fmt_value(v)}" for k, v in values.items() if fmt_value(v))


def describe_changes(before, after):
    return "; ".join(f"{k}: '{fmt_value(before.get(k))}' -> '{fmt_value(v)}'"
                     for k, v in after.items() if before.get(k) != v)


# --------------------------------------------------------------------------- #
#  Demo data
# --------------------------------------------------------------------------- #
def make_demo_db(path):
    """A small CONNECTORS table shaped like the real one."""
    import random
    import sqlite3
    if os.path.exists(path):
        os.remove(path)
    rnd = random.Random(7)
    conn = sqlite3.connect(path)
    conn.execute("""CREATE TABLE CONNECTORS (
        [BOARD/ADAPTER] VARCHAR(50), [CONNECTOR P/N] VARCHAR(50), Qty INTEGER,
        DESCRIPTION VARCHAR(255), Positions INTEGER, Series VARCHAR(50),
        Category VARCHAR(50), [Contact Type] VARCHAR(50), [Housing Type] VARCHAR(50),
        Arrangement VARCHAR(50), Shell VARCHAR(50), Keying VARCHAR(50), Pitch VARCHAR(50),
        Manufacturer VARCHAR(50), [Relevant Assemblies] VARCHAR(255), LOCATION VARCHAR(50),
        OBSOLETE VARCHAR(10), Port VARCHAR(50), CUSTOMER VARCHAR(50),
        [Voltage Rating] VARCHAR(50))""")
    parts = [("D38999/26WD35SN", "Circular Military", 37, "Socket", "Circular", "D38999"),
             ("D38999/20WB98PN", "Circular Military", 6, "Pin", "Circular", "D38999"),
             ("DB25P", "D-Sub", 25, "Pin", "D-Sub", "--"),
             ("DB25S", "D-Sub", 25, "Socket", "D-Sub", "--"),
             ("DE9S", "D-Sub", 9, "Socket", "D-Sub", "--"),
             ("M83513/04-C11N", "Micro-D", 37, "Socket", "Micro-D", "M83513"),
             ("43025-1000", "Plastic Crimp Housing", 10, "Socket", "Plastic", "Micro-Fit"),
             ("1-480698-0", "Plastic Crimp Housing", 4, "Pin", "Plastic", "Mate-N-Lok"),
             ("SMA-J-P", "RF", 1, "Socket", "Coax", "SMA"),
             ("BNC-FB", "RF", 1, "Socket", "Coax", "BNC")]
    mfrs = ["Amphenol", "TE Connectivity", "Molex", "Glenair", "ITT Cannon"]
    rows = []
    for i in range(1, 161):
        adapter = rnd.choice([f"ADAPTER {i}", f"{94000 + i}-1", f"CS-{10100 + i}"])
        scrapped = rnd.random() < .1
        for _ in range(rnd.randint(1, 4)):
            pn, cat, pos, ct, ht, ser = rnd.choice(parts)
            rows.append((adapter, pn, rnd.choice([1, 1, 1, 2]),
                         f"{pos} pos {cat.lower()} {ct.lower()}", pos, ser,
                         "SCRAPPED" if scrapped and rnd.random() < .5 else cat, ct, ht,
                         "--", "--", "--", "--", rnd.choice(mfrs), "--",
                         f"Shelf {rnd.choice('ABCDE')}{rnd.randint(1, 9)}", "No",
                         "SCRAP" if scrapped else f"J{rnd.randint(1, 8)}",
                         rnd.choice(["--", "Boeing", "Raytheon", "Honeywell"]), "--"))
    conn.executemany("INSERT INTO CONNECTORS VALUES (" + ",".join("?" * 20) + ")", rows)
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
#  Small widgets
# --------------------------------------------------------------------------- #
def wheel_step(event):
    return -1 if (getattr(event, "num", 0) == 4 or getattr(event, "delta", 0) > 0) else 1


class ScrollFrame(ttk.Frame):
    def __init__(self, master, style="Card.TFrame", **kw):
        super().__init__(master, style=style, **kw)
        bg = C["surface"] if style == "Card.TFrame" else C["bg"]
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background=bg)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = ttk.Frame(self.canvas, style=style)
        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._win, width=e.width))
        self.canvas.configure(yscrollcommand=vsb.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind_all(seq, self._wheel, add="+")

    def _inside(self, w):
        while w is not None:
            if w is self:
                return True
            w = getattr(w, "master", None)
        return False

    def _wheel(self, event):
        try:
            if not self.winfo_exists() or not self._inside(event.widget):
                return
            self.canvas.yview_scroll(wheel_step(event), "units")
        except tk.TclError:
            pass

    def wheel_and_stop(self, event):
        """Bound on drop-down boxes so the wheel scrolls the panel, not the value."""
        self.canvas.yview_scroll(wheel_step(event), "units")
        return "break"

    def protect(self, widget):
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            widget.bind(seq, self.wheel_and_stop)


class Tooltip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tip, self._job = widget, text, None, None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _):
        self._job = self.widget.after(500, self._show)

    def _show(self):
        if self.tip or not self.text:
            return
        x = self.widget.winfo_rootx() + 8
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, background="#2b3440", foreground="white",
                 padx=8, pady=4, justify="left", wraplength=320).pack()

    def _hide(self, _=None):
        if self._job:
            self.widget.after_cancel(self._job)
            self._job = None
        if self.tip:
            self.tip.destroy()
            self.tip = None


class Placeholder:
    """Grey hint text shown inside an empty entry."""

    def __init__(self, entry, var, text):
        self.entry, self.var = entry, var
        self.label = tk.Label(entry, text=text, background=C["surface"],
                              foreground=C["faint"], cursor="xterm",
                              font=entry.cget("font") or None)
        self.label.bind("<Button-1>", lambda e: entry.focus_set())
        var.trace_add("write", lambda *_: self.update())
        entry.bind("<FocusIn>", lambda e: self.update(), add="+")
        entry.bind("<FocusOut>", lambda e: self.update(), add="+")
        entry.after(10, self.update)

    def update(self):
        if self.var.get():
            self.label.place_forget()
        else:
            self.label.place(x=6, rely=0.5, anchor="w")


def link(parent, text, command, color=None, bg=None, font=None):
    """A clickable text label."""
    lbl = tk.Label(parent, text=text, foreground=color or C["primary"],
                   background=bg or C["surface"], cursor="hand2", font=font)
    lbl.bind("<Button-1>", lambda e: command())
    return lbl


# --------------------------------------------------------------------------- #
#  Filters
# --------------------------------------------------------------------------- #
class FilterRow:
    """One filter in the left panel: a label with an input under it."""

    def __init__(self, app, parent, col, distinct):
        self.app = app
        self.col = col
        self.vars = []
        self.choices = {}
        self.widgets = []
        self.frame = ttk.Frame(parent, style="Card.TFrame")
        self.frame.pack(fill="x", padx=12, pady=(0, 10))
        top = ttk.Frame(self.frame, style="Card.TFrame")
        top.pack(fill="x")
        ttk.Label(top, text=label_for(col.name), style="FieldLabel.TLabel").pack(side="left")
        self.clear_link = link(top, "clear", self._clear_and_search, color=C["faint"],
                               font=app.f_small)
        cell = ttk.Frame(self.frame, style="Card.TFrame")
        cell.pack(fill="x", pady=(3, 0))

        if col.kind == "bool":
            v = tk.StringVar(value="Any")
            cb = ttk.Combobox(cell, textvariable=v, values=["Any", "Yes", "No"],
                              state="readonly", width=8)
            cb.pack(side="left")
            cb.bind("<<ComboboxSelected>>", lambda e: app.run_search())
            self.vars = [v]
            self.widgets = [cb]
        elif col.kind in ("number", "date"):
            lo, hi = tk.StringVar(), tk.StringVar()
            w = 7 if col.kind == "number" else 11
            e1 = ttk.Entry(cell, textvariable=lo, width=w)
            e2 = ttk.Entry(cell, textvariable=hi, width=w)
            e1.pack(side="left", fill="x", expand=True)
            ttk.Label(cell, text="  to  ", style="CardMuted.TLabel").pack(side="left")
            e2.pack(side="left", fill="x", expand=True)
            Placeholder(e1, lo, "min" if col.kind == "number" else "from")
            Placeholder(e2, hi, "max" if col.kind == "number" else "to")
            for e in (e1, e2):
                e.bind("<Return>", lambda ev: app.run_search())
                e.bind("<KeyRelease>", app.schedule_search, add="+")
            self.vars = [lo, hi]
            self.widgets = [e1, e2]
        else:
            v = tk.StringVar()
            if distinct is not None:
                self.choices = {d.lower(): d for d in distinct}
                w = ttk.Combobox(cell, textvariable=v, values=[""] + distinct)
                w.bind("<<ComboboxSelected>>", lambda e: app.run_search())
                Placeholder(w, v, "Any")
            else:
                w = ttk.Entry(cell, textvariable=v)
                Placeholder(w, v, "contains...")
            w.pack(fill="x")
            w.bind("<Return>", lambda e: app.run_search())
            w.bind("<KeyRelease>", app.schedule_search, add="+")
            self.vars = [v]
            self.widgets = [w]
        for var in self.vars:
            var.trace_add("write", lambda *_: self._show_clear())
        for w in self.widgets:
            app.filter_scroll.protect(w)

    def _show_clear(self):
        if self.is_active():
            self.clear_link.pack(side="right")
        else:
            self.clear_link.pack_forget()

    def _clear_and_search(self):
        self.clear()
        self.app.run_search()

    def add_choice(self, value):
        """Make a newly saved value appear in this pick-list."""
        if not self.choices or value in (None, "") or not isinstance(value, str):
            return
        text = clean_text(value)
        if text in PLACEHOLDERS or text.lower() in self.choices:
            return
        self.choices[text.lower()] = text
        values = sorted(self.choices.values(), key=sort_key)
        self.widgets[0].configure(values=[""] + values)

    def is_active(self):
        if self.col.kind == "bool":
            return self.vars[0].get() != "Any"
        return any(v.get().strip() for v in self.vars)

    def clear(self):
        if self.col.kind == "bool":
            self.vars[0].set("Any")
        else:
            for v in self.vars:
                v.set("")

    def chip_text(self):
        name = label_for(self.col.name)
        if self.col.kind == "bool":
            return f"{name}: {self.vars[0].get()}"
        if self.col.kind in ("number", "date"):
            lo, hi = (v.get().strip() for v in self.vars)
            if lo and hi:
                return f"{name}: {lo} to {hi}"
            return f"{name} \u2265 {lo}" if lo else f"{name} \u2264 {hi}"
        text = self.vars[0].get().strip()
        if clean_text(text).lower() in self.choices and "*" not in text:
            return f"{name}: {text}"
        return f"{name} contains \u201c{text}\u201d"

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
                    raise ValueError(f"{label_for(c.name)}: {e}" if c.kind == "date"
                                     else f"{label_for(c.name)}: '{var.get()}' is not a number")
                if val is not None:
                    out.append((c, op, val))
            return out
        text = self.vars[0].get().strip()
        if not text:
            return []
        if clean_text(text).lower() in self.choices and "*" not in text:
            return [(c, "eq", text)]
        return [(c, "contains", text)]


# --------------------------------------------------------------------------- #
#  Add / edit form
# --------------------------------------------------------------------------- #
class RecordDialog(tk.Toplevel):
    """Form for adding or editing one connector row. Calls on_save(values) -> bool."""

    def __init__(self, app, title, columns, values, pick_lists, on_save, intro=""):
        super().__init__(app)
        self.app = app
        self.columns = columns
        self.on_save = on_save
        self.original = dict(values)
        self.pick_lists = pick_lists
        self.inputs = {}
        self.labels = {}
        self.title(title)
        self.configure(background=C["bg"])
        self.transient(app)
        self.geometry("780x700")
        self.minsize(560, 420)

        head = ttk.Frame(self, style="TFrame", padding=(20, 16, 20, 8))
        head.pack(fill="x")
        ttk.Label(head, text=title, style="DialogTitle.TLabel").pack(anchor="w")
        if intro:
            ttk.Label(head, text=intro, style="Muted.TLabel", wraplength=720,
                      justify="left").pack(anchor="w", pady=(4, 0))

        foot = ttk.Frame(self, style="TFrame", padding=(20, 10, 20, 16))
        foot.pack(fill="x", side="bottom")
        self.error = ttk.Label(foot, text="", style="Error.TLabel", wraplength=480,
                               justify="left")
        self.error.pack(side="left", fill="x", expand=True)
        ttk.Button(foot, text="Cancel", command=self.destroy).pack(side="right")
        self.save_btn = ttk.Button(foot, text="Save", style="Accent.TButton",
                                   command=self.save)
        self.save_btn.pack(side="right", padx=(0, 8))

        card = ttk.Frame(self, style="Card.TFrame")
        card.pack(fill="both", expand=True, padx=20)
        self.scroll = ScrollFrame(card)
        self.scroll.pack(fill="both", expand=True, padx=1, pady=1)
        grid = self.scroll.inner
        grid.columnconfigure(0, weight=1, uniform="f")
        grid.columnconfigure(1, weight=1, uniform="f")

        r = c = 0
        for col in columns:
            wide = col.long_text or col.name.upper() in ("DESCRIPTION", "RELEVANT ASSEMBLIES")
            if wide and c == 1:
                r, c = r + 1, 0
            cell = ttk.Frame(grid, style="Card.TFrame")
            cell.grid(row=r, column=c, columnspan=2 if wide else 1, sticky="ew",
                      padx=14, pady=(12 if r == 0 else 4, 6))
            req = col.name in REQUIRED_COLUMNS
            text = label_for(col.name) + (" *" if req else "")
            lbl = ttk.Label(cell, text=text, style="FieldLabel.TLabel")
            lbl.pack(anchor="w")
            self.labels[col.name] = lbl
            self.inputs[col.name] = self._make_input(cell, col, values.get(col.name))
            hint = self._hint(col)
            if hint:
                ttk.Label(cell, text=hint, style="Hint.TLabel").pack(anchor="w")
            if wide:
                r, c = r + 1, 0
            else:
                r, c = (r, 1) if c == 0 else (r + 1, 0)

        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Control-s>", lambda e: self.save())
        self.bind("<Control-Return>", lambda e: self.save())
        self.after(50, self._focus_first)
        self.after(100, self._grab)

    def _grab(self):
        try:
            self.grab_set()
        except tk.TclError:
            pass

    def _hint(self, col):
        bits = []
        if col.kind == "number":
            bits.append("whole number" if col.is_int else "number")
        elif col.kind == "date":
            bits.append("YYYY-MM-DD")
        if col.size and col.size < 255:
            bits.append(f"up to {col.size} characters")
        if col.name in self.pick_lists:
            bits.append("pick or type a new value")
        return ", ".join(bits)

    def _make_input(self, parent, col, value):
        if col.kind == "bool":
            var = tk.BooleanVar(value=bool(value))
            ttk.Checkbutton(parent, text="Yes", variable=var,
                            style="Card.TCheckbutton").pack(anchor="w", pady=(2, 0))
            return ("bool", var)
        if col.long_text:
            t = tk.Text(parent, height=3, wrap="word", relief="solid", borderwidth=1,
                        font=self.app.f_base, highlightthickness=0)
            t.configure(highlightbackground=C["border"])
            t.insert("1.0", "" if value is None else str(value))
            t.pack(fill="x", pady=(3, 0))
            return ("text", t)
        var = tk.StringVar(value=fmt_value(value) if col.kind != "text" else
                           ("" if value is None else str(value).replace("\r", " ")
                            .replace("\n", " ")))
        if col.name in self.pick_lists:
            w = ttk.Combobox(parent, textvariable=var, values=self.pick_lists[col.name])
            self.scroll.protect(w)
        else:
            w = ttk.Entry(parent, textvariable=var)
        w.pack(fill="x", pady=(3, 0))
        var.trace_add("write", lambda *_: self._mark_changed(col.name))
        return ("var", var)

    def _mark_changed(self, name):
        try:
            changed = self._value_text(name) != self._orig_text(name)
        except Exception:
            changed = False
        self.labels[name].configure(style="FieldLabelChanged.TLabel" if changed
                                    else "FieldLabel.TLabel")

    def _orig_text(self, name):
        """The original value as the form shows it."""
        v = self.original.get(name)
        col = next(c for c in self.columns if c.name == name)
        if v is None:
            return ""
        if col.kind != "text":
            return fmt_value(v)
        return str(v).strip() if col.long_text else clean_text(v)

    def _value_text(self, name):
        kind, w = self.inputs[name]
        if kind == "bool":
            return "Yes" if w.get() else "No"
        if kind == "text":
            return w.get("1.0", "end-1c").strip()
        return clean_text(w.get())

    def _focus_first(self):
        for child in self.scroll.inner.winfo_children():
            for sub in child.winfo_children():
                if isinstance(sub, (ttk.Entry, ttk.Combobox)):
                    if not sub.get():
                        sub.focus_set()
                        return
        self.save_btn.focus_set()

    def collect(self):
        """Return {column: value} converted to database types; raises ValueError."""
        out, problems = {}, []
        for col in self.columns:
            kind, w = self.inputs[col.name]
            name = label_for(col.name)
            if kind == "bool":
                out[col.name] = bool(w.get())
                continue
            text = (w.get("1.0", "end-1c").strip() if kind == "text"
                    else clean_text(w.get()))
            if not text and col.name in REQUIRED_COLUMNS:
                problems.append(f"{name} is required.")
                continue
            if col.name in self.original and text == self._orig_text(col.name):
                # Untouched: keep the stored value exactly as it is.
                out[col.name] = self.original[col.name]
                continue
            if not text:
                out[col.name] = None
                continue
            if col.kind == "number":
                try:
                    num = float(text.replace(",", ""))
                except ValueError:
                    problems.append(f"{name} must be a number.")
                    continue
                if col.is_int:
                    if not num.is_integer():
                        problems.append(f"{name} must be a whole number.")
                        continue
                    num = int(num)
                out[col.name] = num
            elif col.kind == "date":
                try:
                    out[col.name] = parse_date(text)
                except ValueError as e:
                    problems.append(f"{name}: {e}")
            else:
                text = " ".join(text.split()) if not col.long_text else text
                # Use the existing spelling when only the capitalisation differs,
                # so the data doesn't gain another "CIRCULAR Military" variant.
                known = {v.lower(): v for v in self.pick_lists.get(col.name, [])}
                text = known.get(text.lower(), text)
                if col.size and len(text) > col.size:
                    problems.append(f"{name} is {len(text)} characters long; "
                                    f"the most it can hold is {col.size}.")
                out[col.name] = text
        if problems:
            raise ValueError("\n".join(problems))
        return out

    def save(self):
        try:
            values = self.collect()
        except ValueError as e:
            self.error.configure(text=str(e))
            self.bell()
            return
        self.error.configure(text="")
        if self.on_save(values, self):
            self.destroy()

    def show_error(self, text):
        self.error.configure(text=text)
        self.bell()


# --------------------------------------------------------------------------- #
#  Main window
# --------------------------------------------------------------------------- #
class App(tk.Tk):
    def __init__(self, db_path=None, demo=False):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1440x860")
        self.minsize(1000, 560)
        self.configure(background=C["bg"])
        self.settings = load_settings()
        self.backend = None
        self.db_path = None
        self.is_sqlite = False
        self.table_map = {}
        self.columns = []
        self.group_col = None
        self.filter_rows = []
        self.pick_lists = {}
        self.result_cols = []
        self.rows = []            # flat mode: matched rows
        self.groups = []          # grouped mode: [(adapter, [(row, matched), ...])]
        self.iid_rows = {}        # tree item id -> list of rows
        self.full_cache = {}      # hide_scrapped flag -> {adapter: [rows]}
        self.sort_state = (None, False)
        self._after_id = None
        self.demo = demo
        self.editing = False
        self.store = None
        self.log = None
        self.undo_stack = []      # [(action, before, after)]
        self.backed_up = False

        self._setup_style()
        self._build_menu()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self.quit_app)
        self.after(50, lambda: self._startup(db_path))

    # ---------- look ----------
    def _setup_style(self):
        families = set(tkfont.families(self))
        fam = next((f for f in ("Segoe UI", "Inter", "Helvetica Neue", "DejaVu Sans")
                    if f in families), tkfont.nametofont("TkDefaultFont").actual("family"))
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
            try:
                tkfont.nametofont(name).configure(family=fam, size=10)
            except tk.TclError:
                pass
        self.f_base = (fam, 10)
        self.f_small = (fam, 9)
        self.f_bold = (fam, 10, "bold")
        self.f_title = (fam, 15, "bold")
        self.f_section = (fam, 11, "bold")
        self.f_search = (fam, 12)
        self._bold = tkfont.Font(self, family=fam, size=10, weight="bold")
        # Results table fonts - resized by zooming (Ctrl+plus / Ctrl+minus / Ctrl+wheel).
        try:
            zoom = int(self.settings.get("zoom", ZOOM_DEFAULT))
        except (TypeError, ValueError):
            zoom = ZOOM_DEFAULT
        self.zoom = max(ZOOM_MIN, min(ZOOM_MAX, zoom))
        self.f_table = tkfont.Font(self, family=fam, size=self.zoom)
        self.f_table_bold = tkfont.Font(self, family=fam, size=self.zoom, weight="bold")
        self.f_table_head = tkfont.Font(self, family=fam, size=self.zoom - 1, weight="bold")

        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=C["bg"], foreground=C["text"], font=self.f_base,
                    bordercolor=C["border"], lightcolor=C["surface"], darkcolor=C["border"],
                    troughcolor=C["bg"], focuscolor=C["primary"])
        s.configure("TFrame", background=C["bg"])
        s.configure("Card.TFrame", background=C["surface"])
        s.configure("Header.TFrame", background=C["header"])
        s.configure("Banner.TFrame", background=C["edit_light"])
        s.configure("TLabel", background=C["bg"], foreground=C["text"])
        s.configure("Muted.TLabel", background=C["bg"], foreground=C["muted"])
        s.configure("Card.TLabel", background=C["surface"])
        s.configure("CardMuted.TLabel", background=C["surface"], foreground=C["muted"])
        s.configure("Hint.TLabel", background=C["surface"], foreground=C["faint"],
                    font=self.f_small)
        s.configure("Section.TLabel", background=C["surface"], font=self.f_section)
        s.configure("Count.TLabel", background=C["surface"], font=self.f_section)
        s.configure("DetailTitle.TLabel", background=C["surface"], font=(self.f_base[0], 13,
                                                                         "bold"))
        s.configure("DialogTitle.TLabel", background=C["bg"], font=self.f_title)
        s.configure("FieldLabel.TLabel", background=C["surface"], foreground=C["muted"],
                    font=(self.f_base[0], 9, "bold"))
        s.configure("FieldLabelChanged.TLabel", background=C["surface"],
                    foreground=C["edit"], font=(self.f_base[0], 9, "bold"))
        s.configure("Error.TLabel", background=C["bg"], foreground=C["danger"])
        s.configure("HeaderTitle.TLabel", background=C["header"],
                    foreground=C["header_text"], font=self.f_title)
        s.configure("HeaderSub.TLabel", background=C["header"], foreground=C["header_sub"])
        s.configure("Banner.TLabel", background=C["edit_light"], foreground=C["edit"])
        s.configure("Status.TLabel", background=C["surface"], foreground=C["muted"],
                    padding=(12, 5))

        s.configure("TButton", background=C["surface"], foreground=C["text"],
                    bordercolor=C["border"], lightcolor=C["surface"], darkcolor=C["surface"],
                    padding=(12, 5), relief="flat", focusthickness=0)
        s.map("TButton", background=[("disabled", C["bg"]), ("pressed", C["heading"]),
                                     ("active", C["primary_light"])],
              foreground=[("disabled", C["faint"])],
              bordercolor=[("active", C["primary"])])
        s.configure("Accent.TButton", background=C["primary"], foreground="white",
                    bordercolor=C["primary"], lightcolor=C["primary"],
                    darkcolor=C["primary"], font=self.f_bold)
        s.map("Accent.TButton",
              background=[("disabled", "#a9bbd6"), ("pressed", C["primary_dark"]),
                          ("active", C["primary_dark"])],
              foreground=[("disabled", "white")],
              bordercolor=[("active", C["primary_dark"])])
        s.configure("Danger.TButton", foreground=C["danger"])
        s.map("Danger.TButton", background=[("disabled", C["bg"]),
                                            ("active", C["danger_light"])],
              bordercolor=[("active", C["danger"])], foreground=[("disabled", C["faint"])])
        s.configure("Header.TButton", background=C["header_btn"], foreground="white",
                    bordercolor=C["header_btn"], lightcolor=C["header_btn"],
                    darkcolor=C["header_btn"], padding=(14, 6))
        s.map("Header.TButton", background=[("active", "#3a5f90")],
              bordercolor=[("active", "#3a5f90")])
        s.configure("EditOn.TButton", background="#d97706", foreground="white",
                    bordercolor="#d97706", lightcolor="#d97706", darkcolor="#d97706",
                    padding=(14, 6), font=self.f_bold)
        s.map("EditOn.TButton", background=[("active", "#b45309")],
              bordercolor=[("active", "#b45309")])

        for w in ("TEntry", "TCombobox"):
            s.configure(w, fieldbackground=C["surface"], bordercolor=C["border"],
                        lightcolor=C["surface"], darkcolor=C["surface"], padding=(6, 4),
                        arrowcolor=C["muted"], insertcolor=C["text"])
            s.map(w, bordercolor=[("focus", C["primary"])],
                  lightcolor=[("focus", C["primary"])],
                  fieldbackground=[("readonly", C["surface"])])
        s.configure("Search.TEntry", padding=(10, 7))
        self.option_add("*TCombobox*Listbox.font", self.f_base)
        self.option_add("*TCombobox*Listbox.selectBackground", C["primary"])

        for w in ("TCheckbutton", "Card.TCheckbutton"):
            s.configure(w, background=C["bg"] if w == "TCheckbutton" else C["surface"],
                        indicatorbackground=C["surface"], indicatorforeground=C["primary"],
                        focusthickness=0)
            s.map(w, background=[("active", C["bg"] if w == "TCheckbutton"
                                  else C["surface"])],
                  indicatorbackground=[("selected", C["surface"])])

        s.configure("Treeview", background=C["surface"], fieldbackground=C["surface"],
                    foreground=C["text"], rowheight=26, borderwidth=0, relief="flat")
        s.map("Treeview", background=[("selected", C["primary"])],
              foreground=[("selected", "white")])
        s.configure("Treeview.Heading", background=C["heading"], foreground=C["muted"],
                    font=(self.f_base[0], 9, "bold"), relief="flat", padding=(8, 6),
                    bordercolor=C["border"], lightcolor=C["heading"], darkcolor=C["border"])
        s.map("Treeview.Heading", background=[("active", "#e2e7ee")])
        s.configure("Results.Treeview", font=self.f_table, rowheight=self._row_height())
        s.configure("Results.Treeview.Heading", font=self.f_table_head)
        s.configure("Details.Treeview", rowheight=24)
        s.configure("TPanedwindow", background=C["bg"])
        s.configure("Sash", sashthickness=8, gripcount=0, background=C["bg"])
        s.configure("Vertical.TScrollbar", background="#cfd6df", troughcolor=C["surface"],
                    bordercolor=C["surface"], lightcolor="#cfd6df", darkcolor="#cfd6df",
                    arrowcolor=C["muted"], gripcount=0)
        s.configure("Horizontal.TScrollbar", background="#cfd6df", troughcolor=C["surface"],
                    bordercolor=C["surface"], lightcolor="#cfd6df", darkcolor="#cfd6df",
                    arrowcolor=C["muted"], gripcount=0)

    # ---------- layout ----------
    def _build_menu(self):
        m = tk.Menu(self)
        f = tk.Menu(m, tearoff=False)
        f.add_command(label="Open database...", command=self.choose_db, accelerator="Ctrl+O")
        f.add_command(label="Reload", command=self.reload, accelerator="F5")
        f.add_separator()
        f.add_command(label="Export results to CSV...", command=self.export_csv,
                      accelerator="Ctrl+E")
        f.add_command(label="Back up database now...", command=self.backup_db)
        f.add_separator()
        f.add_command(label="Exit", command=self.quit_app)
        m.add_cascade(label="File", menu=f)

        e = tk.Menu(m, tearoff=False)
        e.add_command(label="Turn editing on / off", command=self.toggle_edit)
        e.add_separator()
        e.add_command(label="New connector...", command=self.new_record, accelerator="Ctrl+N")
        e.add_command(label="Edit selected...", command=self.edit_selected, accelerator="F2")
        e.add_command(label="Duplicate selected...", command=self.duplicate_selected,
                      accelerator="Ctrl+D")
        e.add_command(label="Delete selected...", command=self.delete_selected,
                      accelerator="Del")
        e.add_separator()
        e.add_command(label="Undo last change", command=self.undo, accelerator="Ctrl+Z")
        e.add_command(label="Open change log", command=self.open_log)
        m.add_cascade(label="Edit", menu=e)
        self.edit_menu = e

        v = tk.Menu(m, tearoff=False)
        v.add_command(label="Expand all adapters", command=lambda: self._expand(True))
        v.add_command(label="Collapse all adapters", command=lambda: self._expand(False))
        v.add_separator()
        v.add_command(label="Show / hide details panel", command=self.toggle_details)
        v.add_separator()
        v.add_command(label="Zoom in", command=lambda: self.set_zoom(self.zoom + 1),
                      accelerator="Ctrl++")
        v.add_command(label="Zoom out", command=lambda: self.set_zoom(self.zoom - 1),
                      accelerator="Ctrl+-")
        v.add_command(label="Reset zoom", command=lambda: self.set_zoom(ZOOM_DEFAULT),
                      accelerator="Ctrl+0")
        m.add_cascade(label="View", menu=v)
        self.config(menu=m)

        self.bind("<Control-o>", lambda e: self.choose_db())
        self.bind("<Control-e>", lambda e: self.export_csv())
        self.bind("<Control-f>", lambda e: (self.quick_entry.focus_set(),
                                            self.quick_entry.select_range(0, "end")))
        self.bind("<F5>", lambda e: self.reload())
        self.bind("<Control-n>", lambda e: self.new_record())
        self.bind("<Control-z>", lambda e: self._shortcut(self.undo, e))
        self.bind("<Control-d>", lambda e: self.duplicate_selected())
        for seq in ("<Control-plus>", "<Control-equal>", "<Control-KP_Add>"):
            self.bind(seq, lambda e: self.set_zoom(self.zoom + 1))
        for seq in ("<Control-minus>", "<Control-KP_Subtract>"):
            self.bind(seq, lambda e: self.set_zoom(self.zoom - 1))
        self.bind("<Control-0>", lambda e: self.set_zoom(ZOOM_DEFAULT))

    def _shortcut(self, func, event):
        # Leave Ctrl+Z inside text boxes alone.
        if isinstance(event.widget, (tk.Entry, ttk.Entry, tk.Text)):
            return
        func()

    def _build_ui(self):
        # Header bar
        header = ttk.Frame(self, style="Header.TFrame", padding=(18, 12, 18, 12))
        header.pack(fill="x")
        ttk.Label(header, text=APP_TITLE, style="HeaderTitle.TLabel").pack(side="left")
        self.db_label = ttk.Label(header, text="No database open", style="HeaderSub.TLabel")
        self.db_label.pack(side="left", padx=(14, 0), pady=(5, 0))
        self.edit_btn = ttk.Button(header, text="\u270e  Enable editing",
                                   style="Header.TButton", command=self.toggle_edit)
        self.edit_btn.pack(side="right")
        Tooltip(self.edit_btn, "Turn on to add, change or delete connector rows. "
                               "Every change is logged and can be undone (Ctrl+Z).")

        # Edit-mode banner (hidden until editing is on)
        self.banner = ttk.Frame(self, style="Banner.TFrame", padding=(18, 6))
        self.banner_label = ttk.Label(self.banner, style="Banner.TLabel", text="")
        self.banner_label.pack(side="left")
        self.undo_link = link(self.banner, "Undo last change", self.undo, color=C["edit"],
                              bg=C["edit_light"], font=self.f_bold)
        self.undo_link.pack(side="right")

        # Search bar
        bar = ttk.Frame(self, padding=(18, 14, 18, 6))
        bar.pack(fill="x")
        self.bar = bar
        self.table_label = ttk.Label(bar, text="Table")
        self.table_var = tk.StringVar()
        self.table_cb = ttk.Combobox(bar, textvariable=self.table_var, state="readonly",
                                     width=24)
        self.table_cb.bind("<<ComboboxSelected>>", lambda e: self.load_table())

        self.search_box = ttk.Frame(bar)
        self.search_box.pack(side="left", fill="x", expand=True)
        self.quick_var = tk.StringVar()
        self.quick_entry = ttk.Entry(self.search_box, textvariable=self.quick_var,
                                     font=self.f_search, style="Search.TEntry")
        self.quick_entry.pack(side="left", fill="x", expand=True)
        self.quick_ph = Placeholder(self.quick_entry, self.quick_var,
                                    "Search adapters, part numbers, descriptions...  "
                                    "(* = wildcard)")
        self.quick_ph.label.configure(font=self.f_search)
        self.quick_entry.bind("<KeyRelease>", self._quick_typed)
        self.quick_entry.bind("<Return>", lambda e: self.run_search())
        self.quick_entry.bind("<Escape>", lambda e: (self.quick_var.set(""), self.run_search()))
        ttk.Label(bar, text="in", style="Muted.TLabel").pack(side="left", padx=(10, 6))
        self.in_var = tk.StringVar(value="All columns")
        self.in_cb = ttk.Combobox(bar, textvariable=self.in_var, state="readonly", width=20)
        self.in_cb.pack(side="left")
        self.in_cb.bind("<<ComboboxSelected>>", lambda e: self.run_search())
        ttk.Button(bar, text="Clear all", command=self.clear_all).pack(side="left",
                                                                         padx=(10, 0))

        opts = ttk.Frame(self, padding=(18, 0, 18, 4))
        opts.pack(fill="x")
        self.group_var = tk.BooleanVar(value=self.settings.get("group", True))
        self.group_chk = ttk.Checkbutton(opts, text="Group by adapter",
                                         variable=self.group_var,
                                         command=self._options_changed)
        self.group_chk.pack(side="left")
        Tooltip(self.group_chk, "Show every connector on each matching adapter. "
                                "The connectors that matched are highlighted.")
        self.scrap_var = tk.BooleanVar(value=self.settings.get("hide_scrapped", True))
        self.scrap_chk = ttk.Checkbutton(opts, text="Hide scrapped", variable=self.scrap_var,
                                         command=self._options_changed)
        self.scrap_chk.pack(side="left", padx=(18, 0))
        # Active-filter chips
        self.chips = ttk.Frame(opts)
        self.chips.pack(side="left", fill="x", expand=True, padx=(22, 0))

        # Main area: filters | results | details
        self.pane = ttk.PanedWindow(self, orient="horizontal")
        self.pane.pack(fill="both", expand=True, padx=18, pady=(6, 10))
        self._build_filter_panel()
        self._build_results()
        self._build_details()

        # Status bar
        status = ttk.Frame(self, style="Card.TFrame")
        status.pack(fill="x", side="bottom")
        zoom = ttk.Frame(status, style="Card.TFrame")
        zoom.pack(side="right", padx=(0, 10))
        zoom_out = link(zoom, "−", lambda: self.set_zoom(self.zoom - 1), font=self.f_bold)
        zoom_out.pack(side="left", padx=4)
        self.zoom_var = tk.StringVar(value=self._zoom_text())
        zoom_pct = link(zoom, "", lambda: self.set_zoom(ZOOM_DEFAULT), color=C["muted"],
                        font=self.f_small)
        zoom_pct.configure(textvariable=self.zoom_var, width=5)
        zoom_pct.pack(side="left")
        zoom_in = link(zoom, "+", lambda: self.set_zoom(self.zoom + 1), font=self.f_bold)
        zoom_in.pack(side="left", padx=4)
        Tooltip(zoom_out, "Zoom out of the results table (Ctrl+minus)")
        Tooltip(zoom_pct, "Results table zoom - click to reset to 100% (Ctrl+0)")
        Tooltip(zoom_in, "Zoom in on the results table (Ctrl+plus, or Ctrl+mouse wheel)")
        self.status_var = tk.StringVar(value="No database open.")
        ttk.Label(status, textvariable=self.status_var, style="Status.TLabel",
                  anchor="w").pack(side="left", fill="x", expand=True)

        self.menu = tk.Menu(self, tearoff=False)
        self.menu.add_command(label="Open details window", command=self.show_details)
        self.menu.add_command(label="Copy row(s)", command=self.copy_rows)
        self.menu.add_command(label="Copy adapter number(s)",
                              command=lambda: self.copy_column(self.group_col, unique=True))
        self.menu.add_command(label="Copy connector P/N(s)",
                              command=lambda: self.copy_column(self._part_col()))
        self.menu.add_separator()
        self.menu.add_command(label="Edit...", command=self.edit_selected)
        self.menu.add_command(label="Duplicate...", command=self.duplicate_selected)
        self.menu.add_command(label="Add connector to this adapter...",
                              command=self.add_to_adapter)
        self.menu.add_command(label="Delete...", command=self.delete_selected)
        self._set_edit_ui()

    def _build_filter_panel(self):
        left = ttk.Frame(self.pane, style="Card.TFrame", width=280)
        head = ttk.Frame(left, style="Card.TFrame", padding=(12, 10, 12, 8))
        head.pack(fill="x")
        ttk.Label(head, text="Filters", style="Section.TLabel").pack(side="left")
        link(head, "Reset all", self.clear_all).pack(side="right")
        ttk.Separator(left).pack(fill="x")
        self.filter_scroll = ScrollFrame(left)
        self.filter_scroll.pack(fill="both", expand=True, pady=(8, 0))
        self.pane.add(left, weight=0)

    def _build_results(self):
        card = ttk.Frame(self.pane, style="Card.TFrame")
        top = ttk.Frame(card, style="Card.TFrame", padding=(14, 8, 10, 8))
        top.pack(fill="x")
        self.count_label = ttk.Label(top, text="", style="Count.TLabel")
        self.count_label.pack(side="left")
        self.sub_count = ttk.Label(top, text="", style="CardMuted.TLabel")
        self.sub_count.pack(side="left", padx=(10, 0), pady=(2, 0))
        ttk.Button(top, text="Export CSV", command=self.export_csv).pack(side="right")
        self.expand_btns = ttk.Frame(top, style="Card.TFrame")
        link(self.expand_btns, "Expand all", lambda: self._expand(True)).pack(side="left")
        ttk.Label(self.expand_btns, text=" \u00b7 ", style="CardMuted.TLabel").pack(side="left")
        link(self.expand_btns, "Collapse all", lambda: self._expand(False)).pack(side="left")

        # Edit toolbar (shown in edit mode)
        self.edit_bar = ttk.Frame(top, style="Card.TFrame")
        self.b_new = ttk.Button(self.edit_bar, text="+  New connector",
                                style="Accent.TButton", command=self.new_record)
        self.b_new.pack(side="left")
        self.b_edit = ttk.Button(self.edit_bar, text="Edit", command=self.edit_selected)
        self.b_edit.pack(side="left", padx=(6, 0))
        self.b_dup = ttk.Button(self.edit_bar, text="Duplicate",
                                command=self.duplicate_selected)
        self.b_dup.pack(side="left", padx=(6, 0))
        self.b_del = ttk.Button(self.edit_bar, text="Delete", style="Danger.TButton",
                                command=self.delete_selected)
        self.b_del.pack(side="left", padx=(6, 0))
        Tooltip(self.b_new, "Add a connector row (Ctrl+N). If a connector or adapter is "
                            "selected, the adapter number is filled in for you.")
        Tooltip(self.b_dup, "Copy the selected connector into a new row you can change "
                            "(Ctrl+D).")
        ttk.Separator(card).pack(fill="x")

        body = ttk.Frame(card, style="Card.TFrame")
        body.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(body, show="headings", selectmode="extended",
                                 style="Results.Treeview")
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(body, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background=C["alt"])
        self.tree.tag_configure("group", background=C["group"], font=self.f_table_bold)
        self.tree.tag_configure("match", font=self.f_table_bold, background=C["match"])
        self.tree.tag_configure("other", foreground=C["muted"])
        self.tree.bind("<Double-1>", self._double_click)
        self.tree.bind("<Return>", self._double_click)
        self.tree.bind("<F2>", lambda e: self.edit_selected())
        self.tree.bind("<Delete>", lambda e: self.delete_selected())
        self.tree.bind("<Control-c>", lambda e: self.copy_rows())
        self.tree.bind("<Button-3>", self._context_menu)
        self.tree.bind("<Control-MouseWheel>", self._zoom_wheel)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self._update_details())

        self.empty = ttk.Frame(body, style="Card.TFrame")
        ttk.Label(self.empty, text="No matches", style="Section.TLabel").pack()
        ttk.Label(self.empty, text="Try a shorter search, or remove a filter.",
                  style="CardMuted.TLabel").pack(pady=(4, 10))
        ttk.Button(self.empty, text="Clear all filters", command=self.clear_all).pack()
        self.pane.add(card, weight=1)

    def _build_details(self):
        card = ttk.Frame(self.pane, style="Card.TFrame", width=330)
        self.details_card = card
        head = ttk.Frame(card, style="Card.TFrame", padding=(14, 10, 14, 6))
        head.pack(fill="x")
        ttk.Label(head, text="DETAILS", style="FieldLabel.TLabel").pack(anchor="w")
        self.d_title = ttk.Label(head, text="Nothing selected", style="DetailTitle.TLabel",
                                 wraplength=300)
        self.d_title.pack(anchor="w", pady=(2, 0))
        self.d_sub = ttk.Label(head, text="Click a row to see all of its fields.",
                               style="CardMuted.TLabel", wraplength=300, justify="left")
        self.d_sub.pack(anchor="w")

        btns = ttk.Frame(card, style="Card.TFrame", padding=(14, 4, 14, 8))
        btns.pack(fill="x")
        self.d_copy = ttk.Button(btns, text="Copy", command=self.copy_rows)
        self.d_copy.pack(side="left")
        self.d_window = ttk.Button(btns, text="Open window", command=self.show_details)
        self.d_window.pack(side="left", padx=(6, 0))
        self.d_edit_btns = ttk.Frame(card, style="Card.TFrame", padding=(14, 0, 14, 8))
        self.d_edit = ttk.Button(self.d_edit_btns, text="Edit", style="Accent.TButton",
                                 command=self.edit_selected)
        self.d_edit.pack(side="left")
        self.d_add = ttk.Button(self.d_edit_btns, text="+ Add connector to adapter",
                                command=self.add_to_adapter)
        self.d_add.pack(side="left", padx=(6, 0))
        ttk.Separator(card).pack(fill="x", side="top")

        body = ttk.Frame(card, style="Card.TFrame")
        body.pack(fill="both", expand=True)
        self.dtree = ttk.Treeview(body, columns=("field", "value"), show="headings",
                                  style="Details.Treeview", selectmode="browse")
        self.dtree.heading("field", text="Field", anchor="w")
        self.dtree.heading("value", text="Value", anchor="w")
        self.dtree.column("field", width=120, stretch=False)
        self.dtree.column("value", width=200, stretch=True)
        self.dtree.tag_configure("odd", background=C["alt"])
        self.dtree.tag_configure("key", font=self._bold)
        dsb = ttk.Scrollbar(body, orient="vertical", command=self.dtree.yview)
        self.dtree.configure(yscrollcommand=dsb.set)
        self.dtree.pack(side="left", fill="both", expand=True)
        dsb.pack(side="right", fill="y")
        self.dtree.bind("<Control-c>", self._copy_detail_value)
        self.dtree.bind("<Double-1>", self._copy_detail_value)
        self.details_shown = self.settings.get("details", True)
        if self.details_shown:
            self.pane.add(card, weight=0)

    def toggle_details(self):
        if self.details_shown:
            self.pane.forget(self.details_card)
        else:
            self.pane.add(self.details_card, weight=0)
        self.details_shown = not self.details_shown
        self.settings["details"] = self.details_shown
        save_settings(self.settings)

    # ---------- zoom (results table only) ----------
    def _zoom_scale(self):
        return self.zoom / ZOOM_DEFAULT

    def _zoom_text(self):
        return f"{round(self._zoom_scale() * 100)}%"

    def _row_height(self):
        return max(self.f_table.metrics("linespace") + 6, round(26 * self._zoom_scale()))

    def set_zoom(self, size):
        size = max(ZOOM_MIN, min(ZOOM_MAX, size))
        if size == self.zoom:
            return
        ratio = size / self.zoom
        self.zoom = size
        self.f_table.configure(size=size)
        self.f_table_bold.configure(size=size)
        self.f_table_head.configure(size=size - 1)
        ttk.Style(self).configure("Results.Treeview", rowheight=self._row_height())
        for col in ("#0",) + tuple(self.tree["columns"]):
            width = self.tree.column(col, "width")
            self.tree.column(col, width=max(40, round(width * ratio)))
        self.zoom_var.set(self._zoom_text())
        self.settings["zoom"] = size
        save_settings(self.settings)

    def _zoom_wheel(self, event):
        self.set_zoom(self.zoom + (1 if event.delta > 0 else -1))
        return "break"

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
        elif os.path.exists(DEFAULT_DB):
            self.open_db(DEFAULT_DB)
        elif len(found) == 1:
            self.open_db(found[0])
        else:
            self.choose_db()

    def choose_db(self):
        if self.editing and not self._stop_editing():
            return
        path = filedialog.askopenfilename(
            title="Choose the adapter Access database",
            initialdir=APP_DIR,
            filetypes=[("Access databases", "*.accdb *.mdb"), ("All files", "*.*")])
        if path:
            self.open_db(path)

    def open_db(self, path, sqlite=False):
        try:
            backend = open_backend(path, sqlite, writable=self.editing)
            # Go straight to the connectors table; only list the tables (a driver
            # catalog call that can fail to decode) when it isn't there.
            tables = ([(DEFAULT_TABLE, False)] if backend.has_table(DEFAULT_TABLE)
                      else backend.tables())
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Couldn't open the database:\n{path}\n\n{e}")
            return
        if self.backend:
            self.backend.close()
        if path != self.db_path:
            self.undo_stack = []
            self.backed_up = False
        self.backend = backend
        self.db_path = path
        self.is_sqlite = sqlite
        self.log = ChangeLog(path)
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
                self.table_label.pack(side="left", before=self.search_box, padx=(0, 6))
                self.table_cb.pack(side="left", padx=(0, 14), before=self.search_box)
        labels = list(self.table_map)
        self.table_cb["values"] = labels
        if not sqlite:
            self.settings["last_db"] = path
            save_settings(self.settings)
        self.title(f"{APP_TITLE} - {os.path.basename(path)}")
        self.db_label.configure(text=path)
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
            self.open_db(self.db_path, sqlite=self.is_sqlite)

    def _table(self):
        return self.table_map.get(self.table_var.get())

    def load_table(self):
        table = self._table()
        if not table:
            return
        self.settings.setdefault("tables", {})[self.db_path] = self.table_var.get()
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
            self.store = RecordStore(self.backend, table, self.columns)
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
        self.pick_lists = {}
        cols = [c for c in self.columns if c.kind != "binary"]
        cols = sorted(cols, key=lambda c: priority(c.name))  # stable: keeps db order otherwise
        common = ttk.Frame(inner, style="Card.TFrame")
        common.pack(fill="x")
        toggle_row = ttk.Frame(inner, style="Card.TFrame", padding=(12, 2, 12, 10))
        more = ttk.Frame(inner, style="Card.TFrame")
        self.more_frame = more
        self.more_link = link(toggle_row, "", self._toggle_more, font=self.f_bold)
        self.more_link.pack(side="left")
        for i, c in enumerate(cols):
            distinct = self._distinct_values(table, c)
            if distinct is not None:
                self.pick_lists[c.name] = distinct
            parent = common if i < COMMON_FILTERS else more
            self.filter_rows.append(FilterRow(self, parent, c, distinct))
        self.n_more = max(0, len(cols) - COMMON_FILTERS)
        if self.n_more:
            toggle_row.pack(fill="x")
            if self.settings.get("more_filters", False):
                more.pack(fill="x")
        self._update_more_link()
        self.filter_scroll.canvas.yview_moveto(0)

    def _toggle_more(self):
        shown = bool(self.more_frame.winfo_ismapped())
        if shown:
            self.more_frame.pack_forget()
        else:
            self.more_frame.pack(fill="x")
        self.settings["more_filters"] = not shown
        save_settings(self.settings)
        self.after(10, self._update_more_link)

    def _update_more_link(self):
        if not self.n_more:
            return
        shown = bool(self.settings.get("more_filters", False))
        hidden_active = sum(fr.is_active() for fr in self.filter_rows[COMMON_FILTERS:])
        text = ("\u25be  Fewer filters" if shown else
                f"\u25b8  More filters ({self.n_more})")
        if hidden_active and not shown:
            text += f"  \u2022 {hidden_active} active"
        self.more_link.configure(text=text)

    # ---------- searching ----------
    def _quick_typed(self, event):
        if event.keysym in ("Return", "Escape", "Tab", "Shift_L", "Shift_R", "Control_L",
                            "Control_R", "Up", "Down", "Left", "Right"):
            return
        self.schedule_search()

    def schedule_search(self, event=None):
        if event is not None and getattr(event, "keysym", "") in (
                "Return", "Tab", "Shift_L", "Shift_R", "Control_L", "Control_R", "Up",
                "Down", "Left", "Right", "Escape"):
            return
        if self._after_id:
            self.after_cancel(self._after_id)
        self._after_id = self.after(400, self.run_search)

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
        """All (non-scrapped) rows, grouped by adapter. Cached until data changes."""
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
        if self._after_id:
            try:
                self.after_cancel(self._after_id)
            except Exception:
                pass
        self._after_id = None
        if not self.backend or not self.columns:
            return
        table = self._table()
        filters = []
        try:
            for fr in self.filter_rows:
                filters.extend(fr.conditions())
        except ValueError as e:
            self.status_var.set(f"Check the filter - {e}")
            self.bell()
            return
        self._update_chips()
        self._update_more_link()
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
        keep = self._selected_connector()
        keep = self._as_dict(keep) if keep is not None else None
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
        if keep:
            self._select_values(keep)

        if full is not None:
            self.count_label.configure(text=plural(len(self.groups), "adapter"))
            self.sub_count.configure(text=f"{plural(len(rows), 'matching connector')} "
                                          f"highlighted")
            self.expand_btns.pack(side="right", padx=(0, 14))
        else:
            self.count_label.configure(text=plural(len(rows), "connector") +
                                       ("+" if truncated else ""))
            self.sub_count.configure(text="")
            self.expand_btns.pack_forget()
        if truncated:
            self.sub_count.configure(text=f"only the first {MAX_ROWS:,} are shown")
        if rows:
            self.empty.place_forget()
        else:
            self.empty.place(relx=0.5, rely=0.4, anchor="center")
        tips = "Double-click a row to open it" if not self.editing else \
            "Double-click a row to edit it"
        self.status_var.set(f"{tips}  \u00b7  right-click to copy  \u00b7  "
                            "Ctrl+F to search  \u00b7  F5 to reload")
        self._update_details()

    def _normalize(self, names, rows):
        bools = {i for i, n in enumerate(names)
                 if any(c.name == n and c.kind == "bool" for c in self.columns)}
        return [tuple(bool(v) if i in bools and v is not None else v
                      for i, v in enumerate(r)) for r in rows]

    def _update_chips(self):
        for w in self.chips.winfo_children():
            w.destroy()
        items = []
        q = self.quick_var.get().strip()
        if q:
            where = "" if self.in_var.get() == "All columns" else f" in {self.in_var.get()}"
            items.append((f"\u201c{q}\u201d{where}",
                          lambda: (self.quick_var.set(""), self.run_search())))
        for fr in self.filter_rows:
            if fr.is_active():
                items.append((fr.chip_text(),
                              lambda fr=fr: (fr.clear(), self.run_search())))
        for text, remove in items:
            chip = tk.Frame(self.chips, background=C["chip"])
            chip.pack(side="left", padx=(0, 6))
            tk.Label(chip, text=text, background=C["chip"], foreground=C["chip_text"],
                     font=self.f_small, padx=8, pady=3).pack(side="left")
            x = tk.Label(chip, text="\u2715", background=C["chip"],
                         foreground=C["chip_text"], font=self.f_small, padx=6,
                         cursor="hand2")
            x.pack(side="left")
            x.bind("<Button-1>", lambda e, r=remove: r())
        if len(items) > 1:
            link(self.chips, "Clear all", self.clear_all, bg=C["bg"],
                 font=self.f_small).pack(side="left", padx=(4, 0))

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
                width = round(max(60, min(320, longest * 7 + 24)) * self._zoom_scale())
                tree.heading(name, text=name, anchor="w",
                             command=lambda n=name: self.sort_by(n))
                tree.column(name, width=width, minwidth=40, stretch=False, anchor="w")
        idx = [self.result_cols.index(n) for n in display]
        for name in display:
            arrow = ""
            if self.sort_state[0] == name:
                arrow = "  \u25bc" if self.sort_state[1] else "  \u25b2"
            tree.heading(name, text=name + arrow)

        if grouped:
            tree.configure(show="tree headings")
            head = self.group_col
            if self.sort_state[0] == head:
                head += "  \u25bc" if self.sort_state[1] else "  \u25b2"
            tree.heading("#0", text=head, anchor="w",
                         command=lambda: self.sort_by(self.group_col))
            tree.column("#0", width=round(200 * self._zoom_scale()), minwidth=90,
                        stretch=False)
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

    # ---------- selection & details ----------
    def _selected(self):
        out, seen = [], set()
        for iid in self.tree.selection():
            for r in self.iid_rows.get(iid, []):
                if id(r) not in seen:
                    seen.add(id(r))
                    out.append(r)
        return out

    def _selected_connector(self):
        """The one connector row selected (not an adapter heading), or None."""
        sel = self.tree.selection()
        if len(sel) != 1 or sel[0].startswith("g") and "_" not in sel[0]:
            return None
        rows = self.iid_rows.get(sel[0], [])
        return rows[0] if len(rows) == 1 else None

    def _selected_adapter(self):
        if not self.group_col:
            return None
        rows = self._selected()
        if not rows:
            return None
        gi = self.result_cols.index(self.group_col)
        adapters = {r[gi] for r in rows}
        return rows[0][gi] if len(adapters) == 1 else None

    def _as_dict(self, row):
        return dict(zip(self.result_cols, row))

    def _update_details(self):
        t = self.dtree
        t.delete(*t.get_children())
        sel = self.tree.selection()
        rows = self._selected()
        part = self._part_col()
        is_group = len(sel) == 1 and sel[0].startswith("g") and "_" not in sel[0]
        state = "normal" if rows else "disabled"
        for b in (self.d_copy, self.d_window):
            b.configure(state=state)
        self._set_edit_buttons()
        if not rows:
            self.d_title.configure(text="Nothing selected")
            self.d_sub.configure(text="Click a row to see all of its fields.")
            return
        if is_group or len(rows) > 1:
            adapter = self._selected_adapter()
            self.d_title.configure(text=fmt_value(adapter) if adapter is not None
                                   else f"{len(rows)} connectors")
            self.d_sub.configure(text=plural(len(rows), "connector") +
                                 ("" if adapter is not None else " on several adapters"))
            pi = self.result_cols.index(part) if part else 0
            di = next((i for i, c in enumerate(self.result_cols)
                       if c.upper() == "DESCRIPTION"), None)
            for k, r in enumerate(rows):
                desc = fmt_value(r[di]) if di is not None else ""
                t.insert("", "end", values=(fmt_value(r[pi]), desc),
                         tags=("odd",) if k % 2 else ())
            t.heading("field", text=label_for(part) if part else "Row")
            t.heading("value", text="Description")
            return
        r = rows[0]
        t.heading("field", text="Field")
        t.heading("value", text="Value")
        self.d_title.configure(text=fmt_value(r[self.result_cols.index(part)]) if part
                               else fmt_value(r[0]))
        if self.group_col:
            self.d_sub.configure(text=f"on {fmt_value(r[self.result_cols.index(self.group_col)])}")
        else:
            self.d_sub.configure(text="")
        for k, (name, v) in enumerate(zip(self.result_cols, r)):
            tags = ["odd"] if k % 2 else []
            if name in (self.group_col, part):
                tags.append("key")
            t.insert("", "end", values=(label_for(name), fmt_value(v)), tags=tags)

    def _copy_detail_value(self, event=None):
        sel = self.dtree.selection()
        if sel:
            value = self.dtree.item(sel[0], "values")[1]
            self.clipboard_clear()
            self.clipboard_append(value)
            self.status_var.set(f"Copied \u201c{value}\u201d")
        return "break"

    def _double_click(self, event=None):
        if event is not None and getattr(event, "y", None) is not None and \
                self.tree.identify_region(event.x, event.y) == "heading":
            return
        if self.editing and self._selected_connector() is not None:
            self.edit_selected()
        elif self._selected_connector() is not None:
            self.show_details()

    def _context_menu(self, event):
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)
            self.menu.entryconfigure(2, state="normal" if self.group_col else "disabled")
            one = self._selected_connector() is not None
            ed = self.editing
            self.menu.entryconfigure(5, state="normal" if ed and one else "disabled")
            self.menu.entryconfigure(6, state="normal" if ed and one else "disabled")
            self.menu.entryconfigure(7, state="normal" if ed and self._selected_adapter()
                                     is not None else "disabled")
            self.menu.entryconfigure(8, state="normal" if ed else "disabled")
            self.menu.tk_popup(event.x_root, event.y_root)

    def copy_rows(self):
        rows = self._selected()
        if not rows:
            return
        text = "\t".join(self.result_cols) + "\n" + \
            "\n".join("\t".join(fmt_value(v) for v in r) for r in rows)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.status_var.set(f"Copied {plural(len(rows), 'row')} - paste into Excel.")

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
            self.status_var.set(f"Copied {plural(len(vals), 'value')} from {name}.")

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
        win.configure(background=C["bg"])
        win.title(f"Details - {fmt_value(rows[0][ti])}")
        win.geometry(f"{min(1200, 260 + 230 * len(rows))}x620")
        win.transient(self)
        ttk.Label(win, text=fmt_value(rows[0][ti]), style="DialogTitle.TLabel",
                  padding=(16, 12, 16, 4)).pack(anchor="w")
        frame = ttk.Frame(win, style="Card.TFrame")
        frame.pack(fill="both", expand=True, padx=16, pady=(4, 8))
        cols = ["field"] + [f"c{k}" for k in range(len(rows))]
        t = ttk.Treeview(frame, columns=cols, show="headings")
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=t.xview)
        t.configure(xscrollcommand=hsb.set)
        t.heading("field", text="Field", anchor="w")
        t.column("field", width=150, stretch=False)
        for k in range(len(rows)):
            t.heading(f"c{k}", text=f"Connector {k + 1}" if len(rows) > 1 else "Value",
                      anchor="w")
            t.column(f"c{k}", width=220, stretch=len(rows) == 1)
        t.tag_configure("odd", background=C["alt"])
        for i, name in enumerate(self.result_cols):
            t.insert("", "end", values=[label_for(name)] + [fmt_value(r[i]) for r in rows],
                     tags=("odd",) if i % 2 else ())
        t.pack(fill="both", expand=True)
        hsb.pack(fill="x")

        def copy_all():
            win.clipboard_clear()
            win.clipboard_append("\n".join(
                "\t".join([name] + [fmt_value(r[i]) for r in rows])
                for i, name in enumerate(self.result_cols)))
        b = ttk.Frame(win, padding=(16, 0, 16, 12))
        b.pack(fill="x")
        ttk.Button(b, text="Copy all", command=copy_all).pack(side="left")
        ttk.Button(b, text="Close", command=win.destroy).pack(side="right")
        win.bind("<Escape>", lambda e: win.destroy())

    def export_csv(self):
        if not self.result_cols:
            return
        table = self._table() or "results"
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
        self.status_var.set(f"Exported {plural(n, 'row')} to {path}")

    # ---------- editing ----------
    def _set_edit_ui(self):
        if self.editing:
            self.edit_btn.configure(text="\u270e  Editing ON  \u2013  click to stop",
                                    style="EditOn.TButton")
            where = self.log.path if self.log else "the change log"
            self.banner_label.configure(
                text="Editing is on. Changes are saved straight to the database and "
                     f"logged to {where}.")
            self.banner.pack(fill="x", after=self.edit_btn.master)
            self.edit_bar.pack(side="right", padx=(0, 14))
            self.d_edit_btns.pack(fill="x", after=self.d_copy.master)
        else:
            self.edit_btn.configure(text="\u270e  Enable editing", style="Header.TButton")
            self.banner.pack_forget()
            self.edit_bar.pack_forget()
            self.d_edit_btns.pack_forget()
        for label in ("New connector...", "Edit selected...", "Duplicate selected...",
                      "Delete selected...", "Undo last change"):
            try:
                self.edit_menu.entryconfigure(label, state="normal" if self.editing
                                              else "disabled")
            except tk.TclError:
                pass
        self._set_edit_buttons()

    def _set_edit_buttons(self):
        one = self._selected_connector() is not None
        any_ = bool(self.tree.selection())
        adapter = self._selected_adapter() is not None
        self.b_edit.configure(state="normal" if one else "disabled")
        self.b_dup.configure(state="normal" if one else "disabled")
        self.b_del.configure(state="normal" if any_ else "disabled")
        self.d_edit.configure(state="normal" if one else "disabled")
        self.d_add.configure(state="normal" if adapter else "disabled")
        if self.undo_stack:
            self.undo_link.configure(text=f"Undo last change ({len(self.undo_stack)})",
                                     cursor="hand2", foreground=C["edit"])
        else:
            self.undo_link.configure(text="Nothing to undo", cursor="",
                                     foreground=C["faint"])

    def toggle_edit(self):
        if self.editing:
            self._stop_editing()
            return
        if not self.backend or not self.store:
            messagebox.showinfo(APP_TITLE, "Open a database first.")
            return
        if not self.is_sqlite and not self.backed_up:
            ans = messagebox.askyesnocancel(
                "Turn on editing",
                "Changes you make will be saved straight into\n"
                f"{self.db_path}\n\n"
                "Every change is written to a change log, and you can undo your "
                "changes from this session with Ctrl+Z.\n\n"
                "Make a backup copy of the database first? (Recommended - it can take "
                "a minute over the network.)",
                icon="warning", parent=self)
            if ans is None:
                return
            if ans:
                self.backup_db(then=self._start_editing)
                return
        self._start_editing()

    def _start_editing(self):
        try:
            backend = open_backend(self.db_path, self.is_sqlite, writable=True)
        except Exception as e:
            messagebox.showerror(
                APP_TITLE, "Couldn't open the database for editing. It may be open "
                           "in Access in exclusive mode, or you may not have permission "
                           f"to change the file or its folder.\n\n{e}")
            return
        old = self.backend
        self.backend = backend
        self.store = RecordStore(backend, self._table(), self.columns)
        if old:
            old.close()
        self.editing = True
        self._set_edit_ui()
        self.run_search()

    def _stop_editing(self):
        try:
            backend = open_backend(self.db_path, self.is_sqlite, writable=False)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Couldn't reopen the database:\n\n{e}")
            return False
        old = self.backend
        self.backend = backend
        self.store = RecordStore(backend, self._table(), self.columns)
        if old:
            old.close()
        self.editing = False
        self._set_edit_ui()
        self.run_search()
        return True

    def _need_edit(self):
        if self.editing:
            return True
        if messagebox.askyesno(APP_TITLE, "Editing is off. Turn it on now?", parent=self):
            self.toggle_edit()
        return self.editing

    def _edit_columns(self):
        return [c for c in self.store.columns if c.name in self.result_cols]

    def _open_form(self, title, values, on_save, intro=""):
        RecordDialog(self, title, self._edit_columns(), values, self.pick_lists, on_save,
                     intro)

    def new_record(self, preset=None):
        if not self._need_edit():
            return
        values = dict(preset or {})
        if not values and self.group_col:
            adapter = self._selected_adapter()
            if adapter is not None:
                values[self.group_col] = adapter
        intro = "Fields marked * are required. Pick from a list or type a new value."
        if values.get(self.group_col) is not None:
            intro = (f"Adding a connector to {fmt_value(values[self.group_col])}. " + intro)
        self._open_form("New connector", values, self._save_new, intro)

    def add_to_adapter(self):
        adapter = self._selected_adapter()
        if adapter is None:
            return
        self.new_record({self.group_col: adapter})

    def duplicate_selected(self):
        row = self._selected_connector()
        if row is None or not self._need_edit():
            return
        self._open_form("Duplicate connector", self._as_dict(row), self._save_new,
                        "A copy of the selected connector. Change what you need "
                        "(for example the adapter or the P/N) and click Save.")

    def edit_selected(self):
        row = self._selected_connector()
        if row is None:
            if self.tree.selection():
                self.status_var.set("Select a single connector row to edit it.")
            return
        if not self._need_edit():
            return
        before = self._as_dict(row)
        title = "Edit connector"
        part = self._part_col()
        if part:
            title += f" \u2013 {fmt_value(before.get(part))}"
        self._open_form(title, before,
                        lambda values, dlg: self._save_edit(before, values, dlg),
                        "Changed fields turn orange. Click Save to write them to the database.")

    def _save_new(self, values, dlg):
        try:
            saved = self.store.insert(values)
        except WriteError as e:
            dlg.show_error(str(e))
            return False
        self._record("insert", None, saved)
        self._after_write(saved, "Added")
        return True

    def _save_edit(self, before, values, dlg):
        changes = {k: v for k, v in values.items() if before.get(k) != v}
        if not changes:
            return True
        try:
            after = self.store.update(before, changes)
        except WriteError as e:
            dlg.show_error(str(e))
            return False
        self._record("update", before, after)
        self._after_write(after, f"Saved {plural(len(changes), 'change')} to")
        return True

    def delete_selected(self):
        rows = self._selected()
        if not rows or not self._need_edit():
            return
        part = self._part_col()
        lines = []
        for r in rows[:12]:
            d = self._as_dict(r)
            lines.append(f"  \u2022 {fmt_value(d.get(part)) if part else ''}"
                         + (f"  on  {fmt_value(d.get(self.group_col))}"
                            if self.group_col else ""))
        if len(rows) > 12:
            lines.append(f"  ... and {len(rows) - 12} more")
        if not messagebox.askyesno(
                "Delete connectors" if len(rows) > 1 else "Delete connector",
                f"Delete {plural(len(rows), 'connector row')} from the database?\n\n"
                + "\n".join(lines) + "\n\nYou can undo this with Ctrl+Z until you close "
                "the app.", icon="warning", parent=self):
            return
        done, problems = 0, []
        for r in rows:
            before = self._as_dict(r)
            try:
                self.store.delete(before)
            except WriteError as e:
                problems.append(f"{fmt_value(before.get(part))}: {e}")
                continue
            self._record("delete", before, None)
            done += 1
        self._after_write(None, None)
        self.status_var.set(f"Deleted {plural(done, 'connector row')}." +
                            ("  Undo with Ctrl+Z." if done else ""))
        if problems:
            messagebox.showwarning(APP_TITLE, "Some rows weren't deleted:\n\n" +
                                   "\n\n".join(problems[:5]), parent=self)

    def _record(self, action, before, after, undoable=True):
        if undoable:
            self.undo_stack.append((action, before, after))
        row = after or before
        adapter = fmt_value(row.get(self.group_col)) if self.group_col else ""
        part = self._part_col()
        pn = fmt_value(row.get(part)) if part else ""
        if action == "update":
            details = describe_changes(before, after)
        else:
            details = describe_values(row)
        if undoable:
            label = {"insert": "Added", "update": "Changed", "delete": "Deleted"}[action]
        else:
            label = {"insert": "Undo (added back)", "update": "Undo (changed back)",
                     "delete": "Undo (removed again)"}[action]
        if self.log and not self.log.write(label, adapter, pn, details):
            self.status_var.set("Saved, but the change log couldn't be written.")

    def _after_write(self, values, verb):
        """Refresh the results and select the row that was just saved."""
        self.full_cache = {}
        for fr in self.filter_rows:
            if values and fr.col.name in values:
                fr.add_choice(values[fr.col.name])
                if fr.choices:
                    self.pick_lists[fr.col.name] = sorted(fr.choices.values(), key=sort_key)
        self.run_search()
        if values:
            found = self._select_values(values)
            what = fmt_value(values.get(self._part_col())) if self._part_col() else "row"
            msg = f"{verb} {what}."
            if not found:
                msg += "  (It doesn't match the current search, so it isn't shown.)"
            self.status_var.set(msg + "  Undo with Ctrl+Z.")
        self._set_edit_buttons()

    def _select_values(self, values):
        want = [fmt_value(values.get(c)) for c in self.result_cols]
        for iid, rows in self.iid_rows.items():
            if len(rows) == 1 and not (iid.startswith("g") and "_" not in iid):
                if [fmt_value(v) for v in rows[0]] == want:
                    parent = self.tree.parent(iid)
                    if parent:
                        self.tree.item(parent, open=True)
                    self.tree.selection_set(iid)
                    self.tree.focus(iid)
                    self.tree.see(iid)
                    return True
        return False

    def undo(self):
        if not self.editing or not self.undo_stack:
            if self.editing:
                self.status_var.set("Nothing to undo.")
            return
        action, before, after = self.undo_stack[-1]
        try:
            if action == "insert":
                self.store.delete(after)
                self._record("delete", after, None, undoable=False)
                shown = None
            elif action == "update":
                restore = self.store.update(after, {k: before.get(k) for k in after})
                self._record("update", after, restore, undoable=False)
                shown = restore
            else:
                self.store.insert(before)
                self._record("insert", None, before, undoable=False)
                shown = before
        except WriteError as e:
            messagebox.showerror(APP_TITLE, f"Couldn't undo the last change:\n\n{e}",
                                 parent=self)
            return
        self.undo_stack.pop()
        word = {"insert": "adding", "update": "editing", "delete": "deleting"}[action]
        self._after_write(shown, "Restored")
        self.status_var.set(f"Undid {word} a connector.  {len(self.undo_stack)} more "
                            "change(s) can be undone.")

    def open_log(self):
        if not self.log or not os.path.exists(self.log.path):
            messagebox.showinfo(APP_TITLE, "No changes have been logged yet.", parent=self)
            return
        try:
            os.startfile(self.log.path)   # Windows: opens in Excel
        except Exception:
            messagebox.showinfo(APP_TITLE, f"The change log is at:\n\n{self.log.path}",
                                parent=self)

    def backup_db(self, then=None):
        if not self.db_path:
            return
        folder = os.path.dirname(os.path.abspath(self.db_path))
        base, ext = os.path.splitext(os.path.basename(self.db_path))
        stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
        path = filedialog.asksaveasfilename(
            title="Save a backup copy of the database", initialdir=folder,
            initialfile=f"{base}_backup_{stamp}{ext}", defaultextension=ext,
            filetypes=[("Access database", f"*{ext}")], parent=self)
        if not path:
            return
        if os.path.abspath(path) == os.path.abspath(self.db_path):
            messagebox.showerror(APP_TITLE, "Pick a different name for the backup.")
            return
        result = {}

        def work():
            try:
                shutil.copy2(self.db_path, path)
            except Exception as e:
                result["error"] = e
            result["done"] = True

        threading.Thread(target=work, daemon=True).start()
        self.config(cursor="watch")
        self.status_var.set("Backing up the database...")

        def poll():
            if not result.get("done"):
                self.after(300, poll)
                return
            self.config(cursor="")
            if "error" in result:
                messagebox.showerror(APP_TITLE, f"The backup failed:\n\n{result['error']}")
                return
            self.backed_up = True
            self.status_var.set(f"Backup saved to {path}")
            if then:
                then()
        poll()

    def quit_app(self):
        try:
            if self.backend:
                self.backend.close()
        finally:
            self.destroy()


def main():
    args = sys.argv[1:]
    demo = "--demo" in args
    paths = [a for a in args if not a.startswith("--")]
    app = App(db_path=paths[0] if paths else None, demo=demo)
    app.mainloop()


if __name__ == "__main__":
    main()
