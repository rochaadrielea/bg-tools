"""
Create the SAS tables inside the shared quality.db.

v2: composite PK (source_system, notification), import_id FK, reference tables.
No generated ids. Every value is either from the source file or from the load
itself (timestamp, file name, SHA-256).

Tables:
  sas_import      one row per load (file, SHA, timestamp, rows, mode)
  sas_nc          one row per NC, PK = (source_system, notification)
  sas_ref_batch   known batch labels (reported, not enforced)
  sas_ref_defect  known defect codes
  sas_ref_vendor  known vendors
  feedback        shared across pages
"""
from __future__ import annotations
import sqlite3
from pathlib import Path

DEFAULT_DB = Path.home() / "bgtools" / "dash" / "quality.db"

SAS_NC_COLS = [
    ("source_system", "TEXT NOT NULL"),
    ("source_file", "TEXT"),
    ("import_id", "INTEGER REFERENCES sas_import(id)"),
    ("notification", "TEXT NOT NULL"),
    ("notif_type", "TEXT"),
    ("notif_year", "INTEGER"),
    ("status", "TEXT"),
    ("wbs", "TEXT"),
    ("wbs_text", "TEXT"),
    ("batch", "TEXT"),
    ("defect_class", "TEXT"),
    ("defect_class_label", "TEXT"),
    ("defect_code", "TEXT"),
    ("disposition", "TEXT"),
    ("cause", "TEXT"),
    ("notif_text", "TEXT"),
    ("material", "TEXT"),
    ("model", "TEXT"),
    ("vendor", "TEXT"),
    ("vendor_clean", "TEXT"),
    ("opened", "TEXT"),
    ("closed", "TEXT"),
    ("month", "TEXT"),
    ("leadtime", "REAL"),
    ("copq", "REAL"),
    ("copq_booked", "INTEGER"),
]


def connect(db_path: Path = DEFAULT_DB) -> sqlite3.Connection:
    con = sqlite3.connect(str(db_path))
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA busy_timeout=5000")
    con.execute("PRAGMA foreign_keys=ON")
    return con


def build(con: sqlite3.Connection, rebuild: bool = True) -> None:
    """Create the tables. rebuild=True drops sas_nc first (default)."""
    cur = con.cursor()

    # import log — never dropped; migrate if needed
    cur.execute("""
        CREATE TABLE IF NOT EXISTS sas_import (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT, file TEXT, sha256 TEXT,
            rows INTEGER, mode TEXT, project TEXT
        )""")
    # migration: add sha256 column if table was created by v1
    try:
        cur.execute("SELECT sha256 FROM sas_import LIMIT 0")
    except sqlite3.OperationalError:
        cur.execute("ALTER TABLE sas_import ADD COLUMN sha256 TEXT")

    if rebuild:
        cur.execute("DROP TABLE IF EXISTS sas_nc")

    cols = ", ".join(f'"{n}" {t}' for n, t in SAS_NC_COLS)
    cur.execute(f"""
        CREATE TABLE IF NOT EXISTS sas_nc (
            {cols},
            PRIMARY KEY (source_system, notification)
        )""")

    # reference tables — append-only, never dropped
    for ref in ("sas_ref_batch", "sas_ref_defect", "sas_ref_vendor"):
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {ref} (
                value TEXT PRIMARY KEY,
                first_seen TEXT,
                source TEXT
            )""")

    # feedback — shared
    cur.execute("""
        CREATE TABLE IF NOT EXISTS feedback (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ts TEXT, app TEXT, version TEXT, category TEXT,
            message TEXT, user TEXT, machine TEXT, context TEXT
        )""")

    con.commit()


def update_refs(con: sqlite3.Connection, df, source_file: str) -> dict:
    """Insert any new vocabulary values into the reference tables.
    Returns {table: [new_values]} so the caller can report them."""
    from datetime import datetime
    ts = datetime.now().isoformat(timespec="seconds")
    news = {}
    for ref, col in (("sas_ref_batch", "batch"),
                     ("sas_ref_defect", "defect_code"),
                     ("sas_ref_vendor", "vendor_clean")):
        if col not in df.columns:
            continue
        vals = set(str(v).strip() for v in df[col].dropna().unique()
                   if str(v).strip() and str(v).strip().lower()
                   not in ("nan", "none", "unassigned", "not recorded", "(not recorded)"))
        existing = set(r[0] for r in con.execute(
            f"SELECT value FROM {ref}").fetchall())
        new = vals - existing
        if new:
            con.executemany(
                f"INSERT OR IGNORE INTO {ref} (value, first_seen, source) "
                f"VALUES (?, ?, ?)",
                [(v, ts, source_file) for v in sorted(new)])
            news[ref] = sorted(new)
    con.commit()
    return news


if __name__ == "__main__":
    con = connect()
    build(con, rebuild=True)
    print(f"sas tables ready in {DEFAULT_DB}")
    con.close()