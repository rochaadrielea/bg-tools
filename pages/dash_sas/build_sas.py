"""
Create the SAS tables inside the shared quality.db.

These tables are OWNED by dash_sas. dash/ingest.py never touches them and this
never touches dash's tables. WAL is enabled so the FastAPI reader and the ingest
writer do not block each other on the one SQLite file.

Tables:
  sas_nc      - one row per notification (the parsed export)
  sas_import  - append-only log of every ingest run (file, rows, when)
"""
from __future__ import annotations
import sqlite3
from pathlib import Path

DEFAULT_DB = Path.home() / "bgtools" / "dash" / "quality.db"

SAS_NC_COLS = [
    ("notification", "TEXT PRIMARY KEY"),
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
    return con


def build(con: sqlite3.Connection, rebuild: bool = True) -> None:
    """Create the tables. rebuild=True drops sas_nc first (default)."""
    cur = con.cursor()
    if rebuild:
        cur.execute("DROP TABLE IF EXISTS sas_nc")
    cols = ", ".join(f'"{n}" {t}' for n, t in SAS_NC_COLS)
    cur.execute(f"CREATE TABLE IF NOT EXISTS sas_nc ({cols})")
    cur.execute(
        "CREATE TABLE IF NOT EXISTS sas_import ("
        "id INTEGER PRIMARY KEY AUTOINCREMENT, ts TEXT, file TEXT, "
        "rows INTEGER, mode TEXT, project TEXT)"
    )
    con.commit()


if __name__ == "__main__":
    con = connect()
    build(con, rebuild=True)
    print(f"sas tables ready in {DEFAULT_DB}")
    con.close()
