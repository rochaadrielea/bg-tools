"""
Ingest a SAS NC export into quality.db.

  python ingest_sas.py path/to/export.xlsx              # rebuild (default)
  python ingest_sas.py path/to/export.xlsx --append     # keep rows, upsert

--rebuild : DROP + recreate sas_nc, then load. (default for now)
--append  : keep existing rows; upsert on notification (re-exporting the same
            period updates those NCs instead of duplicating them).

Same function is called by the web upload, so CLI and API load identically.
"""
from __future__ import annotations
import argparse
import sqlite3
from datetime import datetime
from pathlib import Path

import parse
import build_sas


def ingest(con: sqlite3.Connection, path_or_buffer, source_name: str,
           rebuild: bool = True, project: str = parse.PROJECT) -> int:
    df = parse.load(path_or_buffer, project=project)
    build_sas.build(con, rebuild=rebuild)

    cols = [c for c, _ in build_sas.SAS_NC_COLS]
    rows = []
    for _, r in df.iterrows():
        rows.append((
            r["notification"], r["notif_type"], _int(r["notif_year"]), r["status"],
            r["wbs"], r["wbs_text"], r["batch"], str(r["defect_class"]),
            r["defect_class_label"], r["defect_code"], r["disposition"], r["cause"],
            r["notif_text"], r["material"], r["model"], r["vendor"], r["vendor_clean"],
            _dt(r["opened"]), _dt(r["closed"]), r["month"],
            _num(r["leadtime"]), _num(r["copq"]), int(bool(r["copq_booked"])),
        ))

    placeholders = ", ".join("?" for _ in cols)
    collist = ", ".join(f'"{c}"' for c in cols)
    con.executemany(
        f"INSERT OR REPLACE INTO sas_nc ({collist}) VALUES ({placeholders})", rows
    )
    con.execute(
        "INSERT INTO sas_import (ts, file, rows, mode, project) VALUES (?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"), source_name, len(rows),
         "rebuild" if rebuild else "append", project),
    )
    con.commit()
    return len(rows)


def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def _num(v):
    try:
        f = float(v)
        return None if f != f else f  # NaN guard
    except (TypeError, ValueError):
        return None


def _dt(v):
    if v is None or str(v) == "NaT":
        return None
    try:
        return v.strftime("%Y-%m-%d")
    except AttributeError:
        return None


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("export", help="path to the NC's Overview export (.xlsx)")
    ap.add_argument("--append", action="store_true", help="keep rows, upsert")
    ap.add_argument("--db", default=str(build_sas.DEFAULT_DB))
    args = ap.parse_args()

    con = build_sas.connect(Path(args.db))
    n = ingest(con, args.export, Path(args.export).name, rebuild=not args.append)
    con.close()
    mode = "appended/updated" if args.append else "rebuilt with"
    print(f"{mode} {n} SAS NCs in {args.db}")
