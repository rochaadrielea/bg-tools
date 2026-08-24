"""
Ingest a SAS NC export into quality.db.

  uv run ingest_sas.py path/to/export.xlsx              # rebuild (default)
  uv run ingest_sas.py path/to/export.xlsx --append      # keep rows, upsert

--rebuild : DROP + recreate sas_nc, then load (default for now).
--append  : keep existing rows; upsert on (source_system, notification).

Validation runs BEFORE any write. RED findings block the load entirely.
"""
from __future__ import annotations
import argparse
import hashlib
import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd
import parse
import build_sas
import validate as val


def _sha(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for c in iter(lambda: f.read(65536), b""):
            h.update(c)
    return h.hexdigest()


def _source_system(nc: str) -> str:
    """Derive from the notification number — not generated, observed."""
    s = str(nc).strip()
    if s.startswith("IR-"):
        return "Teamcenter"
    return "SAP"


def ingest(con: sqlite3.Connection, path_or_buffer, source_name: str,
           rebuild: bool = True, project: str = parse.PROJECT,
           sha: str = "", current_count: int = 0) -> dict:
    """Returns {ok, rows, findings, new_refs, error}."""
    try:
        df = parse.load(path_or_buffer, project=project)
    except Exception as e:
        return {"ok": False, "rows": 0, "findings": [],
                "new_refs": {}, "error": str(e)}

    # --- validate BEFORE writing anything ---
    findings = val.validate(df, current_count=current_count)
    reds = val.reds(findings)
    if reds:
        return {"ok": False, "rows": 0,
                "findings": [(s, r, nc, f, m) for s, r, nc, f, m in findings],
                "new_refs": {},
                "error": val.summary(findings)}

    # --- write ---
    build_sas.build(con, rebuild=rebuild)

    # register the import
    cur = con.cursor()
    cur.execute(
        "INSERT INTO sas_import (ts, file, sha256, rows, mode, project) "
        "VALUES (?,?,?,?,?,?)",
        (datetime.now().isoformat(timespec="seconds"), source_name, sha,
         len(df), "rebuild" if rebuild else "append", project))
    import_id = cur.lastrowid

    # stamp provenance
    df["source_system"] = df["notification"].map(_source_system)
    df["source_file"] = source_name
    df["import_id"] = import_id

    cols = [c for c, _ in build_sas.SAS_NC_COLS]
    rows = []
    for _, r in df.iterrows():
        rows.append(tuple(
            _cell(r, c) for c in cols
        ))

    placeholders = ", ".join("?" for _ in cols)
    collist = ", ".join(f'"{c}"' for c in cols)
    con.executemany(
        f"INSERT OR REPLACE INTO sas_nc ({collist}) VALUES ({placeholders})",
        rows)
    con.commit()

    # update reference tables and report new values
    new_refs = build_sas.update_refs(con, df, source_name)

    return {"ok": True, "rows": len(rows),
            "findings": [(s, r, nc, f, m) for s, r, nc, f, m in findings],
            "new_refs": new_refs, "error": None}


def _cell(r, col):
    v = r.get(col)
    if v is None:
        return None
    if isinstance(v, float) and v != v:  # NaN
        return None
    if col in ("opened", "closed"):
        try:
            return v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v)
        except Exception:
            return None
    if col == "copq_booked":
        return int(bool(v))
    if col in ("notif_year", "import_id"):
        try:
            return int(v)
        except (TypeError, ValueError):
            return None
    if col in ("leadtime", "copq"):
        try:
            f = float(v)
            return None if f != f else f
        except (TypeError, ValueError):
            return None
    return str(v) if v is not None else None


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("export", help="path to the NC export (.xlsx)")
    ap.add_argument("--append", action="store_true", help="keep rows, upsert")
    ap.add_argument("--db", default=str(build_sas.DEFAULT_DB))
    args = ap.parse_args()

    path = Path(args.export)
    sha = _sha(str(path))
    con = build_sas.connect(Path(args.db))

    current = 0
    try:
        current = con.execute("SELECT COUNT(*) FROM sas_nc").fetchone()[0]
    except Exception:
        pass

    result = ingest(con, str(path), path.name,
                    rebuild=not args.append, sha=sha,
                    current_count=current)
    con.close()

    if not result["ok"]:
        print(f"BLOCKED: {result['error']}")
        for f in result["findings"]:
            print(f"  {f[0]:5s} row {f[1]} NC {f[2]}: {f[4]}")
    else:
        mode = "appended/updated" if args.append else "rebuilt with"
        print(f"{mode} {result['rows']} SAS NCs in {args.db}")
        if result["new_refs"]:
            for table, vals in result["new_refs"].items():
                print(f"  new in {table}: {', '.join(vals[:10])}")
        for f in result["findings"]:
            if f[0] != "INFO":
                print(f"  {f[0]:5s} row {f[1]} NC {f[2]}: {f[4]}")