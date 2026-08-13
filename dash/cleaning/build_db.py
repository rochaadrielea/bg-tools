"""
build_db.py — repopulate quality.db as clean tables, in order:
    STAGE A: project        canonical project list + NO PROJECT
    STAGE B: flight_unit     valid flight units, each tied to a project
    STAGE C: nc              the NC records — SAME shape as the NC tracker, but
                             CLEAN (canonical project / flight unit, no SAP)

Rules (Adriele, 11.08.2026):
  - start from project, then flight unit; reference tables kept separate;
  - project blank / 'All' / 'TBD' / off-list  ->  'NO PROJECT' (never guessed);
  - a flight unit not in Set_Up is blanked;
  - only the tracker + the new system feed nc — the SAP export is NOT read;
  - CAPA is left untouched.

Reuses the tracker/new-system loaders from ingest.py and the Set_Up cleaning
from clean_tracker.py rather than re-implementing them. The field-content rules
themselves (project/flight-unit canon) stay explicit, per the house rule.

Run:  cd ~/bgtools/dash && ./quality/bin/python cleaning/build_db.py
"""
import os
import sys
import sqlite3

import pandas as pd

# make ingest.py (one level up) and clean_tracker.py (same folder) importable
_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(_HERE))   # dash/
sys.path.insert(0, _HERE)                     # dash/cleaning/

import ingest
import clean_tracker as ct

DB = "quality.db"

# STAGE A — canonical projects and launcher class. NO PROJECT is id 0.
PROJECTS = [
    ("Ariane", "LLV"), ("Vega", "MLV"), ("Vega-C", "MLV"), ("MHI_H3", "LLV"),
    ("Relativity", "LLV"), ("SAS", "SAS"), ("Vulcan", "LLV"), ("Flexline", "SLV"),
]
NO_PROJECT = "NO PROJECT"
PROJ_ID = {NO_PROJECT: 0, **{name: i for i, (name, _) in enumerate(PROJECTS, 1)}}
PROJ_CLASS = {NO_PROJECT: "(no class)", **{name: cls for name, cls in PROJECTS}}


def build_project_table(con):
    cur = con.cursor()
    cur.execute("DROP TABLE IF EXISTS project")
    cur.execute("CREATE TABLE project (project_id INTEGER PRIMARY KEY, "
                "project TEXT UNIQUE NOT NULL, launcher_class TEXT NOT NULL)")
    cur.execute("INSERT INTO project VALUES (0,?,?)", (NO_PROJECT, "(no class)"))
    for i, (name, cls) in enumerate(PROJECTS, start=1):
        cur.execute("INSERT INTO project VALUES (?,?,?)", (i, name, cls))
    con.commit()


def build_flight_unit_table(con, fu_vocab):
    """One row per (flight unit, project). An ambiguous unit (FM10 -> MHI_H3,
    Relativity, Vega) gets one row per project it can belong to."""
    cur = con.cursor()
    cur.execute("DROP TABLE IF EXISTS flight_unit")
    cur.execute("CREATE TABLE flight_unit (flight_unit_id INTEGER PRIMARY KEY, "
                "flight_unit TEXT NOT NULL, project_id INTEGER NOT NULL, "
                "project TEXT NOT NULL)")
    rid = 1
    for code_upper in sorted(fu_vocab):
        for proj in sorted(fu_vocab[code_upper]):
            cur.execute("INSERT INTO flight_unit VALUES (?,?,?,?)",
                        (rid, code_upper, PROJ_ID.get(proj, 0), proj))
            rid += 1
    con.commit()


# Set_Up single-column vocabularies -> one reference table each. These are the
# allowed values for the matching nc fields; cleaning the nc values against them
# is the next step, but the lookup tables are built here, together.
REFERENCE_COLUMNS = {
    "Detection": "detection",
    "Failure": "failure",
    "Q_Responsible": "owner",
    "Disposition": "disposition",
    "Classification": "classification",
    "Status": "status",
    "System": "system",
    "NC Type": "nc_type",
}


def build_reference_tables(con, setup_path):
    """Build one lookup table per Set_Up vocabulary column (detection, failure,
    owner, …). Each is (id, value) with the canonical values in sheet order."""
    su = pd.read_excel(setup_path, sheet_name=ct.SETUP_SHEET, header=None, dtype=str)
    headers = {str(v).strip(): i for i, v in enumerate(su.iloc[0])
               if str(v).strip() and str(v) != "nan"}
    cur = con.cursor()
    built = {}
    for hdr, table in REFERENCE_COLUMNS.items():
        if hdr not in headers:
            continue
        col = headers[hdr]
        seen = []
        for x in su.iloc[1:, col].dropna():
            v = str(x).strip()
            if v and v not in seen:
                seen.append(v)
        cur.execute(f"DROP TABLE IF EXISTS {table}")
        cur.execute(f"CREATE TABLE {table} "
                    f"(id INTEGER PRIMARY KEY, value TEXT UNIQUE NOT NULL)")
        for i, v in enumerate(seen, start=1):
            cur.execute(f"INSERT INTO {table} VALUES (?,?)", (i, v))
        built[table] = len(seen)
    con.commit()
    return built


def build_nc_table(con, tracker_path, fu_vocab):
    """The clean NC table: tracker + new system only, project/flight-unit
    canonicalised against Set_Up, no SAP."""
    tracker = ingest.load_tracker()          # all tracker fields, cleaned

    # --- canonicalise project + flight unit against Set_Up (explicit rules) ---
    def _clean_pf(row):
        p, _ = ct.canon_project(row.get("project"))
        f, projs, _ = ct.canon_flight_unit(row.get("flight_unit"), fu_vocab)
        if not p and projs and len(projs) == 1:   # recover project from an
            p = list(projs)[0]                     # unambiguous flight unit
        return pd.Series({"project": p or NO_PROJECT, "flight_unit": f})

    tracker[["project", "flight_unit"]] = tracker.apply(_clean_pf, axis=1)

    # --- fold in the new system (dedup by IR / TC ID; no double counting) ---
    merged = ingest.merge_new_system(tracker, ingest.load_new_system())

    # new-system rows carry no project -> NO PROJECT; keep flight unit blank
    merged["project"] = merged["project"].fillna("").replace("", NO_PROJECT)
    if "flight_unit" not in merged.columns:
        merged["flight_unit"] = ""
    merged["flight_unit"] = merged["flight_unit"].fillna("")
    merged["launcher_class"] = merged["project"].map(PROJ_CLASS).fillna("(no class)")
    merged["project_id"] = merged["project"].map(PROJ_ID).fillna(0).astype(int)

    # SAP is no longer read, but the dashboard still references these columns.
    # Add them empty so no query errors — the CoPQ/leadtime/defect tiles simply
    # show nothing (that data lived only in the SAP export).
    for col in ("copq", "leadtime", "defect_code_text", "defect_class",
                "notification_type", "notif_year", "profit_center",
                "business_unit", "material_key", "system_status"):
        if col not in merged.columns:
            merged[col] = None

    merged.to_sql("nc", con, if_exists="replace", index=False)
    cur = con.cursor()
    for col in ("project", "launcher_class", "flight_unit", "status",
                "owner", "source", "created_on"):
        if col in merged.columns:
            cur.execute(f"CREATE INDEX IF NOT EXISTS idx_nc_{col} ON nc({col})")
    con.commit()
    return merged


def main():
    tracker_path = ct._find_tracker() if hasattr(ct, "_find_tracker") else None
    if not tracker_path:
        import re
        cands = [f for f in os.listdir(os.path.join(os.path.dirname(_HERE), "data"))
                 if re.search(r"(NCR_Cutover_Tracker|NCtracker|NC_Tracker).*\.xls[xm]$",
                              f, re.I) and not f.startswith("~$")]
        tracker_path = os.path.join(os.path.dirname(_HERE), "data", sorted(cands)[0])

    fu_vocab = ct.load_setup_vocab(tracker_path)
    con = sqlite3.connect(DB)

    build_project_table(con)
    build_flight_unit_table(con, fu_vocab)
    refs = build_reference_tables(con, tracker_path)
    nc = build_nc_table(con, tracker_path, fu_vocab)

    print("STAGE A — project:")
    print(pd.read_sql("SELECT * FROM project ORDER BY project_id", con).to_string(index=False))
    print("\nSTAGE B — flight_unit (first 12 of "
          f"{pd.read_sql('SELECT COUNT(*) n FROM flight_unit', con).iloc[0]['n']} rows):")
    print(pd.read_sql("SELECT * FROM flight_unit ORDER BY flight_unit_id LIMIT 12",
                      con).to_string(index=False))
    print("\nreference tables built (id, value):")
    for tbl, k in refs.items():
        print(f"  {tbl:14s} {k} values")

    print("\nSTAGE C — nc (clean, tracker + new system, no SAP):")
    print(f"  total NCs : {len(nc)}")
    print(f"  open      : {int((nc['is_open']==1).sum())}")
    print("\n  NCs per project:")
    print(pd.read_sql("SELECT project, COUNT(*) ncs, "
                      "SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) open "
                      "FROM nc GROUP BY project ORDER BY ncs DESC", con).to_string(index=False))
    try:
        n = pd.read_sql("SELECT COUNT(*) n FROM capa", con).iloc[0]["n"]
        print(f"\ncapa: {n} rows (left untouched)")
    except Exception:
        print("\ncapa: (no capa table)")
    con.close()


if __name__ == "__main__":
    main()