"""
build_systems.py — ONE STEP: from the NC tracker, take only the SAP and Blackout
rows (EZ1 left out for now) and write them to TWO SEPARATE tables, cleaned:
    sap        System = SAP        (158)
    blackout   System = Blackout   (27)

Project and Flight Unit are canonicalised against Set_Up (blank/'All'/off-list ->
NO PROJECT; invalid flight unit blanked). EZ1 rows are NOT included. CAPA and any
other table are left untouched.

Reuses ingest.load_tracker (field cleaning) and clean_tracker (Set_Up canon).

Run:  cd ~/bgtools/dash && ./quality/bin/python cleaning/build_systems.py
"""
import os
import re
import sys
import sqlite3

import pandas as pd

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(_HERE))   # dash/
sys.path.insert(0, _HERE)                     # dash/cleaning/

import ingest
import clean_tracker as ct

DB = "quality.db"
NO_PROJECT = "NO PROJECT"
PROJ_CLASS = {"Ariane": "LLV", "Vega": "MLV", "Vega-C": "MLV", "MHI_H3": "LLV",
              "Relativity": "LLV", "SAS": "SAS", "Vulcan": "LLV",
              "Flexline": "SLV", NO_PROJECT: "(no class)"}

SYSTEMS = {"SAP": "sap", "Blackout": "blackout"}   # System value -> table name


def _find_tracker():
    data = os.path.join(os.path.dirname(_HERE), "data")
    cands = [f for f in os.listdir(data)
             if re.search(r"(NCR_Cutover_Tracker|NCtracker|NC_Tracker).*\.xls[xm]$",
                          f, re.I) and not f.startswith("~$")]
    return os.path.join(data, sorted(cands)[0]) if cands else None


def main():
    path = _find_tracker()
    fu_vocab = ct.load_setup_vocab(path)

    df = ingest.load_tracker()      # all tracker fields, cleaned

    # canonicalise project + flight unit against Set_Up (explicit rules)
    def _clean_pf(row):
        p, _ = ct.canon_project(row.get("project"))
        f, projs, _ = ct.canon_flight_unit(row.get("flight_unit"), fu_vocab)
        if not p and projs and len(projs) == 1:
            p = list(projs)[0]
        return pd.Series({"project": p or NO_PROJECT, "flight_unit": f})

    df[["project", "flight_unit"]] = df.apply(_clean_pf, axis=1)
    df["launcher_class"] = df["project"].map(PROJ_CLASS).fillna("(no class)")

    # The dashboard still references these SAP-era columns; add them empty so the
    # nc view has them and no query errors (that data lived only in the SAP export).
    for col in ("copq", "leadtime", "defect_code_text", "defect_class",
                "notification_type", "notif_year", "profit_center",
                "business_unit", "material_key", "system_status", "ir_number"):
        if col not in df.columns:
            df[col] = None

    _sys = df["system"].astype(str).str.strip()

    con = sqlite3.connect(DB)
    for value, table in SYSTEMS.items():
        part = df[_sys == value].copy()
        part.to_sql(table, con, if_exists="replace", index=False)
        cur = con.cursor()
        for col in ("project", "launcher_class", "flight_unit", "status", "owner"):
            if col in part.columns:
                cur.execute(f"CREATE INDEX IF NOT EXISTS idx_{table}_{col} "
                            f"ON {table}({col})")
        con.commit()
        print(f"table '{table}' : {len(part)} rows "
              f"({int((part['is_open']==1).sum())} open)")

    # to_fix — the NCs with NO PROJECT, pulled out into their own table to work
    # on. They stay OUT of the nc view until they get a real project (fix them in
    # the tracker, rebuild — they leave to_fix and rejoin the view automatically).
    # Criterion is 'no project' only: blank flight unit is the norm and would pull
    # almost everything out, so it is not used here.
    tofix = df[(df["project"] == NO_PROJECT) & _sys.isin(SYSTEMS.keys())].copy()
    tofix.to_sql("to_fix", con, if_exists="replace", index=False)
    print(f"\ntable 'to_fix' : {len(tofix)} rows (NO PROJECT — kept out of the view)")

    # nc VIEW = ALL sap + blackout rows (project or not). Open-now metrics count
    # everything; project-breakdown charts filter to a real project themselves.
    # build_ez1.py replaces this with a TABLE that also folds in EZ1 (Matched).
    cur = con.cursor()
    # nc may already be a view OR a table (build_ez1 makes it a table). Drop with
    # the correct keyword for whichever it is — DROP VIEW on a table errors.
    _r = cur.execute("SELECT type FROM sqlite_master WHERE name='nc'").fetchone()
    if _r:
        cur.execute(f"DROP {'VIEW' if _r[0]=='view' else 'TABLE'} nc")
    cur.execute("CREATE VIEW nc AS "
                "SELECT * FROM sap UNION ALL SELECT * FROM blackout")
    con.commit()
    _ncn = pd.read_sql("SELECT COUNT(*) n FROM nc", con).iloc[0]["n"]
    print(f"view 'nc' = sap + blackout (all) : {_ncn} rows "
          "(EZ1 added by build_ez1.py)")

    # what we deliberately left out
    left = df[~_sys.isin(SYSTEMS.keys())]
    print(f"\nleft out (EZ1 etc.): {len(left)} rows — "
          + " | ".join(f"{k}:{v}" for k, v in
                       left['system'].astype(str).str.strip().value_counts().items()))
    try:
        n = pd.read_sql("SELECT COUNT(*) n FROM capa", con).iloc[0]["n"]
        print(f"capa: {n} rows (left untouched)")
    except Exception:
        print("capa: (no capa table)")

    # quick per-project view of each system table
    for table in SYSTEMS.values():
        print(f"\n{table} — NCs per project:")
        print(pd.read_sql(
            f"SELECT project, COUNT(*) ncs, "
            f"SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) open "
            f"FROM {table} GROUP BY project ORDER BY ncs DESC", con
        ).to_string(index=False))
    con.close()


if __name__ == "__main__":
    main()