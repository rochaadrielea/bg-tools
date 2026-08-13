"""
build_ez1.py — load the EZ1 / OpCenter–Teamcenter export into a SEPARATE `ez1`
table. NOT wired into the dashboard analytics: EZ1 has no project / flight unit,
so it stays out of the `nc` view (Adriele). This table is just a clean store we
can look at and, later, enrich.

Source file: nc_opcenter_tc_*.xlsx with sheets:
    Matched          IR + NC + Status + Severity + Detection + …   (in both systems)
    OpCenter only    NC + Status + Severity                         (no IR)
    Teamcenter only  IR + NC + Detection + …                        (no OpCenter status)

Every row is kept, tagged with which sheet it came from.

Run:  cd ~/bgtools/dash && ./quality/bin/python cleaning/build_ez1.py
"""
import os
import re
import sqlite3

import pandas as pd

_HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(os.path.dirname(_HERE), "data")
DB = "quality.db"


def _find_ez1():
    for f in sorted(os.listdir(DATA_DIR)):
        if f.startswith("~$") or not f.lower().endswith((".xlsx", ".xlsm")):
            continue
        if re.search(r"opcenter|nc_opcenter", f, re.I):
            return os.path.join(DATA_DIR, f)
        try:                       # or detect by the 'Matched' sheet layout
            xl = pd.ExcelFile(os.path.join(DATA_DIR, f))
            if "Matched" in xl.sheet_names and "Teamcenter only" in xl.sheet_names:
                return os.path.join(DATA_DIR, f)
        except Exception:
            continue
    return None


def _pick(df, *names):
    for n in names:
        if n in df.columns:
            return df[n]
    return pd.Series([None] * len(df))


def main():
    path = _find_ez1()
    if not path:
        print(f"No EZ1/OpCenter export (nc_opcenter_tc*.xlsx) in {DATA_DIR}/")
        return
    print(f"EZ1 export: {os.path.basename(path)}")
    xl = pd.ExcelFile(path)

    # Only the 'Matched' tab (Adriele): NCs present in BOTH OpCenter and
    # Teamcenter — the ones with a real IR + NC + Status. OpCenter-only and
    # Teamcenter-only tabs are ignored.
    frames = []
    for sheet in ("Matched",):
        if sheet not in xl.sheet_names:
            continue
        d = pd.read_excel(path, sheet_name=sheet, dtype=str)
        d.columns = [str(c).strip() for c in d.columns]
        out = pd.DataFrame()
        out["source_sheet"] = [sheet] * len(d)
        out["ir_number"] = _pick(d, "IR")
        out["nc_ref"] = _pick(d, "NC", "Identifier")
        out["status"] = _pick(d, "Status")
        out["severity"] = _pick(d, "Severity")
        out["detection"] = _pick(d, "Detection")
        out["description"] = _pick(d, "Description")
        out["owner"] = _pick(d, "Issue Owner", "User")
        out["problem_item"] = _pick(d, "Problem Item")
        out["containment_action"] = _pick(d, "Containment Action")
        out["failure"] = _pick(d, "Failure Linked with Symptom Defect")
        out["creation_date"] = _pick(d, "Creation Date", "Start Date")
        out["end_date"] = _pick(d, "End Date")
        frames.append(out)

    ez1 = pd.concat(frames, ignore_index=True)

    # OPEN/CLOSED comes from End Date (Adriele): an EZ1 NC with an End Date is
    # CLOSED, and that date is its closure date. No End Date -> still open.
    def _isodate(v):
        try:
            s = str(v).strip()
            if not s or s.lower() == "nan":
                return None
            return pd.to_datetime(v).date().isoformat()
        except Exception:
            return None
    ez1["closure_date"] = ez1["end_date"].apply(_isodate)
    # closed (is_open=0) exactly when a closure date exists. Use isna() — a bare
    # `if d` treats pandas' NaN as truthy and marks everything closed.
    ez1["is_open"] = ez1["closure_date"].isna().astype(int)

    con = sqlite3.connect(DB)
    ez1.to_sql("ez1", con, if_exists="replace", index=False)
    con.execute("CREATE INDEX IF NOT EXISTS idx_ez1_status ON ez1(status)")
    con.commit()

    print(f"\ntable 'ez1' : {len(ez1)} rows (SEPARATE — not in the dashboard/analytics)")
    print("\nby source sheet:")
    print(ez1["source_sheet"].value_counts().to_string())
    print("\nby status:")
    print(ez1["status"].fillna("(blank)").value_counts().to_string())
    print(f"\nopen/closed (from End Date): open {int((ez1['is_open']==1).sum())} | "
          f"closed {int((ez1['is_open']==0).sum())} "
          f"(closed = has an End Date; that date is the closure date)")

    # ---- assemble the nc TABLE = sap + blackout + ez1 (Matched), so open-now
    # metrics count ALL systems. EZ1 has no project -> NO PROJECT / (no class);
    # project-breakdown charts still filter to a real project themselves. ----
    try:
        sap = pd.read_sql("SELECT * FROM sap", con)
        black = pd.read_sql("SELECT * FROM blackout", con)
    except Exception:
        print("\n! sap/blackout not found — run build_systems.py first. nc not built.")
        con.close(); return

    cols = list(sap.columns)
    def _iso(v):
        try:
            return pd.to_datetime(v).date().isoformat()
        except Exception:
            return None
    m = pd.DataFrame({c: [None] * len(ez1) for c in cols})
    m["nc_id"] = ez1["ir_number"].values
    if "tc_id" in cols: m["tc_id"] = ez1["nc_ref"].values
    if "ir_number" in cols: m["ir_number"] = ez1["ir_number"].values
    m["system"] = "EZ1"
    m["project"] = "NO PROJECT"
    if "launcher_class" in cols: m["launcher_class"] = "(no class)"
    if "flight_unit" in cols: m["flight_unit"] = ""
    m["is_open"] = ez1["is_open"].values
    if "status" in cols: m["status"] = ez1["status"].values
    if "status_state" in cols:
        m["status_state"] = (ez1["is_open"].map({1: "Open", 0: "Closed"})
                             .fillna("(no status)").values)
    if "owner" in cols: m["owner"] = ez1["owner"].values
    if "detection_area" in cols: m["detection_area"] = ez1["detection"].values
    if "description" in cols: m["description"] = ez1["description"].values
    if "created_on" in cols: m["created_on"] = [_iso(x) for x in ez1["creation_date"]]
    if "closure_date" in cols: m["closure_date"] = ez1["closure_date"].values
    if "source" in cols: m["source"] = "ez1"

    ncall = pd.concat([sap, black, m[cols]], ignore_index=True)
    cur = con.cursor()
    _r = cur.execute("SELECT type FROM sqlite_master WHERE name='nc'").fetchone()
    if _r:
        cur.execute(f"DROP {'VIEW' if _r[0]=='view' else 'TABLE'} nc")
    ncall.to_sql("nc", con, if_exists="replace", index=False)
    con.commit()

    print(f"\nnc TABLE = sap + blackout + ez1 : {len(ncall)} rows")
    print(pd.read_sql("SELECT system, COUNT(*) rows, "
                      "SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) open "
                      "FROM nc GROUP BY system", con).to_string(index=False))
    print(f"TOTAL OPEN (all systems): "
          f"{pd.read_sql('SELECT COUNT(*) n FROM nc WHERE is_open=1', con).iloc[0]['n']}")
    con.close()


if __name__ == "__main__":
    main()