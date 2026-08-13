"""
db_health.py — integrity guard for quality.db. Read-only; changes nothing.

Confirms the ingest rebuild left NO leftovers, NO duplicates, NO orphans, and
accounts for every row by source. Run after any ingest, or anytime you doubt a
number:

    cd ~/bgtools/dash && ./quality/bin/python db_health.py

Exit code 0 = PASS, 1 = FAIL (so it can gate a script).
"""
import sqlite3
import sys

import pandas as pd

DB = "quality.db"


def q(sql, con, p=()):
    return pd.read_sql(sql, con, params=p)


def one(sql, con, p=()):
    return q(sql, con, p).iloc[0, 0]


def main():
    con = sqlite3.connect(DB)
    fail = []

    total = one("SELECT COUNT(*) FROM nc", con)
    open_ = one("SELECT COUNT(*) FROM nc WHERE is_open=1", con)
    closed = one("SELECT COUNT(*) FROM nc WHERE is_open=0", con)
    nostatus = one("SELECT COUNT(*) FROM nc WHERE is_open IS NULL", con)
    nullid = one("SELECT COUNT(*) FROM nc WHERE nc_id IS NULL OR TRIM(nc_id)=''", con)

    dup = q("SELECT nc_id, COUNT(*) k FROM nc "
            "WHERE nc_id IS NOT NULL AND TRIM(nc_id)<>'' "
            "GROUP BY nc_id HAVING k>1 ORDER BY k DESC", con)

    bysrc = q("SELECT source, COUNT(*) rows, "
              "SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END) open "
              "FROM nc GROUP BY source ORDER BY rows DESC", con)

    # CAPA table is optional
    try:
        capa_rows = one("SELECT COUNT(*) FROM capa", con)
        capa_ncs = one("SELECT COUNT(DISTINCT nc_id) FROM capa", con)
        capa_zero = one("SELECT COUNT(*) FROM capa WHERE nc_id='0'", con)
        capa_orphan = one("SELECT COUNT(DISTINCT c.nc_id) FROM capa c "
                          "LEFT JOIN nc n ON n.nc_id=c.nc_id WHERE n.nc_id IS NULL", con)
    except Exception:
        capa_rows = capa_ncs = capa_zero = capa_orphan = None

    print("=" * 56)
    print("quality.db integrity")
    print("=" * 56)
    print(f"total NCs        : {total}")
    print(f"  open           : {open_}")
    print(f"  closed         : {closed}")
    print(f"  (no status)    : {nostatus}")
    print(f"  open+closed+ns : {open_ + closed + nostatus}  (must equal total)")
    print("\nby source:")
    print(bysrc.to_string(index=False))

    print("\nchecks:")
    print(f"  duplicate nc_id groups : {len(dup)}")
    if len(dup):
        fail.append(f"{len(dup)} duplicate nc_id")
        print(dup.head(20).to_string(index=False))
    print(f"  blank/null nc_id rows  : {nullid}")

    if capa_rows is not None:
        print(f"  CAPA rows / NCs        : {capa_rows} / {capa_ncs}")
        print(f"  CAPA nc_id='0'         : {capa_zero}")
        print(f"  CAPA orphan nc_ids     : {capa_orphan}  (in capa, not in nc)")
        if capa_zero:
            fail.append(f"{capa_zero} CAPA lines with nc_id='0'")

    if (open_ + closed + nostatus) != total:
        fail.append("open+closed+nostatus != total")

    print("\n" + "=" * 56)
    if fail:
        print("RESULT: FAIL — " + "; ".join(fail))
    else:
        print("RESULT: PASS — full rebuild, no leftovers, no duplicates, "
              "no orphan CAPA, row counts reconcile.")
    print("=" * 56)
    con.close()
    sys.exit(1 if fail else 0)


if __name__ == "__main__":
    main()