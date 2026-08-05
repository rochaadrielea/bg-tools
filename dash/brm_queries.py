"""
brm_queries.py
--------------
Runs the queries that feed BRM Slides 2, 3, 4.

Each function returns a pandas DataFrame ready to drop into the .pptx.

Usage:
    python brm_queries.py           # prints all
    from brm_queries import *       # import into other scripts
"""

import sqlite3
import pandas as pd

DB_FILE = "quality.db"


def q(sql):
    """Run SQL and return DataFrame."""
    with sqlite3.connect(DB_FILE) as conn:
        return pd.read_sql(sql, conn)


# ============================================================
# SLIDE 2 — Quality Dashboard (headline numbers)
# ============================================================
def slide2_headline_numbers():
    return q("""
        SELECT
            SUM(CASE WHEN is_open=1 THEN 1 ELSE 0 END)                 AS ncs_open_wip,
            SUM(CASE WHEN is_open=0 THEN 1 ELSE 0 END)                 AS ncs_closed_current,
            SUM(CASE WHEN classification LIKE 'Major%' AND is_open=1
                     THEN 1 ELSE 0 END)                                AS major_ncs_open,
            SUM(CASE WHEN is_supplier_nc=1 AND is_open=1 THEN 1 ELSE 0 END) AS supplier_ncs_open,
            SUM(CASE WHEN is_supplier_nc=0 AND is_open=1 THEN 1 ELSE 0 END) AS production_ncs_open
        FROM nc
    """)


# ============================================================
# SLIDE 3 — NC by Project (Top 6), NC by Area (YTD)
# ============================================================
def slide3_top6_projects_wip():
    """Overview of NCs WIP per Project – Top 6."""
    return q("""
        SELECT COALESCE(project, '(no project)') AS project,
               COUNT(*) AS open_ncs
        FROM nc
        WHERE is_open = 1
        GROUP BY project
        ORDER BY open_ncs DESC
        LIMIT 6
    """)


def slide3_top6_projects_ytd():
    """NC Overview - Opened in 2026 by Project (YTD) – Top 6."""
    return q("""
        SELECT COALESCE(project, '(no project)') AS project,
               COUNT(*) AS ncs_opened_ytd
        FROM nc
        WHERE substr(created_on, 1, 4) = '2026'
        GROUP BY project
        ORDER BY ncs_opened_ytd DESC
        LIMIT 6
    """)


def slide3_by_area_ytd():
    """NC Overview - Opened in 2026 by Area (YTD).

    Uses Detection column. 186 rows are blank -> shown as 'BLANK - to clean'.
    This chart will be misleading until Simone's CW26 clean-up is done.
    """
    return q("""
        SELECT COALESCE(detection_area, 'BLANK - to clean') AS area,
               COUNT(*) AS ncs
        FROM nc
        WHERE substr(created_on, 1, 4) = '2026'
        GROUP BY area
        ORDER BY ncs DESC
    """)


# ============================================================
# SLIDE 4 — Production vs Supplier trends
# ============================================================
def slide4_production_vs_supplier():
    return q("""
        SELECT
            CASE WHEN is_supplier_nc = 1 THEN 'Supplier' ELSE 'Production' END AS source,
            SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_ncs,
            SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed_ncs
        FROM nc
        GROUP BY source
    """)


def slide4_monthly_trend():
    """Monthly opens vs closes (for the line chart)."""
    return q("""
        WITH opens AS (
            SELECT substr(created_on, 1, 7) AS month, COUNT(*) AS opened
            FROM nc WHERE created_on IS NOT NULL
            GROUP BY month
        ),
        closes AS (
            SELECT substr(closure_date, 1, 7) AS month, COUNT(*) AS closed
            FROM nc WHERE closure_date IS NOT NULL
            GROUP BY month
        )
        SELECT o.month, o.opened, COALESCE(c.closed, 0) AS closed
        FROM opens o
        LEFT JOIN closes c USING (month)
        ORDER BY o.month
    """)


# ============================================================
# BONUS — Owner breakdown (for weekly review meeting)
# ============================================================
def owner_workload():
    return q("""
        SELECT COALESCE(owner, '(no owner)') AS owner,
               SUM(CASE WHEN is_open = 1 THEN 1 ELSE 0 END) AS open_ncs,
               SUM(CASE WHEN is_open = 0 THEN 1 ELSE 0 END) AS closed_ncs,
               MAX(days_open) AS oldest_days
        FROM nc
        GROUP BY owner
        ORDER BY open_ncs DESC
    """)


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    sections = [
        ("SLIDE 2 — Headline numbers",       slide2_headline_numbers),
        ("SLIDE 3 — Top 6 Projects (WIP)",   slide3_top6_projects_wip),
        ("SLIDE 3 — Top 6 Projects (YTD)",   slide3_top6_projects_ytd),
        ("SLIDE 3 — By Area (YTD)",          slide3_by_area_ytd),
        ("SLIDE 4 — Production vs Supplier", slide4_production_vs_supplier),
        ("SLIDE 4 — Monthly trend",          slide4_monthly_trend),
        ("BONUS — Owner workload",           owner_workload),
    ]
    for title, fn in sections:
        print("\n" + "=" * 60)
        print(title)
        print("=" * 60)
        print(fn().to_string(index=False))
