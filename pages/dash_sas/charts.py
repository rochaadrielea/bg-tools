"""
Turn the sas_nc rows into the payload the page renders.

One function, build_payload(df), takes an already-filtered DataFrame and returns
a dict of every block's data. The API applies global + section filters, then
calls this. Each block also declares the exact filter (column + value) a click
produces, so Rule 1 (click -> raw view) is driven from here, not the frontend.

Raw view columns (Rule 1), the same set everywhere, including the document each
row came from (source_file, stamped by ingest_sas at load time).
"""
from __future__ import annotations
import pandas as pd
import parse

RAW_COLS = [
    "notification", "notif_text", "notif_type", "origin", "status", "opened",
    "closed", "leadtime", "batch", "material", "defect_class_label",
    "defect_code", "cause", "disposition", "vendor_clean", "copq",
    "source_file", "status_note",
]
RAW_LABELS = {
    "notification": "NC", "notif_text": "Title", "notif_type": "Type",
    "origin": "Origin", "status": "Status", "opened": "Opened",
    "closed": "Closed", "leadtime": "Leadtime", "batch": "Batch",
    "material": "Material", "defect_class_label": "Class",
    "defect_code": "Defect code", "cause": "Cause",
    "disposition": "Disposition", "vendor_clean": "Vendor", "copq": "CoPQ",
    "source_file": "Source",
    "status_note": "Status Note",
}

MONTHS_ORDER = None  # chronological from data


def _months(df):
    return sorted(m for m in df["month"].dropna().unique())


def _clean_cat(s):
    """Turn nan / None / '-' text into a single readable '(not recorded)' label
    so a missing value never shows up as the literal string 'nan' on a chart."""
    out = s.astype("string")
    out = out.str.strip()
    out = out.mask(out.isna() | out.str.lower().isin(
        ["nan", "none", "", "-", "null"]), "(not recorded)")
    return out


def _count_by(df, col, order=None):
    vc = _clean_cat(df[col]).value_counts()
    if order:
        idx = [x for x in order if x in vc.index] + \
              [x for x in vc.index if x not in order]
        vc = vc.reindex(idx).fillna(0)
    return {"labels": list(vc.index), "values": [int(v) for v in vc.values]}


def _stack_by(df, row_col, stack_col, row_order=None, stack_order=None):
    ct = pd.crosstab(_clean_cat(df[row_col]), _clean_cat(df[stack_col]))
    if row_order:
        ct = ct.reindex([r for r in row_order if r in ct.index])
    rows = list(ct.index)
    stacks = stack_order or list(ct.columns)
    series = {s: [int(ct.loc[r, s]) if s in ct.columns else 0 for r in rows]
              for s in stacks}
    return {"rows": rows, "series": series}


def build_payload(df: pd.DataFrame) -> dict:
    df = df.copy()
    # rows may arrive from SQLite (all text) or from the parser (typed) - coerce
    for c in ("opened", "closed"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
    for c in ("leadtime", "copq"):
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["copq"] = df["copq"].fillna(0)
    if "copq_booked" in df.columns:
        df["copq_booked"] = pd.to_numeric(df["copq_booked"], errors="coerce").fillna(0)

    months = _months(df)
    df = parse.tag_lanes(df)
    batch_order = sorted(df["batch"].dropna().unique(), key=parse.batch_sort_key)
    # Binary split: supplier vs everything else (production). Customer
    # complaints have their own chart, so they stay out of production views.
    prod = df[df["lane"] == "Production"]
    supp = df[df["lane"] == "Supplier"]
    cc = df[df["lane"] == "Customer complaint"]
    openn = prod[prod["status"] == "Open"]

    p = {}

    # ---- 1 Overview KPI strip ----
    closed = df[df["status"] == "Closed"]
    p["kpi"] = {
        "total": int(len(df)),
        "open": int((df["status"] == "Open").sum()),
        "closed": int((df["status"] == "Closed").sum()),
        "deleted": int((df["status"] == "Deleted").sum()),
        "copq_total": round(float(df["copq"].sum()), 0),
        "copq_ytd": round(float(df[df["notif_year"] == 2026]["copq"].sum()), 0),
        "lead_median": round(float(closed["leadtime"].median()), 0) if len(closed) else 0,
        "scrap": int((df["disposition"] == "Scrap").sum()),
    }

    # ---- 2 Status — production and supplier are never on the same chart ----
    p["status_prod"] = {
        **_count_by(prod, "status", order=["Closed", "Open", "Deleted"]),
        "click": "status", "scope": {"lane": "Production"},
    }
    p["status_supp"] = {
        **_count_by(supp, "status", order=["Closed", "Open", "Deleted"]),
        "click": "status", "scope": {"lane": "Supplier"},
    }
    # "NCs per part" — every NC has a part, not just production
    partc = _count_by(df, "material")
    p["by_part"] = {**partc, "click": "material",
                    "labels": partc["labels"][:20], "values": partc["values"][:20]}

    # ---- 3 Batches ----
    p["batch_z3_status"] = {
        **_stack_by(prod, "batch", "status", row_order=batch_order,
                    stack_order=["Closed", "Open", "Deleted"]),
        "click": "batch", "scope": {"lane": "Production"},
    }
    p["batch_cc"] = {**_count_by(cc, "batch", order=batch_order), "click": "batch",
                     "scope": {"lane": "Customer complaint"}}
    # Defect class applies to EVERY NC. The chart has an All/Supplier/Production
    # chip: All keeps Unclassified; Supplier/Production show Minor+Major only.
    # ORIGIN is the two-value field (Production incl. customer complaints, vs Supplier).
    p["batch_class"] = {
        **_stack_by(df, "batch", "defect_class_label", row_order=batch_order,
                    stack_order=["Minor", "Major", "Unclassified"]),
        "click": "batch",
    }
    # For the Supplier/Production multi-select on this chart, give the client the
    # class breakdown per lane so it can include whichever lanes are toggled on.
    # Both on (default) = the full batch_class above (all 180, incl. Unclassified).
    df_supp = df[df["origin"] == "Supplier"]
    df_prod = df[df["origin"] == "Production"]
    def _class_by_batch(d):
        # Align counts to the FULL batch_order (pad missing batches with 0), so the
        # client can index series[i] against rows[i] without misalignment. Without
        # this, _stack_by drops empty batches and the series is shorter than rows,
        # which makes Lager's value land on the wrong bar (e.g. Batch 5).
        st = _stack_by(d, "batch", "defect_class_label", row_order=batch_order,
                       stack_order=["Minor", "Major", "Unclassified"])
        idx = {b: i for i, b in enumerate(st["rows"])}
        series = {}
        for cls in ["Minor", "Major", "Unclassified"]:
            src = st["series"].get(cls, [])
            series[cls] = [src[idx[b]] if b in idx else 0 for b in batch_order]
        return series
    p["batch_class_lanes"] = {
        "rows": batch_order,
        "supplier": _class_by_batch(df_supp),
        "production": _class_by_batch(df_prod),
        "click": "batch",
    }
    # CoPQ per batch — cost applies to every NC, supplier or production
    copq_batch = df.groupby("batch")["copq"].sum().reindex(batch_order).fillna(0)
    p["batch_copq"] = {"labels": list(copq_batch.index),
                       "values": [round(float(v), 0) for v in copq_batch.values],
                       "click": "batch"}
    # Batch x defect code — every NC has a defect code, not just production
    p["batch_defect_heat"] = _heatmap(df, "batch", "defect_code", batch_order)

    # ---- 4 Suppliers (Z2) ----
    p["z2_month_status"] = {
        **_stack_by(supp, "month", "status", row_order=months,
                    stack_order=["Closed", "Open", "Deleted"]),
        "click": "month", "scope": {"lane": "Supplier"},
    }
    vend_order = list(supp["vendor_clean"].value_counts().index)
    p["z2_vendor_month"] = _stack_by(supp, "vendor_clean", "month",
                                    row_order=vend_order, stack_order=months)
    p["z2_vendor_month"]["click"] = "vendor_clean"
    p["z2_vendor_month"]["scope"] = {"lane": "Supplier"}
    vp = supp.groupby("vendor_clean").agg(count=("notification", "size"),
                                         copq=("copq", "sum"))
    vp = vp.sort_values("count", ascending=False)
    p["vendor_pareto"] = {
        "labels": list(vp.index),
        "count": [int(v) for v in vp["count"].values],
        "copq": [round(float(v), 0) for v in vp["copq"].values],
        "click": "vendor_clean", "scope": {"lane": "Supplier"},
    }

    # ---- 5 Springs ----
    # Springs NCs by status — all NCs on the Springs batch, any lane
    spr = df[df["batch"] == "Springs"]
    p["springs"] = {
        **_count_by(spr, "status", order=["Closed", "Open", "Deleted"]),
        "click": "status",
        "scope": {"batch": "Springs"},
    }

    # ---- 6 Open backlog — OPEN production NCs only ----
    # Open NCs by age — every open NC, any lane
    openall = df[df["status"] == "Open"]
    p["backlog_age"] = {**_age_buckets(openall),
                        "scope": {"status": "Open"}}
    oldest = openall.sort_values("opened").head(10)
    p["backlog_oldest"] = _raw_rows(oldest)

    # ---- 7 Defects & Cost ----
    p["defect_pareto"] = _pareto(df, "defect_code")
    p["disposition"] = {**_count_by(df, "disposition"), "click": "disposition",
                        "highlight": "Scrap"}
    top = df.sort_values("copq", ascending=False).head(10)
    p["top_copq"] = {
        "labels": list(top["notification"] + " · " + top["batch"].fillna("")),
        "values": [round(float(v), 0) for v in top["copq"].values],
        "ncs": list(top["notification"]),
        "click": "notification",
    }
    copq_month = df.groupby("month")["copq"].sum().reindex(months).fillna(0)
    p["copq_month"] = {"labels": list(copq_month.index),
                       "values": [round(float(v), 0) for v in copq_month.values],
                       "click": "month"}
    p["cause_month_heat"] = _heatmap(df, "cause", "month", None, col_order=months)

    # ---- 8 Data quality — overall and per source system ----
    def _blank(s):
        return s.isna() | s.astype(str).str.strip().isin(
            ["", "nan", "None", "-", "Not assigned"])

    src = ["Teamcenter" if str(n).startswith("IR-") else "SAP"
           for n in df["notification"]]
    df = df.assign(source_system=src)

    def _gaps(d):
        return {
            "rows": int(len(d)),
            "defect_class_blank": int((d["defect_class_label"] == "Unclassified").sum()),
            "disposition_missing": int(_blank(d["disposition"]).sum()),
            "copq_unbooked": int((~d["copq_booked"].astype(bool)).sum()),
            "vendor_missing": int((d["vendor_clean"] == "Not recorded").sum()),
            "batch_missing": int((d["batch"] == "Unassigned").sum()),
            "model_missing": int(_blank(d["model"]).sum()),
            "closed_no_date": int(((d["status"] == "Closed") &
                                   d["closed"].isna()).sum()),
        }

    p["dq"] = {**_gaps(df), "total": int(len(df))}
    p["dq_by_source"] = {s: _gaps(g) for s, g in df.groupby("source_system")}

    # ---- 8b Data quality per DOCUMENT (which file each row came from) ----
    # Separate from dq_by_source: that one is the SYSTEM (SAP / Teamcenter),
    # this one is the uploaded file. Only built when the column carries values,
    # so rows loaded before source_file existed do not create a blank group.
    if "source_file" in df.columns and df["source_file"].notna().any():
        byfile = df.assign(
            source_file=df["source_file"].fillna("(not recorded)").astype(str))
        p["dq_by_file"] = {f: _gaps(g) for f, g in byfile.groupby("source_file")}
    else:
        p["dq_by_file"] = {}

    p["months"] = months
    return p


def _heatmap(df, row_col, col_col, row_order, col_order=None):
    ct = pd.crosstab(df[row_col].fillna("(blank)"), df[col_col].fillna("(blank)"))
    rows = ([r for r in row_order if r in ct.index] if row_order else list(ct.index))
    cols = ([c for c in col_order if c in ct.columns] if col_order else list(ct.columns))
    z = [[int(ct.loc[r, c]) if (r in ct.index and c in ct.columns) else 0
          for c in cols] for r in rows]
    # Plotly draws y[0] at the bottom; flip so Batch 1 / first part sits at the top.
    rows, z = list(reversed(rows)), list(reversed(z))
    return {"rows": rows, "cols": cols, "z": z,
            "row_col": row_col, "col_col": col_col}


def _pareto(df, col):
    vc = _clean_cat(df[col]).value_counts()
    total = int(vc.sum())
    cum, running = [], 0
    for v in vc.values:
        running += int(v)
        cum.append(round(100 * running / total, 1) if total else 0)
    return {"labels": list(vc.index), "values": [int(v) for v in vc.values],
            "cum": cum, "click": col}


def _age_buckets(openn):
    if openn.empty:
        return {"labels": ["0-30", "31-60", "61-90", "90+"], "values": [0, 0, 0, 0],
                "click": "age_bucket"}
    today = pd.Timestamp.today().normalize()
    age = (today - openn["opened"]).dt.days
    buckets = ["0-30", "31-60", "61-90", "90+"]
    counts = [
        int(((age >= 0) & (age <= 30)).sum()),
        int(((age > 30) & (age <= 60)).sum()),
        int(((age > 60) & (age <= 90)).sum()),
        int((age > 90).sum()),
    ]
    return {"labels": buckets, "values": counts, "click": "age_bucket"}


def _raw_rows(df):
    out = []
    for _, r in df.iterrows():
        row = {}
        for c in RAW_COLS:
            v = r.get(c)
            if pd.isna(v):
                v = ""
            elif c == "copq":
                v = round(float(v), 0)
            elif c == "leadtime":
                v = "" if pd.isna(r.get(c)) else int(v)
            row[c] = v
        out.append(row)
    return out