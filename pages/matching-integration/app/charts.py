"""
matching-integration dashboard — chart/KPI payload builder.

SCOPE: this dashboard is Integration and Matching only. Every row outside
areas I and M is dropped before anything else happens (_scope), so the KPIs,
the filter lists, the charts and the raw tables all speak about the same
population. The Area filter therefore only ever offers those two.

build_payload(df, filters) applies the active filters, then returns a dict the
front end renders.

Sections served:
  1 Status          nc_status
  2 Root cause      pareto_rc, pareto_det          (drill-down)
  3 Cost            copq_disposition, copq_project, copq_pareto, top_costly
  4 Heat maps       heat_rc_det, heat_proj_rc, heat_month_rc
  5 Time            trend_monthly, backlog, leadtime_hist
  6 Supplier        supplier_monthly, supplier_compare

Drill state travels inside the filters dict under the key "_drill", e.g.
{"_drill": {"rc": ["Machining"], "det": []}}. _apply_filters only reads the
keys it knows, so the extra key passes through harmlessly.
"""
from __future__ import annotations
import pandas as pd
import numpy as np
import parse


SCOPE_AREAS = ["I", "M"]       # the whole point of this dashboard
AREA_NAMES = {"I": "Integration", "M": "Matching"}

PLACE = "(not coded)"          # shown wherever a code is missing
TOP_N = 12                     # bars in a Pareto before "Other"
HEAT_ROWS = 12
HEAT_COLS = 10

RC_KEYS = ["rc1", "rc2"]           # NC root cause drill path
DET_KEYS = ["origin1", "origin2"]  # real detection drill path


# ---------- small helpers ----------

def _clean(s: pd.Series) -> pd.Series:
    """Trim, and turn every flavour of empty into ''."""
    return (s.astype(str).str.strip()
            .replace({"nan": "", "None": "", "NaT": "", "-": ""}))


def _labelled(d: pd.DataFrame, col: str) -> pd.Series:
    """Cleaned column with blanks shown as the placeholder."""
    return _clean(d[col]).replace({"": PLACE})


def _money(d: pd.DataFrame) -> pd.Series:
    return pd.to_numeric(d["copq"], errors="coerce").fillna(0.0)


def _top(series: pd.Series, n: int = TOP_N) -> pd.Series:
    """Sorted descending, everything past n folded into 'Other'."""
    s = series.sort_values(ascending=False)
    if len(s) <= n:
        return s
    head = s.iloc[:n].copy()
    head["Other"] = float(s.iloc[n:].sum())
    return head


def scope(df: pd.DataFrame) -> pd.DataFrame:
    """Integration + Matching only. Called by everything, including main.py's
    raw endpoint, so no view can ever show another area."""
    if df.empty or "area" not in df.columns:
        return df
    return df[_clean(df["area"]).isin(SCOPE_AREAS)]


# ---------- filtering ----------

def _apply_filters(df: pd.DataFrame, f: dict) -> pd.DataFrame:
    """Scope to I/M, then apply the filters. An empty list for a key means
    'no restriction'. Unknown keys (such as "_drill") are ignored."""
    d = scope(df)
    if not f:
        return d

    def keep(col, key):
        vals = f.get(key)
        if vals:
            return d[d[col].astype(str).isin([str(v) for v in vals])]
        return d

    d = keep("area", "area")
    d = keep("notif_type", "notif_type")
    d = keep("rc1", "rc1")
    d = keep("rc2", "rc2")
    d = keep("project", "project")

    # date ranges (year-based, like the Excel timelines)
    for col, key in (("opened", "opened"), ("closed", "closed")):
        rng = f.get(key + "_range")
        if rng and isinstance(rng, (list, tuple)) and len(rng) == 2:
            lo, hi = rng
            s = pd.to_datetime(d[col], errors="coerce")
            if lo:
                d = d[s.dt.year >= int(lo)]
            if hi:
                d = d[s.dt.year <= int(hi)]
    return d


# ---------- option lists for the filters ----------

def _options(df: pd.DataFrame) -> dict:
    """Built from the SCOPED frame, so the filter lists never offer a value
    that would return nothing."""
    df = scope(df)

    def uniq(col, drop_blank=True):
        vals = _clean(df[col])
        seen = {str(v) for v in vals.tolist()}
        return sorted(v for v in seen if (v or not drop_blank))

    def counts(col):
        s = _clean(df[col])
        vc = s.value_counts()
        out = {str(k): int(v) for k, v in vc.items() if str(k) != ""}
        blank = int((s == "").sum())
        if blank:
            out[PLACE] = blank
        return out

    years = pd.to_datetime(df["opened"], errors="coerce").dt.year.dropna()
    cyears = pd.to_datetime(df["closed"], errors="coerce").dt.year.dropna()

    return {
        "area": [a for a in SCOPE_AREAS if a in set(_clean(df["area"]))],
        "notif_type": uniq("notif_type"),
        "rc1": uniq("rc1") + ([PLACE] if (_clean(df["rc1"]) == "").any() else []),
        "rc2": uniq("rc2") + ([PLACE] if (_clean(df["rc2"]) == "").any() else []),
        "project": uniq("project"),
        "opened_min": int(years.min()) if len(years) else None,
        "opened_max": int(years.max()) if len(years) else None,
        "closed_min": int(cyears.min()) if len(cyears) else None,
        "closed_max": int(cyears.max()) if len(cyears) else None,
        "counts": {
            "area": counts("area"),
            "notif_type": counts("notif_type"),
            "rc1": counts("rc1"),
            "rc2": counts("rc2"),
            "project": counts("project"),
        },
    }


# ---------- KPIs ----------

def _kpis(d: pd.DataFrame) -> dict:
    total = len(d)
    open_n = int((d["status"] == "Open").sum())
    closed_n = int((d["status"] == "Closed").sum())
    copq = float(_money(d).sum())
    lead = pd.to_numeric(d["leadtime"], errors="coerce").dropna()
    med_lead = int(lead.median()) if len(lead) else 0
    rework = int((d["disposition"] == "Rework").sum())
    scrap = int((d["disposition"] == "Scrap").sum())
    unclass = int(d["defect_class"].map(
        lambda v: parse.class_label(v) == "Unclassified").sum())

    oldest = 0
    if open_n:
        opened = pd.to_datetime(d.loc[d["status"] == "Open", "opened"],
                                errors="coerce").dropna()
        if len(opened):
            oldest = int((pd.Timestamp.now() - opened.min()).days)

    return {
        "total": total,
        "open": open_n,
        "closed": closed_n,
        "copq": round(copq),
        "med_lead": med_lead,
        "rework_pct": round(100 * rework / total) if total else 0,
        "scrap": scrap,
        "unclass_pct": round(100 * unclass / total) if total else 0,
        "oldest_open": oldest,
    }


# ---------- a simple count-by helper ----------

def _count_by(d: pd.DataFrame, col: str, order=None, placeholder="(blank)"):
    s = _clean(d[col]).replace({"": placeholder})
    vc = s.value_counts()
    labels = [o for o in order if o in vc.index] if order else list(vc.index)
    return {"labels": labels, "values": [int(vc[l]) for l in labels]}


# ---------- 2 · Pareto with drill-down ----------

def _pareto(d: pd.DataFrame, col: str, n: int = TOP_N) -> dict:
    """Bars = NC count (descending), plus the CoPQ of each bar and the
    cumulative share of NCs. 'Other' always sits last."""
    if d.empty:
        return {"labels": [], "counts": [], "copq": [], "cum_pct": []}
    key = _labelled(d, col)
    counts = _top(key.value_counts().astype(float), n)
    money = d.assign(_k=key, _v=_money(d)).groupby("_k")["_v"].sum()

    labels = list(counts.index)
    if "Other" in labels:
        labels = [l for l in labels if l != "Other"] + ["Other"]
    named = [l for l in labels if l != "Other"]

    vals, cash = [], []
    for l in labels:
        vals.append(int(counts[l]))
        if l == "Other":
            cash.append(round(float(money.drop(index=named, errors="ignore").sum())))
        else:
            cash.append(round(float(money.get(l, 0.0))))

    total = sum(vals) or 1
    run, cum = 0, []
    for v in vals:
        run += v
        cum.append(round(100 * run / total, 1))
    return {"labels": labels, "counts": vals, "copq": cash, "cum_pct": cum}


def _drilled(d: pd.DataFrame, keys: list, path: list) -> pd.DataFrame:
    for col, val in zip(keys, path):
        d = d[_labelled(d, col) == val]
    return d


def _pareto_block(d: pd.DataFrame, keys: list, path) -> dict:
    """One Pareto at the level the drill path points to."""
    path = [p for p in (path or []) if p][:len(keys) - 1]
    level = len(path)
    col = keys[level]
    scoped = _drilled(d, keys, path)
    p = _pareto(scoped, col)
    p.update({
        "col": col,
        "level": level,
        "path": path,
        "last": level == len(keys) - 1,
        "scope": {keys[i]: path[i] for i in range(len(path))},
        "in_scope": len(scoped),
    })
    return p


# ---------- 3 · Cost ----------

def _copq_by(d: pd.DataFrame, col: str, n: int = TOP_N) -> dict:
    if d.empty:
        return {"labels": [], "values": [], "counts": []}
    key = _labelled(d, col)
    money = _top(d.assign(_k=key, _v=_money(d)).groupby("_k")["_v"].sum(), n)
    labels = list(money.index)
    if "Other" in labels:
        labels = [l for l in labels if l != "Other"] + ["Other"]
    cnt = key.value_counts()
    return {
        "labels": labels,
        "values": [round(float(money[l])) for l in labels],
        "counts": [int(cnt.get(l, 0)) for l in labels],
    }


def _copq_pareto(d: pd.DataFrame, col: str = "rc1", n: int = TOP_N) -> dict:
    """Bars = CoPQ, line = cumulative share of CoPQ."""
    b = _copq_by(d, col, n)
    total = sum(b["values"]) or 1
    run, cum = 0, []
    for v in b["values"]:
        run += v
        cum.append(round(100 * run / total, 1))
    b["cum_pct"] = cum
    b["col"] = col
    return b


def _top_costly(d: pd.DataFrame, n: int = 10) -> dict:
    if d.empty:
        return {"cols": [], "rows": []}
    cols = [c for c in ["notification", "title", "project", "rc1",
                        "disposition", "status", "copq"] if c in d.columns]
    t = d.assign(_v=_money(d)).sort_values("_v", ascending=False).head(n)
    rows = []
    for _, r in t.iterrows():
        row = {c: ("" if pd.isna(r[c]) else str(r[c]))
               for c in cols if c != "copq"}
        row["copq"] = round(float(r["_v"]))
        rows.append(row)
    labels = {"notification": "NC", "title": "Title", "project": "Project",
              "rc1": "Root cause L1", "disposition": "Disposition",
              "status": "Status", "copq": "CoPQ (CHF)"}
    return {"cols": cols, "rows": rows, "labels": labels}


# ---------- 4 · Heat maps ----------

def _heat(d: pd.DataFrame, row_col: str, col_col: str,
          nr: int = HEAT_ROWS, nc: int = HEAT_COLS) -> dict:
    """NC counts, rows = the busiest row_col values, columns likewise."""
    if d.empty:
        return {"x": [], "y": [], "z": []}
    r = _labelled(d, row_col)
    c = _labelled(d, col_col)
    top_r = list(r.value_counts().head(nr).index)
    top_c = list(c.value_counts().head(nc).index)
    ct = pd.crosstab(r, c).reindex(index=top_r, columns=top_c, fill_value=0)
    return {
        "y": [str(v) for v in ct.index],
        "x": [str(v) for v in ct.columns],
        "z": [[int(v) for v in row] for row in ct.values],
        "row_col": row_col, "col_col": col_col,
    }


def _heat_month(d: pd.DataFrame, col: str, months: int = 18,
                nr: int = HEAT_ROWS) -> dict:
    """Rows = category, columns = the last N months by opened date."""
    if d.empty:
        return {"x": [], "y": [], "z": []}
    m = pd.to_datetime(d["opened"], errors="coerce").dt.to_period("M").astype(str)
    keep = m != "NaT"
    if not keep.any():
        return {"x": [], "y": [], "z": []}
    dd = d[keep]
    cat = _labelled(dd, col)
    mm = m[keep]
    cols = sorted(mm.unique())[-months:]
    top_r = list(cat.value_counts().head(nr).index)
    ct = pd.crosstab(cat, mm).reindex(index=top_r, columns=cols, fill_value=0)
    return {
        "y": [str(v) for v in ct.index],
        "x": [str(v) for v in ct.columns],
        "z": [[int(v) for v in row] for row in ct.values],
        "row_col": col, "col_col": "__month_opened",
    }


# ---------- 5 · Time ----------

def _monthly(d: pd.DataFrame) -> dict:
    """Opened vs closed per month, over the union of both date columns."""
    if d.empty:
        return {"months": [], "opened": [], "closed": []}
    op = pd.to_datetime(d["opened"], errors="coerce").dt.to_period("M")
    cl = pd.to_datetime(d["closed"], errors="coerce").dt.to_period("M")
    months = sorted({str(p) for p in op.dropna()} | {str(p) for p in cl.dropna()})
    o = op.dropna().astype(str).value_counts()
    c = cl.dropna().astype(str).value_counts()
    return {
        "months": months,
        "opened": [int(o.get(m, 0)) for m in months],
        "closed": [int(c.get(m, 0)) for m in months],
    }


def _backlog(monthly: dict) -> dict:
    """Running open count: everything opened so far minus everything closed."""
    run, out = 0, []
    for o, c in zip(monthly["opened"], monthly["closed"]):
        run += o - c
        out.append(run)
    return {"months": monthly["months"], "values": out}


def _leadtime_hist(d: pd.DataFrame,
                   bins=(0, 7, 14, 30, 60, 90, 180, 365)) -> dict:
    lead = pd.to_numeric(d["leadtime"], errors="coerce").dropna()
    edges = list(bins)
    labels, values, ranges = [], [], []
    for i in range(len(edges) - 1):
        lo, hi = edges[i], edges[i + 1]
        labels.append(f"{lo}-{hi}d")
        values.append(int(((lead >= lo) & (lead < hi)).sum()))
        ranges.append([lo, hi])
    labels.append(f"{edges[-1]}d+")
    values.append(int((lead >= edges[-1]).sum()))
    ranges.append([edges[-1], None])
    return {"labels": labels, "values": values, "ranges": ranges,
            "median": int(lead.median()) if len(lead) else 0}


# ---------- 6 · Supplier vs production ----------

def _side(d: pd.DataFrame) -> pd.Series:
    """Z2 = supplier (procurement complaint), Z3 = internal production."""
    t = _clean(d["notif_type"])
    return pd.Series(
        np.where(t.str.startswith("Z2"), "Supplier",
                 np.where(t.str.startswith("Z3"), "Production", "Other")),
        index=d.index)


def _supplier_monthly(d: pd.DataFrame) -> dict:
    if d.empty:
        return {"months": [], "supplier": [], "production": []}
    op = pd.to_datetime(d["opened"], errors="coerce").dt.to_period("M").astype(str)
    dd = d.assign(_m=op, _s=_side(d))
    dd = dd[dd["_m"] != "NaT"]
    if dd.empty:
        return {"months": [], "supplier": [], "production": []}
    months = sorted(dd["_m"].unique())
    piv = pd.crosstab(dd["_m"], dd["_s"]).reindex(index=months, fill_value=0)

    def grab(name):
        return [int(piv[name][m]) if name in piv.columns else 0 for m in months]

    return {"months": months, "supplier": grab("Supplier"),
            "production": grab("Production")}


def _supplier_compare(d: pd.DataFrame) -> dict:
    """Side-by-side numbers for supplier vs production."""
    out = {"sides": ["Supplier", "Production"], "count": [], "copq": [],
           "med_lead": [], "open": []}
    if d.empty:
        for k in ("count", "copq", "med_lead", "open"):
            out[k] = [0, 0]
        return out
    dd = d.assign(_s=_side(d), _v=_money(d))
    for s in out["sides"]:
        part = dd[dd["_s"] == s]
        lead = pd.to_numeric(part["leadtime"], errors="coerce").dropna()
        out["count"].append(int(len(part)))
        out["copq"].append(round(float(part["_v"].sum())))
        out["med_lead"].append(int(lead.median()) if len(lead) else 0)
        out["open"].append(int((part["status"] == "Open").sum()))
    return out


# ---------- payload ----------

def build_payload(df: pd.DataFrame, filters: dict | None = None) -> dict:
    filters = filters or {}
    in_scope = scope(df)
    d = _apply_filters(df, filters)
    drill = filters.get("_drill") or {}

    monthly = _monthly(d)

    p = {
        "all": len(in_scope),          # the I/M population, not the raw import
        "filtered": len(d),
        "options": _options(df),
        "kpi": _kpis(d),
    }

    # 1 · Status
    p["nc_status"] = {
        **_count_by(d, "status", order=["Closed", "Open", "Deleted"]),
        "click": "status",
    }

    # 2 · Root cause & detection (drill-down Paretos)
    p["pareto_rc"] = _pareto_block(d, RC_KEYS, drill.get("rc"))
    p["pareto_det"] = _pareto_block(d, DET_KEYS, drill.get("det"))

    # 3 · Cost
    p["copq_disposition"] = _copq_by(d, "disposition")
    p["copq_project"] = _copq_by(d, "project")
    p["copq_pareto"] = _copq_pareto(d, "rc1")
    p["top_costly"] = _top_costly(d)

    # 4 · Heat maps
    p["heat_rc_det"] = _heat(d, "rc1", "origin1")
    p["heat_proj_rc"] = _heat(d, "project", "rc1")
    p["heat_month_rc"] = _heat_month(d, "rc1")

    # 5 · Time
    p["trend_monthly"] = monthly
    p["backlog"] = _backlog(monthly)
    p["leadtime_hist"] = _leadtime_hist(d)

    # 6 · Supplier vs production
    p["supplier_monthly"] = _supplier_monthly(d)
    p["supplier_compare"] = _supplier_compare(d)

    return p