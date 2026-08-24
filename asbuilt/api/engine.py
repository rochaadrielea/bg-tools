"""The comparison engine.

Pure functions: no database, no FastAPI. Everything here is testable on lists
of dicts alone (see test_engine.py).

The rules, in the order they decide:
    material   joins       — nothing matches without it
    quantity   reconciles  — per material, per unit
    revision   judges      — same material, other revision is a deviation
    batch      witnesses   — compared only where it exists on BOTH sides
    Traceable  gates       — a missing batch is a flag only for traceable parts

A row is never dropped for a batch or revision mismatch, and a material present
on one side only still occupies a full row.
"""
from __future__ import annotations

import datetime as dt
import re
import unicodedata

VERDICTS = ["OK", "MISSING", "SHORT", "OVER", "REVISION", "BATCH", "EXTRA"]
VERDICT_LABEL = {
    "OK": "Matched", "MISSING": "Missing", "SHORT": "Short qty", "OVER": "Over qty",
    "REVISION": "Revision differs", "BATCH": "Batch not recorded", "EXTRA": "Extra",
}


def norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


def match_key(material) -> str:
    """Type-, case- and space-blind. SUFFIX PRESERVING on purpose:
    C3529115-C is a different part from C3529115."""
    s = unicodedata.normalize("NFKC", norm(material))
    if re.fullmatch(r"\d+\.0+", s):          # 7004369.0 === 7004369
        s = s.split(".")[0]
    return re.sub(r"\s+", "", s).upper()




def trace_key(v) -> str:
    """Batch/serial normalisation. Strip separators and uppercase, keep the
    suffix. Used for tracing a specific batch or serial across projects."""
    import re, unicodedata as U
    s = U.normalize("NFKC", norm(v))
    s = re.sub(r"[\s\-.\/_]", "", s)   # separators the humans put in
    return s.upper()


def to_num(v) -> float:
    s = norm(v).replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def group(rows: list) -> dict:
    """Collapse to one entry per material, keeping every witness.

    Quantity is summed across positions — the same material sits at several
    places in the tree, and 'short by 2' must not be confused with 'used twice'.
    The position count is kept so the tree is still visible.
    """
    out: dict = {}
    for r in rows:
        k = match_key(r.get("material"))
        if not k:
            continue
        g = out.setdefault(k, {"material": norm(r.get("material")),
                               "description": "", "revisions": [], "qty": 0.0,
                               "batches": [], "serials": [], "traceable": False,
                               "positions": 0})
        g["qty"] += to_num(r.get("qty"))
        g["positions"] += 1
        rev = norm(r.get("revision"))
        if rev and rev not in g["revisions"]:
            g["revisions"].append(rev)
        b = norm(r.get("batch"))
        if b and b not in g["batches"]:
            g["batches"].append(b)
        s = norm(r.get("serial"))
        if s and s not in g["serials"]:
            g["serials"].append(s)
        if r.get("traceable"):
            g["traceable"] = True
        if not g["description"]:
            g["description"] = norm(r.get("description"))
    return out


def judge(a: dict | None, b: dict | None) -> tuple:
    """One pair -> (verdict, detail). `a` expected, `b` present."""
    if a is None:
        return "EXTRA", "present, not expected"
    if b is None:
        return "MISSING", "expected, never appeared"

    if b["qty"] < a["qty"]:
        return "SHORT", f"short by {fmt_qty(a['qty'] - b['qty'])}"

    ra, rb = "/".join(a["revisions"]), "/".join(b["revisions"])
    if ra and rb and ra != rb:
        return "REVISION", f"{ra} → {rb}"

    # batch is compared only where it exists on both sides; missing is a fact,
    # and only a flag when the part is traceable
    if (a["traceable"] or b["traceable"]) and not b["batches"]:
        return "BATCH", "traceable part, no batch recorded"

    if b["qty"] > a["qty"]:
        return "OVER", f"over by {fmt_qty(b['qty'] - a['qty'])}"
    return "OK", ""


def fmt_qty(v: float) -> str:
    return str(int(v)) if float(v).is_integer() else f"{v:g}"


def compare(expected: list, present: list) -> dict:
    """Side by side, never merged. Returns rows plus a count per verdict."""
    A, B = group(expected), group(present)
    keys = sorted(set(A) | set(B))
    rows, counts = [], {v: 0 for v in VERDICTS}
    for k in keys:
        a, b = A.get(k), B.get(k)
        verdict, detail = judge(a, b)
        counts[verdict] += 1
        rows.append({"key": k, "a": a, "b": b, "verdict": verdict, "detail": detail})
    counts["TOTAL"] = len(rows)
    counts["ATTENTION"] = sum(counts[v] for v in VERDICTS if v != "OK")
    return {"rows": rows, "counts": counts}


# --------------------------------------------------------------------------
# reading an uploaded sheet
# --------------------------------------------------------------------------
ALIASES = {
    "material": ["material", "material number", "materialnummer", "part number",
                 "part no", "id", "item number", "sach-nr", "sachnummer", "artikel"],
    "revision": ["revision", "rev", "revisionsstand", "index", "revision level",
                 "rev.", "revstand"],
    "description": ["description", "designation", "bezeichnung", "materialkurztext",
                    "revision name", "short text", "text"],
    "batch": ["batch", "charge", "lot", "lot number", "chargennummer"],
    "serial": ["serial", "serialnummer", "serial number", "seriennummer"],
    "qty": ["qty", "quantity", "menge", "anzahl", "quantity required", "req qty"],
    "work_order": ["work order", "wo", "auftrag", "order", "auftragsnummer"],
    "traceable": ["traceable", "lot", "lot managed", "chargenpflichtig"],
    "parent": ["parent", "parent material", "higher level", "baugruppe"],
    "position": ["position", "pos", "item", "posnr"],
}
_LOOKUP = {a: k for k, names in ALIASES.items() for a in names}


def header_key(h) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", str(h or ""))).strip().lower()


def map_headers(headers: list) -> dict:
    out: dict = {}
    for i, h in enumerate(headers):
        f = _LOOKUP.get(header_key(h))
        if f and f not in out.values():
            out[i] = f
    return out


TRUEISH = {"yes", "y", "true", "1", "x", "ja", "traceable"}


def read_rows(sheet_rows: list) -> tuple:
    """(rows, recognised_fields). Picks the header row that maps the most
    columns — a fed-back report has a note banner above its real headers."""
    best = (0, 0, {})
    for i, row in enumerate(sheet_rows[:15]):
        m = map_headers([norm(c) for c in row])
        if len(m) > best[0]:
            best = (len(m), i, m)
    score, head_i, cmap = best
    if score < 2 or "material" not in cmap.values():
        found = [norm(c) for c in (sheet_rows[0] if sheet_rows else []) if norm(c)]
        raise ValueError("no material column recognised. Columns found: "
                         + (", ".join(found[:25]) or "none"))
    rows = []
    for r in sheet_rows[head_i + 1:]:
        if not any(norm(c) for c in r):
            continue
        d = {f: (r[i] if i < len(r) else "") for i, f in cmap.items()}
        row = {k: norm(v) for k, v in d.items()}
        if not row.get("material"):
            continue
        row["qty"] = to_num(d.get("qty", 0)) or 1
        row["traceable"] = norm(d.get("traceable", "")).lower() in TRUEISH
        rows.append(row)
    return rows, sorted(set(cmap.values()))