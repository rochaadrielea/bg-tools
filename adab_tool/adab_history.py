#!/usr/bin/env python3
"""
ADAB history — the comparison memory behind the Research page.

Every ADAB run is recorded so any material can later be looked up across ALL
past comparisons: when it was checked, against which documents, the result, the
batch/revision, and a link back to the exact As-Design / As-Built / report files
from that run.

Design goals (Adriele):
  - ADDITIVE. Nothing here changes how ADAB compares. It only READS the report
    the engine already writes (the "Parts Reconciliation" + "Matched" tabs) and
    stores the rows. If recording ever fails, the comparison itself is untouched.
  - Nothing is deleted — the database only grows.
  - Search by material, batch/charge, document/project, or serial.

Storage (default, override in record_run):
  <base>/adab_history.db          SQLite: runs + run_materials
  <base>/history/<run_id>/        a copy of the As-Design, As-Built, report files

Public API:
  init_db(db_path)
  record_run(design_paths, built_path, report_path, built_label, run_label, base_dir=...)
  search(db_path, text=None, field="material")     -> list of rows (dict)
  material_summary(db_path, material)               -> dict
"""
import os
import re
import shutil
import sqlite3
import datetime
import openpyxl


# --------------------------------------------------------------------------- #
#  paths / db
# --------------------------------------------------------------------------- #
def default_base():
    """Where the history lives. Next to this file by default; overridable."""
    return os.environ.get("ADAB_HISTORY_DIR",
                          os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       "adab_data"))


def db_path_for(base_dir=None):
    base = base_dir or default_base()
    return os.path.join(base, "adab_history.db")


def init_db(db_path):
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    con = sqlite3.connect(db_path)
    con.executescript("""
    CREATE TABLE IF NOT EXISTS runs (
        run_id       TEXT PRIMARY KEY,
        ts           TEXT,
        run_label    TEXT,
        design_name  TEXT,
        built_name   TEXT,
        built_label  TEXT,
        n_materials  INTEGER,
        n_matched    INTEGER,
        n_missing    INTEGER,
        n_extra      INTEGER,
        design_file  TEXT,
        built_file   TEXT,
        report_file  TEXT
    );
    CREATE TABLE IF NOT EXISTS run_materials (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        run_id       TEXT,
        material     TEXT,     -- normalised key (searchable)
        material_raw TEXT,     -- as printed
        rev_name     TEXT,     -- SAP short text (design Revision Name)
        description  TEXT,
        status       TEXT,     -- Matched / In Design, not built / In Built, not design
        revision     TEXT,     -- built revision
        batch        TEXT,     -- Charge / Batch  (normalised for search in batch_key)
        batch_key    TEXT,
        serial       TEXT,
        serial_key   TEXT,
        design_qty   REAL,
        built_qty    REAL
    );
    CREATE INDEX IF NOT EXISTS ix_mat  ON run_materials(material);
    CREATE INDEX IF NOT EXISTS ix_bat  ON run_materials(batch_key);
    CREATE INDEX IF NOT EXISTS ix_ser  ON run_materials(serial_key);
    CREATE INDEX IF NOT EXISTS ix_run  ON run_materials(run_id);
    """)
    con.commit()
    con.close()


# --------------------------------------------------------------------------- #
#  helpers
# --------------------------------------------------------------------------- #
def _norm_key(v):
    """Same spirit as the engine: case/space-invariant, drop a trailing '.0',
    but keep suffixes (C3529115-C != C3529115)."""
    if v is None:
        return ""
    s = str(v).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return re.sub(r"\s+", " ", s).upper()


def _to_float(v):
    try:
        return float(str(v).strip().replace(",", "."))
    except (ValueError, AttributeError, TypeError):
        return None


def _sheet_by_name(wb, *wanted):
    low = {s.title.strip().lower(): s for s in wb.worksheets}
    for w in wanted:
        if w.lower() in low:
            return low[w.lower()]
    # loose contains-match (e.g. "Matched" vs "Matched (As-Design)")
    for name, ws in low.items():
        if any(w.lower() in name for w in wanted):
            return ws
    return None


def _rows(ws):
    return list(ws.iter_rows(values_only=True)) if ws else []


def _header_row(rows, tokens, scan=6):
    """Index of the header row — the row (within `scan`) whose cells match the
    most of `tokens` (a set of expected header names, lower-case). This beats a
    plain 'first row with Status', because the Matched tab has 'Status' in BOTH
    the colour-band row and the real field-header row below it."""
    best_i, best = 0, -1
    for i, r in enumerate(rows[:scan]):
        score = sum(1 for c in r
                    if c is not None and str(c).strip().lower() in tokens)
        if score > best:
            best, best_i = score, i
    return best_i


_MATCHED_TOKENS = {"status", "revision name", "description (design)",
                   "material (built)", "material", "revision (built)",
                   "charge / batch", "charge/batch", "serial", "qty",
                   "quantity", "lot", "equipment"}
_RECON_TOKENS = {"material", "part number", "description", "status", "in design",
                 "in built", "design qty", "built qty", "design positions",
                 "built copies", "fingerprint"}


def _col_index(header, *predicates):
    """First column index whose header matches any predicate (a str = exact-ish
    contains, or a callable(header_lower)->bool)."""
    hl = [("" if h is None else str(h)).strip().lower() for h in header]
    for pred in predicates:
        for i, h in enumerate(hl):
            if callable(pred):
                if pred(h):
                    return i
            elif pred.lower() == h:
                return i
    for pred in predicates:                      # then loose contains
        if isinstance(pred, str):
            for i, h in enumerate(hl):
                if pred.lower() in h:
                    return i
    return None


# --------------------------------------------------------------------------- #
#  report parser  (reads what the engine already writes)
# --------------------------------------------------------------------------- #
def parse_report(report_path):
    """Return (materials, counts) parsed from an ADAB report workbook.

    materials: list of dicts with material/description/status/revision/batch/
               serial/rev_name/design_qty/built_qty.
    Reads 'Parts Reconciliation' for the per-distinct-part status + quantities,
    then enriches with revision/batch/serial from the 'Matched' tab (joined on
    the material number). Both tabs are produced by adab_batch_compare.py.
    """
    wb = openpyxl.load_workbook(report_path, read_only=True, data_only=True)

    # ---- Matched tab: material -> {revision, batch, serial, rev_name} -------
    enrich = {}
    mt = _sheet_by_name(wb, "Matched")
    rows = _rows(mt)
    if rows:
        hi = _header_row(rows, _MATCHED_TOKENS)
        hdr = rows[hi]
        # built-side material = the As-Built "Material" column (design side is
        # shown as "Material (Design)"); prefer an exact "material" that is NOT
        # the design one.
        c_mat = _col_index(hdr, lambda h: h == "material",
                           lambda h: h.startswith("material") and "design" not in h,
                           "part number")
        c_bat = _col_index(hdr, "charge / batch", "charge/batch", "batch",
                           "lot number", "lot")
        c_ser = _col_index(hdr, lambda h: h == "serial", "serial number", "serial")
        c_rev = _col_index(hdr, "revision (built)",
                           lambda h: h.startswith("revision") and "design" not in h
                           and "name" not in h)
        c_rn = _col_index(hdr, "revision name")
        for r in rows[hi + 1:]:
            def g(i):
                return r[i] if (i is not None and i < len(r)) else None
            mat = _norm_key(g(c_mat))
            if not mat:
                continue
            enrich[mat] = {
                "revision": g(c_rev), "batch": g(c_bat),
                "serial": g(c_ser), "rev_name": g(c_rn),
            }

    # ---- Parts Reconciliation: the per-distinct-part status + quantities ----
    materials = []
    counts = {"matched": 0, "missing": 0, "extra": 0}
    rt = _sheet_by_name(wb, "Parts Reconciliation", "Reconciliation")
    rows = _rows(rt)
    if rows:
        hi = _header_row(rows, _RECON_TOKENS)
        hdr = rows[hi]
        c_key = _col_index(hdr, "material", "part number")
        c_dsc = _col_index(hdr, "description")
        c_st = _col_index(hdr, "status")
        c_dq = _col_index(hdr, "design qty")
        c_bq = _col_index(hdr, "built qty")
        for r in rows[hi + 1:]:
            def g(i):
                return r[i] if (i is not None and i < len(r)) else None
            raw = g(c_key)
            mat = _norm_key(raw)
            if not mat:
                continue
            st_raw = str(g(c_st) or "").strip().upper()
            if st_raw.startswith("MISSING"):
                status = "In Design, not built"; counts["missing"] += 1
            elif st_raw.startswith("EXTRA"):
                status = "In Built, not design"; counts["extra"] += 1
            else:                                # MATCHED / SHORT / OVER
                status = "Matched"; counts["matched"] += 1
                if st_raw in ("SHORT", "OVER"):
                    status = f"Matched ({st_raw.lower()} qty)"
            ex = enrich.get(mat, {})
            materials.append({
                "material": mat, "material_raw": str(raw).strip(),
                "description": g(c_dsc), "status": status,
                "revision": ex.get("revision"), "batch": ex.get("batch"),
                "serial": ex.get("serial"), "rev_name": ex.get("rev_name"),
                "design_qty": _to_float(g(c_dq)), "built_qty": _to_float(g(c_bq)),
            })
    wb.close()
    return materials, counts


# --------------------------------------------------------------------------- #
#  record a run
# --------------------------------------------------------------------------- #
def _run_id(ts, label):
    stamp = ts.strftime("%Y%m%d-%H%M%S")
    slug = re.sub(r"[^A-Za-z0-9]+", "-", (label or "run")).strip("-")[:40] or "run"
    return f"{stamp}_{slug}"


def record_run(design_paths, built_path, report_path, built_label="As Built",
               run_label="", base_dir=None, ts=None, copy_files=True):
    """Record ONE comparison into the history DB (and copy its 3 files).

    design_paths : the As-Design file (str) or list of files.
    built_path   : the As-Built source file (str).
    report_path  : the report workbook the engine just wrote.
    run_label    : the short nickname the user typed on the Build page.
    Returns the run_id (or None if the report couldn't be parsed).
    """
    base = base_dir or default_base()
    db = db_path_for(base)
    init_db(db)
    ts = ts or datetime.datetime.now()

    design_list = design_paths if isinstance(design_paths, (list, tuple)) else [design_paths]
    design_name = (os.path.basename(design_list[0]) if len(design_list) == 1
                   else f"{len(design_list)} design files (merged)")
    built_name = os.path.basename(built_path) if built_path else "(as-built)"

    # Build a run_id from timestamp + nickname; when no nickname is given fall
    # back to the As-Built file name so two reports written in the SAME second
    # (e.g. one design vs several built files) don't collide. Then guarantee it
    # is unique in the DB (append -2, -3, ...) so a run is never overwritten.
    slug_src = run_label or os.path.splitext(built_name)[0]
    run_id = _run_id(ts, slug_src)
    try:
        con0 = sqlite3.connect(db)
        existing = {row[0] for row in
                    con0.execute("SELECT run_id FROM runs WHERE run_id LIKE ?",
                                 (run_id + "%",))}
        con0.close()
        if run_id in existing:
            k = 2
            while f"{run_id}-{k}" in existing:
                k += 1
            run_id = f"{run_id}-{k}"
    except Exception:
        pass

    try:
        materials, counts = parse_report(report_path)
    except Exception as e:                       # never let recording break a run
        print(f"[adab_history] could not parse report ({e}); run not recorded.")
        return None

    # copy the three artifacts into history/<run_id>/
    d_store = b_store = r_store = None
    if copy_files:
        run_dir = os.path.join(base, "history", run_id)
        os.makedirs(run_dir, exist_ok=True)

        def _copy(src, tag):
            if not src or not os.path.exists(src):
                return None
            dst = os.path.join(run_dir, tag + "_" + os.path.basename(src))
            try:
                shutil.copy2(src, dst)
                return dst
            except Exception:
                return None
        d_store = _copy(design_list[0], "as_design")
        b_store = _copy(built_path, "as_built")
        r_store = _copy(report_path, "report")

    con = sqlite3.connect(db)
    con.execute(
        "INSERT OR REPLACE INTO runs (run_id, ts, run_label, design_name, "
        "built_name, built_label, n_materials, n_matched, n_missing, n_extra, "
        "design_file, built_file, report_file) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
        (run_id, ts.isoformat(timespec="seconds"), run_label or built_name,
         design_name, built_name, built_label, len(materials),
         counts["matched"], counts["missing"], counts["extra"],
         d_store, b_store, r_store))
    con.executemany(
        "INSERT INTO run_materials (run_id, material, material_raw, rev_name, "
        "description, status, revision, batch, batch_key, serial, serial_key, "
        "design_qty, built_qty) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
        [(run_id, m["material"], m["material_raw"], m["rev_name"], m["description"],
          m["status"], m["revision"],
          (str(m["batch"]).strip() if m["batch"] not in (None, "") else None),
          _norm_key(m["batch"]),
          (str(m["serial"]).strip() if m["serial"] not in (None, "") else None),
          _norm_key(m["serial"]),
          m["design_qty"], m["built_qty"]) for m in materials])
    con.commit()
    con.close()
    return run_id


# --------------------------------------------------------------------------- #
#  search
# --------------------------------------------------------------------------- #
_FIELD_SQL = {
    "material": "m.material LIKE ?",
    "batch":    "m.batch_key LIKE ?",
    "serial":   "m.serial_key LIKE ?",
    "document": "(r.run_label LIKE ? OR r.design_name LIKE ? OR r.built_name LIKE ?)",
}


def search(db_path, text, field="material"):
    """Return matching history rows (material joined to its run), newest first."""
    if not os.path.exists(db_path):
        return []
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    q = f"%{_norm_key(text) if field in ('material','batch','serial') else text.strip()}%"
    where = _FIELD_SQL.get(field, _FIELD_SQL["material"])
    params = [q, q, q] if field == "document" else [q]
    sql = f"""
        SELECT m.*, r.ts, r.run_label, r.design_name, r.built_name, r.built_label,
               r.design_file, r.built_file, r.report_file
        FROM run_materials m JOIN runs r ON r.run_id = m.run_id
        WHERE {where}
        ORDER BY r.ts DESC, m.material
    """
    out = [dict(row) for row in con.execute(sql, params)]
    con.close()
    return out


def material_summary(db_path, material):
    """Headline stats for one material across all runs."""
    rows = search(db_path, material, "material")
    if not rows:
        return {"material": material, "times": 0}
    docs = {r["run_label"] or r["built_name"] for r in rows}
    ts = sorted(r["ts"] for r in rows if r["ts"])
    return {
        "material": rows[0]["material_raw"] or material,
        "description": next((r["description"] for r in rows if r["description"]), ""),
        "rev_name": next((r["rev_name"] for r in rows if r["rev_name"]), ""),
        "times": len(rows),
        "documents": len(docs),
        "first": ts[0][:10] if ts else "",
        "last": ts[-1][:10] if ts else "",
        "latest_status": rows[0]["status"],
        "rows": rows,
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) >= 3 and sys.argv[1] == "search":
        for r in search(db_path_for(), sys.argv[2],
                        sys.argv[3] if len(sys.argv) > 3 else "material"):
            print(r["ts"][:10], r["material_raw"], "|", r["status"], "|",
                  r["run_label"], "| batch", r["batch"])
    else:
        print("usage: adab_history.py search <text> [material|batch|serial|document]")