"""
validate.py
-----------
Checks the source Excel files BEFORE ingest writes anything, and keeps a
persistent log of every problem it finds.

Why this exists
---------------
Owners noticed their NC counts were wrong before the pipeline did. The tracker
has no naming standard (the same person appears as 'Vitaly' and 'Vitaly Meshin'),
some rows carry no status, some carry no ID, and nothing in the pipeline said so.
This module says so.

Warning log behaviour (by Adriele's spec)
-----------------------------------------
  - Warnings live in their OWN database (warnings.db), not quality.db, because
    quality.db is dropped and rebuilt on every ingest and would take the log
    with it.
  - A warning is identified by (check, subject) — NOT by its message text. So
    'Elisa Martin: 20 in tracker, 18 in db' and '...19 in db' are the SAME
    warning, and the row is not duplicated when the number moves.
  - First occurrence is written with a timestamp and never overwritten. The
    timestamp answers 'when did this first appear', which is the useful question.
  - Repeat occurrences update last_seen and times_seen only. The original
    first_seen and its message are preserved verbatim.
  - Nothing is ever deleted. A warning that stops occurring stays in the log
    with its last_seen frozen, so you can see it was fixed and when.

This module WARNS. It never blocks. ingest.py runs regardless.

Usage:
    python validate.py              # check and print
    from validate import run_checks # or call from ingest.py
"""

import sqlite3
from datetime import datetime
from pathlib import Path

import pandas as pd

WARN_DB = "warnings.db"

# The column lists come from ingest.py itself — NOT a copy kept here. A local
# copy silently drifts: an earlier version of this file had its own list and
# reported 'NC Type', 'Failure' and 'Material' as unrecognised new columns while
# load_tracker was reading all three. Importing the real map makes that
# impossible by construction.
try:
    from ingest import TRACKER_COL_MAP, TRACKER_REQUIRED_COLS
    KNOWN_TRACKER_COLS = set(TRACKER_COL_MAP.keys())
    EXPECTED_TRACKER_COLS = list(TRACKER_REQUIRED_COLS)
except ImportError:
    KNOWN_TRACKER_COLS = set()
    EXPECTED_TRACKER_COLS = ["System", "ID-Blackout", "Issue Owner (QA/PA)",
                             "Created On", "Status"]

EXPECTED_SAP_COLS = [
    "Notification", "Notification Type", "Notification Date", "Closing date",
    "Notif. Status", "Project Text  (Notification)", "Defect Class",
    "Disposition Action TEXT", "Notification Cause C TEXT", "Material Key",
    "Batch", "[Vendor NC]", "WBS Element Id (CoPQ)",
]


# ══════════════════════════════════════════════════════════════════════════════
# The log
# ══════════════════════════════════════════════════════════════════════════════
def _connect():
    conn = sqlite3.connect(WARN_DB)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS warnings (
            check_name  TEXT NOT NULL,
            subject     TEXT NOT NULL,
            owner       TEXT,
            severity    TEXT,
            message     TEXT,
            first_seen  TEXT NOT NULL,
            last_seen   TEXT NOT NULL,
            times_seen  INTEGER NOT NULL DEFAULT 1,
            PRIMARY KEY (check_name, subject)
        )
    """)
    # Existing logs predate the owner column — add it rather than rebuild, so
    # first_seen history survives the upgrade.
    cols = [r[1] for r in conn.execute("PRAGMA table_info(warnings)")]
    if "owner" not in cols:
        conn.execute("ALTER TABLE warnings ADD COLUMN owner TEXT")
    conn.commit()
    return conn


def log_warning(conn, check_name, subject, message, severity="warn", owner=None):
    """Record a warning. First time: insert with a timestamp. After that: only
    bump last_seen/times_seen — first_seen and the original message are never
    overwritten, so the log keeps the moment the problem appeared.

    `owner` is stored as its own column rather than left buried in the message
    text, so warnings can be handed to the person responsible without parsing
    prose or joining back to nc. It is NULL for pipeline-level problems (a new
    column, an unreadable file) which belong to nobody's NC list.
    """
    now = datetime.now().isoformat(timespec="seconds")
    cur = conn.cursor()
    cur.execute("SELECT times_seen FROM warnings WHERE check_name=? AND subject=?",
                [check_name, str(subject)])
    row = cur.fetchone()
    if row is None:
        cur.execute(
            "INSERT INTO warnings (check_name, subject, owner, severity, message, "
            "first_seen, last_seen, times_seen) VALUES (?,?,?,?,?,?,?,1)",
            [check_name, str(subject), owner, severity, message, now, now])
        return True          # newly seen
    # owner is refreshed on repeat: if a row's owner is corrected in the tracker,
    # the warning should follow them. first_seen and message stay frozen.
    cur.execute("UPDATE warnings SET last_seen=?, times_seen=times_seen+1, owner=? "
                "WHERE check_name=? AND subject=?", [now, owner, check_name, str(subject)])
    return False             # already known


def view_warnings(only_new_since=None):
    """Return the whole log as a DataFrame, newest first."""
    with sqlite3.connect(WARN_DB) as conn:
        try:
            df = pd.read_sql("SELECT * FROM warnings ORDER BY first_seen DESC", conn)
        except Exception:
            return pd.DataFrame()
    if only_new_since is not None and not df.empty:
        df = df[df["first_seen"] >= only_new_since]
    return df


# ══════════════════════════════════════════════════════════════════════════════
# Helpers shared with ingest
# ══════════════════════════════════════════════════════════════════════════════
def _blank(v):
    """True when a cell is empty in any of the ways this tracker manages to be
    empty: NaN, '', a non-breaking space (\\xa0 — Excel produces these), or '-'."""
    if pd.isna(v):
        return True
    return str(v).strip() in ("", "\xa0", "-", "nan")


def _load_tracker_raw(path, sheet="NC_Tracker_Black_Out"):
    df = pd.read_excel(path, sheet_name=sheet)
    df = df[df["System"].notna()].copy()
    df = df[df["System"].astype(str).str.strip() != "\xa0"].copy()
    return df


# Set by run_checks so the row-level checks can canonicalise owner names without
# threading clean_owner through every signature. Warnings must be filed under the
# same spelling the dashboard shows, or 'Vitaly' and 'Vitaly Meshin' get separate
# warning lists — the exact bug this module exists to catch.
_CANON_FN = None


def _canon(name):
    if _CANON_FN is None or name is None:
        return name
    try:
        return _CANON_FN(name) or name
    except Exception:
        return name


# ══════════════════════════════════════════════════════════════════════════════
# Checks
# ══════════════════════════════════════════════════════════════════════════════
def check_schema(conn, df, expected, source_name):
    """Columns the pipeline reads must still exist. A rename here is the single
    most damaging thing that can happen quietly."""
    n = 0
    have = set(str(c).strip() for c in df.columns)
    for col in expected:
        if col not in have:
            if log_warning(conn, "missing_column", f"{source_name}:{col}",
                           f"Expected column '{col}' not found in {source_name}. "
                           f"Anything downstream that reads it will be NULL.",
                           severity="error"):
                n += 1
    return n


def check_new_columns(conn, df, known, source_name):
    """New columns are not a problem, but they are worth knowing about — they
    usually mean someone changed the file and may expect the change to show up.

    `known` must be every column the pipeline maps, not just the required ones,
    or this warns about columns that are being read perfectly well."""
    n = 0
    for col in df.columns:
        c = str(col).strip()
        if c and c not in known and not c.lower().startswith("unnamed"):
            if log_warning(conn, "new_column", f"{source_name}:{c}",
                           f"New column '{c}' appeared in {source_name}. "
                           f"Not read by the pipeline — add a mapping if it matters.",
                           severity="info"):
                n += 1
    return n


def check_blank_status(conn, tr):
    """A blank status is not 'closed'. It used to become is_open=0 and vanish
    from every open-NC view — which is how an owner's count silently shrank."""
    n = 0
    blank = tr[tr["Status"].apply(_blank)]
    for _, r in blank.iterrows():
        # Same uniqueness problem as check_missing_id: several rows can lack
        # both IDs, so fall back to the Excel row number rather than '(no id)'.
        nc, tc = r.get("ID-Blackout"), r.get("TC ID")
        if not _blank(nc):
            subj = str(nc).strip()
        elif not _blank(tc):
            subj = str(tc).strip()
        else:
            subj = f"tracker-row-{int(r.name) + 2}"
        owner = r.get("Issue Owner (QA/PA)")
        owner = None if _blank(owner) else _canon(str(owner).strip())
        if log_warning(conn, "blank_status", subj,
                       f"NC '{subj}' (owner: {owner or '(no owner)'}) has no Status. It is "
                       f"neither open nor closed — shown as '(no status)' in the dashboard.",
                       severity="warn", owner=owner):
            n += 1
    return n


def check_missing_id(conn, tr):
    """Rows with no ID cannot be matched to SAP and cannot be addressed."""
    n = 0
    noid = tr[tr["ID-Blackout"].apply(_blank)]
    for _, r in noid.iterrows():
        tc = r.get("TC ID")
        has_tc = not _blank(tc)
        tc = str(tc).strip() if has_tc else "(no TC ID either)"
        owner = r.get("Issue Owner (QA/PA)")
        owner = None if _blank(owner) else _canon(str(owner).strip())
        # Subject must be unique per row, or several ID-less rows collapse onto
        # one key and the log reports 1 problem where there are several. When
        # there is no TC ID either, fall back to the Excel row number — the only
        # thing left that distinguishes them.
        subj = tc if has_tc else f"tracker-row-{int(r.name) + 2}"
        where = tc if has_tc else f"Excel row {int(r.name) + 2}"
        if log_warning(conn, "missing_nc_id", subj,
                       f"Tracker row (owner: {owner or '(no owner)'}, {where}) has no "
                       f"ID-Blackout. It cannot be matched against the SAP overview.",
                       severity="warn", owner=owner):
            n += 1
    return n


def check_owner_spellings(conn, tr, clean_owner):
    """The bug owners noticed themselves: one person, two spellings, two counts.

    Flags any raw spelling that clean_owner does NOT map, but which looks like
    the first name of a mapped full name. Deliberately does not flag the shared
    'A / B' owners — those are a known, accepted state."""
    n = 0
    raw = tr["Issue Owner (QA/PA)"].dropna().astype(str).str.strip()
    raw = raw[raw != "\xa0"]
    canon = {clean_owner(v) for v in raw.unique() if clean_owner(v)}

    for spelling in sorted(raw.unique()):
        mapped = clean_owner(spelling)
        if mapped == spelling:              # not aliased
            if "/" in spelling:             # shared owner — known and accepted
                continue
            # A bare first name that prefixes exactly one canonical full name
            # is almost certainly the same person, unaliased.
            hits = [c for c in canon if c != spelling and c.startswith(spelling + " ")]
            if len(hits) == 1:
                if log_warning(conn, "unmapped_owner_spelling", spelling,
                               f"'{spelling}' is not mapped but looks like '{hits[0]}'. "
                               f"Their NCs are being counted as two different people. "
                               f"Add to clean_owner() aliases if they are the same person.",
                               severity="warn", owner=hits[0]):
                    n += 1
    return n


def check_owner_counts(conn, tr, clean_owner, db_file="quality.db"):
    """The reconciliation that answers 'I should have 20 and I see 8'.

    Counts each owner in the tracker, counts them in the database, and warns on
    any difference. Runs only when quality.db already exists — on the first run
    there is nothing to compare against."""
    n = 0
    if not Path(db_file).exists():
        return 0
    owners = tr["Issue Owner (QA/PA)"].apply(
        lambda v: None if _blank(v) else clean_owner(v))
    src = owners.value_counts()

    try:
        with sqlite3.connect(db_file) as c:
            db = pd.read_sql(
                "SELECT owner, COUNT(*) n FROM nc WHERE source IN ('tracker','both') "
                "GROUP BY owner", c).set_index("owner")["n"]
    except Exception as e:
        log_warning(conn, "reconcile_failed", db_file,
                    f"Could not read owner counts from {db_file}: {e}", severity="error")
        return 1

    for owner, cnt in src.items():
        got = int(db.get(owner, 0))
        if got != cnt:
            if log_warning(conn, "owner_count_mismatch", owner,
                           f"{owner}: {cnt} row(s) in the tracker but {got} in the "
                           f"database. Rows are being lost or relabelled between "
                           f"the file and the table.",
                           severity="error", owner=owner):
                n += 1
    return n


def check_row_counts(conn, tr, sap, db_file="quality.db"):
    """Count in vs count out, at the whole-file level."""
    n = 0
    if not Path(db_file).exists():
        return 0
    try:
        with sqlite3.connect(db_file) as c:
            got = pd.read_sql("SELECT source, COUNT(*) n FROM nc GROUP BY source",
                              c).set_index("source")["n"]
    except Exception:
        return 0
    in_db = int(got.sum())
    expected = len(tr) + len(sap)
    # 'both' rows are one NC from two files, so the db is legitimately smaller.
    both = int(got.get("both", 0))
    if in_db + both != expected:
        if log_warning(conn, "row_count_mismatch", "total",
                       f"{len(tr)} tracker + {len(sap)} SAP = {expected} source rows, "
                       f"but the database holds {in_db} ({both} merged). "
                       f"Difference: {expected - in_db - both}.",
                       severity="warn"):
            n += 1
    return n


# ══════════════════════════════════════════════════════════════════════════════
# Runner
# ══════════════════════════════════════════════════════════════════════════════
def check_capa_orphans(conn, db_file="quality.db"):
    """CAPA rows whose NC number matches nothing in `nc`.

    These are left unmatched deliberately — no reconstruction, no guessing at
    what '1376' or '0' was meant to be. They contribute to no coverage figure,
    so without this check they would simply be invisible. Logged one per NC
    number so each can be chased individually."""
    n = 0
    if not Path(db_file).exists():
        return 0
    try:
        with sqlite3.connect(db_file) as c:
            cols = [r[1] for r in c.execute("PRAGMA table_info(capa)")]
            if not cols:
                return 0
            orph = pd.read_sql("""
                SELECT DISTINCT c.nc_id FROM capa c
                LEFT JOIN nc n ON n.nc_id = c.nc_id
                WHERE n.nc_id IS NULL ORDER BY c.nc_id""", c)
    except Exception:
        return 0
    for _, r in orph.iterrows():
        if log_warning(conn, "capa_orphan_nc", str(r["nc_id"]),
                       f"CAPA tracker has row(s) for NC '{r['nc_id']}', but no such NC "
                       f"exists in the NC data. Left unmatched — it counts towards no "
                       f"coverage figure. Either the number is mistyped, or the NC is "
                       f"missing from the tracker and SAP export.",
                       severity="warn"):
            n += 1
    return n


def run_checks(tracker_path, sap_path, clean_owner, tracker_sheet="NC_Tracker_Black_Out",
               db_file="quality.db", verbose=True):
    """Run every check. Returns (n_new, n_total). Never raises, never blocks."""
    started = datetime.now().isoformat(timespec="seconds")
    global _CANON_FN
    _CANON_FN = clean_owner
    conn = _connect()
    new = 0

    try:
        tr = _load_tracker_raw(tracker_path, tracker_sheet)
    except Exception as e:
        log_warning(conn, "unreadable_file", str(tracker_path),
                    f"Could not read the tracker: {e}", severity="error")
        conn.commit(); conn.close()
        return 1, 1

    try:
        sap = pd.read_excel(sap_path)
        sap = sap[sap.iloc[:, 0] != "Totals"].copy()
    except Exception as e:
        log_warning(conn, "unreadable_file", str(sap_path),
                    f"Could not read the SAP overview: {e}", severity="error")
        sap = pd.DataFrame()

    new += check_schema(conn, tr, EXPECTED_TRACKER_COLS, "tracker")
    new += check_new_columns(conn, tr, KNOWN_TRACKER_COLS, "tracker")
    if not sap.empty:
        new += check_schema(conn, sap, EXPECTED_SAP_COLS, "sap_overview")
    new += check_blank_status(conn, tr)
    new += check_missing_id(conn, tr)
    new += check_owner_spellings(conn, tr, clean_owner)
    new += check_owner_counts(conn, tr, clean_owner, db_file)
    new += check_row_counts(conn, tr, sap, db_file)

    conn.commit()
    total = pd.read_sql("SELECT COUNT(*) n FROM warnings", conn).iloc[0]["n"]
    conn.close()

    if verbose:
        _print_summary(started, new, int(total))
    return new, int(total)


def _print_summary(started, n_new, n_total):
    print(f"\n  {n_new} new warning(s) this run · {n_total} in the log ({WARN_DB})")
    df = view_warnings()
    if df.empty:
        print("  Nothing to report.")
        return
    fresh = df[df["first_seen"] >= started]
    if not fresh.empty:
        print("\n  --- NEW since last run ---")
        for _, r in fresh.iterrows():
            print(f"    [{r['severity']:5}] {r['check_name']}: {r['message'][:110]}")
    counts = df.groupby(["check_name", "severity"]).size().reset_index(name="n")
    print("\n  --- log summary (all time) ---")
    for _, r in counts.iterrows():
        print(f"    {r['n']:4} × {r['check_name']}  [{r['severity']}]")
    print(f"\n  Full log: python -c \"import validate; print(validate.view_warnings().to_string())\"")


def main():
    from ingest import find_file, clean_owner, DATA_DIR, TRACKER_SHEET
    t = find_file("NCR_Cutover_Tracker*.xlsx")
    s = find_file("NC_s_Overview*.xlsx")
    if not t:
        print(f"No NCR_Cutover_Tracker*.xlsx in {DATA_DIR}/ — nothing to check.")
        return
    print("=" * 60)
    print("Validating source files")
    print("=" * 60)
    print(f"  tracker: {t.name}")
    print(f"  sap    : {s.name if s else '(missing)'}")
    run_checks(t, s, clean_owner, TRACKER_SHEET)


if __name__ == "__main__":
    main()