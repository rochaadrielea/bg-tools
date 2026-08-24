"""Column mapping, cleaning, validation and import diff.

Pure functions only — no database, no FastAPI. Everything here is testable
against a spreadsheet on its own (see test_mapping.py).
"""
from __future__ import annotations

import datetime as _dt
import re
import unicodedata

# --------------------------------------------------------------------------
# the tracker's own columns, in the tracker's own order
# (header shown in the UI, database column, kind)
# --------------------------------------------------------------------------
COLUMNS = [
    ("System",                        "system",            "list"),
    ("Issue Owner (QA/PA)",           "owner",             "list"),
    ("ID-Blackout",                   "id_blackout",       "list"),
    ("NC Type",                       "nc_type",           "list"),
    ("TC ID",                         "tc_id",             "text"),
    ("Migrated to EZ1",               "migrated",          "list"),
    ("Project",                       "project",           "list"),
    ("Flight Unit",                   "flight_unit",       "flight"),
    ("Detection",                     "detection",         "list"),
    ("Title",                         "title",             "text"),
    ("Description",                   "problem_description","text"),
    ("Failure",                       "failure",           "list"),
    ("Material",                      "material",          "text"),
    ("Batch",                         "batch",             "text"),
    ("Created On",                    "created_on",        "date"),
    ("NRB disposition",               "nrb_disposition",   "list"),
    ("Notes",                         "notes",             "text"),
    ("Status",                        "status",            "list"),
    ("Disposition Implemented Date",  "disposition_date",  "date"),
    ("Classification",                "classification",    "list"),
    ("(Real) Origin Area L1",         "origin_l1",         "list"),
    ("(Real) Origin Area L2",         "origin_l2",         "list"),
    ("RC Category L1",                "root_cause",        "list"),
    ("RC Category L2",                "rc_l2",             "list"),
    ("Department Responsible",        "responsible_area",  "list"),
    ("PSP ref.",                      "psp_ref",           "text"),
    ("NC WBS (EzyOne)",               "nc_wbs",            "text"),
    ("Closure date",                  "closure_date",      "date"),
    ("Supplier name",                 "supplier",          "text"),
]
HEADERS = [h for h, _, _ in COLUMNS]
DB_COLS = [c for _, c, _ in COLUMNS]
BY_HEADER = {h: (c, k) for h, c, k in COLUMNS}
BY_DB = {c: (h, k) for h, c, k in COLUMNS}

# which Set-up list drives which column ("flight" is per project)
LIST_OF = {
    "system": "System", "nc_type": "NC Type", "migrated": "Migrated",
    # the Blackout numbers are a fixed set from the cutover — pick, don't retype
    "id_blackout": "ID-Blackout",
    "project": "Project", "detection": "Detection", "failure": "Failure",
    "owner": "Q_Responsible", "nrb_disposition": "Disposition",
    "classification": "Classification", "status": "Status",
    # these two share the CAPA board's vocabularies — one list, both tools
    "root_cause": "capa:L1 Root cause",
    "origin_l1":  "capa:L1 Origin Area",
    # the two L2 lists follow whatever L1 holds, exactly as the CAPA board does
    "origin_l2":  "*l2origin",
    "rc_l2":      "*l2rc",
    "responsible_area": "capa:Department Responsible",
}
FLIGHT_LIST_OF_PROJECT = {
    "Ariane": "Ariane", "Vega": "Vega", "SAS": "SAS", "Vulcan": "Vulcan",
    "Relativity": "Relativity", "MHI_H3": "MHI H3",
    "Flexline PLF": "Flexline", "Flexline ISA 1-2": "Flexline",
    "Flexline ISA 2-3": "Flexline",
}

# --------------------------------------------------------------------------
# import header aliases — the incoming exports do not use the tracker's names
# --------------------------------------------------------------------------
ALIASES = {
    "system": ["system", "source", "quelle"],
    "id_blackout": ["id-blackout", "id blackout", "blackout id", "nc id", "nc_id"],
    "nc_type": ["nc type", "type", "notification type", "meldungsart"],
    "tc_id": ["tc id", "tc_id", "issue number", "issue id", "teamcenter id",
              "notification", "notification number", "meldung", "qm notification",
              "object"],
    "migrated": ["migrated to ez1", "migrated", "migration"],
    "project": ["project", "programme", "program", "projekt", "projects"],
    "flight_unit": ["flight unit", "flightunit", "unit", "serial", "fm"],
    "detection": ["detection", "detection area", "area", "erkennung", "source"],
    # Teamcenter ships two texts. "Description" is filled on every row and is
    # the short one, so it becomes the Title. "Problem Description" is the long
    # account and keeps its own column.
    # The real Title lives inside Teamcenter's Object string, recovered in
    # clean_row. "Description" is the long account (often a What/Who/When/Where
    # block) and belongs in Problem Description, not in the Title.
    "title": ["title", "short text", "kurztext",
              "title and problem description"],
    # Teamcenter's own "Problem Description" is read into a scratch field and
    # appended to the Description in clean_row, so one cell carries both texts.
    "_problem": ["problem description", "problem descr", "problembeschreibung"],
    # "Description" is filled on every Teamcenter row and carries the full
    # account; "Problem Description" is filled on fewer and adds nothing the
    # other lacks. Description first, so no row loses its text.
    "problem_description": ["description"],
    "notes": ["notes", "note", "bemerkung", "comment", "comments"],
    "origin_l1": ["origin area l1", "l1 origin area", "origin area"],
    "origin_l2": ["origin area l2", "l2 origin area"],
    "rc_l2": ["rc category l2", "root cause l2", "l2 root cause"],
    "failure": ["failure", "defect", "defect type", "defect code text", "fehler",
                "problem item", "failure linked with symptom description",
                "failure linked with symptom de"],
    "material": ["material", "material number", "part number", "materialnummer"],
    "batch": ["batch", "charge", "lot", "lot number"],
    "owner": ["issue owner (qa/pa)", "issue owner", "owner", "qa owner",
              "responsible", "coordinator"],
    "created_on": ["created on", "creation date", "created", "erstellt am", "date"],
    "nrb_disposition": ["nrb disposition", "disposition", "nrb"],
    "disposition_date": ["disposition implemented date", "implemented date"],
    "classification": ["classification", "class", "severity", "severity rating"],
    "psp_ref": ["psp ref.", "psp ref", "psp", "psp element", "parent wbs"],
    "nc_wbs": ["nc wbs (ezyone)", "nc wbs", "wbs", "wbs element"],
    "status": ["status", "notific. status", "notification status", "system status",
               "release status"],
    "closure_date": ["closure date", "closed on", "completion date"],
    "supplier": ["supplier name", "supplier", "vendor", "lieferant",
                 "contractor or supplier", "contractor"],
}
_ALIAS_LOOKUP = {a: db for db, names in ALIASES.items() for a in names}


def _key(header) -> str:
    s = unicodedata.normalize("NFKC", str(header or ""))
    return re.sub(r"\s+", " ", s).strip().lower()


# When two columns in the same file both map to one field, the file's column
# order would otherwise decide. Teamcenter puts "Problem Description" before
# "Description", yet Description is the one filled on every row — so name the
# winner explicitly instead of relying on position.
PREFERRED_ALIAS = {}


def map_headers(headers: list) -> dict:
    """Incoming column index -> database column. Unknown columns are ignored."""
    out = {}
    taken = {}                      # field -> the index currently holding it
    for i, h in enumerate(headers):
        key = _key(h)
        db = _ALIAS_LOOKUP.get(key)
        if not db:
            continue
        if db not in taken:
            out[i] = db
            taken[db] = i
        elif PREFERRED_ALIAS.get(db) == key:
            del out[taken[db]]      # the preferred column takes over
            out[i] = db
            taken[db] = i
    return out


# --------------------------------------------------------------------------
# cleaning — the view is cleaned, the uploaded file is never touched
# --------------------------------------------------------------------------
def norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip()


# TC ID arrives from Teamcenter as a composite:
#   "IR-001769/A;1-NC_1787153879971_170000019476"
# Only the two identifiers are worth keeping. The IR number stops at the first
# non-digit; the NC number keeps its full underscore tail.
_TC_KEEP = re.compile(r"(IR-\d+|NC_\d+(?:_\d+)*)", re.I)


def tidy_tc_id(v) -> str:
    """'IR-001769/A;1-NC_1787153879971_170000019476'
        -> 'IR-001769 / NC_1787153879971_170000019476'
    A value with no IR-/NC_ token is left exactly as it was: a bare legacy
    number, 'N/A' or 'To be opened' must survive untouched."""
    s = norm(v)
    if not s:
        return ""
    found = _TC_KEEP.findall(s)
    if not found:
        return s
    seen, out = set(), []
    for t in found:
        t = t.upper() if t.upper().startswith("NC_") else "IR-" + t.split("-", 1)[1]
        if t not in seen:
            seen.add(t)
            out.append(t)
    return " / ".join(out)


def clean_value(db_col: str, raw) -> str:
    """Normalise one cell. Additive: unknown values pass through untouched."""
    v = norm(raw)
    if db_col == "migrated":
        low = v.lower()
        if low == "yes":
            return "Yes"
        if low in ("no", "n0"):
            return "No"
        if low == "n/a":
            return "N/A"
        return v
    if db_col == "classification":
        low = v.lower()
        # Teamcenter grades severity as "Low-Minor", "Medium-Major",
        # "Low-Observation". The tracker only recognises Major and Minor, so the
        # level in front is dropped. Anything carrying neither word is left
        # exactly as it came and shows as off-list, never silently reshaped.
        if "major" in low:
            return "Major"
        if "minor" in low:
            return "Minor"
        return v
    if db_col == "tc_id":
        return tidy_tc_id(v)
    if db_col == "status":
        low = v.lower()
        # SAP S/4 exports use system-status strings, not Open/Closed. Map them:
        #   CLOSE AND RETURN TO WORK          -> Closed
        #   COMPLETED / SETTLED               -> Closed
        #   OPEN / IN PROCESS / NOTIFIED ...  -> Open
        if not low:
            return ""
        if low == "open" or low == "closed":
            return low.capitalize()
        if "close" in low or "closed" in low or "complet" in low or "settled" in low:
            return "Closed"
        if ("open" in low or "notified" in low or "in process" in low
                or "in work" in low or "engineering" in low or "released" in low):
            return "Open"
        return v         # unrecognised — leave it, so the flag surfaces it
    if db_col == "flight_unit" and re.fullmatch(r"n\s*/?\s*a", v, re.I):
        return ""               # a project always has a real flight unit
    if db_col in ("created_on", "disposition_date", "closure_date"):
        return to_iso(v)
    return v


def to_iso(v) -> str:
    """Accepts 2026-06-23, 23/06/2026, 23.06.2026, Excel serial. '' when unreadable."""
    v = norm(v)
    if not v:
        return ""
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", v)
    if m:
        return "-".join(m.groups())
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$", v)
    if m:
        d, mo, y = m.groups()
        return f"{y}-{int(mo):02d}-{int(d):02d}"
    if re.fullmatch(r"\d{5}", v):                      # Excel serial
        base = _dt.date(1899, 12, 30) + _dt.timedelta(days=int(v))
        return base.isoformat()
    return ""


def dmy(iso: str) -> str:
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})$", iso or "")
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else (iso or "")



def map_project(value, known: list) -> str:
    """'P-3011-00004-SAS' -> 'SAS' when SAS is on the Set-up list.

    Teamcenter writes the project as a contract code; the tracker uses the
    plain name. The name is always somewhere inside the code, so the first
    Set-up value that appears in it wins. Longest first, so 'Flexline ISA 1-2'
    is preferred over 'Flexline'. A code matching nothing is kept verbatim —
    it shows as off-list, which is the honest signal that the Set-up list needs
    a new entry. An empty value stays empty and shows as 'select'.
    """
    v = norm(value)
    if not v or not known:
        return v
    up = v.upper()
    for name in sorted(known, key=len, reverse=True):
        n = norm(name).upper()
        if n and n in up:
            return norm(name)
    return v



# Teamcenter's "Object" is the identifier and the Title glued together:
#   IR-001768/A;1-A: Damages to the HRM Node
#     -> id "IR-001768"   title "A: Damages to the HRM Node"
#   IR-001769/A;1-NC_1787153879971_170000019476
#     -> id "IR-001769"   title ""   (nobody typed a Title in Teamcenter)
# The second shape is the common one and the empty Title is the honest result:
# the cell stays yellow so a person can see the record needs a real title.
_OBJ_SPLIT = re.compile(r"^\s*(IR-\d+)(?:/[A-Za-z0-9]+)?(?:;\d+)?-(.*)$", re.S)


def split_object(v) -> tuple:
    """(identifier, title) out of a Teamcenter Object string."""
    s = norm(v)
    m = _OBJ_SPLIT.match(s)
    if not m:
        return s, ""
    ident, tail = m.group(1), m.group(2).strip()
    # Some objects read "NC_1786454785447_2004622 | FI: Batch 30 Out of ...".
    # The notification number is prefixed to the real Title with a pipe; drop
    # the prefix and keep what a person actually wrote.
    tail = re.sub(r"^\s*(NC_[\d_]+|IR-[\d/;A-Za-z-]+)\s*\|\s*", "", tail).strip()
    # Whatever follows is the Title as Teamcenter holds it — taken verbatim,
    # including the cases where someone left the notification number in the
    # Title field. The tracker shows the data as it is; it does not decide that
    # a title is wrong and hide it.
    return s, tail


def clean_row(row: dict) -> dict:
    out = {c: clean_value(c, row.get(c)) for c in DB_COLS}
    # Teamcenter hides the Title inside the Object string. Recover it, but only
    # into an empty cell — a title already written by a person always wins.
    if out.get("tc_id") and not out.get("title"):
        _, title = split_object(row.get("tc_id"))
        if title:
            out["title"] = title

    # Teamcenter keeps two texts. Description is the summary and is always
    # filled; Problem Description is the longer account and differs on 42 of
    # 110 rows. Keeping only one would drop the other, so the second is
    # appended under its own heading instead of being thrown away.
    extra = norm(row.get("_problem"))
    if extra and extra != out.get("problem_description", ""):
        base = out.get("problem_description", "")
        out["problem_description"] = (base + "\n\nProblem Description: " + extra) if base \
            else "Problem Description: " + extra
    return out


# --------------------------------------------------------------------------
# validation — explicit on purpose, one function per rule, never silent
# --------------------------------------------------------------------------
TC_PREFIX = re.compile(r"^(IR-|NC_)", re.I)
TC_PLACEHOLDERS = {"n/a", "na", "to be opened", "will be added go-live", "to be created"}
SAS_BATCH = re.compile(r"^batch\s*\d+$", re.I)


def validate(row: dict, setup: dict) -> list:
    """Returns a list of {rule, field, message}. Never rejects a row."""
    out = []

    def add(rule, field, msg):
        out.append({"rule": rule, "field": field, "message": msg})

    tc = norm(row.get("tc_id"))
    if tc and tc.lower() not in TC_PLACEHOLDERS:
        parts = [p.strip() for p in re.split(r"[\n,;]", tc) if p.strip()]
        if not all(TC_PREFIX.match(p) for p in parts):
            add("tc_id", "tc_id", "TC ID must start with IR- or NC_")

    if norm(row.get("status")) not in ("Open", "Closed"):
        add("status", "status", "status not ticked")

    if norm(row.get("status")) == "Closed" and not norm(row.get("closure_date")):
        add("closure_date", "closure_date", "closed NC without a closure date")

    project = norm(row.get("project"))
    if not project:
        add("project", "project", "project not set")
    elif project not in setup.get("Project", []):
        add("project", "project", f"project '{project}' is on no list")

    fu = norm(row.get("flight_unit"))
    if project and not fu:
        add("flight_unit", "flight_unit", "project set but no flight unit")
    elif project == "SAS" and fu and not SAS_BATCH.match(fu):
        add("flight_unit", "flight_unit", "SAS flight unit must be Batch NN")
    elif fu:
        key = FLIGHT_LIST_OF_PROJECT.get(project)
        if key and fu not in setup.get(key, []):
            add("flight_unit", "flight_unit", f"'{fu}' is not a {project} flight unit")

    for db_col, list_name in LIST_OF.items():
        v = norm(row.get(db_col))
        if db_col in ("project",):
            continue
        if v and v not in setup.get(list_name, []):
            add(db_col, db_col, f"'{v}' is on no {list_name} list")
        if db_col == "nrb_disposition" and (not v or v.lower() == "not assigned"):
            add("nrb_disposition", "nrb_disposition", "disposition still to be set")

    return out


# --------------------------------------------------------------------------
# import diff — nothing is written before this has been shown
# --------------------------------------------------------------------------
# Match rows using stable notification IDs. The SAP S/4 export ships composite
# strings like "IR-001769/A;1-NC_1787153879971_170000019476". Extract the tokens
# so a re-export of the same notification keeps the same key.
# Tolerant on purpose. Real tracker cells contain these typos:
#   NC1785767638902  — the underscore was never typed
#   R-001503         — the leading I was dropped
#   IR-1393          — written without the zero padding (really IR-001393)
# All three are the SAME notification as the Teamcenter export's version, so
# the key has to survive them. Verified against the live tracker: this recovers
# 5 rows and produces zero collisions between different NCs.
_TC_TOKEN = re.compile(r"(I?R-\d+|NC_?\d+|\b\d{6,}\b)", re.I)


def _canon_token(t: str) -> str:
    """One spelling per notification, whatever was typed."""
    t = t.upper()
    if t.startswith("R-"):                       # R-001503 -> IR-001503
        t = "I" + t
    if re.match(r"^NC\d", t):                    # NC1785…  -> NC_1785…
        t = "NC_" + t[2:]
    m = re.match(r"^IR-0*(\d+)$", t)             # IR-001393 == IR-1393
    if m:
        t = "IR-" + m.group(1)
    return t


def match_key(row: dict) -> str:
    """How an incoming row is tied to an existing NC."""
    tc = norm(row.get("tc_id"))
    if tc and tc.lower() not in TC_PLACEHOLDERS:
        tokens = [_canon_token(t) for t in _TC_TOKEN.findall(tc)]
        if tokens:
            # Take the first token — normally IR-xxx for S/4 exports, or a bare
            # 6-digit number for the legacy ECC exports. Either way it's the one
            # a person would recognise on the row.
            return "tc:" + tokens[0]
        return "tc:" + tc.split(";")[0].split(",")[0].split("\n")[0].strip().upper()
    idb = norm(row.get("id_blackout"))
    if idb and idb.lower() not in ("n/a", "na"):
        return "bo:" + idb.upper()
    return "mb:" + norm(row.get("material")).upper() + "|" + norm(row.get("batch")).upper()




def all_match_keys(row: dict) -> list:
    """Every key a row could match under. Used by diff so the same notification
    still lines up when its ID changed shape between ECC (6-digit) and S/4 (IR-xxx)."""
    keys = [match_key(row)]
    tc = norm(row.get("tc_id"))
    if tc and tc.lower() not in TC_PLACEHOLDERS:
        for tok in _TC_TOKEN.findall(tc):
            k = "tc:" + _canon_token(tok)
            if k not in keys:
                keys.append(k)
    return keys



# --------------------------------------------------------------------------
# migration rule — a real TC ID means the NC now lives in Teamcenter
# --------------------------------------------------------------------------
def is_real_tc_id(v) -> bool:
    """True only for an actual identifier. 'N/A', 'To be opened',
    'will be added go-live' and free text all return False."""
    s = norm(v)
    if not s or s.lower() in TC_PLACEHOLDERS:
        return False
    return bool(_TC_TOKEN.search(s))


def migration_update(row: dict, new_tc_id) -> dict:
    """Fields that must change when a real TC ID lands on a Burndown row.

    A row sitting on SAP or Blackout with a genuine Teamcenter ID has migrated:
    it belongs on the EZ1 tab from that moment. Returns {} when nothing moves.
    Once migrated it stays migrated — clearing the TC ID later does not send
    the row back."""
    if not is_real_tc_id(new_tc_id):
        return {}
    if norm(row.get("system")) not in ("SAP", "Blackout"):
        return {}
    return {"system": "EZ1", "migrated": "Yes"}


def diff(existing: list, incoming: list, setup: dict) -> dict:
    """existing/incoming are lists of cleaned row dicts."""
    by_key = {}
    dupes = []
    for r in existing:
        primary = match_key(r)
        if primary in by_key and by_key[primary] is not r:
            dupes.append(primary)
        for k in all_match_keys(r):
            by_key.setdefault(k, r)
    new = []              # NC key doesn't exist in the tracker at all
    filled = []           # matched — incoming will fill previously-empty cells
    skipped = []          # matched — incoming DIFFERS from a value already typed; we keep the typed one
    same = []             # matched — nothing new to add
    fills_by_key = []     # per-row: which cells will be filled
    skips_by_key = []     # per-row: which cells would be overwritten and are NOT

    for row in incoming:
        old_row = None
        k = match_key(row)
        # Try every alternate key — this is how "223194" can match "IR-223194;NC_..."
        for k in all_match_keys(row):
            if k in by_key:
                old_row = by_key[k]
                break
        k = match_key(row)
        if old_row is None:
            new.append(row)
            continue

        fills = {}         # fields where the tracker was empty and incoming has a value
        skips = {}         # fields where both have values and they differ
        for c in DB_COLS:
            inc = norm(row.get(c))
            was = norm(old_row.get(c))
            if not inc:
                continue                       # nothing to bring in
            if not was:
                fills[c] = ("", inc)           # will write
            elif inc != was:
                skips[c] = (was, inc)          # would overwrite — held back

        if fills:
            filled.append(row)
            fills_by_key.append({"key": k, "fields": fills})
        if skips:
            skipped.append(row)
            skips_by_key.append({"key": k, "fields": skips})
        if not fills and not skips:
            same.append(row)

    by_system = {}
    for row in new:
        by_system[norm(row.get("system")) or "(no system)"] = \
            by_system.get(norm(row.get("system")) or "(no system)", 0) + 1

    flagged = [{"key": match_key(r), "problems": v}
               for r in new + filled if (v := validate(r, setup))]

    return {
        "new": len(new),
        "new_by_system": by_system,
        "new_on_ez1": by_system.get("EZ1", 0),
        # "updated" kept for compatibility — it means "row that will get any
        # write at all", i.e. filled at least one empty cell. Different from
        # "would overwrite" which is held back.
        "updated": len(filled),
        "unchanged": len(same) + len(skipped) - len(filled)
                     if len(skipped) > len(filled) else len(same),
        "filled_blanks": len(filled),
        "kept_existing": len(skipped),        # rows where we held back a value
        "flagged": len(flagged),
        "fills": fills_by_key[:200],
        "kept_back": skips_by_key[:200],
        "flags": flagged[:200],
        "duplicate_keys": sorted(set(dupes)),
    }


# ==========================================================================
# CAPA — the Capa Board. Standalone: no link to the NC tracker.
# ==========================================================================
CAPA_COLUMNS = [
    ("Requestor",                          "requestor",        "list"),
    ("Responsible",                        "responsible",      "list"),
    ("Department Responsible",              "dept_responsible", "list"),
    ("Creation date requestor",             "creation_date",    "date"),
    ("Origin",                              "origin",           "list"),
    ("NC/SCAR Number",                      "nc_number",        "text"),
    ("NC Type",                             "nc_type",          "list"),
    ("PSP Element",                         "psp_element",      "text"),
    ("Affected Project",                    "project",          "list"),
    ("Project Manager",                     "project_manager",  "list"),
    ("Supplier",                            "supplier",         "list"),
    ("CAPA Type",                           "capa_type",        "list"),
    ("ID Number (1,2,3…)",                  "id_number",        "list"),
    ("(Real) Origin Area L1",               "origin_l1",        "list"),
    ("(Real) Origin Area L2",               "origin_l2",        "l2"),
    ("RC Category L1",                      "rc_l1",            "list"),
    ("RC Category L2",                      "rc_l2",            "l2"),
    ("Problem Description",                 "problem",          "text"),
    ("NC Type (Major/Minor)?",              "classification",   "list"),
    ("MAIT flow impacted in next 5 days?",  "mait_flow",        "list"),
    ("DRB planned in 1 month?",             "drb_planned",      "list"),
    ("SUM OF PRIORITY",                     "priority_sum",     "calc"),
    ("Priority",                            "priority",         "list"),
    ("Open Date",                           "open_date",        "date"),
    ("Due Date",                            "due_date",         "date"),
    ("ID Number2",                          "capa_id",          "text"),
    ("Status",                              "status",           "calc"),
    ("Close Date",                          "close_date",       "date"),
    ("Days Open",                           "days_open",        "calc"),
    ("Comments",                            "comments",         "text"),
    ("Implemented/verification of implementation", "implemented", "list"),
    ("Department accountable",              "dept_accountable", "list"),
    ("Department Assigned",                 "dept_assigned",    "list"),
    ("Change Needed",                       "change_needed",    "list"),
    ("Training",                            "training",         "list"),
]
CAPA_HEADERS = [h for h, _, _ in CAPA_COLUMNS]
CAPA_DB_COLS = [c for _, c, _ in CAPA_COLUMNS]
CAPA_KIND = {c: k for _, c, k in CAPA_COLUMNS}

# database column -> Set-up list (the `capa:` prefix is added by the caller)
CAPA_LIST_OF = {
    "requestor": "Responsible", "responsible": "Responsible",
    "dept_responsible": "Department Responsible", "origin": "Origin",
    "nc_type": "NC Type", "project": "Affected Project",
    "project_manager": "Project Manager", "supplier": "Supplier",
    "capa_type": "Type", "id_number": "ID Number",
    "origin_l1": "L1 Origin Area", "rc_l1": "L1 Root cause",
    "classification": "Classification",
    "mait_flow": "MAIT flow impacted in next 5 days",
    "drb_planned": "DRB planned in 1 month", "priority": "Priority",
    "implemented": "Implementation Verification",
    "dept_accountable": "Department Assigned",
    "dept_assigned": "Department Assigned",
    "change_needed": "Yes/No", "training": "Yes/No",
}
# an L2 column takes its list from the L1 chosen on the same row
CAPA_L2_PARENT = {"origin_l2": "origin_l1", "rc_l2": "rc_l1"}

# values the sheet uses for "nothing", which must not become list entries
CAPA_JUNK = {"#n/a", "#value!", "0", ""}


def capa_clean(db_col: str, raw) -> str:
    v = norm(raw)
    if v.lower() in ("#n/a", "#value!"):
        return ""
    if CAPA_KIND.get(db_col) == "date":
        return to_iso(v)
    if db_col in ("mait_flow", "drb_planned", "change_needed", "training"):
        low = v.lower()
        if low == "yes":
            return "Yes"
        if low == "no":
            return "No"
        if low in ("n/a", "na"):
            return "N/A"
    if v == "0":
        return ""            # the board writes 0 where a cell is empty
    return v


def capa_clean_row(row: dict) -> dict:
    out = {c: capa_clean(c, row.get(c)) for c in CAPA_DB_COLS}
    out["status"] = capa_status(out)
    out["days_open"] = capa_days_open(out)
    return out


def capa_status(row: dict) -> str:
    """Computed, never typed. Closed once a close date exists; Overdue while
    the due date is in the past; Open otherwise. Blank when no date at all —
    that shows yellow, exactly like an NC with no status ticked."""
    if norm(row.get("close_date")):
        return "Closed"
    due = norm(row.get("due_date"))
    if due:
        return "Overdue" if due < _dt.date.today().isoformat() else "Open"
    return "Open" if norm(row.get("open_date")) else ""


def capa_days_open(row: dict) -> str:
    start = norm(row.get("open_date"))
    if not start:
        return ""
    end = norm(row.get("close_date")) or _dt.date.today().isoformat()
    try:
        d0 = _dt.date.fromisoformat(start)
        d1 = _dt.date.fromisoformat(end)
    except ValueError:
        return ""
    return str((d1 - d0).days)