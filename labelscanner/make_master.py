#!/usr/bin/env python3
"""
make_master.py - build master.json for the WAE label scanner.

Reads every Teamcenter BOM export in data/source/ and writes data/master.json,
the list the scanner checks a scanned material number against.

ONE LIST. MBOM and EBOM are merged into a single set of material numbers,
because the scanner asks one question and only one:

    is this material in the master list?  yes -> ML   no -> RV

A design part is still a real part, so there is no separate design_only list and
no third verdict. RV never blocks a row in the scanner - it marks it for review.

WHAT IS NOT HERE, AND WILL NOT BE
    Batch / Charge validation. A batch is assigned by the production order at
    build time; no BOM states which batch was consumed, so any batch-vs-master
    check would reject every new batch. A correct Charge depends entirely on
    reading the right label, which the scanner enforces with its own eviction
    and conflict rules - not with data from here.

The master is a GENERATED artifact. Never edit master.json by hand; every run
overwrites it.

File classification is by FILENAME, and it is a guard, not a convenience:
    *MBOM*  -> merged into the list
    *EBOM*  -> merged into the list
Anything else is SKIPPED LOUDLY, so an unrelated spreadsheet dropped into
data/source/ can never quietly become part of the master.

Run:
    cd ~/bgtools/labelscanner
    ~/bgtools/dash/quality/bin/python make_master.py            # full, internal
    ~/bgtools/dash/quality/bin/python make_master.py --public   # numbers only

    A new export in data/source/ changes NOTHING until this is re-run. Finish
    with --public before pushing, or the published file carries part names.

PUBLIC MODE
    Writes the same data/master.json containing ONLY the material numbers, the
    mode and the date. No SAP short texts, no descriptions, no revisions, no
    unit of measure, no traceability flags, no assembly levels, no source
    filenames.

    This exists because the scanner is served from GitHub Pages, where
    master.json is fetchable by anyone with the URL. The scanner confirms the
    material NUMBER and nothing else, so every other field would be published
    for no benefit. A bare list of numbers reveals no part names, no structure
    and no programme.

    It is a reduction of exposure, not an elimination of it. Publishing any
    company data outside the tenant is a decision for Adriele and Information
    Security, not for this script.
"""

import json
import glob
import os
import sys
import datetime
import re

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas not found. Run with ~/bgtools/dash/quality/bin/python")

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE_DIR = os.path.join(HERE, "data", "source")
OUT_PATH = os.path.join(HERE, "data", "master.json")

# Columns we read. Everything else in the 74-column export is ignored.
COL_ID = "ID"
COL_NAME = "Revision Name"      # SAP short text - matches SAP/mb51 100%
COL_DESC = "Description"        # English translation - 67% match, secondary
COL_REV = "Revision"
COL_UOM = "Unit Of Measure"
COL_TRACE = "Traceable"
COL_SERIAL_PROF = "Serialized"
COL_LEVEL = "Level"

HEADER_TOKENS = {"id", "revision", "description", "quantity", "level",
                 "revision name", "part number", "unit of measure"}

# THE GATE IS THE CONTENT, NOT THE FILENAME.
# A Teamcenter BOM export is recognised by its columns: an ID column plus at
# least MIN_SIGNATURE of the structure columns below. A stock list, a scan
# report or a random spreadsheet does not carry these, so it is still skipped -
# but a real export no longer has to be renamed to be accepted.
SIGNATURE_COLS = {"level", "revision", "revision name", "description",
                  "quantity", "find number", "unit of measure"}
MIN_SIGNATURE = 3

# "English SAP name: FOO" -> "FOO". Same cleanup as adab_batch_compare.
LANG_DESC = re.compile(
    r"^\s*(english|german|french|italian|deutsch)\s+sap\s+name\s*:\s*",
    re.I)


def norm_id(v):
    """
    Material numbers must survive Excel.

    Excel hands back 7004369 as the float 7004369.0, and a stray space or a
    lowercase letter must not create a second identity. Suffixes ARE preserved
    on purpose: C3529115-C is a different part from C3529115.

    Must stay in step with normId() in index.html - the scanner normalises the
    scanned value the same way before looking it up.

    Returns "" for anything that is not a usable id.
    """
    if v is None:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v.is_integer():
            return str(int(v))
        return str(v)
    if isinstance(v, int):
        return str(v)
    s = str(v).strip()
    if s.lower() in ("", "nan", "none", "-", "n/a"):
        return ""
    # a float that arrived as text
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s.upper()


def clean_text(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() in ("", "nan", "none"):
        return ""
    return LANG_DESC.sub("", s).strip()


def find_header_row(path):
    """
    Teamcenter exports normally put headers on row 1, but a tab fed back from
    an ADAB report carries a note banner first. Scan the first 15 rows and pick
    the one that looks most like a header. Returns a 0-based index for pandas.
    """
    probe = pd.read_excel(path, header=None, nrows=15)
    best_row, best_hits = 0, -1
    for i in range(len(probe)):
        cells = [str(c).strip().lower() for c in probe.iloc[i].tolist()]
        hits = sum(1 for c in cells if c in HEADER_TOKENS)
        if hits > best_hits:
            best_row, best_hits = i, hits
    if best_hits < 2:
        print("    ! no clear header row found, assuming row 1")
        return 0
    return best_row


def read_bom(path):
    header = find_header_row(path)
    if header:
        print(f"    header found on row {header + 1}")
    df = pd.read_excel(path, header=header)
    if COL_ID not in df.columns:
        print(f"    ! no '{COL_ID}' column - this is not a BOM export, SKIPPED")
        return None
    cols = {str(c).strip().lower() for c in df.columns}
    hits = SIGNATURE_COLS & cols
    if len(hits) < MIN_SIGNATURE:
        print(f"    ! has an ID column but only {len(hits)} BOM structure column(s) "
              f"({', '.join(sorted(hits)) or 'none'}) - not a BOM export, SKIPPED")
        return None
    return df


def root_ids(df):
    """
    The level-0 rows are the assembly this export is a tree of. Reported per
    file because it is the only thing that shows you two exports cover two
    different products - a filename will not, and merging two assemblies makes
    ML mean 'exists somewhere' instead of 'belongs to this build'.
    """
    if COL_LEVEL not in df.columns:
        return set()
    out = set()
    for _, r in df.iterrows():
        lvl = str(r.get(COL_LEVEL)).strip()
        if lvl in ("0", "0.0"):
            mid = norm_id(r.get(COL_ID))
            if mid:
                out.add(mid)
    return out


def kind_of(base):
    """
    Label only. MBOM/EBOM in the filename is recorded when present because it
    is useful when reconciling two exports, but it no longer decides whether a
    file is read - see SIGNATURE_COLS.
    """
    up = base.upper()
    if "MBOM" in up:
        return "MBOM"
    if "EBOM" in up:
        return "EBOM"
    return "BOM"


def as_bool(v):
    s = clean_text(v).lower()
    if s in ("true", "yes", "y", "1", "x"):
        return True
    if s in ("false", "no", "n", "0"):
        return False
    return None


def collect(df, kind, into):
    """
    Merge one export's materials into `into` (id -> entry).

    First non-empty value wins, so a part appearing in both the MBOM and the
    EBOM keeps whichever text was found first rather than losing it. `sources`
    records which BOM kinds the number came from - useful when reconciling the
    two exports, and dropped from the published file.

    Returns the number of distinct ids seen in THIS file.
    """
    seen = set()
    for _, r in df.iterrows():
        mid = norm_id(r.get(COL_ID))
        if not mid:
            continue
        seen.add(mid)
        e = into.setdefault(mid, {
            "material": mid, "name": "", "description": "",
            "revision": "", "uom": "", "traceable": None,
            "serialized": None, "levels": set(), "sources": set(),
        })
        e["sources"].add(kind)
        e["name"] = e["name"] or clean_text(r.get(COL_NAME))
        e["description"] = e["description"] or clean_text(r.get(COL_DESC))
        e["revision"] = e["revision"] or clean_text(r.get(COL_REV))
        e["uom"] = e["uom"] or clean_text(r.get(COL_UOM))
        if e["traceable"] is None:
            e["traceable"] = as_bool(r.get(COL_TRACE))
        if e["serialized"] is None:
            e["serialized"] = as_bool(r.get(COL_SERIAL_PROF))
        lvl = clean_text(r.get(COL_LEVEL))
        if lvl:
            e["levels"].add(lvl)
    return len(seen)


def main():
    if not os.path.isdir(SOURCE_DIR):
        sys.exit(f"missing folder: {SOURCE_DIR}")

    files = sorted(glob.glob(os.path.join(SOURCE_DIR, "*.xls*")))
    if not files:
        sys.exit(f"no Excel files in {SOURCE_DIR}")

    mats = {}
    used, skipped = [], []
    roots = {}

    for path in files:
        base = os.path.basename(path)
        print(f"  {base}")
        df = read_bom(path)
        if df is None:
            skipped.append(base)
            continue
        kind = kind_of(base)
        before = len(mats)
        n = collect(df, kind, mats)
        added = len(mats) - before
        rt = root_ids(df)
        for r in rt:
            roots.setdefault(r, []).append(base)
        print(f"    {kind}: {len(df)} rows -> {n} distinct materials"
              f", {added} new to the list")
        print(f"    root assembly: {', '.join(sorted(rt)) if rt else 'no level-0 row found'}")
        used.append(base)

    if not mats:
        print("\n! No materials loaded. master.json NOT written.")
        print("  Every file was skipped: none of them looks like a BOM export")
        print("  (needs an 'ID' column plus BOM structure columns such as")
        print("  Level, Revision Name, Description, Quantity).")
        sys.exit(1)

    # counts per BOM kind, kept for reconciling two exports of the same assembly
    only_m = sorted(k for k, e in mats.items() if e["sources"] == {"MBOM"})
    only_e = sorted(k for k, e in mats.items() if e["sources"] == {"EBOM"})
    both = sorted(k for k, e in mats.items() if len(e["sources"]) > 1)

    public = "--public" in sys.argv
    numbers = sorted(mats)

    if public:
        # Only what the scanner reads. Everything else is withheld.
        master = {
            "generated": datetime.date.today().isoformat(),
            "mode": "warn",
            "materials": numbers,
        }
    else:
        out_mats = []
        for m in numbers:
            e = dict(mats[m])
            e["levels"] = sorted(e["levels"])
            e["sources"] = sorted(e["sources"])
            out_mats.append(e)
        master = {
            "generated": datetime.datetime.now().isoformat(timespec="seconds"),
            "source_files": used,
            "mode": "warn",
            "materials": out_mats,
        }

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(master, f, indent=1, ensure_ascii=False)

    size = os.path.getsize(OUT_PATH) / 1024
    print(f"\nmaster.json written  ({size:.0f} KB)"
          + ("   [PUBLIC - numbers only]" if public else "   [FULL - internal only]"))
    print(f"  materials in list : {len(numbers)}")
    if only_m or only_e:
        print(f"    in MBOM only    : {len(only_m)}")
        print(f"    in EBOM only    : {len(only_e)}")
        print(f"    in both         : {len(both)}")
    print(f"  files merged      : {len(used)}")
    print(f"  assemblies        : {', '.join(sorted(roots)) if roots else 'unknown'}")
    if skipped:
        print(f"  files SKIPPED     : {len(skipped)} -> {', '.join(skipped)}")

    if len(roots) > 1:
        print("\n  ! MORE THAN ONE ASSEMBLY IS IN THIS LIST:")
        for r in sorted(roots):
            print(f"      {r}  <- {', '.join(roots[r])}")
        print("    ML now means 'this part exists in one of these assemblies',")
        print("    NOT 'this part belongs to the build in front of you'. If that")
        print("    is not what you want, move the extra export out of data/source")
        print("    and re-run.")

    if public:
        print("\n  Published : material numbers, mode, date")
        print("  Withheld  : names, descriptions, revisions, UoM,")
        print("              traceable/serialized flags, levels,")
        print("              which BOM each number came from, source filenames")
    else:
        named = sum(1 for m in master["materials"] if m["name"])
        print(f"  with SAP name     : {named}/{len(numbers)}")
        print("\n  This file contains part descriptions. Do NOT commit it.")
        print("  For the published copy run:  make_master.py --public")


if __name__ == "__main__":
    main()