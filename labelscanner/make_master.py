#!/usr/bin/env python3
"""
make_master.py - build master.json for the WAE label scanner.

Reads every Teamcenter BOM export in data/source/ and writes data/master.json,
the lookup the scanner uses to validate what it reads off a physical label.

The master is a GENERATED artifact. Never edit master.json by hand - it is
overwritten on every run. Hand-known facts go in data/manual_pairs.json, which
is merged in and never overwritten.

File classification is by FILENAME:
    *MBOM*  -> materials that are actually built  -> the authoritative list
    *EBOM*  -> design structure                   -> design_only list
Anything else is skipped loudly.

Run:
    cd ~/bgtools/labelscanner
    ~/bgtools/dash/quality/bin/python make_master.py
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
MANUAL_PATH = os.path.join(HERE, "data", "manual_pairs.json")

# Columns we read. Everything else in the 74-column export is ignored.
COL_ID = "ID"
COL_NAME = "Revision Name"      # SAP short text - matches SAP/mb51 100%
COL_DESC = "Description"        # English translation - 67% match, secondary
COL_REV = "Revision"
COL_QTY = "Quantity"
COL_UOM = "Unit Of Measure"
COL_TRACE = "Traceable"
COL_SERIAL_PROF = "Serialized"
COL_LEVEL = "Level"
# Batch columns. The export has a duplicate that pandas suffixes with .1.
# "Lot" is NOT here on purpose: it is an X/blank lot-managed FLAG, not a
# number, and reading it as a batch invents pairings that reject good rows.
LOT_COLS = ["Lot Number", "Lot Number.1"]

HEADER_TOKENS = {"id", "revision", "description", "quantity", "level",
                 "revision name", "part number", "unit of measure"}

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
        print(f"    ! no clear header row found, assuming row 1")
        return 0
    return best_row


def read_bom(path):
    header = find_header_row(path)
    if header:
        print(f"    header found on row {header + 1}")
    df = pd.read_excel(path, header=header)
    if COL_ID not in df.columns:
        print(f"    ! no '{COL_ID}' column - skipped")
        return None
    return df


def rows_to_materials(df):
    """One entry per distinct material id. First non-empty value wins."""
    out = {}
    for _, r in df.iterrows():
        mid = norm_id(r.get(COL_ID))
        if not mid:
            continue
        e = out.setdefault(mid, {
            "material": mid, "name": "", "description": "",
            "revision": "", "uom": "", "traceable": None,
            "serialized": None, "levels": set(),
        })
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
    for e in out.values():
        e["levels"] = sorted(e["levels"])
    return out


def as_bool(v):
    s = clean_text(v).lower()
    if s in ("true", "yes", "y", "1", "x"):
        return True
    if s in ("false", "no", "n", "0"):
        return False
    return None


def rows_to_pairs(df, source):
    """
    Batch -> material pairings found in the BOM.

    This is the check that makes the cross-contamination bug unsavable: if the
    scanner reads material A with a batch that the master says belongs to
    material B, the row is rejected outright.

    A batch that maps to more than one material is dropped and reported. An
    ambiguous pair cannot reject anything, and a wrong rejection is its own
    kind of damage.
    """
    seen = {}
    for _, r in df.iterrows():
        mid = norm_id(r.get(COL_ID))
        if not mid:
            continue
        for col in LOT_COLS:
            if col not in df.columns:
                continue
            batch = norm_id(r.get(col))
            if not batch or batch == mid:
                continue
            seen.setdefault(batch, set()).add(mid)
    pairs, conflicts = [], []
    for batch, mats in sorted(seen.items()):
        if len(mats) == 1:
            pairs.append({"batch": batch, "material": next(iter(mats)),
                          "source": source})
        else:
            conflicts.append({"batch": batch, "materials": sorted(mats)})
    return pairs, conflicts


def load_manual():
    if not os.path.exists(MANUAL_PATH):
        return []
    try:
        with open(MANUAL_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        print(f"  ! could not read manual_pairs.json ({e}) - ignored")
        return []
    out = []
    for p in data.get("pairs", []):
        b, m = norm_id(p.get("batch")), norm_id(p.get("material"))
        if b and m:
            q = {k: v for k, v in p.items() if k not in ("batch", "material")}
            q.update({"batch": b, "material": m, "source": "manual"})
            out.append(q)
    return out


def main():
    if not os.path.isdir(SOURCE_DIR):
        sys.exit(f"missing folder: {SOURCE_DIR}")

    files = sorted(glob.glob(os.path.join(SOURCE_DIR, "*.xls*")))
    if not files:
        sys.exit(f"no Excel files in {SOURCE_DIR}")

    mbom_mats, ebom_mats = {}, {}
    pairs, conflicts = [], []
    used = []

    for path in files:
        base = os.path.basename(path)
        up = base.upper()
        kind = "MBOM" if "MBOM" in up else ("EBOM" if "EBOM" in up else None)
        print(f"  {base}")
        if not kind:
            print("    ! filename says neither MBOM nor EBOM - SKIPPED")
            continue
        df = read_bom(path)
        if df is None:
            continue
        mats = rows_to_materials(df)
        print(f"    {kind}: {len(df)} rows -> {len(mats)} distinct materials")
        (mbom_mats if kind == "MBOM" else ebom_mats).update(mats)
        if kind == "MBOM":
            p, c = rows_to_pairs(df, base)
            pairs.extend(p)
            conflicts.extend(c)
            print(f"    batch pairs: {len(p)}" +
                  (f", {len(c)} ambiguous (dropped)" if c else ""))
        used.append(base)

    if not mbom_mats:
        print("\n! No MBOM materials loaded. master.json NOT written.")
        print("  Rename the built-BOM file so it contains 'MBOM'.")
        sys.exit(1)

    design_only = sorted(set(ebom_mats) - set(mbom_mats))

    # manual pairs last so a hand-verified fact overrides the BOM
    by_batch = {p["batch"]: p for p in pairs}
    for p in load_manual():
        by_batch[p["batch"]] = p
    all_pairs = sorted(by_batch.values(), key=lambda p: p["batch"])

    # a hand-verified pair settles the ambiguity - stop reporting it
    conflicts = [c for c in conflicts if c["batch"] not in by_batch]

    master = {
        "generated": datetime.datetime.now().isoformat(timespec="seconds"),
        "source_files": used,
        # warn = unknown material is flagged but the row is still written.
        # strict = unknown material blocks the row. Do not switch to strict
        # until the master covers every part you expect to scan.
        "mode": "warn",
        "materials": [mbom_mats[m] for m in sorted(mbom_mats)],
        "design_only": design_only,
        "pairs": all_pairs,
        "ambiguous_batches": conflicts,
    }

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(master, f, indent=1, ensure_ascii=False)

    size = os.path.getsize(OUT_PATH) / 1024
    print(f"\nmaster.json written  ({size:.0f} KB)")
    print(f"  materials (MBOM) : {len(master['materials'])}")
    print(f"  design only      : {len(design_only)}")
    print(f"  batch pairs      : {len(all_pairs)}")
    print(f"  ambiguous        : {len(conflicts)}")
    print(f"  mode             : {master['mode']}")
    named = sum(1 for m in master["materials"] if m["name"])
    print(f"  with SAP name    : {named}/{len(master['materials'])}")


if __name__ == "__main__":
    main()