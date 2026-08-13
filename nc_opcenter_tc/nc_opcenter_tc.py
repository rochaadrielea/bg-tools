#!/usr/bin/env python3
"""
nc_opcenter_tc.py - join the OpCenter nonconformance export to the Teamcenter
issue export, and write one Excel report.

WHY THIS JOINS CLEANLY
    OpCenter names each nonconformance in `Identifier`   -> NC_1785928118994
    Teamcenter buries the same id in `Object`            -> IR-001651/A;1-NC_1785928118994_2003833
                                                                          ^^^^^^^^^^^^^^^^
    So the key is exact. No name matching, no fuzzy logic. The trailing number
    after the NC id is the work order and is used only as a fallback key.

WHAT THE JOIN IS FOR
    OpCenter records the person as an account code (RUAGGROUP\\YX2154).
    Teamcenter records a name (Rikard Bjon). The join is what makes the
    OpCenter data readable by anyone outside the system, and it attaches the
    IR number so a production nonconformance can be traced to its issue.

INPUTS - both taken from --data (default: data/)
    OpCenter    NonConformance*.csv   semicolon-delimited, UTF-8 BOM, CRLF.
                                      NOT comma. pandas fails on line 7 with
                                      the default separator.
    Teamcenter  *.xlsm with an `Object` column of IR- values.
                NEWEST ONE WINS: the 11 Aug export was a strict superset of
                the 5 Aug one (all 93 IRs present, 6 added, more fields
                filled), so an older snapshot has nothing to contribute and
                would only re-introduce blanks.

OUTPUT - three tabs, because a non-match is a finding, not a failure:
    Matched            an OpCenter NC that has a Teamcenter issue
    OpCenter only      a nonconformance in production with NO Teamcenter issue
    Teamcenter only    issues with no OpCenter row, split by reason
    Sources            which files were read, so a stale input is visible

Usage:
    python nc_opcenter_tc.py                       # data/ -> nc_opcenter_tc.xlsx
    python nc_opcenter_tc.py --data ~/bgtools/dash/data --out report.xlsx
"""

import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# Object looks like: IR-001651/A;1-NC_1785928118994_2003833
RE_IR = re.compile(r"(IR-\d+)")
RE_NC = re.compile(r"(NC_\d+)")
RE_WO = re.compile(r"NC_\d+_(\d+)")

# Columns kept from OpCenter, in report order. 'Type' is dropped: it is
# 'QUALITY' on every row and carries no information.
OPC_COLS = [
    "Identifier", "Context", "Severity", "Status",
    "Start Date", "End Date",
    "Work Order", "Work Order Operation",
    "Work Order Operation Sequence", "Work Order Operation Status",
]
# Columns added from Teamcenter.
TC_COLS = [
    "Issue Owner", "Problem Item", "Containment Action",
    "Failure Linked with Symptom Defect",
]


def newest(folder, pattern):
    """Newest matching file, ignoring Excel lock files (~$...)."""
    files = [p for p in folder.glob(pattern) if not p.name.startswith("~$")]
    if not files:
        return None
    return max(files, key=lambda p: p.stat().st_mtime)


def read_opcenter(path):
    """Read the OpCenter export from .csv, .xlsx or .xlsm.

    The raw OpCenter CSV is semicolon-delimited with a UTF-8 BOM. But people
    open it in Excel and re-save, which turns it into comma-delimited CSV, or
    into a workbook. All three arrive here, so the separator is detected rather
    than assumed: a wrong guess yields ONE column containing the whole line,
    which is the test used below.

    Encoding varies too - Excel on a German Windows writes cp1252, where a
    non-breaking space is byte 0xA0 and is not valid UTF-8.
    """
    name = str(getattr(path, "name", path)).lower()

    if name.endswith((".xlsx", ".xlsm", ".xls")):
        df = pd.read_excel(path, sheet_name=0, dtype=str)
    else:
        df = None
        for sep in (";", ",", "\t"):
            for enc in ("utf-8-sig", "cp1252", "latin-1"):
                try:
                    if hasattr(path, "seek"):
                        path.seek(0)
                    cand = pd.read_csv(path, sep=sep, dtype=str, encoding=enc)
                except Exception:
                    continue
                if cand.shape[1] > 1:      # separator was right
                    df = cand
                    break
            if df is not None:
                break
        if df is None:
            sys.exit(f"{name}: could not read as CSV with ; , or tab.")

    df.columns = [str(c).strip().lstrip("\ufeff") for c in df.columns]
    if "Identifier" not in df.columns:
        sys.exit(f"{name}: no 'Identifier' column. Found: {list(df.columns)[:8]}")
    return df


def find_teamcenter(folder):
    """The Teamcenter export is found by CONTENT, not by name: its filename is
    a bare timestamp. Any .xlsm/.xlsx whose first sheet has an `Object` column
    full of IR- values qualifies. Newest first, so the freshest wins."""
    cands = [p for p in list(folder.glob("*.xlsm")) + list(folder.glob("*.xlsx"))
             if not p.name.startswith("~$")]
    for p in sorted(cands, key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            head = pd.read_excel(p, sheet_name=0, nrows=5, dtype=str)
        except Exception:
            continue
        cols = [str(c).strip() for c in head.columns]
        if "Object" in cols and head["Object"].astype(str).str.contains("IR-").any():
            return p
    return None


# The Object cell has a consistent grammar. Every one of the 99 rows in the
# 11 Aug export parses with this:
#
#   IR-001622 / A ; 1 - NC_1785149656124 _ 2004385 | A: Matchdrilling poor ...
#     IR      rev seq      NC id           work no  code   description
#
# Revision and sequence are 'A' and '1' on every row, so they are parsed and
# then dropped - they carry nothing.
RE_HEAD = re.compile(r"^(IR-\d+)/([A-Z0-9]+);(\d+)-(.*)$")
RE_BODY = re.compile(r"^(NC_\d+)(?:_(\d+))?\s*(?:\|\s*(.*))?$")
# 'CC : text' with a space before the colon is real and appears in the export,
# so the colon is matched loosely. 'P_SAS:' is a project prefix, not a
# detection code, and is excluded by requiring letters only.
RE_CODE = re.compile(r"^([A-Z]{1,3})\s*:\s*(.*)$")

# The letter codes are the SAME taxonomy as the tracker's Detection column,
# so a row with no NC id is still classifiable by where it was found.
DETECTION = {
    "A": "Assembly & Matchdrilling (SAS)", "B": "Bonding",
    "C": "Completion + Transport", "CC": "Customer Complaint",
    "D": "Delivery", "DP": "Data package prep (SAS)",
    "FI": "Final Inspection (SAS)", "I": "Integration",
    "II": "Incoming Inspection", "M": "Machining", "N": "NDI",
    "P": "Post Delivery", "PD": "Process Deviation", "S": "Supplier",
    "T": "Testing",
}


def parse_object(s):
    """Split one Teamcenter `Object` cell into its parts.

        IR-001622/A;1-NC_1785149656124_2004385 | A: Matchdrilling poor surface

    gives:
        IR          IR-001622/A                  (issue number + revision)
        NC          NC_1785149656124_2004385     (the full NC string, as typed)
        Identifier  NC_1785149656124             (JOIN KEY - see below)
        Work order  2004385
        Detection   Assembly & Matchdrilling (SAS)
        Description Matchdrilling poor surface

    IR and NC are the two columns kept whole for reading. The match is made on
    `Identifier`, not on NC: OpCenter's own Identifier is NC_1785149656124 with
    NO work-order suffix, so joining on the full string would match nothing.

    Anything that does not fit the grammar keeps the whole cell as its
    description rather than being dropped.
    """
    s = str(s).strip()
    out = {"IR": None, "NC": None, "Identifier": None, "Work order": None,
           "Detection code": None, "Detection": None, "Description": None}
    m = RE_HEAD.match(s)
    if not m:
        out["Description"] = s or None
        return out
    out["IR"] = f"{m.group(1)}/{m.group(2)}"      # IR-001622/A
    body = m.group(4).strip()

    b = RE_BODY.match(body)
    if b:
        out["NC"] = b.group(0).split("|")[0].strip()   # NC_1785149656124_2004385
        out["Identifier"] = b.group(1)                 # NC_1785149656124
        out["Work order"] = b.group(2)
        body = (b.group(3) or "").strip()

    c = RE_CODE.match(body)
    if c and c.group(1) in DETECTION:
        out["Detection code"] = c.group(1)
        out["Detection"] = DETECTION[c.group(1)]
        body = c.group(2).strip()

    out["Description"] = body or None
    return out


def read_teamcenter(path):
    df = pd.read_excel(path, sheet_name=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    parsed = pd.DataFrame([parse_object(s) for s in df["Object"]])
    out = parsed.copy()
    out["Creation Date"] = df.get("Creation Date")
    out["Issue Approver"] = df.get("Issue Approver")
    for c in TC_COLS:
        out[c] = df[c] if c in df.columns else None
    out = out[out["IR"].notna()].copy()
    # A repeated IR would silently multiply OpCenter rows on the merge.
    return out.drop_duplicates(subset=["IR"], keep="first").reset_index(drop=True)


def build(opc, tc):
    """Left join from OpCenter: every production nonconformance appears once,
    whether or not it reached Teamcenter."""
    keyed = tc[tc["Identifier"].notna()]
    extra = ["IR", "NC", "Identifier", "Detection", "Description", "Creation Date"]
    matched = opc.merge(keyed[extra + TC_COLS], on="Identifier", how="left")

    hit = matched[matched["IR"].notna()].copy()
    miss = matched[matched["IR"].isna()].copy()

    cols = (["IR", "NC"] + OPC_COLS + ["Detection", "Description"] + TC_COLS
            + ["Creation Date", "User"])
    hit = hit[[c for c in cols if c in hit.columns]]
    miss = miss[[c for c in OPC_COLS + ["User"] if c in miss.columns]]

    # Teamcenter rows with no OpCenter counterpart, and WHY.
    used = set(opc["Identifier"].dropna())
    tc_only = tc[~tc["Identifier"].isin(used)].copy()
    tc_only["Reason"] = tc_only["Identifier"].apply(
        lambda v: "no NC id in Object (free-text issue)" if pd.isna(v)
        else "NC id not present in the OpCenter export")
    tc_only = tc_only.sort_values(["Reason", "IR"])
    tc_only = tc_only[["IR", "NC", "Reason", "Detection", "Description",
                       "Creation Date"] + TC_COLS]

    return hit, miss, tc_only


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="data", help="folder holding both exports")
    ap.add_argument("--out", default="nc_opcenter_tc.xlsx")
    a = ap.parse_args()

    folder = Path(a.data)
    if not folder.is_dir():
        sys.exit(f"missing folder: {folder}")

    opc_path = (newest(folder, "NonConformance*.csv")
                or newest(folder, "NonConformance*.xls*"))
    if opc_path is None:
        sys.exit(f"no NonConformance*.csv / .xlsx in {folder}/")
    tc_path = find_teamcenter(folder)
    if tc_path is None:
        sys.exit(f"no Teamcenter export (xlsm/xlsx with an IR- 'Object' column) in {folder}/")

    print(f"  OpCenter   : {opc_path.name}")
    print(f"  Teamcenter : {tc_path.name}")

    opc = read_opcenter(opc_path)
    tc = read_teamcenter(tc_path)
    hit, miss, tc_only = build(opc, tc)

    sources = pd.DataFrame([
        {"Role": "OpCenter", "File": opc_path.name, "Rows": len(opc),
         "File date": datetime.fromtimestamp(opc_path.stat().st_mtime)
         .strftime("%Y-%m-%d %H:%M")},
        {"Role": "Teamcenter", "File": tc_path.name, "Rows": len(tc),
         "File date": datetime.fromtimestamp(tc_path.stat().st_mtime)
         .strftime("%Y-%m-%d %H:%M")},
        {"Role": "Report generated", "File": "",
         "Rows": "", "File date": datetime.now().strftime("%Y-%m-%d %H:%M")},
    ])

    with pd.ExcelWriter(a.out, engine="openpyxl") as w:
        hit.to_excel(w, sheet_name="Matched", index=False)
        miss.to_excel(w, sheet_name="OpCenter only", index=False)
        tc_only.to_excel(w, sheet_name="Teamcenter only", index=False)
        sources.to_excel(w, sheet_name="Sources", index=False)
        for name, df in [("Matched", hit), ("OpCenter only", miss),
                         ("Teamcenter only", tc_only), ("Sources", sources)]:
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, start=1):
                width = max(len(str(col)),
                            *(len(str(v)) for v in df[col].head(200))) if len(df) else len(str(col))
                ws.column_dimensions[ws.cell(1, i).column_letter].width = min(46, width + 2)

    print(f"\n  Matched          : {len(hit)} of {len(opc)} OpCenter rows")
    print(f"  OpCenter only    : {len(miss)}  (no Teamcenter issue)")
    print(f"  Teamcenter only  : {len(tc_only)}")
    if len(hit):
        for c in TC_COLS:
            print(f"    {c:36s} filled on {int(hit[c].notna().sum())}/{len(hit)} matched rows")
    print(f"\n  -> {a.out}")


if __name__ == "__main__":
    main()