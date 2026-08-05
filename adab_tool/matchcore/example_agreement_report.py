"""
example_agreement_report.py — how to use matchcore on an ADAB-style compare.

Reads a design list and a built list, runs the two-witness engine, and writes
agreement.xlsx with a Verdict per matched material plus DESC_ONLY rescue
candidates (same description, different number) found among the unmatched parts.

    python matchcore/example_agreement_report.py DESIGN.xlsm BUILT.xlsx [out.xlsx]
"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
from matchcore import norm_key, build_vectorizer, agree


def _map(df, key_col, desc_col):
    m = {}
    for _, r in df.iterrows():
        k = norm_key(r.get(key_col))
        if k:
            m.setdefault(k, r.get(desc_col))
    return m


def agreement_report(design_df, built_df, design_key="ID", built_key="Material",
                     desc="Description"):
    dmap = _map(design_df, design_key, desc)
    bmap = _map(built_df, built_key, desc)
    vec = build_vectorizer(list(dmap.values()), list(bmap.values()))

    matched_rows = []
    for k in sorted(set(dmap) & set(bmap)):
        a = agree(k, dmap[k], k, bmap[k], vec)
        matched_rows.append({"Part": k, "Design Desc": dmap[k],
                             "Built Desc": bmap[k], "Similarity": round(a.desc_similarity, 3),
                             "Verdict": a.verdict})

    extra, missing = set(bmap) - set(dmap), set(dmap) - set(bmap)
    rescue = []
    for eb in sorted(extra):
        best = None
        for md in missing:
            a = agree(eb, bmap[eb], md, dmap[md], vec)
            if (a.verdict == "DESC_ONLY" or a.desc_similarity >= 0.6) and \
               (best is None or a.desc_similarity > best["Similarity"]):
                best = {"Built Part": eb, "Built Desc": bmap[eb],
                        "Design Part?": md, "Design Desc": dmap[md],
                        "Similarity": round(a.desc_similarity, 3)}
        if best:
            rescue.append(best)
    return pd.DataFrame(matched_rows), pd.DataFrame(rescue)


def main():
    design, built = sys.argv[1], sys.argv[2]
    out = sys.argv[3] if len(sys.argv) > 3 else "agreement.xlsx"
    dd = pd.read_excel(design, sheet_name=0)
    dd = dd[dd["ID"].astype(str).str.strip().ne("nan")]
    bb = pd.read_excel(built, sheet_name="Parts")
    m, r = agreement_report(dd, bb)
    with pd.ExcelWriter(out) as w:
        m.to_excel(w, sheet_name="Matched Agreement", index=False)
        r.to_excel(w, sheet_name="DESC_ONLY rescue", index=False)
    print(f"{len(m)} matched pairs "
          f"({(m['Verdict']=='CONFLICT').sum()} CONFLICT), "
          f"{len(r)} rescue candidate(s) -> {out}")


if __name__ == "__main__":
    main()
