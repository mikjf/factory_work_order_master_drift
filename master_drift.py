"""Master drift - build the planners' workbook from the two SQL exports.

Inputs (exported to Excel, first sheet, header row as the query names it),
by default in the private folder next to this repo
(<repo>_private\\data\\):
  drift.xlsx    output of master_drift.sql        (level 1, one row per open,
                                                   not-started WO operation)
  masters.xlsx  output of old_revision_masters.sql (master lines of the own
                                                   and the current revision,
                                                   for parts with
                                                   older-revision WOs)
Run both queries on the same day: masters change. The script warns if the
two files were saved on different days.

Output: <repo>_private\\outputs\\master_drift_list_<date>.xlsx by default,
never inside this repo - it names real parts and orders:
  00_Summary          counts, checks, input files
  01_Drift            LEVEL 1: WO copy differs from its OWN revision's master
  02_NoMasterLine     WO lines with no own-revision master line
  03_OldRevisionWOs   LEVEL 2: one row per WO built on an older revision
  04_CrossRevision    LEVEL 3 (INDICATIVE): older-revision WO lines against
                      the CURRENT revision's master
  05_Review           LEVEL 1 for a person to judge: one row per work order
                      x part x differing field, with its PATTERN number and
                      empty VERDICT / NOTE columns - fill in here
  06_ReviewPatterns   the same grouped into numbered patterns (family,
                      project, operation, field, WO value, master value) -
                      overview only

LEVEL 3 - REVIEW AGAIN LATER (agreed with the user, 2026-09-28). A new
revision inserts steps and shifts later ones, so lines cannot be paired by
sequence. Pairing here uses only rules that need no guessing, or rest on
unique evidence; everything else is reported as NOT_COMPARED:
  UNIQUE_NAME            the name occurs exactly once in both revisions
  IDENTICAL_OCCURRENCES  the name repeats, but all its occurrences carry the
                         same values within each revision - any pairing
                         gives the same answer
  BY_RESOURCE            the name repeats, each occurrence on its own
                         resource, same resource set in both revisions
  NOT_IN_CURRENT         the name does not occur in the current revision
The standing assumption is "same name = same step". It is not proven.
"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_DIR = Path(__file__).resolve().parent
PRIVATE_DIR = REPO_DIR.parent / f"{REPO_DIR.name}_private"

WO_KEY = ["WORK_ORDER", "LOT_ID", "SPLIT_ID", "SUB_ID"]
LINE_KEY = WO_KEY + ["OPERATION_SEQUENCE"]

# compared field -> (text field?, flag suffix used by master_drift.sql)
# NULL_AS: value an empty field counts as when compared - kept in step with
# master_drift.sql. Empty crew = 1 (user, 2026-09-28). Output shows raw values.
NULL_AS = {"USER_5": 1}
FIELDS = {
    "RESOURCE_ID": (True, "RESOURCE"),
    "RUN_TYPE": (True, "RUN_TYPE"),
    "RUN": (False, "RUN"),
    "LOAD_SIZE_QTY": (False, "LOAD_SIZE_QTY"),
    "SETUP_HRS": (False, "SETUP_HRS"),
    "USER_5": (False, "USER_5"),
}

DRIFT_COLUMNS = (
    LINE_KEY
    + ["WO_KIND", "WO_STATUS", "PROJECT_CODE", "FAMILY", "PART_ID", "PART_DESCRIPTION",
       "WO_REV", "CURRENT_REV", "ON_CURRENT_REV", "CREATE_DATE", "DESIRED_RLS_DATE",
       "OPERATION", "OP_STATUS", "MASTER_OPS_ROWID", "NO_MASTER_LINE"]
    + [f"WO_{f}" for f in FIELDS] + [f"M_{f}" for f in FIELDS]
    + [f"D_{s}" for _, s in FIELDS.values()]
)
MASTER_COLUMNS = ["PART_ID", "ROLE", "REV", "MASTER_EXISTS", "SEQUENCE_NO", "OPERATION"] + list(FIELDS)


def code(v):
    """Text-like code: None for blank, trailing blanks dropped, 450.0 -> '450'.

    SQL '=' ignores trailing blanks; pandas does not. Revisions and resources
    arrive as numbers or text depending on the column, hence the int step.
    """
    if pd.isna(v):
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).rstrip()
    return s if s else None


def same(a, b, text, null_as=None):
    """Element-wise equality with NULL = NULL, numbers compared as numbers."""
    if text:
        a, b = a.map(code), b.map(code)
        return (a == b) | (a.isna() & b.isna())
    a, b = pd.to_numeric(a, errors="coerce"), pd.to_numeric(b, errors="coerce")
    if null_as is not None:
        a, b = a.fillna(null_as), b.fillna(null_as)
    return (a == b) | (a.isna() & b.isna())


def load(path, required):
    df = pd.read_excel(path, sheet_name=0)
    missing = [c for c in required if c not in df.columns]
    if missing:
        sys.exit(f"{path.name}: missing columns {missing}")
    return df


def normalise_drift(d):
    for c in ["WO_REV", "CURRENT_REV", "OPERATION", "WO_KIND", "WO_STATUS", "OP_STATUS", "PART_ID"]:
        d[c] = d[c].map(code)
    for f, (text, _) in FIELDS.items():
        if text:
            d[f"WO_{f}"] = d[f"WO_{f}"].map(code)
            d[f"M_{f}"] = d[f"M_{f}"].map(code)
    return d


def normalise_masters(m):
    for c in ["PART_ID", "ROLE", "REV", "OPERATION", "RESOURCE_ID", "RUN_TYPE"]:
        m[c] = m[c].map(code)
    return m


def check_level1_flags(d):
    """Recompute every D_ flag and require agreement with the SQL."""
    matched = d[d.NO_MASTER_LINE == 0]
    problems = {}
    for f, (text, s) in FIELDS.items():
        mine = (~same(matched[f"WO_{f}"], matched[f"M_{f}"], text, NULL_AS.get(f))).astype(int)
        bad = int((mine != matched[f"D_{s}"].astype(int)).sum())
        if bad:
            problems[f"D_{s}"] = bad
    return problems


def signature(rows):
    rows = rows[list(FIELDS)].fillna(NULL_AS)
    return rows.apply(lambda r: "|".join(map(str, r)), axis=1)


def pair_current(own_line, own_rev_lines, cur_rev_lines):
    """Return (category, current master line or None). See module docstring."""
    name = own_line.OPERATION
    own_same = own_rev_lines[own_rev_lines.OPERATION == name]
    cur_same = cur_rev_lines[cur_rev_lines.OPERATION == name]
    if cur_same.empty:
        return "NOT_IN_CURRENT", None
    if len(own_same) == 1 and len(cur_same) == 1:
        return "UNIQUE_NAME", cur_same.iloc[0]
    if signature(own_same).nunique() == 1 and signature(cur_same).nunique() == 1:
        return "IDENTICAL_OCCURRENCES", cur_same.iloc[0]
    if (own_same.RESOURCE_ID.is_unique and cur_same.RESOURCE_ID.is_unique
            and set(own_same.RESOURCE_ID) == set(cur_same.RESOURCE_ID)):
        return "BY_RESOURCE", cur_same[cur_same.RESOURCE_ID == own_line.RESOURCE_ID].iloc[0]
    return "NOT_COMPARED", None


def level3(d, m):
    old = d[d.ON_CURRENT_REV == 0].copy()
    lines = m[m.SEQUENCE_NO.notna()]
    by_rev = {k: g for k, g in lines.groupby(["PART_ID", "ROLE", "REV"], dropna=False)}
    extracted = set(m.PART_ID)

    out = []
    for r in old.itertuples(index=False):
        rec = {"CATEGORY": None}
        own = by_rev.get((r.PART_ID, "OWN", r.WO_REV))
        cur = by_rev.get((r.PART_ID, "CUR", r.CURRENT_REV))
        if r.PART_ID not in extracted:
            rec["CATEGORY"] = "MASTERS_NOT_EXTRACTED"
        elif r.NO_MASTER_LINE == 1 or own is None:
            rec["CATEGORY"] = "NO_OWN_MASTER_LINE"
        elif cur is None:
            rec["CATEGORY"] = "NO_CURRENT_MASTER"
        else:
            hit = own[(own.SEQUENCE_NO == r.OPERATION_SEQUENCE) & (own.OPERATION == r.OPERATION)]
            if len(hit) != 1:
                # the SQL matched it, so the export must contain it exactly once
                rec["CATEGORY"] = "OWN_LINE_NOT_IN_EXPORT"
            else:
                cat, cline = pair_current(hit.iloc[0], own, cur)
                rec["CATEGORY"] = cat
                if cline is not None:
                    rec["CUR_SEQUENCE_NO"] = cline.SEQUENCE_NO
                    for f in FIELDS:
                        rec[f"CUR_{f}"] = cline[f]
        out.append(rec)

    res = pd.DataFrame(out, index=old.index)
    for f in FIELDS:
        if f"CUR_{f}" not in res:
            res[f"CUR_{f}"] = pd.NA
    if "CUR_SEQUENCE_NO" not in res:
        res["CUR_SEQUENCE_NO"] = pd.NA
    old = old.join(res)
    compared = old.CUR_SEQUENCE_NO.notna()
    for f, (text, s) in FIELDS.items():
        diff = (~same(old[f"WO_{f}"], old[f"CUR_{f}"], text, NULL_AS.get(f))).astype("Int64")
        old[f"D3_{s}"] = diff.where(compared)
    old["D3_ANY"] = old[[f"D3_{s}" for _, s in FIELDS.values()]].max(axis=1).where(compared)
    return old


def level2(d, l3):
    old = d[d.ON_CURRENT_REV == 0]
    d_any = old[[f"D_{s}" for _, s in FIELDS.values()]].fillna(0).max(axis=1)
    base = old.assign(L1_DIFF=d_any).groupby(WO_KEY, dropna=False).agg(
        WO_KIND=("WO_KIND", "first"), WO_STATUS=("WO_STATUS", "first"),
        PROJECT_CODE=("PROJECT_CODE", "first"), FAMILY=("FAMILY", "first"),
        PART_ID=("PART_ID", "first"), PART_DESCRIPTION=("PART_DESCRIPTION", "first"),
        WO_REV=("WO_REV", "first"), CURRENT_REV=("CURRENT_REV", "first"),
        CREATE_DATE=("CREATE_DATE", "first"), DESIRED_RLS_DATE=("DESIRED_RLS_DATE", "first"),
        N_OPS=("OPERATION", "size"), N_NO_OWN_MASTER_LINE=("NO_MASTER_LINE", "sum"),
        N_L1_DIFF=("L1_DIFF", "sum"))
    cats = pd.crosstab([l3[c] for c in WO_KEY], l3.CATEGORY).add_prefix("N_L3_")
    diffs = l3.groupby(WO_KEY, dropna=False).D3_ANY.sum().rename("N_L3_DIFF")
    return base.join(cats).join(diffs).fillna(0).reset_index()


REVIEW_KEYS = ["FAMILY", "PROJECT_CODE", "WO_KIND", "OPERATION", "FIELD", "WO_VALUE", "MASTER_VALUE",
               "ON_CURRENT_REV"]
REVIEW_LINE_COLUMNS = [
    "PATTERN", "HINT", "FAMILY", "PROJECT_CODE", "WO_KIND", "WORK_ORDER", "PART_ID", "PART_DESCRIPTION",
    "WO_STATUS", "CREATE_DATE", "DESIRED_RLS_DATE", "OPERATION", "SEQ", "FIELD", "WO_VALUE", "MASTER_VALUE",
    "DIRECTION", "ON_CURRENT_REV", "VERDICT", "NOTE"]
REVIEW_PATTERN_COLUMNS = [
    "PATTERN", "FAMILY", "PROJECT_CODE", "WO_KIND", "OPERATION", "SEQS", "FIELD", "WO_VALUE", "MASTER_VALUE",
    "DIRECTION", "N_LINES", "N_WOS", "N_PARTS", "ON_CURRENT_REV", "WO_STATUSES", "CREATED_FROM", "CREATED_TO",
    "HINT", "WORK_ORDERS", "PARTS"]


def review(l1):
    """Level 1 prepared for a person to judge. Returns (lines, patterns).

    lines:    one row per work order x part x differing field - a line that
              differs in several fields appears once per field. VERDICT and
              NOTE are left empty; this is the one place to fill them in.
    patterns: the same rows grouped by family x project x WO kind x operation
              x field x (WO value, master value) x revision state, numbered
              by PATTERN (largest first) - an overview, no verdict columns.
    """
    rows = []
    for r in l1.itertuples(index=False):
        for f, (_, s) in FIELDS.items():
            if getattr(r, f"D_{s}") != 1:
                continue
            wo, ms = getattr(r, f"WO_{f}"), getattr(r, f"M_{f}")
            rows.append({
                "FAMILY": r.FAMILY, "PROJECT_CODE": r.PROJECT_CODE, "WO_KIND": r.WO_KIND,
                "WORK_ORDER": r.WORK_ORDER, "PART_ID": r.PART_ID, "PART_DESCRIPTION": r.PART_DESCRIPTION,
                "WO_STATUS": r.WO_STATUS, "CREATE_DATE": pd.to_datetime(r.CREATE_DATE).date(),
                "DESIRED_RLS_DATE": pd.to_datetime(r.DESIRED_RLS_DATE).date(),
                "OPERATION": r.OPERATION, "SEQ": r.OPERATION_SEQUENCE, "FIELD": f,
                "WO_VALUE": code(wo) or "(empty)", "MASTER_VALUE": code(ms) or "(empty)",
                "WO_NUM": pd.to_numeric(wo, errors="coerce"), "M_NUM": pd.to_numeric(ms, errors="coerce"),
                "ON_CURRENT_REV": r.ON_CURRENT_REV,
            })
    if not rows:
        return pd.DataFrame(columns=REVIEW_LINE_COLUMNS), pd.DataFrame(columns=REVIEW_PATTERN_COLUMNS)
    x = pd.DataFrame(rows)
    x["DIRECTION"] = ""
    x.loc[x.M_NUM > x.WO_NUM, "DIRECTION"] = "master higher"
    x.loc[x.M_NUM < x.WO_NUM, "DIRECTION"] = "master lower"
    x["HINT"] = (x.FIELD == "RESOURCE_ID").map({True: "different machine planned", False: ""})
    # group once, number the groups, and hang the number on every line -
    # no join back on the keys, so empty keys (a missing project) cannot
    # lose their pattern
    x["_G"] = x.groupby(REVIEW_KEYS, dropna=False).ngroup()

    def joined(s):
        return ", ".join(map(str, sorted(set(s))))

    g = x.groupby("_G").agg(
        **{k: (k, "first") for k in REVIEW_KEYS + ["DIRECTION", "HINT"]},
        N_LINES=("WORK_ORDER", "size"), N_WOS=("WORK_ORDER", "nunique"), N_PARTS=("PART_ID", "nunique"),
        SEQS=("SEQ", joined), WO_STATUSES=("WO_STATUS", joined),
        CREATED_FROM=("CREATE_DATE", "min"), CREATED_TO=("CREATE_DATE", "max"),
        WORK_ORDERS=("WORK_ORDER", joined), PARTS=("PART_ID", joined),
    )
    g = g.sort_values(["N_LINES", "FAMILY", "OPERATION"], ascending=[False, True, True])
    g["PATTERN"] = range(1, len(g) + 1)
    x["PATTERN"] = x["_G"].map(g["PATTERN"])
    x["VERDICT"] = ""   # drift / intentional / wrong master / ?
    x["NOTE"] = ""
    lines = x.sort_values(["PATTERN", "WORK_ORDER", "SEQ"])[REVIEW_LINE_COLUMNS]
    return lines, g.reset_index(drop=True)[REVIEW_PATTERN_COLUMNS]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--drift", type=Path, default=PRIVATE_DIR / "data" / "drift.xlsx")
    ap.add_argument("--masters", type=Path, default=PRIVATE_DIR / "data" / "masters.xlsx")
    ap.add_argument("--out", type=Path, help="default: <repo>_private\\outputs\\master_drift_list_<today>.xlsx")
    a = ap.parse_args()

    for p in (a.drift, a.masters):
        if not p.exists():
            sys.exit(f"input not found: {p}")
    out = a.out or PRIVATE_DIR / "outputs" / f"master_drift_list_{datetime.now():%Y-%m-%d}.xlsx"
    if REPO_DIR in out.resolve().parents:
        sys.exit(f"refusing to write {out} inside the repo - the list names real parts and orders")
    out.parent.mkdir(parents=True, exist_ok=True)
    drift_day = datetime.fromtimestamp(a.drift.stat().st_mtime).date()
    masters_day = datetime.fromtimestamp(a.masters.stat().st_mtime).date()

    d = normalise_drift(load(a.drift, DRIFT_COLUMNS))
    m = normalise_masters(load(a.masters, MASTER_COLUMNS))

    dup = int(d.duplicated(LINE_KEY).sum())
    if dup:
        sys.exit(f"{a.drift.name}: {dup} duplicate operation keys - a join fanned out")
    flag_problems = check_level1_flags(d)
    if flag_problems:
        sys.exit(f"level 1 flags disagree with the recomputation: {flag_problems}")

    flags = [f"D_{s}" for _, s in FIELDS.values()]
    matched = d[d.NO_MASTER_LINE == 0]
    l1 = matched[matched[flags].fillna(0).max(axis=1) == 1]
    nomaster = d[d.NO_MASTER_LINE == 1].assign(
        HINT=lambda x: x.OPERATION.str.startswith("S RIL").map({True: "rework step (S RIL ...)", False: ""}))
    l3 = level3(d, m)
    l2 = level2(d, l3)
    rv_lines, rv_patterns = review(l1)
    # counted straight from the flags, independently of review()
    n_field_diffs = int((l1[flags].fillna(0) == 1).sum().sum())
    if not (len(rv_lines) == int(rv_patterns.N_LINES.sum()) == n_field_diffs) or rv_lines.PATTERN.isna().any():
        sys.exit(f"review does not cover every field difference: {len(rv_lines)} lines, "
                 f"{int(rv_patterns.N_LINES.sum())} in patterns, {n_field_diffs} flagged")

    not_extracted = int((l3.CATEGORY == "MASTERS_NOT_EXTRACTED").sum())
    not_in_export = int((l3.CATEGORY == "OWN_LINE_NOT_IN_EXPORT").sum())
    summary = [
        ("Run", f"{datetime.now():%Y-%m-%d %H:%M}"),
        ("Drift input", f"{a.drift.name} ({datetime.fromtimestamp(a.drift.stat().st_mtime):%Y-%m-%d %H:%M})"),
        ("Masters input", f"{a.masters.name} ({datetime.fromtimestamp(a.masters.stat().st_mtime):%Y-%m-%d %H:%M})"),
        ("Open, not-started WO operations", len(d)),
        ("Work orders", d[WO_KEY].drop_duplicates().shape[0]),
        ("Check: duplicate operation keys", dup),
        ("Check: level 1 flags recomputed", "all agree"),
        ("LEVEL 1 - lines differing from own revision", len(l1)),
        ("LEVEL 1 - work orders with a difference", l1[WO_KEY].drop_duplicates().shape[0]),
        ("LEVEL 1 - field differences", n_field_diffs),
        ("LEVEL 1 - review patterns (06_ReviewPatterns)", len(rv_patterns)),
        ("Lines with no own-revision master line", len(nomaster)),
        ("LEVEL 2 - work orders on an older revision", len(l2)),
        ("LEVEL 3 - older-revision lines", len(l3)),
    ]
    summary += [(f"LEVEL 3 - {k}", int(v)) for k, v in l3.CATEGORY.value_counts().items()]
    summary += [
        ("LEVEL 3 - compared lines differing from current revision", int(l3.D3_ANY.fillna(0).sum())),
        ("Check: level 3 parts missing from masters export", not_extracted),
        ("Check: matched own lines missing from masters export", not_in_export),
        ("LEVEL 3 NOTE", "INDICATIVE. Pairing assumes same name = same step. "
                         "To be reviewed again (agreed 2026-09-28)."),
    ]

    with pd.ExcelWriter(out) as xw:
        pd.DataFrame(summary, columns=["ITEM", "VALUE"]).to_excel(xw, sheet_name="00_Summary", index=False)
        l1.to_excel(xw, sheet_name="01_Drift", index=False)
        nomaster.to_excel(xw, sheet_name="02_NoMasterLine", index=False)
        l2.to_excel(xw, sheet_name="03_OldRevisionWOs", index=False)
        l3.to_excel(xw, sheet_name="04_CrossRevision", index=False)
        rv_lines.to_excel(xw, sheet_name="05_Review", index=False)
        rv_patterns.to_excel(xw, sheet_name="06_ReviewPatterns", index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        for sheet, widths in {
            "05_Review": {"A": 8, "B": 24, "C": 22, "D": 11, "E": 8, "F": 14, "G": 16, "H": 30, "I": 9,
                          "J": 11, "K": 11, "L": 24, "M": 7, "N": 14, "O": 10, "P": 12, "Q": 13, "R": 8,
                          "S": 14, "T": 30},
            "06_ReviewPatterns": {"A": 8, "B": 22, "C": 11, "D": 8, "E": 22, "F": 10, "G": 14, "H": 10,
                                  "I": 12, "J": 13, "R": 24},
        }.items():
            for col, w in widths.items():
                xw.sheets[sheet].column_dimensions[col].width = w

    for k, v in summary:
        print(f"{k:<58} {v}")
    print(f"\nwritten: {out}")
    status = 0
    if drift_day != masters_day:
        print(f"!! inputs saved on different days ({drift_day} / {masters_day}) - "
              "masters change; rerun both queries together")
        status = 2
    if not_extracted or not_in_export:
        print("!! masters export does not cover every older-revision line - "
              "check both queries ran on the same day with the same scope")
        status = 2
    return status


if __name__ == "__main__":
    sys.exit(main())
