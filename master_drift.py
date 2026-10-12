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
  00_Summary          counts, checks, input files, verdicts carried
  01_Drift            LEVEL 1: WO copy differs from its OWN revision's master
  02_NoMasterLine     WO lines with no own-revision master line
  03_OldRevisionWOs   LEVEL 2: one row per WO built on an older revision
  04_CrossRevision    LEVEL 3 (INDICATIVE): older-revision WO lines against
                      the CURRENT revision's master
  05_Review           LEVEL 1 for a person to judge: one row per work order
                      x part x differing field, with its PATTERN number and
                      the VERDICT / NOTE columns - fill in here
  06_ReviewPatterns   the same grouped into numbered patterns - overview only
  07_StaleUser8       work orders whose USER_8 names an older revision while
                      their routing is the current one (see below)
  08_VerdictHistory   every line ever given a verdict, carried from list to
                      list - the memory behind SUGGESTED_VERDICT

STALE USER_8 (user's review, 2026-09-30). USER_8 is a label: an order can
carry the current revision's routing while USER_8 still names an older one.
An older-revision order is treated as being on the current revision when
every one of its not-started lines matches the current master on sequence,
operation name and every compared value, and they do not all match its own
revision. Its lines are then compared with the current master, and it is
listed on 07_StaleUser8. Only the not-started lines are visible here, so
"every line" means every line in this export.

VERDICTS. In 05_Review the VERDICT column takes exactly one of VERDICTS
(chosen with the user, 2026-09-28/30). New rows are prefilled with all of
them separated by " / " - delete down to one; a value still containing
" / " counts as undecided. The next run reads the most recently saved
review in the outputs folder (or --verdicts) and carries VERDICT and NOTE
forward to the same work order, sequence, operation, field and the same
two values. When a difference changes, it is judged again.
An open line whose difference (operation, field, both values) was judged
before on any order gets SUGGESTED_VERDICT and SUGGESTION_BASIS from the
verdict history (user, 2026-10-12) - a suggestion only, VERDICT stays open.
INTENTIONAL_CHANGE does not hide a line: it stays listed while the order
and the master differ (user, 2026-10-12) - the master may be the one to
change.

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

# MASTER_DRIFT        the master changed after the order was created (improved)
# ORIG_MASTER_WRONG   the master was already wrong when the order was created
#                     and has been corrected since - same action as MASTER_DRIFT
# INTENTIONAL_CHANGE  the order was changed on purpose (e.g. by supply, the PM)
# WO_WRONG            the order was edited by mistake; the master did not change
# MASTER_WRONG        the order is right; the master is wrong today
# UNCLEAR             asked, no clear answer yet
VERDICTS = ["MASTER_DRIFT", "ORIG_MASTER_WRONG", "INTENTIONAL_CHANGE", "WO_WRONG", "MASTER_WRONG", "UNCLEAR"]
VERDICT_OPTIONS = " / ".join(VERDICTS)
CARRY_KEY = ["WORK_ORDER", "SEQ", "OPERATION", "FIELD", "WO_VALUE", "MASTER_VALUE"]
REVIEW_SHEETS = ("05_Review", "Lines")   # "Lines": the hand-built review of 28 Sep

DRIFT_COLUMNS = (
    LINE_KEY
    + ["WO_KIND", "WO_STATUS", "PROJECT_CODE", "FAMILY", "PART_ID", "PART_DESCRIPTION",
       "WO_REV", "CURRENT_REV", "ON_CURRENT_REV", "CREATE_DATE", "DESIRED_RLS_DATE",
       "OPERATION", "OP_STATUS", "MASTER_OPS_ROWID", "NO_MASTER_LINE"]
    + [f"WO_{f}" for f in FIELDS] + [f"M_{f}" for f in FIELDS]
    + [f"D_{s}" for _, s in FIELDS.values()]
)
# MASTER_OPS_ROWID is optional in masters.xlsx: exports before 2026-10-12 lack it
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
    if "MASTER_OPS_ROWID" not in m:
        m["MASTER_OPS_ROWID"] = pd.NA
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


def apply_stale_user8(d, m):
    """Move older-revision WOs that carry the current routing onto it.

    See STALE USER_8 in the module docstring. Returns (d, stale), where d
    gains USER_8_STALE and ON_CURRENT_ROUTING, and stale lists the WOs.
    """
    d["USER_8_STALE"] = 0
    d["ON_CURRENT_ROUTING"] = d.ON_CURRENT_REV.fillna(0).astype(int)
    cur = m[(m.ROLE == "CUR") & m.SEQUENCE_NO.notna()]
    cur_cols = ["SEQUENCE_NO", "OPERATION", "MASTER_OPS_ROWID"] + list(FIELDS)
    flags = [f"D_{s}" for _, s in FIELDS.values()]
    stale = []
    for _, w in d[d.ON_CURRENT_ROUTING == 0].groupby(WO_KEY, dropna=False):
        c = cur[(cur.PART_ID == w.PART_ID.iloc[0]) & (cur.REV == w.CURRENT_REV.iloc[0])][cur_cols]
        if c.empty or c.duplicated(["SEQUENCE_NO", "OPERATION"]).any():
            continue
        own_all = bool(((w.NO_MASTER_LINE == 0) & (w[flags].fillna(0).max(axis=1) == 0)).all())
        if own_all:
            continue    # consistent with its own revision - the label may well be right
        j = w.reset_index().merge(c.add_prefix("CUR_"), how="left",
                                  left_on=["OPERATION_SEQUENCE", "OPERATION"],
                                  right_on=["CUR_SEQUENCE_NO", "CUR_OPERATION"]).set_index("index")
        if len(j) != len(w) or j.CUR_SEQUENCE_NO.isna().any():
            continue
        ok = pd.Series(True, index=j.index)
        for f, (text, _) in FIELDS.items():
            ok &= same(j[f"WO_{f}"], j[f"CUR_{f}"], text, NULL_AS.get(f))
        if not ok.all():
            continue
        # pandas 3 refuses to upcast silently: numbers go in as floats (NaN for blank)
        for f, (text, _) in FIELDS.items():
            v = j[f"CUR_{f}"]
            d.loc[j.index, f"M_{f}"] = v if text else pd.to_numeric(v, errors="coerce").astype(float)
        d.loc[j.index, "MASTER_OPS_ROWID"] = pd.to_numeric(j["CUR_MASTER_OPS_ROWID"], errors="coerce").astype(float)
        d.loc[j.index, "NO_MASTER_LINE"] = 0
        d.loc[j.index, flags] = 0
        d.loc[j.index, ["USER_8_STALE", "ON_CURRENT_ROUTING"]] = 1
        first = w.iloc[0]
        stale.append({**{k: first[k] for k in WO_KEY}, "PART_ID": first.PART_ID,
                      "PART_DESCRIPTION": first.PART_DESCRIPTION, "WO_KIND": first.WO_KIND,
                      "USER_8": first.WO_REV, "CURRENT_REV": first.CURRENT_REV,
                      "N_LINES": len(w), "LINES_NO_OWN_MASTER_LINE_BEFORE": int(w.NO_MASTER_LINE.sum()),
                      "LINES_DIFFERING_BEFORE": int((w[flags].fillna(0).max(axis=1) == 1).sum())})
    return d, pd.DataFrame(stale)


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
    old = d[d.ON_CURRENT_ROUTING == 0].copy()
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
    if "CATEGORY" not in res:
        res["CATEGORY"] = pd.NA
    old = old.join(res)
    compared = old.CUR_SEQUENCE_NO.notna()
    for f, (text, s) in FIELDS.items():
        diff = (~same(old[f"WO_{f}"], old[f"CUR_{f}"], text, NULL_AS.get(f))).astype("Int64")
        old[f"D3_{s}"] = diff.where(compared)
    old["D3_ANY"] = old[[f"D3_{s}" for _, s in FIELDS.values()]].max(axis=1).where(compared)
    return old


def level2(d, l3):
    old = d[d.ON_CURRENT_ROUTING == 0]
    d_any = old[[f"D_{s}" for _, s in FIELDS.values()]].fillna(0).max(axis=1)
    base = old.assign(L1_DIFF=d_any).groupby(WO_KEY, dropna=False).agg(
        WO_KIND=("WO_KIND", "first"), WO_STATUS=("WO_STATUS", "first"),
        PROJECT_CODE=("PROJECT_CODE", "first"), FAMILY=("FAMILY", "first"),
        PART_ID=("PART_ID", "first"), PART_DESCRIPTION=("PART_DESCRIPTION", "first"),
        WO_REV=("WO_REV", "first"), CURRENT_REV=("CURRENT_REV", "first"),
        CREATE_DATE=("CREATE_DATE", "first"), DESIRED_RLS_DATE=("DESIRED_RLS_DATE", "first"),
        N_OPS=("OPERATION", "size"), N_NO_OWN_MASTER_LINE=("NO_MASTER_LINE", "sum"),
        N_L1_DIFF=("L1_DIFF", "sum"))
    if l3.empty:
        return base.reset_index()
    cats = pd.crosstab([l3[c] for c in WO_KEY], l3.CATEGORY).add_prefix("N_L3_")
    diffs = l3.groupby(WO_KEY, dropna=False).D3_ANY.sum().rename("N_L3_DIFF")
    return base.join(cats).join(diffs).fillna(0).reset_index()


REVIEW_KEYS = ["FAMILY", "PROJECT_CODE", "WO_KIND", "OPERATION", "FIELD", "WO_VALUE", "MASTER_VALUE",
               "ON_CURRENT_ROUTING"]
REVIEW_LINE_COLUMNS = [
    "PATTERN", "HINT", "FAMILY", "PROJECT_CODE", "WO_KIND", "WORK_ORDER", "PART_ID", "PART_DESCRIPTION",
    "WO_STATUS", "CREATE_DATE", "DESIRED_RLS_DATE", "OPERATION", "SEQ", "FIELD", "WO_VALUE", "MASTER_VALUE",
    "CUR_MASTER_VALUE", "DIRECTION", "ON_CURRENT_ROUTING", "VERDICT", "SUGGESTED_VERDICT", "SUGGESTION_BASIS",
    "NOTE", "VERDICT_FROM"]
REVIEW_PATTERN_COLUMNS = [
    "PATTERN", "FAMILY", "PROJECT_CODE", "WO_KIND", "OPERATION", "SEQS", "FIELD", "WO_VALUE", "MASTER_VALUE",
    "DIRECTION", "N_LINES", "N_WOS", "N_PARTS", "ON_CURRENT_ROUTING", "WO_STATUSES", "CREATED_FROM", "CREATED_TO",
    "HINT", "VERDICTS", "WORK_ORDERS", "PARTS"]


def seq_key(v):
    """Sequence as text for matching: 2200, 2200.0 and '2200' are the same."""
    n = pd.to_numeric(v, errors="coerce")
    return str(int(n)) if pd.notna(n) else code(v)


def review(l1, l3):
    """Level 1 prepared for a person to judge. Returns (lines, patterns).

    lines:    one row per work order x part x differing field - a line that
              differs in several fields appears once per field. For an
              older-revision order CUR_MASTER_VALUE shows the current master's
              value where level 3 could pair the line.
    patterns: the same rows grouped by family x project x WO kind x operation
              x field x (WO value, master value) x routing state, numbered by
              PATTERN (largest first) - an overview, no verdict columns.
    VERDICT / NOTE / VERDICT_FROM are filled in by carry_verdicts().
    """
    rows = []
    for idx, r in zip(l1.index, l1.itertuples(index=False)):
        for f, (_, s) in FIELDS.items():
            if getattr(r, f"D_{s}") != 1:
                continue
            wo, ms = getattr(r, f"WO_{f}"), getattr(r, f"M_{f}")
            cur_value, hint = "", []
            if f == "RESOURCE_ID":
                hint.append("different machine planned")
            if idx in l3.index and pd.notna(l3.at[idx, "CUR_SEQUENCE_NO"]):
                cur_value = code(l3.at[idx, f"CUR_{f}"]) or "(empty)"
                if l3.at[idx, f"D3_{s}"] == 0:
                    hint.append("order already matches the current master")
            rows.append({
                "FAMILY": r.FAMILY, "PROJECT_CODE": r.PROJECT_CODE, "WO_KIND": r.WO_KIND,
                "WORK_ORDER": r.WORK_ORDER, "PART_ID": r.PART_ID, "PART_DESCRIPTION": r.PART_DESCRIPTION,
                "WO_STATUS": r.WO_STATUS, "CREATE_DATE": pd.to_datetime(r.CREATE_DATE).date(),
                "DESIRED_RLS_DATE": pd.to_datetime(r.DESIRED_RLS_DATE).date(),
                "OPERATION": r.OPERATION, "SEQ": r.OPERATION_SEQUENCE, "FIELD": f,
                "WO_VALUE": code(wo) or "(empty)", "MASTER_VALUE": code(ms) or "(empty)",
                "CUR_MASTER_VALUE": cur_value, "HINT": "; ".join(hint),
                "WO_NUM": pd.to_numeric(wo, errors="coerce"), "M_NUM": pd.to_numeric(ms, errors="coerce"),
                "ON_CURRENT_ROUTING": r.ON_CURRENT_ROUTING,
            })
    if not rows:
        return pd.DataFrame(columns=REVIEW_LINE_COLUMNS), pd.DataFrame(columns=REVIEW_PATTERN_COLUMNS)
    x = pd.DataFrame(rows)
    x["DIRECTION"] = ""
    x.loc[x.M_NUM > x.WO_NUM, "DIRECTION"] = "master higher"
    x.loc[x.M_NUM < x.WO_NUM, "DIRECTION"] = "master lower"
    # group once, number the groups, and hang the number on every line -
    # no join back on the keys, so empty keys (a missing project) cannot
    # lose their pattern
    x["_G"] = x.groupby(REVIEW_KEYS, dropna=False).ngroup()

    def joined(s):
        return ", ".join(map(str, sorted(set(s))))

    g = x.groupby("_G").agg(
        **{k: (k, "first") for k in REVIEW_KEYS + ["DIRECTION"]},
        HINT=("HINT", lambda s: "; ".join(sorted({h for v in s for h in v.split("; ") if h}))),
        N_LINES=("WORK_ORDER", "size"), N_WOS=("WORK_ORDER", "nunique"), N_PARTS=("PART_ID", "nunique"),
        SEQS=("SEQ", joined), WO_STATUSES=("WO_STATUS", joined),
        CREATED_FROM=("CREATE_DATE", "min"), CREATED_TO=("CREATE_DATE", "max"),
        WORK_ORDERS=("WORK_ORDER", joined), PARTS=("PART_ID", joined),
    )
    g = g.sort_values(["N_LINES", "FAMILY", "OPERATION"], ascending=[False, True, True])
    g["PATTERN"] = range(1, len(g) + 1)
    x["PATTERN"] = x["_G"].map(g["PATTERN"])
    for c in ["VERDICT", "SUGGESTED_VERDICT", "SUGGESTION_BASIS", "NOTE", "VERDICT_FROM"]:
        x[c] = ""
    lines = x.sort_values(["PATTERN", "WORK_ORDER", "SEQ"])[REVIEW_LINE_COLUMNS]
    g["VERDICTS"] = ""
    return lines, g.reset_index(drop=True)[REVIEW_PATTERN_COLUMNS]


def find_previous_review(folder, exclude):
    """Most recently saved workbook in folder with a review sheet, or (None, None)."""
    files = [p for p in folder.glob("*.xlsx")
             if not p.name.startswith("~$") and p.resolve() != exclude.resolve()]
    for p in sorted(files, key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            sheets = pd.ExcelFile(p).sheet_names
        except Exception:
            continue
        for s in REVIEW_SHEETS:
            if s in sheets:
                cols = set(pd.read_excel(p, sheet_name=s, nrows=0).columns)
                if set(CARRY_KEY) | {"VERDICT"} <= cols:
                    return p, s
    return None, None


HISTORY_SHEET = "08_VerdictHistory"
HISTORY_COLUMNS = CARRY_KEY + ["VERDICT", "NOTE", "JUDGED_IN"]


def line_key(df):
    """The carry key: work order, sequence, operation, field and both values."""
    return pd.Series(list(zip(
        df.WORK_ORDER.map(code), df.SEQ.map(seq_key), df.OPERATION.map(code), df.FIELD.map(code),
        df.WO_VALUE.map(code), df.MASTER_VALUE.map(code))), index=df.index)


def pattern_key(df):
    """The suggestion key: the same difference on any order - operation, field, both values."""
    return list(zip(df.OPERATION.map(code), df.FIELD.map(code), df.WO_VALUE.map(code), df.MASTER_VALUE.map(code)))


def carry_verdicts(lines, path, sheet):
    """Fill VERDICT / NOTE / VERDICT_FROM / SUGGESTED_VERDICT from earlier reviews.

    VERDICT is carried when the same line (line_key) was decided in the
    previous review. The verdict history - every decided line ever, carried
    from list to list on HISTORY_SHEET, newest judgement first - suggests a
    verdict for an open line when the same difference (pattern_key) was
    judged before on any order. A suggestion is never written into VERDICT.
    Returns (stats, history).
    """
    stats = {"file": "none found", "previous_decided": 0, "carried": 0, "previous_unknown": [],
             "previous_decided_gone": 0, "history": 0, "suggested": 0}
    lines["VERDICT"] = VERDICT_OPTIONS
    lines["SUGGESTED_VERDICT"] = ""
    lines["SUGGESTION_BASIS"] = ""
    history = pd.DataFrame(columns=HISTORY_COLUMNS)
    if path is None:
        return stats, history
    stats["file"] = f"{path.name} [{sheet}] ({datetime.fromtimestamp(path.stat().st_mtime):%Y-%m-%d %H:%M})"
    prev = pd.read_excel(path, sheet_name=sheet, dtype=str)
    if "NOTE" not in prev:
        prev["NOTE"] = pd.NA

    prev["MATCH_KEY"] = line_key(prev)
    prev["V"] = prev.VERDICT.map(lambda v: (code(v) or "").strip())
    decided = prev.V.isin(VERDICTS)
    undecided = (prev.V == "") | prev.V.str.contains("/", regex=False)
    stats["previous_decided"] = int(decided.sum())
    stats["previous_unknown"] = sorted(set(prev.loc[~decided & ~undecided, "V"]))

    # history: this review's decisions first (they are the newest), then the
    # history that review carried; the first judgement of a line wins
    judged = prev[decided].assign(VERDICT=prev.V, JUDGED_IN=path.name)
    parts = [judged.reindex(columns=HISTORY_COLUMNS)]
    if HISTORY_SHEET in pd.ExcelFile(path).sheet_names:
        parts.append(pd.read_excel(path, sheet_name=HISTORY_SHEET, dtype=str).reindex(columns=HISTORY_COLUMNS))
    history = pd.concat(parts, ignore_index=True)
    history = history[history.VERDICT.isin(VERDICTS)]
    history = history[~line_key(history).duplicated()].reset_index(drop=True)
    stats["history"] = len(history)
    if lines.empty:
        return stats, history

    prev = prev.drop_duplicates("MATCH_KEY")
    # a plain dict: tuple keys in a pandas index invite multi-level lookups
    by_key = {r.MATCH_KEY: r for r in prev.itertuples(index=False)}
    k = line_key(lines)
    for i, kk in k.items():
        p = by_key.get(kk)
        if p is None:
            continue
        note = code(p.NOTE) or ""
        if p.V in VERDICTS:
            lines.at[i, "VERDICT"] = p.V
            lines.at[i, "VERDICT_FROM"] = path.name
            stats["carried"] += 1
        elif p.V and "/" not in p.V:
            # free text in VERDICT: keep it, visibly, so nothing written is lost
            note = f"previous VERDICT text: {p.V}" + (f" | {note}" if note else "")
        lines.at[i, "NOTE"] = note
    current = set(k)
    stats["previous_decided_gone"] = int(sum(1 for kk in prev.loc[prev.V.isin(VERDICTS), "MATCH_KEY"]
                                             if kk not in current))

    by_pattern = {}
    for pk, v, note in zip(pattern_key(history), history.VERDICT, history.NOTE):
        by_pattern.setdefault(pk, []).append((v, code(note) or ""))
    for i, pk in zip(lines.index, pattern_key(lines)):
        if lines.at[i, "VERDICT"] in VERDICTS or pk not in by_pattern:
            continue
        seen = by_pattern[pk]
        counts = pd.Series([v for v, _ in seen]).value_counts()
        lines.at[i, "SUGGESTED_VERDICT"] = (counts.index[0] if len(counts) == 1 else
                                            "MIXED: " + ", ".join(f"{v} x{n}" for v, n in counts.items()))
        example = next((n for _, n in seen if n), "")
        lines.at[i, "SUGGESTION_BASIS"] = (f"same difference judged {len(seen)}x before"
                                           + (f"; note: {example}" if example else ""))
        stats["suggested"] += 1
    return stats, history


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--drift", type=Path, default=PRIVATE_DIR / "data" / "drift.xlsx")
    ap.add_argument("--masters", type=Path, default=PRIVATE_DIR / "data" / "masters.xlsx")
    ap.add_argument("--out", type=Path, help="default: <repo>_private\\outputs\\master_drift_list_<today>.xlsx")
    ap.add_argument("--verdicts", type=Path,
                    help="review to carry verdicts from (default: the most recently saved one in outputs)")
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
    if a.verdicts:
        if not a.verdicts.exists():
            sys.exit(f"verdicts file not found: {a.verdicts}")
        sheet = next((s for s in REVIEW_SHEETS if s in pd.ExcelFile(a.verdicts).sheet_names), None)
        if sheet is None:
            sys.exit(f"{a.verdicts.name}: no sheet named {' or '.join(REVIEW_SHEETS)}")
        verdict_src = (a.verdicts, sheet)
    else:
        verdict_src = find_previous_review(PRIVATE_DIR / "outputs", out)

    d = normalise_drift(load(a.drift, DRIFT_COLUMNS))
    m = normalise_masters(load(a.masters, MASTER_COLUMNS))

    dup = int(d.duplicated(LINE_KEY).sum())
    if dup:
        sys.exit(f"{a.drift.name}: {dup} duplicate operation keys - a join fanned out")
    # against the SQL's own flags, BEFORE the stale-USER_8 rule rewrites any
    flag_problems = check_level1_flags(d)
    if flag_problems:
        sys.exit(f"level 1 flags disagree with the recomputation: {flag_problems}")
    d, stale = apply_stale_user8(d, m)

    flags = [f"D_{s}" for _, s in FIELDS.values()]
    matched = d[d.NO_MASTER_LINE == 0]
    l1 = matched[matched[flags].fillna(0).max(axis=1) == 1]
    nomaster = d[d.NO_MASTER_LINE == 1].assign(
        HINT=lambda x: x.OPERATION.str.startswith("S RIL").map({True: "rework step (S RIL ...)", False: ""}))
    l3 = level3(d, m)
    l2 = level2(d, l3)
    rv_lines, rv_patterns = review(l1, l3)
    # counted straight from the flags, independently of review()
    n_field_diffs = int((l1[flags].fillna(0) == 1).sum().sum())
    if not (len(rv_lines) == int(rv_patterns.N_LINES.sum()) == n_field_diffs) or rv_lines.PATTERN.isna().any():
        sys.exit(f"review does not cover every field difference: {len(rv_lines)} lines, "
                 f"{int(rv_patterns.N_LINES.sum())} in patterns, {n_field_diffs} flagged")
    vs, history = carry_verdicts(rv_lines, *verdict_src)
    # this run's decided lines join the history too: none are new here (they
    # were all carried), but keep the history complete if that ever changes
    now = rv_lines[rv_lines.VERDICT.isin(VERDICTS)].assign(
        JUDGED_IN=lambda x: x.VERDICT_FROM.where(x.VERDICT_FROM != "", out.name))
    history = pd.concat([history, now.reindex(columns=HISTORY_COLUMNS)], ignore_index=True)
    history = history[~line_key(history).duplicated()].reset_index(drop=True)
    if not rv_patterns.empty:
        rv_patterns["VERDICTS"] = rv_patterns.PATTERN.map(
            rv_lines.assign(V=rv_lines.VERDICT.where(rv_lines.VERDICT.isin(VERDICTS), "(open)"))
            .groupby("PATTERN").V.agg(lambda s: ", ".join(f"{k} x{n}" for k, n in s.value_counts().items())))

    not_extracted = int((l3.CATEGORY == "MASTERS_NOT_EXTRACTED").sum()) if not l3.empty else 0
    not_in_export = int((l3.CATEGORY == "OWN_LINE_NOT_IN_EXPORT").sum()) if not l3.empty else 0
    summary = [
        ("Run", f"{datetime.now():%Y-%m-%d %H:%M}"),
        ("Drift input", f"{a.drift.name} ({datetime.fromtimestamp(a.drift.stat().st_mtime):%Y-%m-%d %H:%M})"),
        ("Masters input", f"{a.masters.name} ({datetime.fromtimestamp(a.masters.stat().st_mtime):%Y-%m-%d %H:%M})"),
        ("Open, not-started WO operations", len(d)),
        ("Work orders", d[WO_KEY].drop_duplicates().shape[0]),
        ("Check: duplicate operation keys", dup),
        ("Check: level 1 flags recomputed", "all agree"),
        ("Stale USER_8 - WOs moved to the current routing (07_StaleUser8)", len(stale)),
        ("LEVEL 1 - lines differing from their master", len(l1)),
        ("LEVEL 1 - work orders with a difference", l1[WO_KEY].drop_duplicates().shape[0]),
        ("LEVEL 1 - field differences", n_field_diffs),
        ("LEVEL 1 - review patterns (06_ReviewPatterns)", len(rv_patterns)),
        ("Lines with no own-revision master line", len(nomaster)),
        ("LEVEL 2 - work orders on an older revision", len(l2)),
        ("LEVEL 3 - older-revision lines", len(l3)),
    ]
    summary += [(f"LEVEL 3 - {k}", int(v)) for k, v in l3.CATEGORY.value_counts().items()] if not l3.empty else []
    summary += [
        ("LEVEL 3 - compared lines differing from current revision",
         int(l3.D3_ANY.fillna(0).sum()) if not l3.empty else 0),
        ("Check: level 3 parts missing from masters export", not_extracted),
        ("Check: matched own lines missing from masters export", not_in_export),
        ("LEVEL 3 NOTE", "INDICATIVE. Pairing assumes same name = same step. "
                         "To be reviewed again (agreed 2026-09-28)."),
        ("Verdicts carried from", vs["file"]),
        ("Verdicts - decided in that review", vs["previous_decided"]),
        ("Verdicts - carried to a current difference", vs["carried"]),
        ("Verdicts - decided before, difference gone or changed", vs["previous_decided_gone"]),
        ("Verdicts - still to decide in 05_Review", int((~rv_lines.VERDICT.isin(VERDICTS)).sum())),
        ("Verdicts - of those, with a SUGGESTED_VERDICT", vs["suggested"]),
        (f"Verdicts - history ({HISTORY_SHEET}), every line ever decided", len(history)),
    ]
    summary += [(f"Verdicts - {v}", int((rv_lines.VERDICT == v).sum())) for v in VERDICTS]

    with pd.ExcelWriter(out) as xw:
        pd.DataFrame(summary, columns=["ITEM", "VALUE"]).to_excel(xw, sheet_name="00_Summary", index=False)
        l1.to_excel(xw, sheet_name="01_Drift", index=False)
        nomaster.to_excel(xw, sheet_name="02_NoMasterLine", index=False)
        l2.to_excel(xw, sheet_name="03_OldRevisionWOs", index=False)
        l3.to_excel(xw, sheet_name="04_CrossRevision", index=False)
        rv_lines.to_excel(xw, sheet_name="05_Review", index=False)
        rv_patterns.to_excel(xw, sheet_name="06_ReviewPatterns", index=False)
        stale.to_excel(xw, sheet_name="07_StaleUser8", index=False)
        history.to_excel(xw, sheet_name=HISTORY_SHEET, index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        for sheet, widths in {
            "00_Summary": {"A": 62, "B": 60},
            "05_Review": {"A": 8, "B": 30, "C": 22, "D": 11, "E": 8, "F": 14, "G": 16, "H": 30, "I": 9,
                          "J": 11, "K": 11, "L": 24, "M": 7, "N": 14, "O": 10, "P": 12, "Q": 12, "R": 13,
                          "S": 8, "T": 40, "U": 22, "V": 50, "W": 30, "X": 24},
            HISTORY_SHEET: {"A": 14, "B": 7, "C": 24, "D": 14, "E": 10, "F": 12, "G": 22, "H": 50, "I": 30},
            "06_ReviewPatterns": {"A": 8, "B": 22, "C": 11, "D": 8, "E": 22, "F": 10, "G": 14, "H": 10,
                                  "I": 12, "J": 13, "R": 30, "S": 30},
        }.items():
            for col, w in widths.items():
                xw.sheets[sheet].column_dimensions[col].width = w

    for k, v in summary:
        print(f"{k:<66} {v}")
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
    if vs["previous_unknown"]:
        print(f"!! the previous review has VERDICT text that is not a keyword: {vs['previous_unknown']} - "
              f"kept in NOTE, the row is open again. Keywords: {', '.join(VERDICTS)}")
        status = 2
    return status


if __name__ == "__main__":
    sys.exit(main())
