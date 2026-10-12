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
  00_Summary          counts, checks, one line per LABEL with its verdicts
  01_OrdersToUpdate   same resource, other values differ: the order should
                      follow the master
  02_ResourceCheck    the resource differs: usually intentional, always looked at
  03_Information      order already matches today's master
                      (01-03: one row per work-order line, LABEL naming
                      exactly what differs; rules in THE PLANNERS' LISTS below)
  04_Drift            LEVEL 1 detail: WO copy differs from its master
  05_NoMasterLine     WO lines with no own-revision master line
  06_OldRevisionWOs   LEVEL 2: one row per WO built on an older revision
  07_CrossRevision    LEVEL 3 (INDICATIVE): older-revision WO lines against
                      the CURRENT revision's master
  08_StaleUser8       work orders whose USER_8 names an older revision while
                      their routing is the current one (see below)
  09_VerdictHistory   every line ever given a verdict, carried from list to
                      list - the memory behind SUGGESTED_VERDICT

STALE USER_8 (user's review, 2026-09-30). USER_8 is a label: an order can
carry the current revision's routing while USER_8 still names an older one.
An older-revision order is treated as being on the current revision when
every one of its not-started lines matches the current master on sequence,
operation name and every compared value, and they do not all match its own
revision. Its lines are then compared with the current master, and it is
listed on 08_StaleUser8. Only the not-started lines are visible here, so
"every line" means every line in this export.

VERDICTS. On 01-03 the VERDICT column is empty by default and is filled
only for exceptions, with exactly one of VERDICTS (chosen with the user,
2026-09-28/30); empty or a value containing " / " counts as open. The next
run reads the most recently saved list in the outputs folder (or
--verdicts) and carries VERDICT and NOTE forward to the same work order,
sequence, operation and the same differences (DIFF_KEY). When a difference
changes, it is open again. An open line whose operation and differences
were judged before on any order gets SUGGESTED_VERDICT and
SUGGESTION_BASIS from the verdict history - a suggestion only.
INTENTIONAL_CHANGE does not hide a line: it stays listed while the order
and the master differ (user, 2026-10-12) - the master may be the one to
change. Lists and reviews made before 2026-10-12 (one row per field) are
read too: their verdicts are combined per line.

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


# ---------------------------------------------------------------------------
# THE PLANNERS' LISTS (agreed with the user, 2026-10-12)
#
# One row per work-order line that differs from its master (level 1). Each
# line gets a LABEL naming exactly what differs, and goes to one list, rules
# checked in this order:
#   1. Order already matches today's master -> 03_Information
#      (older-revision line, level 3 paired it and every value equals the
#       current master: the difference is only with its own old revision)
#   2. Same resource, other values differ  -> 01_OrdersToUpdate
#      (28 Sep review: 51 of 51 decided lines - the order follows the master)
#   3. Different resource                  -> 02_ResourceCheck
#      (28 Sep review: 9 of 13 intentional, 4 not - always looked at)
# ---------------------------------------------------------------------------

# label names, in label order; the resource decides "same" / "different"
NAMES = {"RUN": "rate", "SETUP_HRS": "setup", "RUN_TYPE": "run type", "LOAD_SIZE_QTY": "load size",
         "USER_5": "crew"}
LISTS = {  # key -> (sheet, name shown in the LIST column)
    "update": ("01_OrdersToUpdate", "Orders to update"),
    "check": ("02_ResourceCheck", "Resource changed - check"),
    "info": ("03_Information", "Order already matches today's master"),
}
HISTORY_SHEET = "09_VerdictHistory"
OLD_HISTORY_SHEET = "08_VerdictHistory"        # field-level, lists of 2026-10-12 morning
OLD_REVIEW_SHEETS = ("05_Review", "Lines")     # field-level; "Lines" = hand-built review of 28 Sep
LIST_COLUMNS = [
    "LIST", "LABEL", "FAMILY", "PROJECT_CODE", "WO_KIND", "WORK_ORDER", "PART_ID", "PART_DESCRIPTION",
    "WO_STATUS", "CREATE_DATE", "DESIRED_RLS_DATE", "OPERATION", "SEQ", "ON_CURRENT_ROUTING",
    "DIFFERENCES", "TODAY_MASTER", "VERDICT", "SUGGESTED_VERDICT", "SUGGESTION_BASIS", "NOTE",
    "VERDICT_FROM", "DIFF_KEY"]
HISTORY_COLUMNS = ["WORK_ORDER", "SEQ", "OPERATION", "LABEL", "DIFF_KEY", "VERDICT", "NOTE", "JUDGED_IN"]


def seq_key(v):
    """Sequence as text for matching: 2200, 2200.0 and '2200' are the same."""
    n = pd.to_numeric(v, errors="coerce")
    return str(int(n)) if pd.notna(n) else code(v)


def shown(v):
    return code(v) or "(empty)"


def make_diff_key(pairs):
    """Machine-readable difference: 'RUN=60|50; SETUP_HRS=0|1' (order|master), in FIELDS order."""
    return "; ".join(f"{f}={pairs[f][0]}|{pairs[f][1]}" for f in FIELDS if f in pairs)


def label_of(fields):
    others = [n for f, n in NAMES.items() if f in fields]
    if "RESOURCE_ID" in fields:
        return "different resource" + (": " + " + ".join(others) if others else ", nothing else")
    return "same resource: " + " + ".join(others)


def label_of_key(diff_key):
    return label_of({p.split("=")[0] for p in str(diff_key).split("; ") if p})


def build_lists(l1, l3):
    """Level 1 lines with LIST, LABEL, DIFFERENCES and TODAY_MASTER. See the rules above."""
    rows = []
    for idx, r in zip(l1.index, l1.itertuples(index=False)):
        diff = [f for f, (_, s) in FIELDS.items() if getattr(r, f"D_{s}") == 1]
        pairs = {f: (shown(getattr(r, f"WO_{f}")), shown(getattr(r, f"M_{f}"))) for f in diff}
        compared = idx in l3.index and pd.notna(l3.at[idx, "CUR_SEQUENCE_NO"])
        today = ""
        if compared:
            today = "; ".join(f"{NAMES.get(f, 'resource')}: {shown(l3.at[idx, f'CUR_{f}'])}" for f in diff)
        if compared and l3.at[idx, "D3_ANY"] == 0:
            key = "info"
        elif "RESOURCE_ID" in diff:
            key = "check"
        else:
            key = "update"
        rows.append({
            "_LIST": key, "LIST": LISTS[key][1], "LABEL": label_of(diff),
            "FAMILY": r.FAMILY, "PROJECT_CODE": r.PROJECT_CODE, "WO_KIND": r.WO_KIND,
            "WORK_ORDER": r.WORK_ORDER, "PART_ID": r.PART_ID, "PART_DESCRIPTION": r.PART_DESCRIPTION,
            "WO_STATUS": r.WO_STATUS, "CREATE_DATE": pd.to_datetime(r.CREATE_DATE).date(),
            "DESIRED_RLS_DATE": pd.to_datetime(r.DESIRED_RLS_DATE).date(),
            "OPERATION": r.OPERATION, "SEQ": r.OPERATION_SEQUENCE, "ON_CURRENT_ROUTING": r.ON_CURRENT_ROUTING,
            "DIFFERENCES": "; ".join(f"{NAMES.get(f, 'resource')}: order {a} / master {b}"
                                     for f, (a, b) in pairs.items()),
            "TODAY_MASTER": today, "DIFF_KEY": make_diff_key(pairs),
        })
    cols = ["_LIST"] + LIST_COLUMNS
    if not rows:
        return pd.DataFrame(columns=cols)
    x = pd.DataFrame(rows)
    for c in ["VERDICT", "SUGGESTED_VERDICT", "SUGGESTION_BASIS", "NOTE", "VERDICT_FROM"]:
        x[c] = ""
    return x.sort_values(["_LIST", "LABEL", "OPERATION", "WORK_ORDER", "SEQ"])[cols].reset_index(drop=True)


def verdict_state(v):
    """'decided' / 'open' / 'unknown' for a VERDICT cell."""
    v = (code(v) or "").strip()
    if v in VERDICTS:
        return "decided"
    if v == "" or "/" in v:
        return "open"
    return "unknown"


def fields_to_lines(df, by_extra=()):
    """Convert a field-level review (one row per differing field) to one row per line."""
    df = df.copy()
    if "NOTE" not in df:
        df["NOTE"] = pd.NA
    df["_W"], df["_S"], df["_O"] = df.WORK_ORDER.map(code), df.SEQ.map(seq_key), df.OPERATION.map(code)
    out = []
    for keys, g in df.groupby(["_W", "_S", "_O", *by_extra], dropna=False):
        pairs = {code(f): (code(a) or "(empty)", code(b) or "(empty)")
                 for f, a, b in zip(g.FIELD, g.WO_VALUE, g.MASTER_VALUE)}
        texts = sorted({(code(v) or "").strip() for v in g.VERDICT})
        decided = [t for t in texts if t in VERDICTS]
        # one verdict for the line only if its fields agree; otherwise it is open
        verdict = texts[0] if len(texts) == 1 else " / ".join(t for t in texts if t)
        notes = sorted({code(n) for n in g.NOTE if code(n)})
        rec = {"WORK_ORDER": keys[0], "SEQ": keys[1], "OPERATION": keys[2], "DIFF_KEY": make_diff_key(pairs),
               "VERDICT": verdict, "NOTE": " | ".join(notes), "_CONFLICT": len(decided) > 1}
        for name, val in zip(by_extra, keys[3:]):
            rec[name] = val
        out.append(rec)
    return pd.DataFrame(out)


def read_previous(path):
    """(previous lines, previous history), both one row per line, from any of our formats."""
    sheets = pd.ExcelFile(path).sheet_names
    new_fmt = [s for s, _ in LISTS.values() if s in sheets]
    if new_fmt:
        prev = pd.concat([pd.read_excel(path, sheet_name=s, dtype=str) for s in new_fmt], ignore_index=True)
        prev["_CONFLICT"] = False
    else:
        sheet = next(s for s in OLD_REVIEW_SHEETS if s in sheets)
        prev = fields_to_lines(pd.read_excel(path, sheet_name=sheet, dtype=str))
    if HISTORY_SHEET in sheets:
        hist = pd.read_excel(path, sheet_name=HISTORY_SHEET, dtype=str)
    elif OLD_HISTORY_SHEET in sheets:
        hist = fields_to_lines(pd.read_excel(path, sheet_name=OLD_HISTORY_SHEET, dtype=str), ("JUDGED_IN",))
    else:
        hist = pd.DataFrame(columns=HISTORY_COLUMNS)
    return prev, hist


def find_previous(folder, exclude):
    """Most recently saved workbook in folder holding lists or a review, or None."""
    wanted = {s for s, _ in LISTS.values()} | set(OLD_REVIEW_SHEETS)
    files = [p for p in folder.glob("*.xlsx")
             if not p.name.startswith("~$") and p.resolve() != exclude.resolve()]
    for p in sorted(files, key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            if wanted & set(pd.ExcelFile(p).sheet_names):
                return p
        except Exception:
            continue
    return None


def line_key(df):
    return list(zip(df.WORK_ORDER.map(code), df.SEQ.map(seq_key), df.OPERATION.map(code), df.DIFF_KEY.map(code)))


def carry_verdicts(lines, path, stale_wos=()):
    """Carry VERDICT / NOTE from the previous lists, keep the history, suggest verdicts.

    Carry: same work order, sequence, operation and the same differences
    (DIFF_KEY). History: every line ever decided, newest judgement first -
    except lines of work orders now known to have a stale USER_8: those were
    judged against the wrong revision, so they are no experience at all.
    Suggestion: an open line whose operation and differences were judged
    before on any order. A suggestion never fills VERDICT. Returns (stats, history).
    """
    stats = {"file": "none found", "previous_decided": 0, "carried": 0, "previous_unknown": [],
             "previous_decided_gone": 0, "suggested": 0, "conflicts": 0, "history_stale_dropped": 0}
    history = pd.DataFrame(columns=HISTORY_COLUMNS)
    if path is None:
        return stats, history
    stats["file"] = f"{path.name} ({datetime.fromtimestamp(path.stat().st_mtime):%Y-%m-%d %H:%M})"
    prev, hist = read_previous(path)
    state = prev.VERDICT.map(verdict_state)
    stats["previous_decided"] = int((state == "decided").sum())
    stats["previous_unknown"] = sorted({(code(v) or "").strip() for v in prev.VERDICT[state == "unknown"]})
    stats["conflicts"] = int(prev._CONFLICT.fillna(False).astype(bool).sum())

    judged = prev[state == "decided"].assign(VERDICT=lambda x: x.VERDICT.map(lambda v: code(v).strip()),
                                             JUDGED_IN=path.name)
    history = pd.concat([judged.reindex(columns=HISTORY_COLUMNS), hist.reindex(columns=HISTORY_COLUMNS)],
                        ignore_index=True)
    history = history[history.VERDICT.map(verdict_state) == "decided"]
    history = history[~pd.Series(line_key(history), index=history.index).duplicated()].reset_index(drop=True)
    on_stale = history.WORK_ORDER.map(code).isin({code(w) for w in stale_wos})
    stats["history_stale_dropped"] = int(on_stale.sum())
    history = history[~on_stale].reset_index(drop=True)
    history["LABEL"] = history.DIFF_KEY.map(label_of_key)

    by_key = {k: (v, n) for k, v, n in zip(line_key(prev), prev.VERDICT, prev.NOTE.fillna(""))}
    for i, k in zip(lines.index, line_key(lines)):
        if k not in by_key:
            continue
        v, note = by_key[k]
        st, note = verdict_state(v), code(note) or ""
        if st == "decided":
            lines.at[i, "VERDICT"] = code(v).strip()
            lines.at[i, "VERDICT_FROM"] = path.name
            stats["carried"] += 1
        elif st == "unknown":
            # free text in VERDICT: keep it, visibly, so nothing written is lost
            note = f"previous VERDICT text: {code(v).strip()}" + (f" | {note}" if note else "")
        lines.at[i, "NOTE"] = note
    current = set(line_key(lines))
    stats["previous_decided_gone"] = sum(1 for k, s in zip(line_key(prev), state) if s == "decided" and k not in current)

    by_diff = {}
    for op, dk, v, n in zip(history.OPERATION.map(code), history.DIFF_KEY.map(code), history.VERDICT,
                            history.NOTE.fillna("")):
        by_diff.setdefault((op, dk), []).append((v, code(n) or ""))
    for i, op, dk in zip(lines.index, lines.OPERATION.map(code), lines.DIFF_KEY.map(code)):
        if lines.at[i, "VERDICT"] in VERDICTS or (op, dk) not in by_diff:
            continue
        seen = by_diff[(op, dk)]
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
                    help="list to carry verdicts from (default: the most recently saved one in outputs)")
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
    if a.verdicts and not a.verdicts.exists():
        sys.exit(f"verdicts file not found: {a.verdicts}")
    verdict_src = a.verdicts or find_previous(PRIVATE_DIR / "outputs", out)

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
    lines = build_lists(l1, l3)
    if len(lines) != len(l1) or lines.duplicated(["WORK_ORDER", "SEQ"]).any():
        sys.exit(f"the lists do not hold each level 1 line exactly once: {len(lines)} rows for {len(l1)} lines")
    vs, history = carry_verdicts(lines, verdict_src,
                                 stale.WORK_ORDER.tolist() if "WORK_ORDER" in stale else [])
    # this run's decided lines are all carried, hence already in the history;
    # added anyway so the history stays complete if that ever changes
    now = lines[lines.VERDICT.isin(VERDICTS)].assign(
        JUDGED_IN=lambda x: x.VERDICT_FROM.where(x.VERDICT_FROM != "", out.name))
    history = pd.concat([history, now.reindex(columns=HISTORY_COLUMNS)], ignore_index=True)
    history = history[~pd.Series(line_key(history), index=history.index).duplicated()].reset_index(drop=True)

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
        ("Stale USER_8 - WOs moved to the current routing (08_StaleUser8)", len(stale)),
        ("Lines differing from their master", len(lines)),
    ]
    for key, (sheet, name) in LISTS.items():
        part = lines[lines._LIST == key]
        summary.append((f"{sheet} - {name}", f"{len(part)} lines, {part.WORK_ORDER.nunique()} work orders"))
    summary.append(("", ""))
    for lab, g in lines.groupby("LABEL"):
        judged = history[history.LABEL == lab].VERDICT.value_counts()
        summary.append((f"Label - {lab}",
                        f"{len(g)} lines now ({', '.join(f'{LISTS[k][1]} {n}' for k, n in g._LIST.value_counts().items())})"
                        f" | judged before: " + (", ".join(f"{v} x{n}" for v, n in judged.items()) or "never")))
    for lab in sorted(set(history.LABEL) - set(lines.LABEL)):
        judged = history[history.LABEL == lab].VERDICT.value_counts()
        summary.append((f"Label - {lab}", "0 lines now | judged before: "
                        + ", ".join(f"{v} x{n}" for v, n in judged.items())))
    summary += [
        ("", ""),
        ("Verdicts carried from", vs["file"]),
        ("Verdicts - decided there", vs["previous_decided"]),
        ("Verdicts - carried to a current line", vs["carried"]),
        ("Verdicts - decided before, line gone or changed", vs["previous_decided_gone"]),
        ("Verdicts - open lines with a SUGGESTED_VERDICT", vs["suggested"]),
        (f"Verdicts - history ({HISTORY_SHEET}), every line ever decided", len(history)),
        ("Verdicts - dropped from history: judged on a stale-USER_8 comparison", vs["history_stale_dropped"]),
    ]
    summary += [(f"Verdicts - {v}", int((lines.VERDICT == v).sum())) for v in VERDICTS]
    summary += [
        ("", ""),
        ("Lines with no own-revision master line (05_NoMasterLine)", len(nomaster)),
        ("LEVEL 2 - work orders on an older revision (06_OldRevisionWOs)", len(l2)),
        ("LEVEL 3 - older-revision lines (07_CrossRevision)", len(l3)),
    ]
    summary += [(f"LEVEL 3 - {k}", int(v)) for k, v in l3.CATEGORY.value_counts().items()] if not l3.empty else []
    summary += [
        ("LEVEL 3 - compared lines differing from current revision",
         int(l3.D3_ANY.fillna(0).sum()) if not l3.empty else 0),
        ("Check: level 3 parts missing from masters export", not_extracted),
        ("Check: matched own lines missing from masters export", not_in_export),
        ("LEVEL 3 NOTE", "INDICATIVE. Pairing assumes same name = same step. "
                         "To be reviewed again (agreed 2026-09-28)."),
    ]

    with pd.ExcelWriter(out) as xw:
        pd.DataFrame(summary, columns=["ITEM", "VALUE"]).to_excel(xw, sheet_name="00_Summary", index=False)
        for key, (sheet, _) in LISTS.items():
            lines[lines._LIST == key][LIST_COLUMNS].to_excel(xw, sheet_name=sheet, index=False)
        l1.to_excel(xw, sheet_name="04_Drift", index=False)
        nomaster.to_excel(xw, sheet_name="05_NoMasterLine", index=False)
        l2.to_excel(xw, sheet_name="06_OldRevisionWOs", index=False)
        l3.to_excel(xw, sheet_name="07_CrossRevision", index=False)
        stale.to_excel(xw, sheet_name="08_StaleUser8", index=False)
        history.reindex(columns=HISTORY_COLUMNS).to_excel(xw, sheet_name=HISTORY_SHEET, index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        list_widths = {"A": 24, "B": 34, "C": 22, "D": 11, "E": 8, "F": 14, "G": 16, "H": 30, "I": 9, "J": 11,
                       "K": 11, "L": 24, "M": 7, "N": 8, "O": 50, "P": 30, "Q": 20, "R": 22, "S": 50, "T": 40,
                       "U": 24, "V": 30}
        for sheet, widths in {
            "00_Summary": {"A": 70, "B": 110},
            **{s: list_widths for s, _ in LISTS.values()},
            HISTORY_SHEET: {"A": 14, "B": 7, "C": 24, "D": 34, "E": 30, "F": 20, "G": 50, "H": 30},
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
        print(f"!! the previous list has VERDICT text that is not a keyword: {vs['previous_unknown']} - "
              f"kept in NOTE, the line is open again. Keywords: {', '.join(VERDICTS)}")
        status = 2
    if vs["conflicts"]:
        print(f"!! {vs['conflicts']} line(s) had different verdicts on their fields in the previous review - "
              "left open, to decide once for the line")
        status = 2
    return status


if __name__ == "__main__":
    sys.exit(main())
