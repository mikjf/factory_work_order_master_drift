"""Master drift - build the planners' workbook from the two SQL queries.

    python master_drift.py               query the ERP, then build the list
    python master_drift.py --no-extract  build again from the last extract

Inputs: the results of
  master_drift.sql          level 1, one row per open, not-started WO operation
  old_revision_masters.sql  master lines of the own and the current revision,
                            for parts with older-revision WOs
By default extract.py runs both against the ERP (read-only, Windows login,
settings in <repo>_private\\extract.ini) and saves them as Parquet in
<repo>_private\\data\\ - together, so they always come from the same moment
(since 2026-10-15). The manual route still works: --drift and --masters take
exports pasted into Excel, as used until then; the script warns if the two
were saved on different days, or if their resource IDs lost leading zeros.

Output: <repo>_private\\outputs\\master_drift_list_<date>.xlsx by default,
never inside this repo - it names real parts and orders:
  00_Summary          counts per scenario, checks, since last list
  01_SqlUpdater       THE list: one row per line that does not match its
                      master - the ROWIDs, part, order, step, revisions, the
                      order's (WO_) / master's (M_) / current master's (C_)
                      values, SCENARIO, SQL_UPDATE (an UPDATE on the order's
                      operation, "Work order to update" lines only) -
                      see THE DRIFT LIST and THE SQL UPDATER below
  02_SinceLastList    what happened to each line of the previous list that is
                      not listed again unchanged: fixed (order / master
                      updated), started on the old values, or the step
                      replaced - see fates_since_last() (user, 2026-10-12)
  03_OldRevisionWOs   LEVEL 2: one row per WO built on an older revision
  04_CrossRevision    LEVEL 3 (INDICATIVE): older-revision WO lines against
                      the CURRENT revision's master
  05_StaleUser8       work orders whose USER_8 names an older revision while
                      their routing is the current one (see below)

STALE USER_8 (user's review, 2026-09-30). USER_8 is a label: an order can
carry the current revision's routing while USER_8 still names an older one.
An older-revision order is treated as being on the current revision when
every one of its not-started lines matches the current master on sequence,
operation name and every compared value, and they do not all match its own
revision. Its lines are then compared with the current master, and it is
listed on 05_StaleUser8. Only the not-started lines are visible here, so
"every line" means every line in this export.

PREVIOUS LIST. The next run reads the most recently saved list in the
outputs folder (or --previous) and fills 02_SinceLastList from it, matching
lines on work order, step, operation and the same differences (DIFF_KEY).
Lists of every earlier layout are still read. Verdict, LABEL and NOTE
columns existed until 2026-10-12; the user then chose rules alone.

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
import hashlib
import json
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

REPO_DIR = Path(__file__).resolve().parent
PRIVATE_DIR = REPO_DIR.parent / f"{REPO_DIR.name}_private"
# the queries extract.py runs, in step with extract.QUERIES (checked when it
# runs); named here too so --no-extract works without pyodbc
EXTRACT_QUERIES = ("master_drift.sql", "old_revision_masters.sql")


def query_sha256(name):
    """SHA-256 of a query file, as extract.py records it in its manifest."""
    return hashlib.sha256((REPO_DIR / name).read_text(encoding="utf-8-sig").encode("utf-8")).hexdigest()

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
    """An input: the Parquet of extract.py, or an Excel export of the manual route."""
    df = pd.read_parquet(path) if path.suffix == ".parquet" else pd.read_excel(path, sheet_name=0)
    missing = [c for c in required if c not in df.columns]
    if missing:
        sys.exit(f"{path.name}: missing columns {missing}")
    return df


# Resource IDs: read straight from the database (extract.py, since 2026-10-15)
# they arrive as text, '034' included. Excel exports of 13-14 Oct 2026 carry
# them as 'RES:034' - pasted into Excel, '034' had become the number 34 - and
# exactly that marker is stripped. Older Excel exports have neither: their
# leading zeros are lost, and the run says so.
RESOURCE_MARK = "RES:"


def unmark_resources(s, where, warnings):
    vals = s.dropna().astype(str)
    marked = vals.str.startswith(RESOURCE_MARK)
    if marked.all() and len(vals):
        return s.map(lambda v: v if pd.isna(v) else str(v)[len(RESOURCE_MARK):])
    if marked.any():
        sys.exit(f"{where}: {int((~marked).sum())} of {len(vals)} resource IDs lack the '{RESOURCE_MARK}' marker - "
                 "an export mixing two versions of the query")
    if not pd.api.types.is_string_dtype(s) and len(vals):
        warnings.append(f"{where}: resource IDs read as numbers - an Excel export without the marker, "
                        "leading zeros are lost (034 shows as 34); run without --drift/--masters")
    return s


# Numbers compared as numbers. Excel turned every numeric-looking text into a
# number; from the database a text column such as USER_5 (crew) arrives as
# text. Both are made numbers here - a value that is not a number stops the run.
NUMERIC = [f for f, (text, _) in FIELDS.items() if not text]


def to_number(s, where):
    n = pd.to_numeric(s, errors="coerce")
    bad = s[n.isna() & s.notna() & (s.astype(str).str.strip() != "")]
    if len(bad):
        sys.exit(f"{where}: {len(bad)} values that are not numbers (e.g. {bad.iloc[0]!r})")
    return n.astype(float)


def normalise_drift(d, name, warnings):
    # WORKORDER_OPS_ROWID is optional: exports before 2026-10-12 lack it
    if "WORKORDER_OPS_ROWID" not in d:
        d["WORKORDER_OPS_ROWID"] = pd.NA
    for c in ["WO_RESOURCE_ID", "M_RESOURCE_ID"]:
        d[c] = unmark_resources(d[c], f"{name} {c}", warnings)
    for c in ["WO_REV", "CURRENT_REV", "OPERATION", "WO_KIND", "WO_STATUS", "OP_STATUS", "PART_ID"]:
        d[c] = d[c].map(code)
    for f, (text, _) in FIELDS.items():
        for c in (f"WO_{f}", f"M_{f}"):
            d[c] = d[c].map(code) if text else to_number(d[c], f"{name} {c}")
    return d


def normalise_masters(m, name, warnings):
    m["RESOURCE_ID"] = unmark_resources(m["RESOURCE_ID"], f"{name} RESOURCE_ID", warnings)
    for c in ["PART_ID", "ROLE", "REV", "OPERATION", "RESOURCE_ID", "RUN_TYPE"]:
        m[c] = m[c].map(code)
    for f in NUMERIC:
        m[f] = to_number(m[f], f"{name} {f}")
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
                    rec["CUR_MASTER_OPS_ROWID"] = cline.MASTER_OPS_ROWID
                    for f in FIELDS:
                        rec[f"CUR_{f}"] = cline[f]
        out.append(rec)

    res = pd.DataFrame(out, index=old.index)
    for f in FIELDS:
        if f"CUR_{f}" not in res:
            res[f"CUR_{f}"] = pd.NA
    for c in ("CUR_SEQUENCE_NO", "CUR_MASTER_OPS_ROWID"):
        if c not in res:
            res[c] = pd.NA
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
# THE DRIFT LIST (agreed with the user, 2026-10-12)
#
# One row per open, not-started work-order line that does NOT simply match
# its master. Each row has exactly one SCENARIO, decided in this order:
#   No master line            no own-revision master line to compare with
#                             (a step added to the order: rework, test samples)
#   Work order matches current master
#                             an older-revision line that differs from its own
#                             old revision, while every value equals the
#                             current master (level 3 paired it)
#   Resource changed          the resource differs (28 Sep review: 9 of 13
#                             intentional, 4 not - always looked at)
#   Work order to update      same resource, another value differs (28 Sep
#                             review: 51 of 51 - the order follows the master)
#   Differs from current master (WO on older revision)
#                             matches its own old revision, but level 3 found
#                             a different value in the current master
#                             (indicative: cross-revision pairing)
# Lines that match their master are counted, not listed. DIFF_KEY names
# exactly what differs, each value tagged (WO) / (M) / (C).
#
# THE SQL UPDATER (user, 2026-10-12). The list is laid out as the user asked,
# FTS-updater style: per line the order's (WO_), its master's (M_) and the
# current master's (C_) values side by side, and SQL_UPDATE - a spreadsheet FORMULA
# that builds an UPDATE on the work-order operation (WO_OPS_ROWID) from the M_
# cells of its own row, so correcting a value corrects the statement.
#   - only on "Work order to update": the other scenarios need a person
#     (resource), are indicative (current master), or have nothing to copy;
#   - the value is M_, the master the line was compared with (exact pairing by
#     sequence + name), never C_ on an older revision (indicative pairing);
#   - it sets only the fields that differ; never RESOURCE_ID (a "Work order to
#     update" line has the same resource by definition);
#   - empty when a master value or the ROWID is blank, so nothing can blank a
#     field by accident. Nothing here runs anything.
# The cells of each difference in DIFF_KEY are coloured (user, 2026-10-13):
# the order's value (WO_) orange, the value it is compared with green - M_, or
# C_ on "Differs from current master". "Work order to update" rows are light
# grey across the row underneath. Fills are fixed at writing time; they do
# not follow edits made in the sheet.
# M_REV is the revision the line was compared with (= C_REV for an order moved
# by the stale-USER_8 rule). For orders on the current revision C_ = M_.
# ---------------------------------------------------------------------------

SCENARIOS = {  # key -> shown name; dict order = sort order of the list
    "update": "Work order to update",
    "check": "Resource changed",
    "today": "Differs from current master (WO on older revision)",
    "info": "Work order matches current master",
    "nomaster": "No master line",
}
LIST_SHEET = "01_SqlUpdater"
# older workbooks still read for 02_SinceLastList
OLD_DRIFT_LIST_SHEET = "01_DriftList"                                           # 2026-10-12, before the updater
OLD_LIST_SHEETS = ("01_OrdersToUpdate", "02_ResourceCheck", "03_Information")   # 2026-10-12 afternoon
OLD_REVIEW_SHEETS = ("05_Review", "Lines")                                      # one row per field
# compared field -> its short name in the list's column headers (user, 2026-10-12)
SHORT = {"RESOURCE_ID": "RES_ID", "RUN_TYPE": "RUN_TYPE", "RUN": "RUN", "LOAD_SIZE_QTY": "LSQ",
         "SETUP_HRS": "SETUP_HRS", "USER_5": "CREW"}
# the user's column order: resource, then rate and run type paired per source
# (rate first - user, 2026-10-12), then load size, setup and crew
VALUE_COLUMNS = (
    ["WO_RES_ID", "M_RES_ID", "C_RES_ID",
     "WO_RUN", "WO_RUN_TYPE", "M_RUN", "M_RUN_TYPE", "C_RUN", "C_RUN_TYPE"]
    + [f"{p}_{s}" for s in ("LSQ", "SETUP_HRS", "CREW") for p in ("WO", "M", "C")])
LIST_COLUMNS = (
    ["WO_OPS_ROWID", "M_OPS_ROWID", "C_MASTER_OPS_ROWID", "FAMILY", "PART_ID", "PART_DESCRIPTION",
     "WO", "WO_STATUS", "WO_CREATE_DATE", "WO_DESIRED_RLS_DATE", "OPERATION", "STEP",
     "WO_REV", "M_REV", "C_REV"]
    + VALUE_COLUMNS
    + ["DIFF_KEY", "SCENARIO", "SQL_UPDATE"])


def seq_key(v):
    """Sequence as text for matching: 2200, 2200.0 and '2200' are the same."""
    n = pd.to_numeric(v, errors="coerce")
    return str(int(n)) if pd.notna(n) else code(v)


def shown(v):
    return code(v) or "(empty)"


def make_diff_key(pairs, tag="M"):
    """What differs, readable and machine-readable at once (user, 2026-10-12):
    'RUN: 60 (WO) | 70 (M); SETUP_HRS: 0 (WO) | 1 (M)' - short column names, in
    FIELDS order, each value tagged with its source: WO the order, M its master,
    C the current master (tag="C", the "differs from current master" scenario).
    """
    return "; ".join(f"{SHORT[f]}: {pairs[f][0]} (WO) | {pairs[f][1]} ({tag})" for f in FIELDS if f in pairs)


_TAGGED = re.compile(r"^(\w+): (.*) \(WO\) \| (.*) \((?:M|C)\)$")


def parse_diff_key(diff_key):
    """DIFF_KEY -> {'RUN': ('60', '70'), ...}, keyed by the ERP field name.

    Reads the tagged form above and the plain one used until 2026-10-12
    ('RUN=60|70; SETUP_HRS=0|1'), so lists of either kind still match.
    """
    long_name = {s: f for f, s in SHORT.items()}
    out = {}
    for part in str(diff_key or "").split("; "):
        m = _TAGGED.match(part.strip())
        if m:
            out[long_name.get(m.group(1), m.group(1))] = (m.group(2), m.group(3))
        elif "=" in part and "|" in part:
            f, vals = part.split("=", 1)
            o, mm = vals.split("|", 1)
            out[f.strip()] = (o, mm)
    return out


def canonical_diff(diff_key):
    """The comparison form of a DIFF_KEY: same differences, same text, whatever the format."""
    pairs = parse_diff_key(diff_key)
    return "; ".join(f"{f}={pairs[f][0]}|{pairs[f][1]}" for f in FIELDS if f in pairs)



def build_list(d, l3):
    """The drift list: one row per line that does not match its master. See above."""
    flags = {f: f"D_{s}" for f, (_, s) in FIELDS.items()}
    rows, matching = [], 0
    for idx, r in zip(d.index, d.itertuples(index=False)):
        compared = idx in l3.index and pd.notna(l3.at[idx, "CUR_SEQUENCE_NO"])
        own = [f for f in FIELDS if getattr(r, flags[f]) == 1] if r.NO_MASTER_LINE == 0 else []
        today = [f for f, (_, s) in FIELDS.items() if compared and l3.at[idx, f"D3_{s}"] == 1]
        if r.NO_MASTER_LINE == 1:
            key, pairs = "nomaster", {}
        elif own and compared and not today:
            key, pairs = "info", {f: (shown(getattr(r, f"WO_{f}")), shown(getattr(r, f"M_{f}"))) for f in own}
        elif own:
            key = "check" if "RESOURCE_ID" in own else "update"
            pairs = {f: (shown(getattr(r, f"WO_{f}")), shown(getattr(r, f"M_{f}"))) for f in own}
        elif today:
            key = "today"
            pairs = {f: (shown(getattr(r, f"WO_{f}")), shown(l3.at[idx, f"CUR_{f}"])) for f in today}
        else:
            matching += 1
            continue
        # current master: the line itself for an order on (or moved onto) the
        # current routing, the level 3 pairing for an older revision
        on_current = r.ON_CURRENT_ROUTING == 1
        no_line = r.NO_MASTER_LINE == 1
        rec = {
            "_S": list(SCENARIOS).index(key), "_DIFF": list(pairs),
            "WO_OPS_ROWID": r.WORKORDER_OPS_ROWID,
            "M_OPS_ROWID": pd.NA if no_line else r.MASTER_OPS_ROWID,
            "C_MASTER_OPS_ROWID": (pd.NA if no_line else r.MASTER_OPS_ROWID) if on_current else
                                  (l3.at[idx, "CUR_MASTER_OPS_ROWID"] if compared else pd.NA),
            "FAMILY": r.FAMILY, "PART_ID": r.PART_ID, "PART_DESCRIPTION": r.PART_DESCRIPTION,
            "WO": r.WORK_ORDER, "WO_STATUS": r.WO_STATUS,
            "WO_CREATE_DATE": pd.to_datetime(r.CREATE_DATE).date(),
            "WO_DESIRED_RLS_DATE": pd.to_datetime(r.DESIRED_RLS_DATE).date(),
            "OPERATION": r.OPERATION, "STEP": r.OPERATION_SEQUENCE,
            "WO_REV": r.WO_REV, "M_REV": r.CURRENT_REV if r.USER_8_STALE == 1 else r.WO_REV,
            "C_REV": r.CURRENT_REV,
            "DIFF_KEY": make_diff_key(pairs, "C" if key == "today" else "M"),
            "SCENARIO": SCENARIOS[key], "SQL_UPDATE": "",
        }
        for f, s in SHORT.items():
            rec[f"WO_{s}"] = getattr(r, f"WO_{f}")
            rec[f"M_{s}"] = pd.NA if no_line else getattr(r, f"M_{f}")
            rec[f"C_{s}"] = (rec[f"M_{s}"] if on_current else
                             l3.at[idx, f"CUR_{f}"] if compared else pd.NA)
        rows.append(rec)
    x = pd.DataFrame(rows, columns=["_S", "_DIFF"] + LIST_COLUMNS)
    return x.sort_values(["_S", "OPERATION", "WO", "STEP"]).reset_index(drop=True), matching


def add_sql_update(lst):
    """SQL_UPDATE on every 'Work order to update' line, as a live spreadsheet formula.

    Uses the row's position in the written sheet (header = row 1), so lst
    must be in its final order and column order. See THE SQL UPDATER above.
    Returns the number of statements written and of update lines left empty.
    """
    from openpyxl.utils import get_column_letter
    col = {c: get_column_letter(i + 1) for i, c in enumerate(LIST_COLUMNS)}
    written = blank = 0
    for n, (i, r) in enumerate(lst.iterrows(), start=2):
        if r.SCENARIO != SCENARIOS["update"]:
            continue
        fields = [f for f in r._DIFF if f != "RESOURCE_ID"]
        if (pd.isna(r.WO_OPS_ROWID) or not fields
                or any(pd.isna(r[f"M_{SHORT[f]}"]) or code(r[f"M_{SHORT[f]}"]) is None for f in fields)):
            blank += 1
            continue
        sets = ", ".join(f"{f} = '\"&{col['M_' + SHORT[f]]}{n}&\"'" for f in fields)
        lst.at[i, "SQL_UPDATE"] = f'="UPDATE OPERATION SET {sets} WHERE ROWID = "&{col["WO_OPS_ROWID"]}{n}&";"'
        written += 1
    return written, blank


FILL_WO, FILL_TARGET, FILL_UPDATE = "FCE4D6", "E2EFDA", "EDEDED"   # light orange, light green, light grey


def colour_differences(ws, lst):
    """Colour the two cells of every difference in DIFF_KEY. See THE SQL UPDATER above."""
    from openpyxl.styles import PatternFill
    fills = {p: PatternFill("solid", start_color=c, end_color=c)
             for p, c in (("WO", FILL_WO), ("T", FILL_TARGET), ("U", FILL_UPDATE))}
    pos = {c: i + 1 for i, c in enumerate(LIST_COLUMNS)}
    coloured = 0
    for n, (_, r) in enumerate(lst.iterrows(), start=2):
        if r.SCENARIO == SCENARIOS["update"]:   # whole row grey first, the differences on top
            for i in range(1, len(LIST_COLUMNS) + 1):
                ws.cell(n, i).fill = fills["U"]
        target = "C" if r.SCENARIO == SCENARIOS["today"] else "M"
        for f in r._DIFF:
            ws.cell(n, pos[f"WO_{SHORT[f]}"]).fill = fills["WO"]
            ws.cell(n, pos[f"{target}_{SHORT[f]}"]).fill = fills["T"]
            coloured += 1
    return coloured


def line_key(df, seq_col, wo_col="WORK_ORDER"):
    return list(zip(df[wo_col].map(code), df[seq_col].map(seq_key), df.OPERATION.map(code),
                    df.DIFF_KEY.map(canonical_diff)))


def fields_to_lines(df):
    """One-row-per-field review (before 2026-10-12) -> one row per line."""
    df = df.copy()
    df["_W"], df["_S"], df["_O"] = df.WORK_ORDER.map(code), df.SEQ.map(seq_key), df.OPERATION.map(code)
    out = []
    for (w, s, o), g in df.groupby(["_W", "_S", "_O"], dropna=False):
        pairs = {code(f): (code(a) or "(empty)", code(b) or "(empty)")
                 for f, a, b in zip(g.FIELD, g.WO_VALUE, g.MASTER_VALUE)}
        out.append({"WORK_ORDER": w, "SEQ": s, "OPERATION": o, "DIFF_KEY": make_diff_key(pairs),
                    "SCENARIO": ""})
    return pd.DataFrame(out)


def read_previous(path):
    """The previous list as WORK_ORDER / SEQ / OPERATION / DIFF_KEY / SCENARIO, any of our layouts."""
    sheets = pd.ExcelFile(path).sheet_names
    if LIST_SHEET in sheets:
        p = pd.read_excel(path, sheet_name=LIST_SHEET, dtype=str).rename(columns={"WO": "WORK_ORDER", "STEP": "SEQ"})
    elif OLD_DRIFT_LIST_SHEET in sheets:
        p = pd.read_excel(path, sheet_name=OLD_DRIFT_LIST_SHEET, dtype=str).rename(columns={"STEP": "SEQ"})
    elif any(s in sheets for s in OLD_LIST_SHEETS):
        p = pd.concat([pd.read_excel(path, sheet_name=s, dtype=str) for s in OLD_LIST_SHEETS if s in sheets],
                      ignore_index=True)
        p["SCENARIO"] = p.get("LIST", "")
    elif any(s in sheets for s in OLD_REVIEW_SHEETS):
        sheet = next(s for s in OLD_REVIEW_SHEETS if s in sheets)
        p = fields_to_lines(pd.read_excel(path, sheet_name=sheet, dtype=str))
    else:
        sys.exit(f"{path.name} holds no list this script can read (sheets: {', '.join(sheets)})")
    for c in ("SCENARIO", "DIFF_KEY"):
        if c not in p:
            p[c] = ""
    return p[["WORK_ORDER", "SEQ", "OPERATION", "DIFF_KEY", "SCENARIO"]].fillna("")


def find_previous(folder, out):
    """Most recently saved workbook in folder holding a list (any of our formats), or None.

    The output file itself does not count: a second run on the same day
    overwrites the first and still compares with the previous list, not with
    its own earlier run. (It counted while NOTE had to survive a rerun.)
    """
    wanted = {LIST_SHEET, OLD_DRIFT_LIST_SHEET, *OLD_LIST_SHEETS, *OLD_REVIEW_SHEETS}
    files = [p for p in folder.glob("*.xlsx")
             if not p.name.startswith("~$") and p.resolve() != out.resolve()]
    for p in sorted(files, key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            if wanted & set(pd.ExcelFile(p).sheet_names):
                return p
        except Exception:
            continue
    return None


def same_value(a, b, field):
    """One value against another as the comparison does: numbers as numbers, blank crew = 1."""
    def norm(v):
        v = None if v is None or (not isinstance(v, str) and pd.isna(v)) else v
        v = None if v in ("(empty)", "") else v
        if v is None and field in NULL_AS:
            v = NULL_AS[field]
        if v is None:
            return None
        n = pd.to_numeric(v, errors="coerce")
        return float(n) if pd.notna(n) and not FIELDS[field][0] else code(v)
    return norm(a) == norm(b)


def fates_since_last(prev, d, lst, stale_wos):
    """What happened to each line of the previous list that is not listed again unchanged.

    FATE, in this order:
      moved by USER_8 rule           its work order is now on 05_StaleUser8
      step replaced by another       the order still has the step open, but under
        operation                    another operation name - the line was
                                     removed or renamed, not fixed (user, 2026-10-13)
      started, closed or removed     the step is no longer an open, not-started
                                     operation - mostly produced (or closing) on
                                     the values it had; this export cannot tell a
                                     started step from a deleted one
      still listed, changed          listed again, with other differences
      fixed - order updated          the order now carries the old master value
      fixed - master updated         the master now carries the order's old value
      fixed - both changed           order and master agree on a third value
      fixed - now matches            no longer differs (a line without master line
                                     or compared with the current master before)
    """
    cols = ["FATE", "WO", "STEP", "OPERATION", "SCENARIO_BEFORE", "DIFF_KEY_BEFORE", "DIFF_KEY_NOW"]
    if prev is None or prev.empty:
        return pd.DataFrame(columns=cols)
    current = set(line_key(lst, "STEP", "WO"))
    # lines are matched on order, step AND operation name: a step can be
    # replaced by another operation under the same number
    listed_now = {(code(w), seq_key(s), code(o)) for w, s, o in zip(lst.WO, lst.STEP, lst.OPERATION)}
    by_line, steps = {}, set()
    for i, w, s, o in zip(d.index, d.WORK_ORDER.map(code), d.OPERATION_SEQUENCE.map(seq_key), d.OPERATION.map(code)):
        by_line.setdefault((w, s, o), i)
        steps.add((w, s))
    stale = {code(w) for w in stale_wos}
    rows = []
    for (w, s, op, dk), scen in zip(line_key(prev, "SEQ"), prev.SCENARIO):
        if (w, s, op, dk) in current:
            continue
        before = parse_diff_key(dk)
        now = ""
        if w in stale:
            fate = "moved by USER_8 rule"
        elif (w, s, op) not in by_line:
            fate = "step replaced by another operation" if (w, s) in steps else "started, closed or removed"
        elif (w, s, op) in listed_now:
            fate = "still listed, changed"
        else:
            r = d.loc[by_line[(w, s, op)]]
            now = make_diff_key({f: (shown(r[f"WO_{f}"]), shown(r[f"M_{f}"])) for f in before})
            own_compare = before and not str(scen).startswith("Differs from")
            if not own_compare:
                fate = "fixed - now matches"
            elif all(same_value(r[f"WO_{f}"], om, f) for f, (_, om) in before.items()):
                fate = "fixed - order updated"
            elif all(same_value(r[f"M_{f}"], oo, f) for f, (oo, _) in before.items()):
                fate = "fixed - master updated"
            else:
                fate = "fixed - both changed"
        rows.append({"FATE": fate, "WO": w, "STEP": s, "OPERATION": op, "SCENARIO_BEFORE": scen,
                     "DIFF_KEY_BEFORE": make_diff_key(before, "C" if str(scen).startswith("Differs from") else "M"),
                     "DIFF_KEY_NOW": now})
    return pd.DataFrame(rows, columns=cols).sort_values(["FATE", "WO", "STEP"]).reset_index(drop=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-extract", action="store_true",
                    help="build from the last extract in <repo>_private\\data instead of querying the ERP")
    ap.add_argument("--config", type=Path, default=PRIVATE_DIR / "extract.ini",
                    help="database settings for the extract (default: <repo>_private\\extract.ini)")
    ap.add_argument("--drift", type=Path, help="manual route: an export of master_drift.sql (.xlsx or .parquet)")
    ap.add_argument("--masters", type=Path, help="manual route: an export of old_revision_masters.sql")
    ap.add_argument("--out", type=Path, help="default: <repo>_private\\outputs\\master_drift_list_<today>.xlsx")
    ap.add_argument("--previous", type=Path,
                    help="list to compare with (default: the most recently saved one in outputs)")
    a = ap.parse_args()

    warnings = []
    data = PRIVATE_DIR / "data"
    if a.drift or a.masters:
        if not (a.drift and a.masters):
            sys.exit("--drift and --masters go together")
        for p in (a.drift, a.masters):
            if not p.exists():
                sys.exit(f"input not found: {p}")
        drift_day = datetime.fromtimestamp(a.drift.stat().st_mtime).date()
        masters_day = datetime.fromtimestamp(a.masters.stat().st_mtime).date()
        if drift_day != masters_day:
            warnings.append(f"inputs saved on different days ({drift_day} / {masters_day}) - "
                            "masters change; rerun both queries together")
        source = "manual exports"
    else:
        if not a.no_extract:
            import extract   # the database route; needs pyodbc and pyarrow
            if tuple(extract.QUERIES) != EXTRACT_QUERIES:
                sys.exit(f"extract.QUERIES {extract.QUERIES} and EXTRACT_QUERIES {EXTRACT_QUERIES} disagree")
            extract.extract(data, a.config)
        a.drift, a.masters = (data / f"{Path(q).stem}.parquet" for q in EXTRACT_QUERIES)
        manifest_path = data / "extract_manifest.json"
        if not (a.drift.exists() and a.masters.exists() and manifest_path.exists()):
            sys.exit(f"no extract in {data} - run without --no-extract")
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        changed = [q for q in EXTRACT_QUERIES if manifest["queries"][q]["sql_sha256"] != query_sha256(q)]
        if changed:
            warnings.append(f"{', '.join(changed)} changed since the extract of {manifest['started']} - "
                            "run without --no-extract")
        source = f"extract of {manifest['started']}"
    out = a.out or PRIVATE_DIR / "outputs" / f"master_drift_list_{datetime.now():%Y-%m-%d}.xlsx"
    if REPO_DIR in out.resolve().parents:
        sys.exit(f"refusing to write {out} inside the repo - the list names real parts and orders")
    out.parent.mkdir(parents=True, exist_ok=True)
    if a.previous and not a.previous.exists():
        sys.exit(f"previous list not found: {a.previous}")
    prev_path = a.previous or find_previous(PRIVATE_DIR / "outputs", out)

    d = normalise_drift(load(a.drift, DRIFT_COLUMNS), a.drift.name, warnings)
    m = normalise_masters(load(a.masters, MASTER_COLUMNS), a.masters.name, warnings)

    dup = int(d.duplicated(LINE_KEY).sum())
    if dup:
        sys.exit(f"{a.drift.name}: {dup} duplicate operation keys - a join fanned out")
    # against the SQL's own flags, BEFORE the stale-USER_8 rule rewrites any
    flag_problems = check_level1_flags(d)
    if flag_problems:
        sys.exit(f"level 1 flags disagree with the recomputation: {flag_problems}")
    d, stale = apply_stale_user8(d, m)

    l3 = level3(d, m)
    l2 = level2(d, l3)
    lst, matching = build_list(d, l3)
    if len(lst) + matching != len(d) or lst.duplicated(["WO", "STEP"]).any():
        sys.exit(f"the list does not account for every line once: {len(lst)} listed + {matching} matching "
                 f"vs {len(d)} lines")
    prev = read_previous(prev_path) if prev_path else None
    stale_wos = stale.WORK_ORDER.tolist() if "WORK_ORDER" in stale else []
    fates = fates_since_last(prev, d, lst, stale_wos)
    sql_written, sql_blank = add_sql_update(lst)   # lst is in its final order here

    not_extracted = int((l3.CATEGORY == "MASTERS_NOT_EXTRACTED").sum()) if not l3.empty else 0
    not_in_export = int((l3.CATEGORY == "OWN_LINE_NOT_IN_EXPORT").sum()) if not l3.empty else 0
    summary = [
        ("Run", f"{datetime.now():%Y-%m-%d %H:%M}"),
        ("Source", source),
        ("Drift input", f"{a.drift.name} ({datetime.fromtimestamp(a.drift.stat().st_mtime):%Y-%m-%d %H:%M})"),
        ("Masters input", f"{a.masters.name} ({datetime.fromtimestamp(a.masters.stat().st_mtime):%Y-%m-%d %H:%M})"),
        ("Previous list", f"{prev_path.name} ({datetime.fromtimestamp(prev_path.stat().st_mtime):%Y-%m-%d %H:%M})"
         if prev_path else "none found"),
        ("Check: duplicate operation keys", dup),
        ("Check: level 1 flags recomputed", "all agree"),
        ("", ""),
        ("Open, not-started WO operations", len(d)),
        ("Work orders", d[WO_KEY].drop_duplicates().shape[0]),
        ("Lines matching their master (not listed)", matching),
        (f"Lines listed ({LIST_SHEET})", f"{len(lst)} lines, {lst.WO.nunique()} work orders"),
    ]
    for key, name in SCENARIOS.items():
        part = lst[lst.SCENARIO == name]
        summary.append((f"Scenario - {name}", f"{len(part)} lines, {part.WO.nunique()} work orders"))
    summary += [
        ("SQL_UPDATE - statements written (Work order to update)", sql_written),
        ("SQL_UPDATE - Work order to update lines left empty (master value or WO_OPS_ROWID blank)", sql_blank),
    ]
    summary.append(("", ""))
    summary.append(("Since the previous list (02_SinceLastList) - lines not listed again unchanged", len(fates)))
    summary += [(f"Since the previous list - {k}", int(n)) for k, n in fates.FATE.value_counts().sort_index().items()]
    summary += [
        ("", ""),
        ("Stale USER_8 - WOs moved to the current routing (05_StaleUser8)", len(stale)),
        ("LEVEL 2 - work orders on an older revision (03_OldRevisionWOs)", len(l2)),
        ("LEVEL 3 - older-revision lines (04_CrossRevision)", len(l3)),
    ]
    summary += [(f"LEVEL 3 - {k}", int(v)) for k, v in l3.CATEGORY.value_counts().items()] if not l3.empty else []
    summary += [
        ("Check: level 3 parts missing from masters export", not_extracted),
        ("Check: matched own lines missing from masters export", not_in_export),
        ("LEVEL 3 NOTE", "INDICATIVE. Pairing assumes same name = same step. "
                         "To be reviewed again (agreed 2026-09-28)."),
    ]

    with pd.ExcelWriter(out) as xw:
        pd.DataFrame(summary, columns=["ITEM", "VALUE"]).to_excel(xw, sheet_name="00_Summary", index=False)
        lst[LIST_COLUMNS].to_excel(xw, sheet_name=LIST_SHEET, index=False)
        fates.to_excel(xw, sheet_name="02_SinceLastList", index=False)
        l2.to_excel(xw, sheet_name="03_OldRevisionWOs", index=False)
        l3.to_excel(xw, sheet_name="04_CrossRevision", index=False)
        stale.to_excel(xw, sheet_name="05_StaleUser8", index=False)
        for ws in xw.sheets.values():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        xw.sheets["00_Summary"].column_dimensions["A"].width = 78
        xw.sheets["00_Summary"].column_dimensions["B"].width = 60
        ws = xw.sheets[LIST_SHEET]
        widths = {"FAMILY": 22, "PART_ID": 16, "PART_DESCRIPTION": 30, "WO": 14, "OPERATION": 24,
                  "DIFF_KEY": 50, "SCENARIO": 34, "SQL_UPDATE": 70}
        for i, c in enumerate(LIST_COLUMNS):
            ws.column_dimensions[ws.cell(1, i + 1).column_letter].width = widths.get(c, 12)
        colour_differences(ws, lst)
        fs = xw.sheets["02_SinceLastList"]
        for col, w in {"A": 34, "B": 14, "C": 7, "D": 24, "E": 34, "F": 50, "G": 50, "H": 40}.items():
            fs.column_dimensions[col].width = w

    for k, v in summary:
        print(f"{k:<80} {v}")
    print(f"\nwritten: {out}")
    status = 0
    for w in warnings:
        print(f"!! {w}")
        status = 2
    if not_extracted or not_in_export:
        print("!! masters export does not cover every older-revision line - "
              "check both queries ran on the same day with the same scope")
        status = 2
    return status


if __name__ == "__main__":
    sys.exit(main())
