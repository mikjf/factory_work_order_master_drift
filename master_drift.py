"""Master drift - build the planners' workbook from the two SQL queries.

    python master_drift.py               query the ERP, then build the list
    python master_drift.py --no-extract  build again from the last extract

Inputs: the results of
  master_drift.sql          level 1, one row per open, not-started WO operation
  old_revision_masters.sql  master lines of the own and the current revision,
                            for parts with older-revision WOs
By default extract.py runs both against the ERP (read-only, Windows login,
settings in factory_shared_private\\connections\\visual.ini) and saves them as Parquet in
<repo>_private\\data\\ - together, so they always come from the same moment
(since 2026-10-15). The manual route still works: --drift and --masters take
exports pasted into Excel, as used until then; the script warns if the two
were saved on different days, or if their resource IDs lost leading zeros.

Output: <repo>_private\\outputs\\master_drift_list_<date>.xlsx by default,
never inside this repo - it names real parts and orders:
  00_Legend           how to read the list: scenarios, colours, columns
                      (user, 2026-10-23); the workbook opens on 01_SqlUpdater
  00_Summary          counts per scenario, checks, since last list
  01_SqlUpdater       THE list: one row per line that does not match the
                      master it follows (its own; the current revision for an
                      order on an older one) - the ROWIDs, part, order, step,
                      revisions, the order's (WO_) / master's (M_) / current
                      revision's (C_) values, SCENARIO, SQL_UPDATE (an UPDATE
                      on the order's operation, both "Update - ..."
                      scenarios) - see THE DRIFT LIST and THE SQL UPDATER below
  02_SinceLastList    what happened to each line of the previous list that is
                      not listed again unchanged: fixed (order / master
                      updated), started on the old values, or the step
                      replaced - see fates_since_last() (user, 2026-10-12)
  03_OldRevisionWOs   LEVEL 2: one row per WO built on an older revision
  04_CrossRevision    LEVEL 3: older-revision WO lines against the CURRENT
                      revision's master - since 2026-10-21 the reference for
                      those orders (PM and SC)
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
The standing assumption is "same name = same step". Each pairing starts from
the order's line matched to its own revision (by step and name, or by a name
unique in master and order), and the name must be unambiguous in both
revisions. Still to review: the NOT_COMPARED lines (a name
that repeats with different values) - counted, not listed.
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
# master_drift.sql. Until 2026-10-23 an empty crew counted as 1; since then it
# is a difference (user: the crew must be filled in correctly), so nothing is
# filled in for the comparison. Empty against empty is still equal.
NULL_AS = {}
# for RECOGNISING which routing an order carries (stale-USER_8 rule, level 3
# identical occurrences) an empty crew still counts as 1: a missing crew must
# not hide the routing - the values are then compared strictly with NULL_AS.
RECOGNISE_NULL_AS = {"USER_5": 1}
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
# MASTER_OPS_ROWID is optional in the masters input: exports before 2026-10-12 lack it
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
    # WO_BLOCKED (OdP Bloccato) is new on 2026-11-03: older exports cannot
    # tell, so nothing is left out and the run says so
    if "WO_BLOCKED" not in d:
        d["WO_BLOCKED"] = 0
        warnings.append(f"{name}: no WO_BLOCKED column (export before 2026-11-03) - "
                        "blocked orders are not left out; run without --no-extract")
    d["WO_BLOCKED"] = pd.to_numeric(d.WO_BLOCKED, errors="raise").astype(int)
    # pairing by unique name (2026-10-21): exports before it paired by
    # sequence + name only, and did not say why a line had no master line
    matched = d.NO_MASTER_LINE == 0
    if "MATCHED_BY" not in d:
        d["MATCHED_BY"] = pd.Series("SEQUENCE", index=d.index).where(matched)
    if "M_SEQUENCE_NO" not in d:
        d["M_SEQUENCE_NO"] = d.OPERATION_SEQUENCE.where(matched)
    if "NO_MASTER_REASON" not in d:
        d["NO_MASTER_REASON"] = pd.Series("UNKNOWN (export before 2026-10-21)", index=d.index).where(~matched)
    d["MATCHED_BY"] = d.MATCHED_BY.map(code)
    d["NO_MASTER_REASON"] = d.NO_MASTER_REASON.map(code)
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
        # exact matches only: a line paired by name (renumbered master) does
        # not count as "consistent with its own revision" here - otherwise
        # pairing by name (2026-10-21) would stop recognising stale labels
        own_ok = w.MATCHED_BY == "SEQUENCE"
        for f, (text, _) in FIELDS.items():
            own_ok &= same(w[f"WO_{f}"], w[f"M_{f}"], text, RECOGNISE_NULL_AS.get(f))
        own_all = bool(own_ok.all())
        if own_all:
            continue    # consistent with its own revision - the label may well be right
        j = w.reset_index().merge(c.add_prefix("CUR_"), how="left",
                                  left_on=["OPERATION_SEQUENCE", "OPERATION"],
                                  right_on=["CUR_SEQUENCE_NO", "CUR_OPERATION"]).set_index("index")
        if len(j) != len(w) or j.CUR_SEQUENCE_NO.isna().any():
            continue
        ok = pd.Series(True, index=j.index)
        for f, (text, _) in FIELDS.items():
            ok &= same(j[f"WO_{f}"], j[f"CUR_{f}"], text, RECOGNISE_NULL_AS.get(f))
        if not ok.all():
            continue
        # pandas 3 refuses to upcast silently: numbers go in as floats (NaN for blank)
        for f, (text, _) in FIELDS.items():
            v = j[f"CUR_{f}"]
            d.loc[j.index, f"M_{f}"] = v if text else pd.to_numeric(v, errors="coerce").astype(float)
        d.loc[j.index, "MASTER_OPS_ROWID"] = pd.to_numeric(j["CUR_MASTER_OPS_ROWID"], errors="coerce").astype(float)
        d.loc[j.index, "M_SEQUENCE_NO"] = pd.to_numeric(j["CUR_SEQUENCE_NO"], errors="coerce").astype(float)
        d.loc[j.index, "NO_MASTER_LINE"] = 0
        d.loc[j.index, "MATCHED_BY"] = "SEQUENCE"
        d.loc[j.index, "NO_MASTER_REASON"] = None
        # values compared strictly with the current master it now follows
        for f, (text, s) in FIELDS.items():
            d.loc[j.index, f"D_{s}"] = (~same(d.loc[j.index, f"WO_{f}"], d.loc[j.index, f"M_{f}"],
                                              text, NULL_AS.get(f))).astype(int)
        d.loc[j.index, ["USER_8_STALE", "ON_CURRENT_ROUTING"]] = 1
        first = w.iloc[0]
        stale.append({**{k: first[k] for k in WO_KEY}, "PART_ID": first.PART_ID,
                      "PART_DESCRIPTION": first.PART_DESCRIPTION, "WO_KIND": first.WO_KIND,
                      "USER_8": first.WO_REV, "CURRENT_REV": first.CURRENT_REV,
                      "N_LINES": len(w), "LINES_NO_OWN_MASTER_LINE_BEFORE": int(w.NO_MASTER_LINE.sum()),
                      "LINES_DIFFERING_BEFORE": int((w[flags].fillna(0).max(axis=1) == 1).sum())})
    return d, pd.DataFrame(stale)


def signature(rows):
    rows = rows[list(FIELDS)].fillna(RECOGNISE_NULL_AS)
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
            # the master's own step: differs from the WO's when paired by name
            hit = own[(own.SEQUENCE_NO == r.M_SEQUENCE_NO) & (own.OPERATION == r.OPERATION)]
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
# its master. A line WITHOUT a master line is not listed, only counted by
# reason (user, 2026-10-21, after PM and SC: steps added to the order on the
# go - rework S RIL ..., samples - have nothing to compare with). A line
# whose step was renumbered in the master is paired by name in the SQL
# (MATCHED_BY = UNIQUE_NAME) and listed like any other; M_STEP then differs
# from STEP.
#
# THE REFERENCE IS THE CURRENT REVISION (PM and SC, 2026-10-21): an order
# created on an older revision follows the CURRENT revision of the master -
# the one PART_SITE.ENGINEERING_MSTR names - not the one it was created from.
# Their two examples: an order still carrying its old revision's rate is
# brought to the current one; an order that already carries the current rate
# stays as it is, although its old revision said otherwise.
#   - An order on the current revision (or moved onto it by the stale-USER_8
#     rule) is compared with its master line (M_), paired exactly.
#   - An order on an older revision is compared with the current revision
#     (C_), through the level 3 pairing: unique name, identical occurrences or
#     own resource - each a pairing of the order's EXACT own-revision line
#     with one current line. Lines level 3 cannot pair are counted by
#     category, not listed: their current value is unknown.
# Each row has exactly one SCENARIO (names chosen by the user, 2026-10-23:
# factual - which master the order differs from, not why; the cause may be a
# master corrected later, a master wrong at creation or an order edited):
#   Update - WO differs from master (WO rev)
#                             order on its master's revision (the current
#                             one): same resource, another value differs from
#                             that master (28 Sep review: 51 of 51 - the order
#                             follows the master)
#   Update - WO differs from master (current rev)
#                             order on an older revision: same resource,
#                             another value differs from the current revision
#   Review - OPP WO differs from master (WO rev)
#   Review - OPP WO differs from master (current rev)
#                             the same two cases on an OPP order (a new
#                             project): always reviewed by a person before
#                             updating (PM and SC, 2026-10-23). Only the purple
#                             "Update - ..." lines may be automated later.
#   Resource changed          the resource differs from the master compared
#                             with (28 Sep review: 9 of 13 intentional; SC,
#                             2026-10-21: "should be" intentional - kept)
# Lines that match are counted, not listed. DIFF_KEY names exactly what
# differs, each value tagged (WO) / (M) / (C) - (C) for an older-revision
# order. Earlier names: "Work order to update" and "Work order to update
# (current revision)" (12-23 Oct); before 21 Oct also "Differs from current
# master (WO on older revision)" - the new-revision case - and "Work order
# matches current master", now not listed: correct as it is. The 00_Legend
# sheet explains scenarios and colours to the reader.
#
# THE SQL UPDATER (user, 2026-10-12). The list is laid out as the user asked,
# FTS-updater style: per line the order's (WO_), its master's (M_) and the
# current revision's (C_) values side by side, and SQL_UPDATE - a spreadsheet
# FORMULA that builds an UPDATE on the work-order operation (WO_OPS_ROWID)
# from the cells of its own row, so correcting a value corrects the statement.
#   - on the "Update - ..." and "Review - ..." scenarios (a review line gets
#     its statement ready for after the review); "Resource changed" needs a
#     person;
#   - the value is the one compared with: M_ on "(WO rev)", C_ on
#     "(current rev)";
#   - it sets only the fields that differ; never RESOURCE_ID (an update line
#     has the same resource by definition), never the step number;
#   - empty when a master value or the ROWID is blank, so nothing can blank a
#     field by accident. Nothing here runs anything.
# The cells of each difference in DIFF_KEY are coloured (user, 2026-10-13):
# the order's value (WO_) orange, the value it is compared with green (M_ or
# C_, as tagged). Under them the row is coloured by scenario group, strong
# when the line is new since the previous list, light when already listed -
# see ROW_FILLS. Fills are fixed at writing time; they do not follow edits
# made in the sheet.
# M_ / M_REV: the order's own revision (= C_REV for an order moved by the
# stale-USER_8 rule). For orders on the current revision C_ = M_.
# ---------------------------------------------------------------------------

SCENARIOS = {  # key -> shown name; dict order = sort order of the list
    "update": "Update - WO differs from master (WO rev)",
    "update_cur": "Update - WO differs from master (current rev)",
    "review": "Review - OPP WO differs from master (WO rev)",
    "review_cur": "Review - OPP WO differs from master (current rev)",
    "check": "Resource changed",
}
UPDATE_SCENARIOS = ("update", "update_cur")          # purple row: may be automated later
REVIEW_SCENARIOS = ("review", "review_cur")          # blue row: always reviewed (OPP)
SQL_SCENARIOS = UPDATE_SCENARIOS + REVIEW_SCENARIOS   # get an SQL_UPDATE
REVIEW_KINDS = ("OPP",)   # WO_KIND whose updates are always reviewed (PM and SC, 2026-10-23)
# scenarios of earlier lists, still read by 02_SinceLastList
OLD_NOMASTER = "No master line"                                       # until 2026-10-21
OLD_CUR_SCENARIOS = ("Differs from current master (WO on older revision)", "Differs from today's master",
                     "Work order matches current master")             # until 2026-10-21
# level 3 CATEGORY -> how the summary names older-revision lines it could not
# pair with the current revision (counted, not listed)
UNPAIRED = {
    "NOT_COMPARED": "name repeats with different values (not paired)",
    "NOT_IN_CURRENT": "step no longer in the current revision",
    "NO_CURRENT_MASTER": "no master for the current revision",
    "MASTERS_NOT_EXTRACTED": "masters not in the extract",
    "OWN_LINE_NOT_IN_EXPORT": "own master line not in the extract",
}
# NO_MASTER_REASON -> how the summary names the unlisted lines
NO_MASTER_REASONS = {
    "NOT_IN_MASTER": "step not in the master (added to the order on the go)",
    "NAME_NOT_UNIQUE": "name repeats in the master or the order (not paired)",
    "NO_MASTER": "no master for the order's revision",
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
     "WO", "WO_STATUS", "WO_CREATE_DATE", "WO_DESIRED_RLS_DATE", "OPERATION", "STEP", "M_STEP",
     "WO_REV", "M_REV", "C_REV"]
    + VALUE_COLUMNS
    + ["DIFF_KEY", "SCENARIO", "FIRST_LISTED", "SQL_UPDATE"])


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
    """The drift list: one row per line that does not match the master it
    follows - its own on the current revision, the current one otherwise.
    See THE DRIFT LIST above.

    Returns (list, counts): counts["matching"] lines that match; counts
    ["no_line"] {NO_MASTER_REASON: lines without a master line}; counts
    ["unpaired"] {level 3 CATEGORY: older-revision lines not paired with the
    current revision} and counts["unpaired_own_diff"] how many of those differ
    from their own revision; counts["matches_current_only"] older-revision
    lines that differ from their own revision but match the current one.
    """
    flags = {f: f"D_{s}" for f, (_, s) in FIELDS.items()}
    rows = []
    counts = {"matching": 0, "no_line": {}, "unpaired": {}, "unpaired_own_diff": 0, "matches_current_only": 0}
    for idx, r in zip(d.index, d.itertuples(index=False)):
        if r.NO_MASTER_LINE == 1:
            reason = r.NO_MASTER_REASON or "UNKNOWN"
            counts["no_line"][reason] = counts["no_line"].get(reason, 0) + 1
            continue
        own = [f for f in FIELDS if getattr(r, flags[f]) == 1]
        on_current = r.ON_CURRENT_ROUTING == 1
        compared = not on_current and pd.notna(l3.at[idx, "CUR_SEQUENCE_NO"])
        if on_current:
            tag, diff = "M", own
            pairs = {f: (shown(getattr(r, f"WO_{f}")), shown(getattr(r, f"M_{f}"))) for f in own}
        elif compared:
            # older revision: the current revision is the reference (PM and SC)
            tag = "C"
            diff = [f for f, (_, s) in FIELDS.items() if l3.at[idx, f"D3_{s}"] == 1]
            pairs = {f: (shown(getattr(r, f"WO_{f}")), shown(l3.at[idx, f"CUR_{f}"])) for f in diff}
            if own and not diff:
                counts["matches_current_only"] += 1
        else:
            cat = l3.at[idx, "CATEGORY"]
            counts["unpaired"][cat] = counts["unpaired"].get(cat, 0) + 1
            counts["unpaired_own_diff"] += bool(own)
            continue
        if not diff:
            counts["matching"] += 1
            continue
        if "RESOURCE_ID" in diff:
            key = "check"
        else:
            key = "update" if on_current else "update_cur"
            if code(r.WO_KIND) in REVIEW_KINDS:   # new project: a person reviews first
                key = "review" if on_current else "review_cur"
        rec = {
            "_S": list(SCENARIOS).index(key), "_DIFF": diff, "_T": tag,
            "WO_OPS_ROWID": r.WORKORDER_OPS_ROWID,
            "M_OPS_ROWID": r.MASTER_OPS_ROWID,
            "C_MASTER_OPS_ROWID": r.MASTER_OPS_ROWID if on_current else
                                  (l3.at[idx, "CUR_MASTER_OPS_ROWID"] if compared else pd.NA),
            "FAMILY": r.FAMILY, "PART_ID": r.PART_ID, "PART_DESCRIPTION": r.PART_DESCRIPTION,
            "WO": r.WORK_ORDER, "WO_STATUS": r.WO_STATUS,
            "WO_CREATE_DATE": pd.to_datetime(r.CREATE_DATE).date(),
            "WO_DESIRED_RLS_DATE": pd.to_datetime(r.DESIRED_RLS_DATE).date(),
            "OPERATION": r.OPERATION, "STEP": r.OPERATION_SEQUENCE,
            # the own master's step: differs from STEP when paired by name
            "M_STEP": int(r.M_SEQUENCE_NO) if pd.notna(r.M_SEQUENCE_NO) else pd.NA,
            "WO_REV": r.WO_REV, "M_REV": r.CURRENT_REV if r.USER_8_STALE == 1 else r.WO_REV,
            "C_REV": r.CURRENT_REV,
            "DIFF_KEY": make_diff_key(pairs, tag),
            "SCENARIO": SCENARIOS[key], "SQL_UPDATE": "",
        }
        for f, s in SHORT.items():
            rec[f"WO_{s}"] = getattr(r, f"WO_{f}")
            rec[f"M_{s}"] = getattr(r, f"M_{f}")
            rec[f"C_{s}"] = (rec[f"M_{s}"] if on_current else
                             l3.at[idx, f"CUR_{f}"] if compared else pd.NA)
        rows.append(rec)
    x = pd.DataFrame(rows, columns=["_S", "_DIFF", "_T"] + LIST_COLUMNS)
    return x.sort_values(["_S", "OPERATION", "WO", "STEP"]).reset_index(drop=True), counts


def add_sql_update(lst):
    """SQL_UPDATE on every line of both 'Update - ...' scenarios, as a
    live spreadsheet formula reading the value compared with (M_ or C_).

    Uses the row's position in the written sheet (header = row 1), so lst
    must be in its final order and column order. See THE SQL UPDATER above.
    Returns the number of statements written and of update lines left empty.
    """
    from openpyxl.utils import get_column_letter
    col = {c: get_column_letter(i + 1) for i, c in enumerate(LIST_COLUMNS)}
    updates = {SCENARIOS[k] for k in SQL_SCENARIOS}
    written = blank = 0
    for n, (i, r) in enumerate(lst.iterrows(), start=2):
        if r.SCENARIO not in updates:
            continue
        fields = [f for f in r._DIFF if f != "RESOURCE_ID"]
        src = {f: f"{r._T}_{SHORT[f]}" for f in fields}
        if (pd.isna(r.WO_OPS_ROWID) or not fields
                or any(pd.isna(r[src[f]]) or code(r[src[f]]) is None for f in fields)):
            blank += 1
            continue
        sets = ", ".join(f"{f} = '\"&{col[src[f]]}{n}&\"'" for f in fields)
        lst.at[i, "SQL_UPDATE"] = f'="UPDATE OPERATION SET {sets} WHERE ROWID = "&{col["WO_OPS_ROWID"]}{n}&";"'
        written += 1
    return written, blank


# orange, green (Excel's "Good" fill) - stronger than the first pastel tones
# (FCE4D6 / E2EFDA), which some screens showed as white (2026-10-21)
FILL_WO, FILL_TARGET = "F8CBAD", "C6EFCE"
# ROW colours (user, 2026-10-28): per scenario group a HIGHLIGHTED colour for a
# line that is new since the previous list (or whose differences changed), a
# LIGHT one for a line already in the previous list with the same differences.
# Light means "not new", not "approved". Updates were grey until 2026-10-28.
ROW_FILLS = {                    # group -> (new, already listed)
    "update": ("CCC0DA", "E4DFEC"),   # purple: may be automated later
    "review": ("BDD7EE", "DDEBF7"),   # blue: OPP, always reviewed (2026-10-23)
    "check": ("FFE699", "FFF2CC"),    # yellow: resource changed
}


def scenario_group(scenario):
    """The ROW_FILLS group of a scenario name."""
    keys = [k for k, v in SCENARIOS.items() if v == scenario]
    k = keys[0] if keys else ""
    return "update" if k in UPDATE_SCENARIOS else "review" if k in REVIEW_SCENARIOS else "check"


def colour_differences(ws, lst):
    """Colour every row by scenario group and new / already listed, then the two
    cells of every difference in DIFF_KEY. See THE SQL UPDATER above."""
    from openpyxl.styles import PatternFill
    solid = lambda c: PatternFill("solid", start_color=c, end_color=c)
    fills = {"WO": solid(FILL_WO), "T": solid(FILL_TARGET)}
    rows = {(g, new): solid(c[0 if new else 1]) for g, c in ROW_FILLS.items() for new in (True, False)}
    pos = {c: i + 1 for i, c in enumerate(LIST_COLUMNS)}
    coloured = 0
    for n, (_, r) in enumerate(lst.iterrows(), start=2):
        row = rows[(scenario_group(r.SCENARIO), bool(r._NEW))]   # whole row first, the differences on top
        for i in range(1, len(LIST_COLUMNS) + 1):
            ws.cell(n, i).fill = row
        target = r._T   # M_ or C_: the value compared with
        for f in r._DIFF:
            ws.cell(n, pos[f"WO_{SHORT[f]}"]).fill = fills["WO"]
            ws.cell(n, pos[f"{target}_{SHORT[f]}"]).fill = fills["T"]
            coloured += 1
    return coloured


def write_legend(book):
    """00_Legend: how to read the list, for someone who did not build it (user,
    2026-10-23). First sheet; the workbook still opens on the list itself."""
    from openpyxl.styles import Alignment, Font, PatternFill
    colours = {"WO": FILL_WO, "T": FILL_TARGET, "U": ROW_FILLS["update"][0], "u": ROW_FILLS["update"][1],
               "R": ROW_FILLS["review"][0], "r": ROW_FILLS["review"][1],
               "C": ROW_FILLS["check"][0], "c": ROW_FILLS["check"][1]}
    fill = {k: PatternFill("solid", start_color=c, end_color=c) for k, c in colours.items()}
    S = SCENARIOS
    rows = [
        ("title", "How to read 01_SqlUpdater", "", ""),
        ("", "One row per open, not-started operation whose values differ from the master it follows. "
             "Read-only: nothing is written to VISUAL.", "", ""),
        ("head", "SCENARIO", "What it means", "What to do"),
        ("U", S["update"],
         "The order is on its master's revision. Some values differ from that master - the master "
         "corrected later, the master wrong when the order was created, or the order edited. "
         "Green = the master's value (M_).",
         "Update the order to the green value (SQL_UPDATE)."),
        ("U", S["update_cur"],
         "The order was created on a revision that is no longer the current one (WO_REV differs from "
         "C_REV). The current revision is the reference (PM and SC). Some values differ from it. "
         "Green = the current revision's value (C_).",
         "Update the order to the green value (SQL_UPDATE)."),
        ("R", S["review"] + "\n" + S["review_cur"],
         "The same two cases on an OPP order (a new project).",
         "Always review first; then update to the green value (SQL_UPDATE)."),
        ("C", S["check"],
         "The planned resource differs from the master the line follows. Often intentional "
         "(supply, PM).",
         "Check; correct by hand if needed. No SQL_UPDATE."),
        ("", "", "", ""),
        ("head", "COLOUR", "Meaning", ""),
        ("WO", "orange", "the order's value today (WO_ column)", ""),
        ("T", "green", "the value it should have - the one SQL_UPDATE writes (M_ or C_ column)", ""),
        ("", "ROW colour", "strong = NEW since the previous list (or its differences changed); "
                           "light = already in the previous list with the same differences. Light means "
                           "'not new', not 'approved'.", ""),
        ("U", "purple row", "an update line, new: SQL_UPDATE ready - may be automated later", ""),
        ("u", "light purple row", "an update line, already listed", ""),
        ("R", "blue row", "a review line (OPP), new: SQL_UPDATE ready, only after a person has checked it", ""),
        ("r", "light blue row", "a review line, already listed", ""),
        ("C", "yellow row", "a resource change, new", ""),
        ("c", "light yellow row", "a resource change, already listed", ""),
        ("", "", "", ""),
        ("head", "COLUMNS", "Meaning", ""),
        ("", "WO_ / M_ / C_", "the order's value / its own master's / the current revision's", ""),
        ("", "WO_REV / M_REV / C_REV", "the order's revision / its master's / the current one", ""),
        ("", "STEP / M_STEP", "the order's step / the master's (different only when the master was renumbered)", ""),
        ("", "DIFF_KEY", "what differs, e.g. RUN: 60 (WO) | 70 (M) - (C) = compared with the current revision", ""),
        ("", "FIRST_LISTED", "the date of the first saved list that showed this line with these same "
                             "differences - how long it has been waiting", ""),
        ("", "SQL_UPDATE", "the statement for the order, built from the green cells of its row: only the values "
                           "that differ, never the resource or the step. Nothing runs it automatically.", ""),
        ("", "", "", ""),
        ("head", "NOT LISTED", "Counted in 00_Summary", ""),
        ("", "blocked orders", "orders with OdP Bloccato ticked (order header, Altro tab): left out", ""),
        ("", "matching lines", "the order already has the values of the master it follows", ""),
        ("", "no master line", "steps added to the order on the go (rework, samples): nothing to compare", ""),
        ("", "not paired", "older-revision lines that cannot be paired with the current revision without guessing", ""),
        ("", "", "", ""),
        ("", "02_SinceLastList", "what happened to the lines of the previous list: fixed, started, replaced ...", ""),
    ]
    ws = book.create_sheet("00_Legend", 0)
    wrap = Alignment(wrap_text=True, vertical="top")
    for n, (kind, a, b, c) in enumerate(rows, start=1):
        for col, v in enumerate((a, b, c), start=1):
            cell = ws.cell(n, col, v)
            cell.alignment = wrap
            if kind in ("head", "title"):
                cell.font = Font(bold=True, size=13 if kind == "title" else 11)
        if kind in fill:
            ws.cell(n, 1).fill = fill[kind]
    for col, w in (("A", 46), ("B", 90), ("C", 42)):
        ws.column_dimensions[col].width = w
    book.active = book.sheetnames.index(LIST_SHEET)
    for sheet in book.worksheets:   # only the list shows as selected
        sheet.sheet_view.tabSelected = sheet.title == LIST_SHEET


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


def list_day(path):
    """The day a saved list stands for: from its name, else when it was saved."""
    m = re.search(r"(\d{4}-\d{2}-\d{2})", path.name)
    return m.group(1) if m else f"{datetime.fromtimestamp(path.stat().st_mtime):%Y-%m-%d}"


def list_history(folder, out):
    """{line key: day of the first saved list that showed it} over every list in
    folder except out (the one being written). A line key includes its
    differences, so a line whose differences changed starts again. For
    FIRST_LISTED (user, 2026-10-28)."""
    wanted = {LIST_SHEET, OLD_DRIFT_LIST_SHEET, *OLD_LIST_SHEETS, *OLD_REVIEW_SHEETS}
    first = {}
    for p in folder.glob("*.xlsx"):
        if p.name.startswith("~$") or p.resolve() == out.resolve():
            continue
        try:
            if not wanted & set(pd.ExcelFile(p).sheet_names):
                continue
            keys = line_key(read_previous(p), "SEQ")
        except (Exception, SystemExit):
            continue    # not a list this script can read - ignored here
        day = list_day(p)
        for k in keys:
            if k not in first or day < first[k]:
                first[k] = day
    return first


def same_value(a, b, field):
    """One value against another as the comparison does: numbers as numbers, blank = blank."""
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


def fates_since_last(prev, d, lst, stale_wos, l3, blocked):
    """What happened to each line of the previous list that is not listed again unchanged.

    FATE, in this order:
      left out - order blocked       its work order is blocked (OdP Bloccato) and
                                     no longer compared (2026-11-03)
      moved by USER_8 rule           its work order is now on 05_StaleUser8
      step replaced by another       the order still has the step open, but under
        operation                    another operation name - the line was
                                     removed or renamed, not fixed (user, 2026-10-13)
      started, closed or removed     the step is no longer an open, not-started
                                     operation - mostly produced (or closing) on
                                     the values it had; this export cannot tell a
                                     started step from a deleted one
      partly fixed - still listed    listed again with fewer differences, the rest
                                     unchanged - some values corrected (2026-10-23)
      still listed, changed          listed again, with other differences
      fixed - order updated          the order now carries the value the list
                                     proposed (M or C)
      not compared - current         the order is on an older revision that level 3
        revision not paired          cannot pair with the current one: not listed,
                                     but not known to match either
      fixed - master updated         the master now carries the order's old value
      fixed - both changed           order and master agree on a third value
      fixed - now matches            no longer differs (compared with the current
                                     revision before, or the order is on an older
                                     revision and now matches the current one)
    Lines tagged (C) in DIFF_KEY were compared with the current revision; "order
    updated" then means the order now carries that value. Lines of the scenarios
    no longer listed - "No master line", "Work order matches current master" -
    are not followed up.
    """
    cols = ["FATE", "WO", "STEP", "OPERATION", "SCENARIO_BEFORE", "DIFF_KEY_BEFORE", "DIFF_KEY_NOW"]
    if prev is None or prev.empty:
        return pd.DataFrame(columns=cols)
    current = set(line_key(lst, "STEP", "WO"))
    # lines are matched on order, step AND operation name: a step can be
    # replaced by another operation under the same number
    listed_now = {(code(w), seq_key(s), code(o)): dk for w, s, o, dk in zip(lst.WO, lst.STEP, lst.OPERATION, lst.DIFF_KEY)}
    by_line, steps = {}, set()
    for i, w, s, o in zip(d.index, d.WORK_ORDER.map(code), d.OPERATION_SEQUENCE.map(seq_key), d.OPERATION.map(code)):
        by_line.setdefault((w, s, o), i)
        steps.add((w, s))
    stale = {code(w) for w in stale_wos}
    blocked_lines = set(zip(blocked.WORK_ORDER.map(code), blocked.OPERATION_SEQUENCE.map(seq_key),
                            blocked.OPERATION.map(code)))
    rows = []
    for (w, s, op, dk), scen, raw_dk in zip(line_key(prev, "SEQ"), prev.SCENARIO, prev.DIFF_KEY):
        # scenarios no longer listed (2026-10-21) - lines without a master
        # line, orders already matching the current revision - are not followed up
        if (w, s, op, dk) in current or str(scen) == OLD_NOMASTER or str(scen) == OLD_CUR_SCENARIOS[-1]:
            continue
        before = parse_diff_key(dk)
        # compared with the current revision before: tagged (C) in the list's
        # own DIFF_KEY (dk above is the normalised form, without tags - reading
        # the tag from it labelled every line (M) until 2026-10-23), or the old
        # "Differs from ..." scenario, whose earliest lists had no tags
        tag = "C" if "(C)" in str(raw_dk) or str(scen).startswith("Differs from") else "M"
        now = ""
        if (w, s, op) in blocked_lines:
            fate = "left out - order blocked"
        elif w in stale:
            fate = "moved by USER_8 rule"
        elif (w, s, op) not in by_line:
            fate = "step replaced by another operation" if (w, s) in steps else "started, closed or removed"
        elif (w, s, op) in listed_now:
            now = listed_now[(w, s, op)]
            after = parse_diff_key(now)
            # fewer differences, the remaining ones exactly as before
            partly = set(after) < set(before) and all(after[f] == before[f] for f in after)
            fate = "partly fixed - still listed" if partly else "still listed, changed"
        else:
            i = by_line[(w, s, op)]
            r = d.loc[i]
            # the order now carries the value the list proposed - whatever the
            # revision did meanwhile (first corrections, 2026-10-21, landed
            # while the part moved to a new revision)
            updated = bool(before) and all(same_value(r[f"WO_{f}"], v, f) for f, (_, v) in before.items())
            unknown = r.ON_CURRENT_ROUTING == 0 and not (i in l3.index and pd.notna(l3.at[i, "CUR_SEQUENCE_NO"]))
            if updated:
                fate = "fixed - order updated"
            elif unknown:
                # older revision, not paired with the current one: not listed,
                # but not known to match either
                fate = "not compared - current revision not paired"
            elif tag == "C" or r.ON_CURRENT_ROUTING == 0:
                fate = "fixed - now matches"
            elif not before:
                fate = "fixed - now matches"
            elif all(same_value(r[f"M_{f}"], oo, f) for f, (oo, _) in before.items()):
                fate = "fixed - master updated"
            else:
                fate = "fixed - both changed"
            if tag == "M" and r.ON_CURRENT_ROUTING == 1:
                now = make_diff_key({f: (shown(r[f"WO_{f}"]), shown(r[f"M_{f}"])) for f in before})
        rows.append({"FATE": fate, "WO": w, "STEP": s, "OPERATION": op, "SCENARIO_BEFORE": scen,
                     "DIFF_KEY_BEFORE": make_diff_key(before, tag),
                     "DIFF_KEY_NOW": now})
    return pd.DataFrame(rows, columns=cols).sort_values(["FATE", "WO", "STEP"]).reset_index(drop=True)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--no-extract", action="store_true",
                    help="build from the last extract in <repo>_private\\data instead of querying the ERP")
    ap.add_argument("--config", type=Path,
                    default=REPO_DIR.parent / "factory_shared_private" / "connections" / "visual.ini",
                    help="database settings for the extract "
                         "(default: factory_shared_private\\connections\\visual.ini)")
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
    # blocked orders (OdP Bloccato) are left out before anything else and
    # only counted (2026-11-03)
    n_all, n_all_wos = len(d), d[WO_KEY].drop_duplicates().shape[0]
    blocked = d[d.WO_BLOCKED == 1]
    d = d[d.WO_BLOCKED == 0].copy()
    d, stale = apply_stale_user8(d, m)

    l3 = level3(d, m)
    l2 = level2(d, l3)
    lst, counts = build_list(d, l3)
    matching, no_line_lines, unpaired = counts["matching"], counts["no_line"], counts["unpaired"]
    no_line, n_unpaired = sum(no_line_lines.values()), sum(unpaired.values())
    if len(lst) + matching + no_line + n_unpaired != len(d) or lst.duplicated(["WO", "STEP"]).any():
        sys.exit(f"the list does not account for every line once: {len(lst)} listed + {matching} matching "
                 f"+ {no_line} without master line + {n_unpaired} not paired with the current revision "
                 f"vs {len(d)} lines")
    by_name = d[(d.MATCHED_BY == "UNIQUE_NAME") & (d.NO_MASTER_LINE == 0)]
    by_name_listed = int(lst.STEP.map(seq_key).ne(lst.M_STEP.map(seq_key)).sum())
    prev = read_previous(prev_path) if prev_path else None
    stale_wos = stale.WORK_ORDER.tolist() if "WORK_ORDER" in stale else []
    fates = fates_since_last(prev, d, lst, stale_wos, l3, blocked)
    # new since the previous list, or already in it with the same differences
    # (row colour), and the day it was first listed (user, 2026-10-28)
    keys = line_key(lst, "STEP", "WO")
    prev_keys = set(line_key(prev, "SEQ")) if prev is not None and not prev.empty else set()
    lst["_NEW"] = [k not in prev_keys for k in keys]
    history = list_history(PRIVATE_DIR / "outputs", out)
    today = list_day(out)
    lst["FIRST_LISTED"] = [history.get(k, today) for k in keys]
    # final order (user, 2026-10-28): purple, light purple, blue, light blue,
    # yellow, light yellow; inside each block newest FIRST_LISTED first, the
    # longest-waiting at the bottom; then scenario, operation, order, step
    lst["_G"] = lst.SCENARIO.map(lambda s: list(ROW_FILLS).index(scenario_group(s)))
    lst = lst.sort_values(["_G", "_NEW", "FIRST_LISTED", "_S", "OPERATION", "WO", "STEP"],
                          ascending=[True, False, False, True, True, True, True]).reset_index(drop=True)
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
        ("Open, not-started WO operations", n_all),
        ("Work orders", n_all_wos),
        ("Blocked work orders - OdP Bloccato (left out, not listed)",
         f"{len(blocked)} lines, {blocked[WO_KEY].drop_duplicates().shape[0]} work orders"),
        ("Lines compared (not blocked)", len(d)),
        ("Lines matching the master they follow (not listed)", matching),
        ("   of which older-revision lines matching the current revision but not their own",
         counts["matches_current_only"]),
        ("Lines without a master line (not listed)", no_line),
    ]
    summary += [(f"   without master line - {NO_MASTER_REASONS.get(k, k)}", n)
                for k, n in sorted(no_line_lines.items(), key=lambda kv: -kv[1])]
    summary += [("Older-revision lines not paired with the current revision (not listed)", n_unpaired)]
    summary += [(f"   not paired - {UNPAIRED.get(k, k)}", n)
                for k, n in sorted(unpaired.items(), key=lambda kv: -kv[1])]
    summary += [("   of which differ from their own revision", counts["unpaired_own_diff"])]
    summary += [
        ("Lines paired by unique name (step renumbered in the master)",
         f"{len(by_name)} - of which listed with a difference: {by_name_listed}"),
        (f"Lines listed ({LIST_SHEET})", f"{len(lst)} lines, {lst.WO.nunique()} work orders"),
    ]
    for key, name in SCENARIOS.items():
        part = lst[lst.SCENARIO == name]
        summary.append((f"Scenario - {name}", f"{len(part)} lines, {part.WO.nunique()} work orders - "
                                              f"{int(part._NEW.sum())} new, {int((~part._NEW).sum())} already listed"))
    summary += [
        ("SQL_UPDATE - statements written (Update and Review scenarios)", sql_written),
        ("SQL_UPDATE - update lines left empty (value or WO_OPS_ROWID blank)", sql_blank),
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
        ("LEVEL 3 NOTE", "Drives the list for older-revision orders since 2026-10-21 (current revision = "
                         "reference). Pairs only unambiguous names; NOT_COMPARED lines stay unpaired. "
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
        write_legend(xw.book)   # after the loop above: no filter or frozen header there

    for k, v in summary:
        print(f"{k:<80} {v}")
    print(f"\nwritten: {out}")
    status = 0
    for w in warnings:
        print(f"!! {w}")
        status = 2
    if not_extracted or not_in_export:
        print("!! the masters result does not cover every older-revision line - "
              "check both queries ran together and keep the same work-order scope")
        status = 2
    return status


if __name__ == "__main__":
    sys.exit(main())
