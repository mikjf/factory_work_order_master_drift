# Work-order master drift

Lists open, not-started work-order operations whose values no longer match
their master routing.

A work order **copies** its master's operations when it is created and never
looks back. When a master is improved later - a new run rate, a new setup
time - the work orders already created keep the old values, so planning and
production run on an outdated standard until they finish. This tool finds
them. It is **read-only**: it produces a list; nothing is written to the ERP.

It follows on from `factory_time_statistics` (FTS), which proposes new master
values: key the new masters, run this, see which open orders still carry the
old ones.

## Terms

| Term | Meaning here |
|---|---|
| **Master** | the engineering routing of a part: `WORK_ORDER.TYPE = 'M'` and its operations |
| **Work order (WO)** | a production order, `TYPE = 'W'`, with its own copy of the master's operations |
| **Revision** | a version of a part's master, held in the master's `LOT_ID`. A WO's `USER_8` says which revision it was built from; `PART_SITE.ENGINEERING_MSTR` says which one is current. Revision labels (`A`, `7`, `ZZ` ...) cannot be ordered - "older" means *different from the current one* |
| **Open, not started** | WO status firmed or released (`F`, `R`), operation status `F` or `R`, and no labour ticket of any kind on the operation |
| **Drift** | a value on the WO's operation that differs from the master it is compared with |

Compared fields: planned resource, `RUN_TYPE`, `RUN`, `LOAD_SIZE_QTY`,
`SETUP_HRS` and `USER_5` (planned crew; empty counts as 1). `RUN_HRS` and
`CALC_START_QTY` depend on the order quantity and are not compared.

## Scope

- Work orders numbered `.../OP` (production) and `.../OPP` (production on a
  new project). Rework (`OR`) and `OS` are out: rework routings are ad hoc.
- Parts `M.%`.
- Unreleased (`U`) work orders are out - in practice old orders set back.
- External operations (names starting `E `) are out.
- 4-digit sequences only. Work orders from before the masters were
  renumbered keep 2- or 3-digit sequences and are not covered yet.

## The three levels

| Level | Compares | How lines are paired | Trust |
|---|---|---|---|
| **1** | each WO operation with the master of **its own revision**, as that master is today | same sequence and same operation name - exact, since the WO copied that revision line by line | firm |
| **2** | - | one row per WO built on an **older** revision than the current one: the routing itself changed, a planner reviews the whole order | firm |
| **3** | older-revision WO operations with the **current** revision's master | a new revision usually inserts a step and shifts later ones, so sequence numbers do not match. Lines are paired only when the name is unique in both revisions, when all its occurrences carry the same values, or when each occurrence has its own resource. Everything else is reported as not compared | **indicative** - assumes the same name is the same step; to be reviewed |

### When `USER_8` is out of date

`USER_8` is only a label. An order can carry the current revision's routing
while `USER_8` still names an older one. When every not-started line of an
older-revision order matches the current master - sequence, operation name
and every compared value - and they do not all match its own revision, the
order is treated as being on the current revision: it is compared with the
current master and listed on `07_StaleUser8`. Without this, such orders
showed false differences and false "no master line" rows.

## Running it

Every time, on the same day:

| Step | What | Save as |
|---|---|---|
| 1 | run `master_drift.sql` | `..\factory_work_order_master_drift_private\data\drift.xlsx` |
| 2 | run `old_revision_masters.sql` | `..\factory_work_order_master_drift_private\data\masters.xlsx` |
| 3 | `python master_drift.py` (in the `fwomd` environment) | - |

Paste each query's result into the first sheet with its header row, and
overwrite the previous files. Both queries must run on the same day: masters
change.

Both queries write resource IDs as `RES:034`, not `034`: pasted into Excel, an
ID with a leading zero turned into a number and lost the zero. The script
strips exactly the `RES:` marker, so the list shows `034`, and it stops if a
resource has no marker - an export from an older version of the query.

The list is written to
`..\factory_work_order_master_drift_private\outputs\master_drift_list_<date>.xlsx`.
`--drift`, `--masters`, `--out` and `--previous` override the paths.

### The SQL updater list

`01_SqlUpdater` holds one row per open, not-started line that does **not**
simply match its master (lines that match are only counted). Each row has
exactly one `SCENARIO`, decided in this order:

| `SCENARIO` | Meaning | Basis |
|---|---|---|
| `No master line` | no master line to compare with - a step added to the order (rework `S RIL ...`, test samples) | not comparable |
| `Work order matches current master` | an older-revision line that differs from its own old revision, while every value equals the current master | nothing to change on the order |
| `Resource changed` | the resource differs from the master | 28 Sep review: 9 of 13 intentional, 4 not - always looked at |
| `Work order to update` | same resource, another value differs | 28 Sep review: 51 of 51 - the order follows the master |
| `Differs from current master (WO on older revision)` | matches its own old revision, but the current master has a different value | level 3 pairing - indicative |

Columns, in order:

| Columns | Content |
|---|---|
| `WO_OPS_ROWID`, `M_OPS_ROWID`, `C_MASTER_OPS_ROWID` | the order's operation line, the master line it is compared with, the current master line |
| `FAMILY`, `PART_ID`, `PART_DESCRIPTION` | which part |
| `WO`, `WO_STATUS`, `WO_CREATE_DATE`, `WO_DESIRED_RLS_DATE` | which order |
| `OPERATION`, `STEP` | which step (`STEP` = sequence number) |
| `WO_REV`, `M_REV`, `C_REV` | the order's revision (`USER_8`), the one it is compared with, today's |
| `WO_/M_/C_` + `RES_ID`, `RUN_TYPE`, `RUN`, `LSQ` (load size), `SETUP_HRS`, `CREW` | the order's, its master's and the current master's values side by side |
| `DIFF_KEY` | what differs, each value tagged with its source, e.g. `RUN: 60 (WO) \| 70 (M); SETUP_HRS: 0 (WO) \| 1 (M)` - `(C)` for the current master on the "differs from current master" scenario. Also what recognises the line next time (the plain form of lists before 12 Oct 2026, `RUN=60\|70`, is still read) |
| `SCENARIO` | see above |
| `SQL_UPDATE` | the statement for the order - see below |

For every difference in `DIFF_KEY` the two cells involved are coloured: the
order's value (`WO_`) light orange, the value it is compared with light green -
`M_`, or `C_` on `Differs from current master`. `Work order to update` rows
are light grey across the whole row, under those two colours. The colours are set when the
list is written and do not follow later edits in the sheet.

`M_` is the master the line was compared with: its own revision, or the current one
for an order whose `USER_8` is out of date (then `M_REV` = `C_REV`). For an
order on the current revision `C_` = `M_`.

**`SQL_UPDATE`** is written only on `Work order to update` lines:

```
UPDATE OPERATION SET RUN = '90', SETUP_HRS = '0.5' WHERE ROWID = <WO_OPS_ROWID>;
```

- it sets only the fields that differ, to the `M_` value - never the resource;
- it is a **spreadsheet formula**, not text: it reads the `M_` cells and
  `WO_OPS_ROWID` of its own row, so correcting a value in the sheet corrects
  the statement; the column letters are derived from the column order;
- it is left **empty** when a master value or `WO_OPS_ROWID` is blank, so
  nothing can blank a field by accident;
- the other scenarios get no statement: a resource change needs a person,
  the current master is only paired indicatively, and the rest has nothing to
  copy. Nothing here runs anything - it produces the SQL and leaves running
  it to a person who has read it.

Leave the list where it is: the next run reads the most recently saved list
in the `outputs` folder, other than the one it is about to overwrite (or
`--previous <file>`), and fills `02_SinceLastList`,
matching lines on order, step, operation and the same differences. Lists of
every earlier layout are still read.

### Output sheets

| Sheet | Content |
|---|---|
| `00_Summary` | counts per scenario, checks, input files, what happened since the previous list |
| `01_SqlUpdater` | **the list**, with the update statements - see above |
| `02_SinceLastList` | what happened to each line of the previous list that is not listed again unchanged: **fixed** (order updated / master updated / both changed / now matches), **started, closed or removed** (mostly produced on the values it had - the export cannot tell a started step from a deleted one), **step replaced by another operation** (same step number, another operation name: removed or renamed, not fixed), still listed but changed, or moved by the `USER_8` rule. Shows whether the list is acted on |
| `03_OldRevisionWOs` | level 2, one row per work order on an older revision, with its level 1 and level 3 counts |
| `04_CrossRevision` | level 3, every older-revision line with its pairing category and, where compared, the current master's values (`CUR_*`, `D3_*`) |
| `05_StaleUser8` | orders moved onto the current routing because `USER_8` is out of date, with what they showed before |

### Exit codes and checks

| Code | Meaning |
|---|---|
| 0 | clean |
| 1 | stopped: an input is missing, an operation appears twice (a join fanned out), or the level 1 flags from the SQL disagree with the script's own recomputation |
| 2 | the list was written, but something needs a look - explained on a `!!` line: the two exports come from different days, or the masters export does not cover every older-revision line |

The script refuses to write inside this repository.

## Data never goes in this repository

The exports and the list name real parts, work orders and projects. They live
in the private folder next to this one, which is never a repository; the
`.gitignore` and a commit hook keep data files out of here. Commit messages
describe changes by shape, never by part, order, project or family.

## Environment

A conda environment named `fwomd`, built from Python plus `requirements.txt`
alone:

```
conda create -n fwomd -c conda-forge python=3.12.14 -y
conda activate fwomd
python -m pip install -r requirements.txt
```

Outside the Anaconda prompt, call the environment's Python by its full path
(`...\miniconda3\envs\fwomd\python.exe`) - plain PowerShell has no conda hook.

`requirements.txt` holds floors and the proven versions;
`requirements-lock.txt` pins every package `fwomd` was proven on
(2026-09-28: identical output to the previous environment in every cell).
