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

The list is written to
`..\factory_work_order_master_drift_private\outputs\master_drift_list_<date>.xlsx`.
`--drift`, `--masters` and `--out` override the paths.

### Output sheets

| Sheet | Content |
|---|---|
| `00_Summary` | counts, checks, input files |
| `01_Drift` | level 1: lines differing from their own revision's master |
| `02_NoMasterLine` | WO lines with no own-revision master line (steps added to the order, e.g. rework `S RIL ...`) |
| `03_OldRevisionWOs` | level 2, one row per work order, with its level 1 and level 3 counts |
| `04_CrossRevision` | level 3, every older-revision line with its pairing category and, where compared, the current master's values |
| `05_Review` | level 1 grouped into patterns - family, project, operation, field, work-order value, master value - with the lines, orders and parts behind each, and empty `VERDICT` / `NOTE` columns to fill in. Start here |

Columns: `WO_*` is the work order's value, `M_*` the own-revision master's,
`CUR_*` the current master's; `D_*` / `D3_*` are 1 where they differ.

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
