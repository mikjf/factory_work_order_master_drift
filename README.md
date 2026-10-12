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

The list is written to
`..\factory_work_order_master_drift_private\outputs\master_drift_list_<date>.xlsx`.
`--drift`, `--masters`, `--out` and `--previous` override the paths.

### The drift list

`01_DriftList` holds one row per open, not-started line that does **not**
simply match its master (lines that match are only counted). Each row has
exactly one `SCENARIO`, decided in this order:

| `SCENARIO` | Meaning | Basis |
|---|---|---|
| `No master line` | no master line to compare with - a step added to the order (rework `S RIL ...`, test samples) | not comparable |
| `Order already matches today's master` | an older-revision line that differs from its own old revision, while every value equals the current master | nothing to change on the order |
| `Resource changed - check` | the resource differs from the master | 28 Sep review: 9 of 13 intentional, 4 not - always looked at |
| `Order to update` | same resource, another value differs | 28 Sep review: 51 of 51 - the order follows the master |
| `Differs from today's master (older revision, indicative)` | matches its own old revision, but the current master has a different value | level 3 pairing - indicative |

`LABEL` names exactly what differs - resource, rate (`RUN`), setup
(`SETUP_HRS`), run type (`RUN_TYPE`), load size (`LOAD_SIZE_QTY`), crew
(`USER_5`) - e.g. `same resource: rate + setup` or `different resource,
nothing else`; for the last scenario it describes the difference with
today's master. `REVISION` describes the order: `current`, `older`, or
`older - USER_8 out of date`.

Columns, in order: `SCENARIO`, `LABEL`, `REVISION`, `WORKORDER_OPS_ROWID`
(the order's operation line), `MASTER_OPS_ROWID` (the master line it is
compared with), `CUR_MASTER_OPS_ROWID` (today's master line, older
revisions), `FAMILY`, `PART_ID`, `PART_DESCRIPTION`, `PROJECT_CODE`,
`WO_KIND`, `WORK_ORDER`, `WO_STATUS`, `CREATE_DATE`, `DESIRED_RLS_DATE`,
`WO_REV`, `CURRENT_REV`, `OPERATION`, `STEP` (sequence number), then for
each compared field the order's value (`WO_*`), its master's (`M_*`) and
today's master's (`CUR_*`), `DIFFERENCES` in words, `NOTE`, and `DIFF_KEY`
(the differences in machine-readable form, used to recognise the line next
time).

`NOTE` is free text. Save the list where it is: the next run reads the most
recently saved list in the `outputs` folder (or `--previous <file>`) and
carries each note to the same work order, step, operation and the same
differences. Lists made before 12 Oct 2026 had verdict columns; their
verdicts arrive in `NOTE` as `earlier verdict: ...`.

### Output sheets

| Sheet | Content |
|---|---|
| `00_Summary` | counts per scenario and per label, checks, input files, what happened since the previous list |
| `01_DriftList` | **the list** - see above |
| `02_SinceLastList` | what happened to each line of the previous list that is not listed again unchanged: **fixed** (order updated / master updated / both changed / now matches), **started or closed** (produced on the values it had), still listed but changed, or moved by the `USER_8` rule. Shows whether the list is acted on |
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
