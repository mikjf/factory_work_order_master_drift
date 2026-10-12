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
`--drift`, `--masters` and `--out` override the paths.

### The three lists

Every work-order line that differs from its master lands on exactly one
list, one row per line. Its `LABEL` names exactly what differs - resource,
rate (`RUN`), setup (`SETUP_HRS`), run type (`RUN_TYPE`), load size
(`LOAD_SIZE_QTY`), crew (`USER_5`) - e.g. `same resource: rate + setup` or
`different resource, nothing else`. `DIFFERENCES` spells out the values
("rate: order 60 / master 50"). The rules, checked in this order, come from
the first review (28 Sep 2026, 65 judged lines):

| Order | Rule | List | Basis |
|---|---|---|---|
| 1 | Order already matches today's master (an older-revision line whose values equal the current master; the difference is only with its own old revision) | `03_Information` | nothing to change on the order |
| 2 | Same resource, other values differ | `01_OrdersToUpdate` | 51 of 51 decided lines: the order follows the master |
| 3 | Different resource | `02_ResourceCheck` | 9 of 13 were intentional, 4 were not - always looked at |

`00_Summary` counts every label with the verdicts it has received so far,
so a label that starts behaving differently stands out - and can then get
its own rule.

### Verdicts - for exceptions only

`VERDICT` is empty by default. Fill it only where a line is an exception -
an "update" line that should not be updated, a resource change confirmed or
found wrong - with exactly one keyword:

| `VERDICT` | Meaning | What follows |
|---|---|---|
| `MASTER_DRIFT` | the master changed after the order was created | update the order |
| `ORIG_MASTER_WRONG` | the master was already wrong when the order was created and has been corrected since | update the order |
| `INTENTIONAL_CHANGE` | the order was changed on purpose (e.g. by supply or the PM) | nothing on the order - but it stays listed while order and master differ: the master may be the one to change |
| `WO_WRONG` | the order was edited by mistake; the master did not change | correct the order |
| `MASTER_WRONG` | the order is right; the master is wrong today | correct the master |
| `UNCLEAR` | asked, no clear answer yet | ask again |

Anything else goes in `NOTE` (who said it, when). Save the list where it
is. The next run reads the most recently saved list in the `outputs` folder
(or `--verdicts <file>`) and carries each verdict and note forward to the
same work order, sequence, operation and the same differences;
`VERDICT_FROM` says where it came from. When a difference changes, the line
is open again.

Every verdict ever given is kept on `09_VerdictHistory` and passed on from
list to list, also after its line has gone (verdicts given on orders later
found to have an out-of-date `USER_8` are dropped: they were judged against
the wrong revision). When an open line shows a difference already judged on
another order - same operation and same differences - `SUGGESTED_VERDICT`
proposes the earlier verdict (or `MIXED: ...`) and `SUGGESTION_BASIS` says
how often and with which note. It is only a suggestion.

Lists and reviews made before 12 Oct 2026 had one row per field; they are
still read, their verdicts combined per line.

### Output sheets

| Sheet | Content |
|---|---|
| `00_Summary` | counts, checks, input files, one line per label with its verdicts so far |
| `01_OrdersToUpdate` | **for planners:** same resource, other values differ - bring the order to the master |
| `02_ResourceCheck` | the resource differs - check whether intentional |
| `03_Information` | order already matches today's master - nothing to do |
| `04_Drift` | level 1 detail: every differing line with all values (`WO_*` order, `M_*` master, `D_*` = 1 where they differ) |
| `05_NoMasterLine` | WO lines with no own-revision master line (steps added to the order, e.g. rework `S RIL ...`) |
| `06_OldRevisionWOs` | level 2, one row per work order, with its level 1 and level 3 counts |
| `07_CrossRevision` | level 3, every older-revision line with its pairing category and, where compared, the current master's values (`CUR_*`, `D3_*`) |
| `08_StaleUser8` | orders moved onto the current routing because `USER_8` is out of date, with what they showed before |
| `09_VerdictHistory` | every line ever given a verdict, with its note and the list it was judged in - the memory behind the suggestions and the label counts |

### Exit codes and checks

| Code | Meaning |
|---|---|
| 0 | clean |
| 1 | stopped: an input is missing, an operation appears twice (a join fanned out), or the level 1 flags from the SQL disagree with the script's own recomputation |
| 2 | the list was written, but something needs a look - explained on a `!!` line: the two exports come from different days, the masters export does not cover every older-revision line, the previous list has `VERDICT` text that is not one of the keywords (it is kept in `NOTE` and the line is open again), or an old one-row-per-field review gave one line different verdicts on its fields (the line is left open, to decide once) |

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
