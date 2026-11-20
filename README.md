# Work-order master drift

#### -- Project Status: [Completed]

## Objective

Lists open, not-started work-order operations whose values no longer match
their master routing.

A work order **copies** its master's operations when it is created and never
looks back. When a master is improved later - a new run rate, a new setup
time - the work orders already created keep the old values, so planning and
production run on an outdated standard until they finish. This tool finds
them. It is **read-only**: it produces a list; nothing is written to the ERP.

It follows on from
[factory_time_statistics](https://github.com/mikjf/factory_time_statistics),
which proposes new master values: key the new masters, run this, see which
open orders still carry the old ones.

### What this project demonstrates

- **Comparing an ERP's copies against their source:** every open operation is
  paired with the master line it was copied from, even when the master was
  renumbered or revised since, and every pairing that cannot be trusted is
  counted rather than guessed
- **A reviewable output, not an automatic one:** an Excel list where every
  difference is coloured, every row has one scenario, and each update is a
  spreadsheet formula that writes the SQL `UPDATE` from the cells beside it -
  nothing runs anything
- **Follow-up across runs:** each list is compared with the previous one, so it
  shows what was fixed, what is new and how long each line has been waiting
- **Safe database access:** read-only, Windows login, both queries saved
  together or not at all, and every read gives way to ERP users
- **Self-checks:** the SQL's own flags are recomputed in Python, fan-out and
  type errors stop the run, and exit codes say when a person must look

### Tech stack

- Python 3.12, pandas, openpyxl, pyarrow
- pyodbc and SQL Server (T-SQL) on Infor VISUAL ERP

### Data

The extracts, the exports and the list name real parts, work orders and
projects. They live in a private folder next to this repository, which is never
a repository itself; the `.gitignore` and a commit hook keep data files out of
here.

## Terms

| Term | Meaning here |
|---|---|
| **Master** | the engineering routing of a part: `WORK_ORDER.TYPE = 'M'` and its operations |
| **Work order (WO)** | a production order, `TYPE = 'W'`, with its own copy of the master's operations |
| **Revision** | a version of a part's master, held in the master's `LOT_ID`. A WO's `USER_8` says which revision it was built from; `PART_SITE.ENGINEERING_MSTR` says which one is current. Revision labels (`A`, `7`, `ZZ` ...) cannot be ordered - "older" means *different from the current one* |
| **Open, not started** | WO status firmed or released (`F`, `R`), operation status `F` or `R`, and no labour ticket of any kind on the operation |
| **Drift** | a value on the WO's operation that differs from the master it follows: its own master on the current revision, the **current** revision for an order on an older one (PM and SC, 2026-10-21) |

Compared fields: planned resource, `RUN_TYPE`, `RUN`, `LOAD_SIZE_QTY`,
`SETUP_HRS` and `USER_5` (planned crew). An empty crew against a filled one
is a difference (since 2026-10-23: the crew must be filled in correctly;
until then empty counted as 1); empty against empty is equal. To recognise
which routing an order carries (the `USER_8` rule in [DESIGN.md](DESIGN.md), level 3), an empty
crew still counts as 1, so a missing crew does not hide the routing. `RUN_HRS` and
`CALC_START_QTY` depend on the order quantity and are not compared.

## Scope

- Work orders numbered `.../OP` (production) and `.../OPP` (production on a
  new project). Rework (`OR`) and `OS` are out: rework routings are ad hoc.
- Parts `M.%`.
- Unreleased (`U`) work orders are out - in practice old orders set back.
- External operations (names starting `E `) are out.
- Blocked work orders (`OdP Bloccato` ticked in the order header) are out
  (since 2026-11-03). The query flags them as `WO_BLOCKED`, and the script
  leaves them out and counts them in `00_Summary`.
- 4-digit sequences only. Work orders from before the masters were
  renumbered keep 2- or 3-digit sequences and are not covered yet.

## The three levels

| Level | Compares | How lines are paired | Trust |
|---|---|---|---|
| **1** | each WO operation with the master of **its own revision**, as that master is today | same sequence and same operation name - exact, since the WO copied that revision line by line; failing that, the same name when it is unique in the master and in the order (master renumbered) | firm |
| **2** | - | one row per WO built on an **older** revision than the current one: the routing itself changed, a planner reviews the whole order | firm |
| **3** | older-revision WO operations with the **current** revision's master | a new revision usually inserts a step and shifts later ones, so sequence numbers do not match. Lines are paired only when the name is unique in both revisions, when all its occurrences carry the same values, or when each occurrence has its own resource. Everything else is reported as not compared | **firm where paired** - since 2026-10-21 the current revision is the reference, so these pairings drive the list for older-revision orders: the order's line is matched to its own revision (level 1), and the name is unambiguous in both revisions. The "name repeats with different values" lines stay unpaired - to be reviewed |

## Running it

One command, in the `fwomd` environment:

```
python master_drift.py
```

It runs both queries against the ERP through `extract.py`, then builds the
list. Built the same way as `factory_planner`'s extract:

- **Read-only, Windows login.** No user name or password is stored. Server
  and database names are in `..\factory_shared_private\connections\visual.ini`,
  never in this repository:

  ```
  [database]
  server = ...
  database = ...
  driver = ODBC Driver 18 for SQL Server
  ```

- **Both queries together, all or nothing.** The results are saved as
  `master_drift.parquet` and `old_revision_masters.parquet` in
  `..\factory_work_order_master_drift_private\data`, with
  `extract_manifest.json` (when it ran, rows, columns, and a SHA-256 of each
  query's text). The previous extract is replaced only once both succeeded.
- **It gives way to ERP users.** If a read collides with someone saving in the
  ERP, the read is cancelled, never their work, and retried up to three times.
- **Text stays text.** Resource IDs such as `034` keep their leading zero -
  pasting into Excel had turned them into `34`.

`python master_drift.py --no-extract` builds the list again from the last
extract without querying the ERP; it warns if a query file changed since.
`python extract.py` runs only the extract.

The list is written to
`..\factory_work_order_master_drift_private\outputs\master_drift_list_<date>.xlsx`.
Manual Excel exports still work as input, and `--out`, `--previous` and
`--config` override the paths - see [DESIGN.md](DESIGN.md#manual-exports-and-paths).

## Output

| Sheet | Content |
|---|---|
| `00_Legend` | how to read the list - scenarios, colours, columns - for someone who did not build it. The workbook opens on `01_SqlUpdater` |
| `00_Summary` | counts per scenario, the lines not listed by reason (no master line; older revision not paired with the current one), lines paired by name, checks, input files, what happened since the previous list |
| `01_SqlUpdater` | **the list**, with the update statements - see [DESIGN.md](DESIGN.md#the-sql-updater-list) |
| `02_SinceLastList` | what happened to each line of the previous list that is not listed again unchanged: **fixed - order updated** (the order now carries the value the list proposed), **partly fixed - still listed** (some values corrected, the rest still listed - `DIFF_KEY_NOW` shows what is left), fixed - master updated / both changed / now matches, **started, closed or removed** (mostly produced on the values it had - the export cannot tell a started step from a deleted one), **step replaced by another operation** (same step number, another operation name: removed or renamed, not fixed), **not compared - current revision not paired** (an older-revision order level 3 cannot pair: not listed, but not known to match), still listed but changed, or moved by the `USER_8` rule. Lines of scenarios no longer listed are not followed up. Shows whether the list is acted on |
| `03_OldRevisionWOs` | level 2, one row per work order on an older revision, with its level 1 and level 3 counts |
| `04_CrossRevision` | level 3, every older-revision line with its pairing category and, where compared, the current revision's values (`CUR_*`, `D3_*`) - the basis of `Update - WO differs from master (current rev)` |
| `05_StaleUser8` | orders moved onto the current routing because `USER_8` is out of date, with what they showed before |

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

The database route also needs Microsoft's **ODBC Driver 18 for SQL Server**
installed on the machine (`pyodbc` talks to the database through it).

## Design notes

How an order with an out-of-date `USER_8` is handled, the SQL updater list in
full, and the exit codes are in [DESIGN.md](DESIGN.md).
