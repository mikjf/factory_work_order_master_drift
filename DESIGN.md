# Design notes

The detail behind the list described in [README.md](README.md).

## When `USER_8` is out of date

`USER_8` is only a label. An order can carry the current revision's routing
while `USER_8` still names an older one. When every not-started line of an
older-revision order matches the current master - sequence, operation name
and every compared value - and they do not all match its own revision, the
order is treated as being on the current revision: it is compared with the
current master and listed on `05_StaleUser8`. Without this, such orders
showed false differences and false "no master line" rows. The rule counts
only exact matches (same step and name), not lines paired by name.

## Manual exports and paths

The manual route of Excel exports still works, for older exports:
`--drift <file> --masters <file>`. It warns if the two were saved on different
days, or if their resource IDs were turned into numbers. Exports of 13-14 Oct
2026 carry resource IDs as `RES:034`; the marker is stripped.

The list is written to
`..\factory_work_order_master_drift_private\outputs\master_drift_list_<date>.xlsx`.
`--out`, `--previous` and `--config` override the paths.

## The SQL updater list

`01_SqlUpdater` holds one row per open, not-started line that does **not**
match the master it follows (lines that match are only counted).

**Lines without a master line are not listed** (since 2026-10-21, after
PM and SC): steps added to the order on the go - rework `S RIL ...`, samples -
have nothing to compare with. `00_Summary` counts them by reason: the step is
not in the master; its name repeats in the master or the order, so it is not
paired; or the order's revision has no master.

**A step renumbered in the master is still compared.** A master can be
renumbered after an order was created (a step inserted, the later ones moved
by 100 or 200), so the order's step number is no longer in it. When the
operation name occurs exactly once in the master and exactly once among all
the order's operations, the query pairs the two lines by name
(`MATCHED_BY = UNIQUE_NAME`); `M_STEP` then shows the master's step next to
the order's `STEP`. Before, such lines showed as "no master line" and hid
real differences.

**The reference is the current revision** (PM and SC, 2026-10-21). An order
created on an older revision follows the **current** revision of the master -
the one `PART_SITE.ENGINEERING_MSTR` names - not the one it was created from:
an order still carrying its old revision's rate is brought to the current
one, and an order that already carries the current rate is left alone. So:

- an order on the current revision is compared with its master line (`M_`),
  paired exactly;
- an order on an older revision is compared with the current revision (`C_`),
  through the level 3 pairing of its exact own-revision line: unique name,
  identical occurrences, or each occurrence on its own resource. Lines level 3
  cannot pair - the name repeats with different values, the step no longer
  exists in the current revision - are counted by reason in `00_Summary`, not
  listed: their current value is unknown.

Each listed row has exactly one `SCENARIO`:

| `SCENARIO` | Meaning | Basis |
|---|---|---|
| `Update - WO differs from master (WO rev)` | order on its master's revision (the current one); same resource, another value differs from that master | 28 Sep review: 51 of 51 - the order follows the master |
| `Update - WO differs from master (current rev)` | order on an older revision; same resource, another value differs from the current revision | PM and SC, 2026-10-21 |
| `Review - OPP WO differs from master (WO rev)`, `Review - OPP WO differs from master (current rev)` | the same two cases on an **OPP** order (a new project) | PM and SC, 2026-10-23: OPP updates are always reviewed by a person first |
| `Resource changed` | the resource differs from the master the line follows | 28 Sep review: 9 of 13 intentional; SC: "should be" intentional - kept listed |

The names say which master the order differs from, not why: the cause may be
a master corrected later, a master wrong when the order was created, or an
order edited (user, 2026-10-23). Earlier names: `Work order to update` and
`Work order to update (current revision)`; before 2026-10-21 also `Differs
from current master (WO on older revision)` (the current-rev case) and `Work
order matches current master`, now not listed: such an order is correct as
it is. The `00_Legend` sheet explains scenarios, colours and columns to the
reader.

Columns, in order:

| Columns | Content |
|---|---|
| `WO_OPS_ROWID`, `M_OPS_ROWID`, `C_MASTER_OPS_ROWID` | the order's operation line, its own master line, the current revision's line |
| `FAMILY`, `PART_ID`, `PART_DESCRIPTION` | which part |
| `WO`, `WO_STATUS`, `WO_CREATE_DATE`, `WO_DESIRED_RLS_DATE` | which order |
| `OPERATION`, `STEP`, `M_STEP` | which step: `STEP` = the order's sequence number, `M_STEP` = the master line's (different only when paired by name) |
| `WO_REV`, `M_REV`, `C_REV` | the order's revision (`USER_8`), its master's, the current one - `WO_REV` ≠ `C_REV` marks an order on an older revision |
| `WO_/M_/C_` + `RES_ID`, `RUN_TYPE`, `RUN`, `LSQ` (load size), `SETUP_HRS`, `CREW` | the order's, its own master's and the current revision's values side by side |
| `DIFF_KEY` | what differs, each value tagged with its source, e.g. `RUN: 60 (WO) \| 70 (M); SETUP_HRS: 0 (WO) \| 1 (M)` - `(C)` when compared with the current revision (an order on an older revision). Also what recognises the line next time (the plain form of lists before 12 Oct 2026, `RUN=60\|70`, is still read) |
| `SCENARIO` | see above |
| `FIRST_LISTED` | the day of the first saved list (in `outputs`) that showed this line with these same differences - how long it has been waiting. Worked out from the saved lists; nobody fills it in. A line whose differences change starts again |
| `SQL_UPDATE` | the statement for the order - see below |

For every difference in `DIFF_KEY` the two cells involved are coloured: the
order's value (`WO_`) orange, the value it is compared with green - `M_`, or
`C_` when tagged `(C)`. Under them, the whole row is coloured by scenario, in
a **strong** shade when the line is **new** since the previous list (or its
differences changed) and a **light** shade when it was already in the previous
list with the same differences (since 2026-10-28):

| Scenario | New | Already listed |
|---|---|---|
| `Update - ...` | purple | light purple |
| `Review - ...` | blue | light blue |
| `Resource changed` | yellow | light yellow |

The list is ordered by these colours, top to bottom as in the table (new
before already listed in each scenario group), and inside each block by
`FIRST_LISTED`, newest first - the longest-waiting lines at the bottom.

Light means "not new", not "approved". Only the purple lines are candidates
for automating the update later; blue lines always need a person. Until
2026-10-28 update rows were grey and the shade did not change. The colours
are set when the list is written and do not follow later edits in the sheet.

`M_` is the order's own master: its own revision, or the current one for an
order whose `USER_8` is out of date (then `M_REV` = `C_REV`). For an order on
the current revision `C_` = `M_`.

**`SQL_UPDATE`** is written on the `Update - ...` and `Review - ...` scenarios
(on a review line it is ready for after the review):

```
UPDATE OPERATION SET RUN = '90', SETUP_HRS = '0.5' WHERE ROWID = <WO_OPS_ROWID>;
```

- it sets only the fields that differ, to the value the line is compared
  with - `M_` on `(WO rev)`, `C_` on `(current rev)` - never the resource,
  never the step number;
- it is a **spreadsheet formula**, not text: it reads those cells and
  `WO_OPS_ROWID` of its own row, so correcting a value in the sheet corrects
  the statement; the column letters are derived from the column order;
- it is left **empty** when a value or `WO_OPS_ROWID` is blank, so nothing
  can blank a field by accident;
- `Resource changed` gets no statement: a resource change needs a person.
  Nothing here runs anything - it produces the SQL and leaves running it to
  a person who has read it.

Leave the list where it is: the next run reads the most recently saved list
in the `outputs` folder, other than the one it is about to overwrite (or
`--previous <file>`), and fills `02_SinceLastList`,
matching lines on order, step, operation and the same differences. Lists of
every earlier layout are still read.

## Exit codes and checks

| Code | Meaning |
|---|---|
| 0 | clean |
| 1 | stopped: the extract failed (the previous one is kept), an input is missing, an operation appears twice (a join fanned out), a numeric field holds text, or the level 1 flags from the SQL disagree with the script's own recomputation |
| 2 | the list was written, but something needs a look - explained on a `!!` line: a query changed since the extract used (`--no-extract`), manual exports from different days or with resource IDs turned into numbers, or the masters result does not cover every older-revision line |

The script refuses to write inside this repository.
