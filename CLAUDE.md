# factory_work_order_master_drift

Lists open, not-started work-order operations whose values no longer match their master routing.
Read-only: it produces a list; nothing is written to the ERP.

## Shared rules - read first

The rules every factory project follows (databases, SQL, data and people, ERP
facts, how to answer, git) are in
`~\GitHub\factory_shared_private\RULES.md`, imported here:

@~/GitHub/factory_shared_private/RULES.md

If that import did not load, read the file before anything else. What follows
is specific to this project.

## This project

- Read first: `README.md`, `master_drift.sql`; ERP notes in the private `ERP_KNOWLEDGE.md`, open points in `OPEN_QUESTIONS.md`.
- Private twin: `~\GitHub\factory_work_order_master_drift_private\` (data, outputs, notes that
  name real parts; never a git repository).
- Python: conda env `fwomd` (`%LOCALAPPDATA%\miniconda3\envs\fwomd\python.exe`), built from
  `requirements.txt`; `requirements-lock.txt` pins what it was proven on.
- Saved memory: `~\.claude\projects\<encoded path of the repo>-factory-work-order-master-drift\memory\`.
