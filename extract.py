"""Run master drift's two queries against the ERP and save each result as Parquet.

    python extract.py                 both queries, into the private data folder
    python extract.py --out <folder>  somewhere other than the default

master_drift.py calls this itself before building the list, so it is rarely
run on its own. Built the same way as factory_planner/extract.py (2026-10-15).

Replaces pasting query results into Excel, which turned text IDs into numbers:
resource '034' became 34. Read straight from the database, every column keeps
the type SQL Server gives it, so the queries stay exactly as written - no
markers, no casts for Excel's sake.

Connection settings come from
../factory_work_order_master_drift_private/extract.ini (or --config), never
from this repository:

    [database]
    server = ...
    database = ...
    driver = ODBC Driver 18 for SQL Server

Windows authentication only: no password is read or stored. The connection is
opened read-only, and the queries are SELECTs.

Output, by default in ../factory_work_order_master_drift_private/data/:

    master_drift.parquet           result of master_drift.sql
    old_revision_masters.parquet   result of old_revision_masters.sql
    extract_manifest.json          when it ran, and per query its rows, columns
                                   and the SHA-256 of the SQL text that produced it

ALL OR NOTHING. Every query is written to a temporary file first; the previous
extract is replaced only once both have succeeded, so the folder never holds
results from two different runs - the queries are read together, and masters
change during the day.

IT YIELDS TO ERP USERS. The session runs at DEADLOCK_PRIORITY LOW, so when a
read collides with someone saving in the ERP, the read is the one cancelled;
it is retried up to three times, 10 s apart, before the run gives up.

Types: SQL Server DECIMAL/NUMERIC arrive as Python Decimal and are stored as
float64. Everything else is stored as Arrow infers it: text as string,
integers as int64 (NULLs stay NULL), dates as date, datetimes as timestamp.

The script refuses to write inside this repository.
"""
import argparse
import configparser
import datetime as dt
import decimal
import hashlib
import json
import pathlib
import sys
import time

import pyarrow as pa
import pyarrow.parquet as pq
import pyodbc

HERE = pathlib.Path(__file__).resolve().parent
PRIVATE = HERE.parent / f"{HERE.name}_private"
QUERIES = ("master_drift.sql", "old_revision_masters.sql")
MANIFEST = "extract_manifest.json"
FETCH_ROWS = 50_000
DEADLOCK_STATE = "40001"    # SQLSTATE of "chosen as the deadlock victim"
DEADLOCK_TRIES = 4          # the first attempt and three retries
DEADLOCK_WAIT = 10          # seconds between attempts


def parse_args():
    p = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    p.add_argument("--config", type=pathlib.Path, default=PRIVATE / "extract.ini")
    p.add_argument("--out", type=pathlib.Path, default=PRIVATE / "data")
    return p.parse_args()


def connect(config_path):
    if not config_path.exists():
        sys.exit(f"{config_path} not found - it holds the server and database "
                 f"names (see the docstring of extract.py)")
    cfg = configparser.ConfigParser()
    cfg.read(config_path, encoding="utf-8")
    db = cfg["database"]
    cn = pyodbc.connect(
        f"DRIVER={{{db.get('driver', 'ODBC Driver 18 for SQL Server')}}};"
        f"SERVER={db['server']};DATABASE={db['database']};"
        "Trusted_Connection=yes;TrustServerCertificate=yes;",
        readonly=True)
    # The extract always yields to ERP users: when one of its reads deadlocks
    # with someone saving in the ERP, LOW makes SQL Server cancel this session,
    # never the user's work - the extract then retries (see run_query).
    cn.execute("SET DEADLOCK_PRIORITY LOW")
    return cn


def column(values):
    """One result column as an Arrow array; Decimal becomes float64."""
    if any(isinstance(v, decimal.Decimal) for v in values):
        return pa.array([None if v is None else float(v) for v in values],
                        type=pa.float64())
    return pa.array(values)


def run_query(cn, sql, name):
    """Run one query; retry it when it was chosen as a deadlock victim."""
    for attempt in range(1, DEADLOCK_TRIES + 1):
        try:
            return _fetch(cn, sql)
        except pyodbc.Error as e:
            if e.args[0] != DEADLOCK_STATE or attempt == DEADLOCK_TRIES:
                raise
            print(f"  {name}: gave way to an ERP user (deadlock), "
                  f"retry {attempt} of {DEADLOCK_TRIES - 1} in {DEADLOCK_WAIT} s")
            time.sleep(DEADLOCK_WAIT)


def _fetch(cn, sql):
    cur = cn.cursor()
    cur.execute(sql)
    names = [d[0] for d in cur.description]
    if len(set(names)) != len(names):
        sys.exit(f"duplicate column names in the result: {names}")
    rows = []
    while batch := cur.fetchmany(FETCH_ROWS):
        rows.extend(batch)
    cur.close()
    cols = list(zip(*rows)) if rows else [[] for _ in names]
    return pa.table({n: column(list(c)) for n, c in zip(names, cols)})


def extract(out, config):
    out = out.resolve()
    if out == HERE or HERE in out.parents:
        sys.exit(f"refusing to write inside the repository: {out}")
    out.mkdir(parents=True, exist_ok=True)

    started = dt.datetime.now().astimezone()
    manifest = {"started": started.isoformat(timespec="seconds"), "queries": {}}
    written = []
    cn = connect(config)
    try:
        for name in QUERIES:
            q = HERE / name
            sql = q.read_text(encoding="utf-8-sig")
            t0 = time.perf_counter()
            table = run_query(cn, sql, q.name)
            tmp = out / f"{q.stem}.parquet.tmp"
            pq.write_table(table, tmp)
            written.append((tmp, out / f"{q.stem}.parquet"))
            secs = time.perf_counter() - t0
            manifest["queries"][q.name] = {
                "parquet": f"{q.stem}.parquet",
                "rows": table.num_rows,
                "columns": {f.name: str(f.type) for f in table.schema},
                "sql_sha256": hashlib.sha256(sql.encode("utf-8")).hexdigest(),
                "seconds": round(secs, 1),
            }
            print(f"  {q.name:<28} {table.num_rows:>9,} rows  {secs:6.1f} s")
    except BaseException:
        for tmp, _ in written:
            tmp.unlink(missing_ok=True)
        print("extract failed - the previous extract is untouched", file=sys.stderr)
        raise
    finally:
        cn.close()

    for tmp, final in written:
        tmp.replace(final)
    # A query renamed or removed leaves its old Parquet behind, which would sit
    # beside this run's files looking current. The output folder holds exactly
    # one extract, so any Parquet this run did not write goes. (Excel exports
    # of the manual route are left alone.)
    current = {final for _, final in written}
    for stale in out.glob("*.parquet"):
        if stale not in current:
            stale.unlink()
            print(f"  removed {stale.name} - no query produces it any more")
    manifest["finished"] = dt.datetime.now().astimezone().isoformat(timespec="seconds")
    (out / MANIFEST).write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    print(f"written to {out}")
    return manifest


def main():
    args = parse_args()
    extract(args.out, args.config)


if __name__ == "__main__":
    main()
