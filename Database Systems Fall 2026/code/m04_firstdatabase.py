"""Chapter 4 -- what a database costs you before it holds anything.

Students arrive believing a database is a large, heavy, administered thing.
For most of the history of the subject that was true, and for a server DBMS it
still is. This measures the one in front of them instead:

  1. How long from nothing at all to a query that returns a row?
  2. How many bytes is an empty database? A one-table database? A thousand
     rows?
  3. What is the file actually made of -- and why does it grow in jumps
     rather than smoothly?

The third question is the one worth the chapter. A SQLite database grows one
PAGE at a time, and seeing the file sit at exactly 4,096 bytes through the
first several hundred rows, then jump, is the first evidence a student gets
that a database is a designed structure rather than a bag of rows.

    python m04_firstdatabase.py            all of it
    python m04_firstdatabase.py --growth   just the page-by-page growth
    python m04_firstdatabase.py --keep     leave the database behind to poke at

Everything is written to a scratch file in this directory and removed again,
so running it twice gives the same answer.
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(HERE, "m04_scratch.db")


def fresh() -> None:
    for suffix in ("", "-wal", "-shm", "-journal"):
        p = DB + suffix
        if os.path.exists(p):
            os.remove(p)


def size() -> int:
    return os.path.getsize(DB) if os.path.exists(DB) else 0


def time_to_first_row() -> dict:
    """From no file at all to a row coming back."""
    fresh()
    t0 = time.perf_counter()
    con = sqlite3.connect(DB)
    t_connect = time.perf_counter()
    con.execute("CREATE TABLE student("
                "  roll_no INTEGER PRIMARY KEY,"
                "  name    TEXT NOT NULL,"
                "  programme TEXT NOT NULL)")
    t_create = time.perf_counter()
    con.execute("INSERT INTO student VALUES (1, 'Ayesha Khan', 'BSIT')")
    con.commit()
    t_insert = time.perf_counter()
    row = con.execute("SELECT name FROM student WHERE roll_no = 1").fetchone()
    t_select = time.perf_counter()
    con.close()
    return {
        "connect": (t_connect - t0) * 1000,
        "create": (t_create - t_connect) * 1000,
        "insert": (t_insert - t_create) * 1000,
        "select": (t_select - t_insert) * 1000,
        "total": (t_select - t0) * 1000,
        "row": row,
    }


def footprint() -> list[tuple[str, int]]:
    """Bytes on disk at each stage, measured rather than estimated."""
    out = []
    fresh()
    con = sqlite3.connect(DB)
    con.commit()
    out.append(("an empty database, just connected", size()))

    con.execute("CREATE TABLE student("
                "  roll_no INTEGER PRIMARY KEY,"
                "  name    TEXT NOT NULL,"
                "  programme TEXT NOT NULL)")
    con.commit()
    out.append(("after one CREATE TABLE, no rows", size()))

    for n in (1, 10, 100, 1000, 10000):
        have = con.execute("SELECT count(*) FROM student").fetchone()[0]
        con.executemany(
            "INSERT INTO student VALUES (?,?,?)",
            [(i, f"Student {i}", "BSIT") for i in range(have + 1, n + 1)])
        con.commit()
        out.append((f"{n:,} rows", size()))

    con.execute("VACUUM")
    out.append(("the same 10,000 rows after VACUUM", size()))
    con.close()
    return out


def growth() -> list[tuple[int, int, int]]:
    """(rows, bytes, page count) one row at a time, to see the jumps."""
    fresh()
    con = sqlite3.connect(DB)
    con.execute("CREATE TABLE student("
                "  roll_no INTEGER PRIMARY KEY,"
                "  name    TEXT NOT NULL,"
                "  programme TEXT NOT NULL)")
    con.commit()
    page = con.execute("PRAGMA page_size").fetchone()[0]
    out, last = [], -1
    for i in range(1, 401):
        con.execute("INSERT INTO student VALUES (?,?,?)",
                    (i, f"Student {i}", "BSIT"))
        con.commit()
        b = size()
        if b != last:                      # only the moments it CHANGES
            pages = con.execute("PRAGMA page_count").fetchone()[0]
            out.append((i, b, pages))
            last = b
    con.close()
    return out, page


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--growth", action="store_true")
    ap.add_argument("--keep", action="store_true")
    a = ap.parse_args()

    print(f"\n  SQLite {sqlite3.sqlite_version}, Python "
          f"{sys.version.split()[0]}")

    if not a.growth:
        t = time_to_first_row()
        print("\n=== 1. From nothing to a row ===")
        for k in ("connect", "create", "insert", "select"):
            print(f"  {k:<28}{t[k]:>8.2f} ms")
        print(f"  {'TOTAL':<28}{t['total']:>8.2f} ms   -> {t['row'][0]!r}")
        print("  No server was started, no port opened, no password set.")

        print("\n=== 2. What it costs on disk ===")
        print(f"  {'stage':<40}{'bytes':>12}")
        print("  " + "-" * 52)
        for label, b in footprint():
            print(f"  {label:<40}{b:>12,}")

    rows, page = growth()
    print(f"\n=== 3. How the file grows (page size {page:,} bytes) ===")
    print(f"  {'at row':>8}{'bytes':>12}{'pages':>8}{'bytes/row so far':>20}")
    print("  " + "-" * 48)
    for n, b, p in rows:
        print(f"  {n:>8}{b:>12,}{p:>8}{b / n:>19.0f}")
    print("\n  The file does not grow a little with every row. It sits at one")
    print("  size, then jumps by exactly one page. A database is a structure")
    print("  of fixed-size pages, and Chapter 21's index lives in the same")
    print("  pages -- which is the whole reason an index costs disk space.")

    if not a.keep:
        fresh()
    else:
        print(f"\n  left behind for you to poke at: {DB}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
