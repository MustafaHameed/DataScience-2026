"""Chapter 8 -- what a declared constraint actually enforces.

Three findings, each produced by asking SQLite rather than by reading about
it. All three have the shape this course is built on: the declaration is
present, correct and accepted, and nothing happens.

  1. PRAGMA foreign_keys is 0 by default. A REFERENCES clause is parsed,
     stored in the schema, reported by .schema -- and not enforced. Orphan
     rows go in without an error.
  2. The pragma is per CONNECTION, not per database. Setting it once does not
     set it for the next program, or the next run of the same program.
  3. A TEXT PRIMARY KEY column accepts NULL and stores it. Entity integrity
     says a primary key is never null; SQLite enforces that only when the
     table is STRICT or the column is also declared NOT NULL. It is
     documented upstream as a long-standing bug kept for compatibility.
     INTEGER PRIMARY KEY also "accepts" a NULL, for an unrelated reason --
     it is an alias for the rowid, so the NULL is replaced by the next
     integer. Part 3 prints typeof() so the two cannot be confused.

    python m08_constraints.py
"""
from __future__ import annotations

import os
import sqlite3
import tempfile
import time

SCHEMA = """
CREATE TABLE course (
  course_code TEXT PRIMARY KEY,
  title       TEXT NOT NULL
);
CREATE TABLE section (
  course_code TEXT NOT NULL REFERENCES course(course_code),
  section_no  INTEGER NOT NULL,
  room        TEXT,
  PRIMARY KEY (course_code, section_no)
);
"""

ORPHAN = ("CS-999", 1, "Room 7")      # CS-999 does not exist in course


def fresh(path: str, fk: bool) -> sqlite3.Connection:
    """A new database with the schema, and the pragma set or not."""
    if os.path.exists(path):
        os.remove(path)
    cx = sqlite3.connect(path)
    if fk:
        cx.execute("PRAGMA foreign_keys = ON")
    cx.executescript(SCHEMA)
    cx.execute("INSERT INTO course VALUES ('CS-103','Database Systems')")
    cx.commit()
    return cx


def insert_orphan(cx: sqlite3.Connection) -> str:
    """Try the orphan insert; return 'accepted' or the error message."""
    try:
        cx.execute("INSERT INTO section VALUES (?,?,?)", ORPHAN)
        cx.commit()
        return "accepted"
    except sqlite3.IntegrityError as e:
        return f"refused: {e}"


def part1(tmp: str) -> None:
    print("\n=== 1. The same schema, the same INSERT, two settings ===")
    print(f"  {'PRAGMA foreign_keys':<26}{'orphan row':<12}outcome")
    print("  " + "-" * 62)
    for fk in (False, True):
        path = os.path.join(tmp, f"fk{int(fk)}.db")
        cx = fresh(path, fk)
        got = cx.execute("PRAGMA foreign_keys").fetchone()[0]
        out = insert_orphan(cx)
        rows = cx.execute("SELECT COUNT(*) FROM section").fetchone()[0]
        cx.close()
        print(f"  {'ON' if fk else 'default (untouched)':<26}"
              f"{rows:<12}{out[:40]}")
        assert got == (1 if fk else 0), "pragma did not take the value set"
    print("\n  The REFERENCES clause is identical in both. Only the")
    print("  connection differs.")


def part2(tmp: str) -> None:
    print("\n=== 2. The pragma is per connection, not per database ===")
    path = os.path.join(tmp, "perconn.db")
    cx = fresh(path, True)
    print(f"  connection A, after PRAGMA foreign_keys = ON   -> "
          f"{cx.execute('PRAGMA foreign_keys').fetchone()[0]}")
    print(f"  connection A, orphan insert                    -> "
          f"{insert_orphan(cx)[:40]}")
    cx.close()

    cx2 = sqlite3.connect(path)       # same file, new connection
    print(f"  connection B, same FILE, nothing set           -> "
          f"{cx2.execute('PRAGMA foreign_keys').fetchone()[0]}")
    print(f"  connection B, orphan insert                    -> "
          f"{insert_orphan(cx2)}")
    n = cx2.execute("SELECT COUNT(*) FROM section").fetchone()[0]
    cx2.close()
    print(f"\n  orphan rows now in the table: {n}")
    print("  The setting did not survive the connection. Nothing in the")
    print("  file remembers that enforcement was ever wanted.")


def part3(tmp: str) -> None:
    print("\n=== 3. A PRIMARY KEY column that accepts NULL ===")
    path = os.path.join(tmp, "pknull.db")
    if os.path.exists(path):
        os.remove(path)
    cx = sqlite3.connect(path)
    cases = [
        ("TEXT PRIMARY KEY",           "CREATE TABLE a (k TEXT PRIMARY KEY)"),
        ("INTEGER PRIMARY KEY",        "CREATE TABLE b (k INTEGER PRIMARY KEY)"),
        ("TEXT PRIMARY KEY, STRICT",
         "CREATE TABLE c (k TEXT PRIMARY KEY) STRICT"),
        ("TEXT PRIMARY KEY NOT NULL",
         "CREATE TABLE d (k TEXT PRIMARY KEY NOT NULL)"),
    ]
    # "Accepted" is not one outcome but two, and the difference is the whole
    # finding. INTEGER PRIMARY KEY is an alias for the rowid, so a NULL there
    # is not stored as NULL -- SQLite substitutes the next integer, which is
    # ordinary AUTOINCREMENT behaviour and no violation at all. TEXT PRIMARY
    # KEY really does store the NULL. Reporting only accepted/refused would
    # have put those two in the same column and taught the wrong lesson.
    print(f"  {'declared as':<30}{'INSERT NULL':<13}{'stored':<9}rows")
    print("  " + "-" * 62)
    for label, ddl in cases:
        cx.executescript(ddl)
        tbl = ddl.split()[2]
        try:
            cx.execute(f"INSERT INTO {tbl} VALUES (NULL)")
            cx.commit()
            verdict = "accepted"
        except sqlite3.IntegrityError:
            verdict = "refused"
        rows = cx.execute(f"SELECT k, typeof(k) FROM {tbl}").fetchall()
        stored = rows[0][1] if rows else "--"
        print(f"  {label:<30}{verdict:<13}{stored:<9}{len(rows)}")
    cx.close()
    print("\n  Entity integrity says a primary key is never null. Read the")
    print("  'stored' column, not the 'INSERT NULL' one: INTEGER PRIMARY KEY")
    print("  accepted the NULL and stored an integer, which is the rowid")
    print("  being assigned and is not a violation. TEXT PRIMARY KEY accepted")
    print("  it and stored null -- and that is the declaration every table in")
    print("  this course uses.")


def part4(tmp: str) -> None:
    # A single run of this put the two settings 0.6 ms apart on 5,000 rows,
    # which is smaller than the run-to-run variation and would have been
    # reported as a 3% cost. Take the median of several runs and print the
    # spread, so the reader can see whether the difference is a result or
    # noise -- the same discipline Chapter 4 applies to its own timings.
    print("\n=== 4. What enforcement costs ===")
    n, reps = 20000, 5
    timing: dict[bool, list[float]] = {False: [], True: []}
    for r in range(reps):
        for fk in (False, True):
            path = os.path.join(tmp, f"cost{int(fk)}_{r}.db")
            cx = fresh(path, fk)
            cx.executemany("INSERT INTO course VALUES (?,?)",
                           [(f"CS-{i:05d}", f"Course {i}") for i in range(n)])
            cx.commit()
            rows = [(f"CS-{i:05d}", 1, "R1") for i in range(n)]
            t0 = time.perf_counter()
            cx.executemany("INSERT INTO section VALUES (?,?,?)", rows)
            cx.commit()
            timing[fk].append((time.perf_counter() - t0) * 1000)
            cx.close()
            os.remove(path)

    def med(v: list[float]) -> float:
        return sorted(v)[len(v) // 2]

    off, on = med(timing[False]), med(timing[True])
    print(f"  {n:,} valid child rows, {reps} runs, median of each")
    print(f"  {'':<34}{'median':>9}{'min':>9}{'max':>9}")
    for fk, label in ((False, "foreign_keys OFF"), (True, "foreign_keys ON")):
        v = timing[fk]
        print(f"  {label:<34}{med(v):8.1f}{min(v):9.1f}{max(v):9.1f}")
    print(f"  {'cost of enforcement':<34}{on - off:8.1f} ms"
          f"   ({(on / off - 1) * 100:+.0f}%)")
    print("\n  Every child row costs one lookup in the parent's primary key")
    print("  index -- an index that already exists, which is why the cost is")
    print("  what it is. Compare it against the spread above before calling")
    print("  it expensive.")


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}")
    with tempfile.TemporaryDirectory() as tmp:
        part1(tmp)
        part2(tmp)
        part3(tmp)
        part4(tmp)
    print()


if __name__ == "__main__":
    main()
