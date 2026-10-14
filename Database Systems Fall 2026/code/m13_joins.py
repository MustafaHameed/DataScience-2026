"""Chapter 13 -- what a join costs, and what the database does about it.

Four parts:
  1. Four ways to write the same inner join: same plan, same rows.
  2. Inner against left join: the sections with no instructor, which an
     inner join silently drops -- the promise Chapter 7 made about nulls.
  3. A join on an unindexed column, at three sizes, three ways:
       - as written, letting SQLite do what it likes
       - with PRAGMA automatic_index = OFF, which is the real cost
       - with a proper index on the column
     The middle line is quadratic. The first is not, because SQLite quietly
     builds a throwaway index for the query and discards it afterwards.
  4. What EXPLAIN QUERY PLAN says for each of the three.

    python m13_joins.py
    python m13_joins.py --quick       smaller sizes, for run_all
"""
from __future__ import annotations

import os
import sqlite3
import sys
import tempfile
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "schema"))
from seed import build  # noqa: E402

QUICK = "--quick" in sys.argv


def plan(cx: sqlite3.Connection, sql: str) -> list[str]:
    return [r[3] for r in cx.execute("EXPLAIN QUERY PLAN " + sql)]


def timed(cx: sqlite3.Connection, sql: str, reps: int = 3) -> tuple[float, int]:
    best, n = float("inf"), 0
    for _ in range(reps):
        t0 = time.perf_counter()
        n = cx.execute(sql).fetchone()[0]
        best = min(best, (time.perf_counter() - t0) * 1000)
    return best, n


FORMS = [
    ("implicit, comma + WHERE",
     "SELECT COUNT(*) FROM student s, enrolment e "
     "WHERE s.roll_no = e.roll_no"),
    ("explicit JOIN ... ON",
     "SELECT COUNT(*) FROM student s JOIN enrolment e "
     "ON s.roll_no = e.roll_no"),
    ("JOIN ... USING",
     "SELECT COUNT(*) FROM student JOIN enrolment USING (roll_no)"),
    ("NATURAL JOIN",
     "SELECT COUNT(*) FROM student NATURAL JOIN enrolment"),
]


def part1(cx: sqlite3.Connection) -> None:
    # The first draft of this part assumed NATURAL JOIN would agree with the
    # other three on this schema, and planned to break it later by adding a
    # contrived column to both tables. It did not agree: it returned 0 on the
    # schema exactly as Chapter 9 built it. The reason was already there --
    # student and enrolment BOTH have a column called "semester", meaning two
    # different things -- and the harness was printing "same count" directly
    # under a zero. The real schema is a better demonstration than any
    # contrived one, so this part now reports what it finds.
    print("\n=== 1. Four spellings of one inner join ===")
    print(f"  {'written as':<28}{'rows':>9}{'ms':>9}   plan")
    print("  " + "-" * 74)
    counts = []
    for label, sql in FORMS:
        ms, n = timed(cx, sql)
        counts.append(n)
        p = " / ".join(plan(cx, sql))
        print(f"  {label:<28}{n:>9,}{ms:>9.1f}   {p[:60]}")

    shared = sorted(
        {r[1] for r in cx.execute("PRAGMA table_info(student)")}
        & {r[1] for r in cx.execute("PRAGMA table_info(enrolment)")})
    print(f"\n  The first three agree: {counts[0]:,} rows, one plan.")
    print(f"  NATURAL JOIN returned {counts[3]:,}.")
    print(f"\n  Columns the two tables share BY NAME: {', '.join(shared)}")
    for t in ("student", "enrolment"):
        ty = [r[2] for r in cx.execute(f"PRAGMA table_info({t})")
              if r[1] == "semester"][0]
        ex = cx.execute(f"SELECT semester FROM {t} LIMIT 1").fetchone()[0]
        print(f"    {t}.semester   {ty:<8}  e.g. {ex!r}")
    print("\n  NATURAL JOIN matches on EVERY shared name, so it also demanded")
    print("  student.semester = enrolment.semester -- a year-of-study against")
    print("  a term label, which is never true. Every row vanished, with no")
    print("  error, on the schema exactly as Chapter 9 designed it.")


def part2(cx: sqlite3.Connection) -> None:
    print("\n=== 2. Inner against left join: who disappears ===")
    total = cx.execute("SELECT COUNT(*) FROM section").fetchone()[0]
    nulls = cx.execute("SELECT COUNT(*) FROM section "
                       "WHERE emp_id IS NULL").fetchone()[0]
    inner = cx.execute("SELECT COUNT(*) FROM section s JOIN instructor i "
                       "ON s.emp_id = i.emp_id").fetchone()[0]
    left = cx.execute("SELECT COUNT(*) FROM section s LEFT JOIN instructor i "
                      "ON s.emp_id = i.emp_id").fetchone()[0]
    print(f"  sections in the database                      {total:>7,}")
    print(f"  sections with no instructor yet (emp_id null) {nulls:>7,}")
    print(f"  timetable as an INNER JOIN                    {inner:>7,}")
    print(f"  timetable as a LEFT JOIN                      {left:>7,}")
    print(f"\n  {total - inner} sections are missing from the inner-join "
          "timetable.")
    print("  They are exactly the ones nobody has been assigned to teach --")
    print("  the rows a timetable most needs to show.")


def selfjoin_sql() -> str:
    # Pairs of students who share a name. name is not indexed.
    return ("SELECT COUNT(*) FROM student a JOIN student b "
            "ON a.name = b.name AND a.roll_no < b.roll_no")


def part3(tmp: str) -> list[tuple]:
    print("\n=== 3. A join on an UNINDEXED column, at three sizes ===")
    sizes = [500, 1000, 2000] if QUICK else [1000, 2000, 4000]
    rows = []
    print(f"  {'students':>9}{'pairs':>10}{'as written':>13}"
          f"{'auto idx OFF':>15}{'real index':>13}")
    print("  " + "-" * 60)
    for n in sizes:
        path = os.path.join(tmp, f"j{n}.db")
        cx = build(path, students=n)
        sql = selfjoin_sql()
        a, pairs = timed(cx, sql)
        cx.execute("PRAGMA automatic_index = OFF")
        b, _ = timed(cx, sql, reps=1)
        cx.execute("PRAGMA automatic_index = ON")
        cx.execute("CREATE INDEX idx_name ON student(name)")
        c, _ = timed(cx, sql)
        cx.close()
        rows.append((n, pairs, a, b, c))
        print(f"  {n:>9,}{pairs:>10,}{a:>11.1f}ms{b:>13.1f}ms{c:>11.1f}ms")

    # Do NOT read the growth ratios as the result. The first draft printed
    # them and said the unindexed join was "n-squared" -- but the indexed
    # join grew about x4 per doubling as well, because the OUTPUT grows x4:
    # names come from a fixed pool, so the number of matching pairs is itself
    # quadratic. A growth ratio mixes up how much the database had to look at
    # with how much it had to return. Compare at the same size instead: the
    # same pairs come back, so any difference is pure search cost.
    print("\n  The pairs column grows about x4 per doubling for every method,")
    print("  because with names drawn from a fixed pool the number of real")
    print("  matches is itself quadratic. So growth ratios mislead here;")
    print("  compare the methods at the SAME size, where the output is equal:")
    print(f"\n  {'students':>9}{'pairs examined':>17}{'pairs found':>13}"
          f"{'no index / index':>19}")
    for n, pairs, a, b, c in rows:
        examined = n * (n - 1) // 2
        print(f"  {n:>9,}{examined:>17,}{pairs:>13,}{b / c:>17.0f}x")
    print("\n  Without an index, every pair of students is examined to find")
    print("  the few that match -- about 8 million comparisons at 4,000 to")
    print("  return 28 thousand pairs. With one, only the matches are visited.")
    print("\n  'As written' is close to 'real index' because SQLite builds an")
    print("  index on name for this ONE query, uses it, and discards it. You")
    print("  pay to build an index on every run and never get to keep it.")
    return rows


def part4(tmp: str) -> None:
    print("\n=== 4. What EXPLAIN QUERY PLAN says for each ===")
    path = os.path.join(tmp, "plan.db")
    cx = build(path, students=500)
    sql = selfjoin_sql()
    print("  as written (automatic_index ON, no index on name):")
    for p in plan(cx, sql):
        print(f"      {p}")
    cx.execute("PRAGMA automatic_index = OFF")
    print("  automatic_index OFF:")
    for p in plan(cx, sql):
        print(f"      {p}")
    cx.execute("PRAGMA automatic_index = ON")
    cx.execute("CREATE INDEX idx_name ON student(name)")
    print("  with CREATE INDEX idx_name ON student(name):")
    for p in plan(cx, sql):
        print(f"      {p}")
    cx.close()
    print("\n  'AUTOMATIC ... INDEX' in a plan is the optimiser telling you an")
    print("  index is missing. It is the cheapest diagnosis in this course.")


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        cx = build(os.path.join(tmp, "u.db"), students=4000)
        part1(cx)
        cx.close()
        cx = build(os.path.join(tmp, "u2.db"), students=4000)
        part2(cx)
        cx.close()
        part3(tmp)
        part4(tmp)
    print()


if __name__ == "__main__":
    main()
