"""Chapter 15 -- sub-queries: which forms agree, and what correlation costs.

Three parts:
  1. A semi-join written three ways -- IN, EXISTS, JOIN + DISTINCT.
     "Students with at least one A." Same answer; what does each cost?
  2. An anti-join written four ways -- NOT IN, NOT EXISTS, LEFT JOIN ...
     IS NULL, EXCEPT. "Instructors who teach nothing", after three are hired.
     One of the four gets it wrong, and Chapter 14 said which.
  3. A correlated sub-query in the SELECT list, against GROUP BY, for
     "enrolments per course", at three sizes -- with and without an index on
     the column the sub-query is correlated on.

    python m15_subqueries.py
    python m15_subqueries.py --quick
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


def timed(cx: sqlite3.Connection, sql: str, reps: int = 3):
    best, out = float("inf"), None
    for _ in range(reps):
        t0 = time.perf_counter()
        out = cx.execute(sql).fetchall()
        best = min(best, (time.perf_counter() - t0) * 1000)
    return best, out


SEMI = [
    ("IN (uncorrelated)",
     "SELECT COUNT(*) FROM student WHERE roll_no IN "
     "(SELECT roll_no FROM enrolment WHERE grade = 'A')"),
    ("EXISTS (correlated)",
     "SELECT COUNT(*) FROM student s WHERE EXISTS "
     "(SELECT 1 FROM enrolment e WHERE e.roll_no = s.roll_no "
     "AND e.grade = 'A')"),
    ("JOIN + DISTINCT",
     "SELECT COUNT(DISTINCT s.roll_no) FROM student s JOIN enrolment e "
     "ON e.roll_no = s.roll_no WHERE e.grade = 'A'"),
]

ANTI = [
    ("NOT IN",
     "SELECT COUNT(*) FROM instructor WHERE emp_id NOT IN "
     "(SELECT emp_id FROM section)"),
    ("NOT EXISTS",
     "SELECT COUNT(*) FROM instructor i WHERE NOT EXISTS "
     "(SELECT 1 FROM section s WHERE s.emp_id = i.emp_id)"),
    ("LEFT JOIN ... IS NULL",
     "SELECT COUNT(*) FROM instructor i LEFT JOIN section s "
     "ON s.emp_id = i.emp_id WHERE s.emp_id IS NULL"),
    ("EXCEPT",
     "SELECT COUNT(*) FROM (SELECT emp_id FROM instructor "
     "EXCEPT SELECT emp_id FROM section)"),
]


def part1(cx: sqlite3.Connection) -> None:
    print("\n=== 1. One semi-join, three spellings ===")
    print("  'students with at least one A'")
    print(f"  {'written as':<22}{'answer':>8}{'ms':>9}   plan")
    print("  " + "-" * 74)
    for label, sql in SEMI:
        ms, out = timed(cx, sql)
        p = " / ".join(plan(cx, sql))
        print(f"  {label:<22}{out[0][0]:>8,}{ms:>9.1f}   {p[:52]}")


def part2(cx: sqlite3.Connection) -> None:
    print("\n=== 2. One anti-join, four spellings ===")
    cx.execute("PRAGMA foreign_keys = ON")
    cx.executemany("INSERT INTO instructor VALUES (?,?,?,?)",
                   [("E-9001", "New Hire One", "Lecturer", "CS"),
                    ("E-9002", "New Hire Two", "Lecturer", "IT"),
                    ("E-9003", "New Hire Three", "Lecturer", "SE")])
    nulls = cx.execute("SELECT COUNT(*) FROM section "
                       "WHERE emp_id IS NULL").fetchone()[0]
    print("  'instructors who teach nothing' -- after hiring three who")
    print(f"  teach nothing yet. section.emp_id has {nulls} nulls.")
    print(f"  {'written as':<26}{'answer':>8}{'ms':>9}   right?")
    print("  " + "-" * 56)
    for label, sql in ANTI:
        ms, out = timed(cx, sql)
        n = out[0][0]
        # EXCEPT is a set operation, not a comparison: it removes each
        # left-hand row that also appears on the right. The NULL that section
        # contributes on the right could only ever remove a NULL on the left,
        # and instructor.emp_id is a primary key and has none -- so, unlike
        # NOT IN, it cannot be emptied by the nulls.
        print(f"  {label:<26}{n:>8,}{ms:>9.2f}   "
              f"{'yes' if n == 3 else 'NO -- one null empties it'}")


def corr_sql() -> str:
    return ("SELECT c.course_code, (SELECT COUNT(*) FROM enrolment e "
            "WHERE e.course_code = c.course_code) FROM course c")


def group_sql() -> str:
    return ("SELECT c.course_code, COUNT(e.course_code) FROM course c "
            "LEFT JOIN enrolment e ON e.course_code = c.course_code "
            "GROUP BY c.course_code")


def part3(tmp: str) -> None:
    print("\n=== 3. A correlated sub-query in the SELECT list ===")
    print("  'how many enrolments does each course have'")
    print("  enrolment's primary key is (roll_no, course_code, ...), so it")
    print("  cannot be searched by course_code alone -- course_code is its")
    print("  SECOND column. That is the only thing that changes below.")
    sizes = [1000, 2000, 4000] if QUICK else [2000, 4000, 8000]
    print(f"\n  {'students':>9}{'enrolments':>12}{'correlated':>13}"
          f"{'GROUP BY':>11}{'corr+index':>13}{'GROUP+index':>13}")
    print("  " + "-" * 71)
    agree = True
    for n in sizes:
        cx = build(os.path.join(tmp, f"s{n}.db"), students=n)
        e = cx.execute("SELECT COUNT(*) FROM enrolment").fetchone()[0]
        a, ra = timed(cx, corr_sql())
        b, rb = timed(cx, group_sql())
        cx.execute("CREATE INDEX idx_enr_course ON enrolment(course_code)")
        c, rc = timed(cx, corr_sql())
        d, rd = timed(cx, group_sql())
        agree &= sorted(ra) == sorted(rb) == sorted(rc) == sorted(rd)
        cx.close()
        print(f"  {n:>9,}{e:>12,}{a:>11.1f}ms{b:>9.1f}ms"
              f"{c:>11.1f}ms{d:>11.1f}ms")
    print(f"\n  all four agree on every count: {'yes' if agree else 'NO'}")

    cx = build(os.path.join(tmp, "p.db"), students=500)
    print("\n  plan for the correlated form, no index on course_code:")
    for p in plan(cx, corr_sql()):
        print(f"      {p}")
    cx.execute("CREATE INDEX idx_enr_course ON enrolment(course_code)")
    print("  the same, with CREATE INDEX ON enrolment(course_code):")
    for p in plan(cx, corr_sql()):
        print(f"      {p}")
    cx.close()


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        cx = build(os.path.join(tmp, "u.db"), students=4000)
        part1(cx)
        part2(cx)
        cx.close()
        part3(tmp)
    print()


if __name__ == "__main__":
    main()
