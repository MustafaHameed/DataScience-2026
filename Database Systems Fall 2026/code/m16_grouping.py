"""Chapter 16 -- grouping, aggregation, and window functions.

Four parts:
  1. What the aggregates do with nulls: COUNT(*) against COUNT(col), and
     AVG over a column that is a third null.
  2. GROUP BY with and without an index that supplies the grouping order --
     whether the plan needs a temporary B-tree, and what that costs.
  3. The same filter in WHERE and in HAVING: same answer, different work.
  4. A rank within each programme: a window function against the correlated
     sub-query people wrote before window functions existed.

    python m16_grouping.py
    python m16_grouping.py --quick
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

POINTS = ("CASE grade WHEN 'A' THEN 4 WHEN 'B' THEN 3 WHEN 'C' THEN 2 "
          "WHEN 'D' THEN 1 WHEN 'F' THEN 0 END")


def plan(cx: sqlite3.Connection, sql: str) -> list[str]:
    return [r[3] for r in cx.execute("EXPLAIN QUERY PLAN " + sql)]


def timed(cx: sqlite3.Connection, sql: str, reps: int = 3):
    best, out = float("inf"), None
    for _ in range(reps):
        t0 = time.perf_counter()
        out = cx.execute(sql).fetchall()
        best = min(best, (time.perf_counter() - t0) * 1000)
    return best, out


def part1(cx: sqlite3.Connection) -> None:
    print("\n=== 1. Aggregates and nulls ===")
    star, col, nul = cx.execute(
        "SELECT COUNT(*), COUNT(grade), COUNT(*) - COUNT(grade) "
        "FROM enrolment").fetchone()
    print(f"  COUNT(*)                         {star:>8,}   every row")
    print(f"  COUNT(grade)                     {col:>8,}   rows where grade "
          "is not null")
    print(f"  difference                       {nul:>8,}   the ungraded "
          "FALL2026 rows")
    avg_skip = cx.execute(f"SELECT AVG({POINTS}) FROM enrolment").fetchone()[0]
    avg_zero = cx.execute(
        f"SELECT AVG(COALESCE({POINTS}, 0)) FROM enrolment").fetchone()[0]
    print(f"\n  AVG(points)                      {avg_skip:>8.3f}   nulls "
          "skipped -- divides by "
          f"{col:,}")
    print(f"  AVG(COALESCE(points, 0))         {avg_zero:>8.3f}   nulls as "
          f"zero -- divides by {star:,}")
    print(f"  the second is {(1 - avg_zero / avg_skip) * 100:.0f}% lower, "
          "and it has decided that an ungraded course is an F.")

    print("\n  GROUP BY grade -- null is a group of its own:")
    for g, n in cx.execute("SELECT grade, COUNT(*) FROM enrolment "
                           "GROUP BY grade ORDER BY grade"):
        print(f"      {str(g):<6}{n:>8,}")


GROUP = ("SELECT course_code, COUNT(*) FROM enrolment "
         "GROUP BY course_code")


def part2(tmp: str) -> None:
    print("\n=== 2. GROUP BY with and without a supporting index ===")
    sizes = [2000, 4000] if QUICK else [4000, 8000, 16000]
    print(f"  {'students':>9}{'enrolments':>12}{'no index':>12}"
          f"{'with index':>13}{'speed-up':>11}")
    print("  " + "-" * 57)
    last = None
    for n in sizes:
        cx = build(os.path.join(tmp, f"g{n}.db"), students=n)
        e = cx.execute("SELECT COUNT(*) FROM enrolment").fetchone()[0]
        a, ra = timed(cx, GROUP)
        p_no = plan(cx, GROUP)
        cx.execute("CREATE INDEX idx_enr_course ON enrolment(course_code)")
        b, rb = timed(cx, GROUP)
        p_ix = plan(cx, GROUP)
        assert sorted(ra) == sorted(rb)
        cx.close()
        last = (p_no, p_ix)
        print(f"  {n:>9,}{e:>12,}{a:>10.1f}ms{b:>11.1f}ms{a / b:>10.1f}x")
    print("\n  plan without the index:")
    for p in last[0]:
        print(f"      {p}")
    print("  plan with CREATE INDEX ON enrolment(course_code):")
    for p in last[1]:
        print(f"      {p}")
    print("\n  'TEMP B-TREE FOR GROUP BY' means the rows had to be sorted into")
    print("  groups first. An index on the grouping column is ALREADY in that")
    print("  order, so the sort disappears -- and so does most of the cost.")


def program(cx: sqlite3.Connection, sql: str) -> list[tuple]:
    """The compiled bytecode, as comparable tuples."""
    return [(r[1], r[2], r[3], r[4], str(r[5]))
            for r in cx.execute("EXPLAIN " + sql)]


def first(prog: list[tuple], pred) -> int:
    return next((i for i, o in enumerate(prog) if pred(o)), -1)


def part3(cx: sqlite3.Connection) -> None:
    # The first draft of this part asserted that HAVING "groups every row and
    # then throws groups away", and so is slower than WHERE. The timings came
    # back identical to a tenth of a millisecond. The reason is that SQLite
    # moves a HAVING term that mentions no aggregate into the WHERE clause
    # itself -- and the way to establish that is not to argue from timings
    # but to compare the compiled programs, which is what this now does.
    print("\n=== 3. The same filter in WHERE and in HAVING ===")
    w = ("SELECT course_code, COUNT(*) FROM enrolment "
         "WHERE course_code LIKE 'CS-%' GROUP BY course_code")
    h = ("SELECT course_code, COUNT(*) FROM enrolment "
         "GROUP BY course_code HAVING course_code LIKE 'CS-%'")
    agg = ("SELECT course_code, COUNT(*) FROM enrolment "
           "GROUP BY course_code HAVING COUNT(*) > 1000")
    a, ra = timed(cx, w)
    b, rb = timed(cx, h)
    print(f"  filter in WHERE    {a:7.1f} ms   {len(ra)} groups")
    print(f"  filter in HAVING   {b:7.1f} ms   {len(rb)} groups")
    print(f"  same answer: {'yes' if sorted(ra) == sorted(rb) else 'NO'}")

    pw, ph, pa = program(cx, w), program(cx, h), program(cx, agg)
    print(f"\n  compiled programs: WHERE {len(pw)} ops, HAVING {len(ph)} ops,"
          f" identical: {'YES' if pw == ph else 'no'}")

    def where_is(prog, test):
        f = first(prog, test)
        s = first(prog, lambda o: o[0] == "SorterInsert")
        return f, s, ("BEFORE" if 0 <= f < s else "AFTER")

    is_like = (lambda o: o[0] == "Function" and "like" in o[4].lower())
    is_cmp = (lambda o: o[0] in ("Le", "Lt", "Ge", "Gt"))
    for label, prog, test in (("WHERE  course_code LIKE", pw, is_like),
                              ("HAVING course_code LIKE", ph, is_like),
                              ("HAVING COUNT(*) > 1000 ", pa, is_cmp)):
        f, s, when = where_is(prog, test)
        print(f"    {label}   tested at op {f:>2}, grouping sort at op {s:>2}"
              f"  ->  {when} grouping")
    print("\n  SQLite moved the HAVING condition into the WHERE clause itself:")
    print("  the two queries compiled to the same program, and the test runs")
    print("  before the rows are grouped in both. It could do that only")
    print("  because the condition mentions no aggregate. COUNT(*) > 1000")
    print("  cannot be moved -- the count does not exist until the grouping")
    print("  is done -- and it is tested afterwards, as it must be.")
    print("\n  So: put a condition on a column in WHERE because that is what")
    print("  it MEANS, not because it is faster here. Here the optimiser")
    print("  fixed it for you, which is an optimisation, not a promise.")


def rank_window() -> str:
    return (f"WITH a AS (SELECT s.roll_no, s.programme, "
            f"SUM(e.grade = 'A') AS n_a FROM student s JOIN enrolment e "
            f"ON e.roll_no = s.roll_no GROUP BY s.roll_no) "
            f"SELECT roll_no, programme, n_a, "
            f"RANK() OVER (PARTITION BY programme ORDER BY n_a DESC) "
            f"FROM a")


def rank_corr() -> str:
    return (f"WITH a AS (SELECT s.roll_no, s.programme, "
            f"SUM(e.grade = 'A') AS n_a FROM student s JOIN enrolment e "
            f"ON e.roll_no = s.roll_no GROUP BY s.roll_no) "
            f"SELECT x.roll_no, x.programme, x.n_a, "
            f"1 + (SELECT COUNT(*) FROM a y WHERE y.programme = x.programme "
            f"AND y.n_a > x.n_a) FROM a x")


def part4(tmp: str) -> None:
    print("\n=== 4. A rank within each programme ===")
    print("  'each student's rank in their programme by number of A grades'")
    sizes = [1000, 2000] if QUICK else [1000, 2000, 4000]
    print(f"  {'students':>9}{'window RANK()':>16}{'correlated':>14}"
          f"{'ratio':>9}   agree?")
    print("  " + "-" * 58)
    for n in sizes:
        cx = build(os.path.join(tmp, f"r{n}.db"), students=n)
        a, ra = timed(cx, rank_window())
        b, rb = timed(cx, rank_corr(), reps=1)
        cx.close()
        same = sorted(ra) == sorted(rb)
        print(f"  {n:>9,}{a:>14.1f}ms{b:>12.1f}ms{b / a:>8.0f}x   "
              f"{'yes' if same else 'NO'}")
    print("\n  The correlated version counts, for every student, how many in")
    print("  the same programme did better -- every student against every")
    print("  other in the programme. The window function sorts each")
    print("  programme once and reads the rank off the sorted order.")


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        cx = build(os.path.join(tmp, "u.db"), students=4000)
        part1(cx)
        part3(cx)
        cx.close()
        part2(tmp)
        part4(tmp)
    print()


if __name__ == "__main__":
    main()
