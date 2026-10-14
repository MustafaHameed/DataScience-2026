"""Chapter 18 -- what normalisation buys, and what it costs.

The same information held two ways:
  * report: one wide table, one row per enrolment, every fact repeated
    (Chapter 17's fourteen columns)
  * the normalised schema Chapter 9 built: eight tables, every fact once

Measured:
  1. bytes on disk, and how many copies of one fact each holds
  2. the update anomaly: rename a department -- rows touched, and time
  3. what a half-finished update leaves behind in each
  4. the insertion and deletion anomalies, demonstrated
  5. read cost: a transcript and a department report, with and without
     joins -- the one place the wide table is expected to win

    python m18_normalforms.py
    python m18_normalforms.py --quick
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import sys
import tempfile
import time

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "schema"))
from seed import build  # noqa: E402

QUICK = "--quick" in sys.argv
STUDENTS = 2000 if QUICK else 4000

REPORT = """
CREATE TABLE report AS
SELECT s.roll_no, s.name, s.programme,
       c.course_code, c.title, c.credit_hours,
       d.dept_code, d.name AS dept_name,
       e.semester, e.section_no, x.room, x.emp_id,
       i.name AS instructor, e.grade
  FROM enrolment e
  JOIN student    s ON s.roll_no     = e.roll_no
  JOIN course     c ON c.course_code = e.course_code
  JOIN department d ON d.dept_code   = c.dept_code
  JOIN section    x ON x.course_code = e.course_code
                   AND x.semester    = e.semester
                   AND x.section_no  = e.section_no
  LEFT JOIN instructor i ON i.emp_id = x.emp_id
"""

# the tables that hold the SAME information as report -- no more
NORMAL = ["department", "instructor", "course", "section", "student",
          "enrolment"]


def timed(cx, sql, params=(), reps=3):
    best, out = float("inf"), None
    for _ in range(reps):
        t0 = time.perf_counter()
        out = cx.execute(sql, params).fetchall()
        best = min(best, (time.perf_counter() - t0) * 1000)
    return best, out


def size_of(path: str, keep: list[str]) -> int:
    """Bytes of a copy of the database holding only the named tables."""
    tmp = path + ".only.db"
    shutil.copy(path, tmp)
    cx = sqlite3.connect(tmp)
    cx.execute("PRAGMA foreign_keys = OFF")
    have = [r[0] for r in cx.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")]
    for t in have:
        if t not in keep:
            cx.execute(f"DROP TABLE {t}")
    cx.commit()
    cx.execute("VACUUM")
    cx.close()
    n = os.path.getsize(tmp)
    os.remove(tmp)
    return n


def part1(path: str, cx: sqlite3.Connection) -> None:
    print("\n=== 1. The same information, two ways: bytes and copies ===")
    # A like-for-like comparison needs the wide table to have a key too.
    # CREATE TABLE AS copies rows and NOTHING else -- no primary key, no
    # index -- while every normalised table carries the index its PRIMARY
    # KEY implies. The first draft compared those directly and reported the
    # wide table as only 1.2x the size, which mostly measured the missing
    # index. Both figures are printed so the difference is visible.
    rows = cx.execute("SELECT COUNT(*) FROM report").fetchone()[0]
    rb = size_of(path, ["report"])
    nb = size_of(path, NORMAL)
    nrows = sum(cx.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                for t in NORMAL)
    cx.execute("CREATE UNIQUE INDEX report_key ON report"
               "(roll_no, course_code, semester, section_no)")
    cx.commit()
    rk = size_of(path, ["report"])
    print(f"  {'':<38}{'rows':>9}{'bytes':>13}")
    print(f"  {'report, as CREATE TABLE AS made it':<38}{rows:>9,}{rb:>13,}"
          "   no index")
    print(f"  {'report, with its key indexed':<38}{rows:>9,}{rk:>13,}"
          "   like for like")
    print(f"  {'normalised, ' + str(len(NORMAL)) + ' tables':<38}"
          f"{nrows:>9,}{nb:>13,}   PK indexes")
    print(f"\n  like for like, the wide table is {rk / nb:.1f}x the size "
          "for the same facts")
    print(f"  (as created it looks like {rb / nb:.1f}x -- because it has no "
          "index at all)")
    print("\n  copies of ONE fact:")
    facts = [
        ("a student's name", "SELECT MAX(n) FROM (SELECT COUNT(*) n FROM "
         "report GROUP BY roll_no)"),
        ("a course's title", "SELECT MAX(n) FROM (SELECT COUNT(*) n FROM "
         "report GROUP BY course_code)"),
        ("a department's name", "SELECT MAX(n) FROM (SELECT COUNT(*) n FROM "
         "report GROUP BY dept_code)"),
        ("an instructor's name", "SELECT MAX(n) FROM (SELECT COUNT(*) n FROM "
         "report WHERE emp_id IS NOT NULL GROUP BY emp_id)"),
    ]
    print(f"  {'':<24}{'report':>10}{'normalised':>13}")
    for label, sql in facts:
        n = cx.execute(sql).fetchone()[0]
        print(f"  {label:<24}{n:>10,}{1:>13}")


def part2(cx: sqlite3.Connection) -> None:
    print("\n=== 2. The update anomaly: rename one department ===")
    a, _ = timed(cx, "SELECT 1", reps=1)
    t0 = time.perf_counter()
    cur = cx.execute("UPDATE report SET dept_name = 'Computing' "
                     "WHERE dept_code = 'CS'")
    wide_rows = cur.rowcount
    cx.commit()
    wide_ms = (time.perf_counter() - t0) * 1000
    t0 = time.perf_counter()
    cur = cx.execute("UPDATE department SET name = 'Computing' "
                     "WHERE dept_code = 'CS'")
    norm_rows = cur.rowcount
    cx.commit()
    norm_ms = (time.perf_counter() - t0) * 1000
    print(f"  {'':<14}{'rows changed':>14}{'ms':>10}")
    print(f"  {'report':<14}{wide_rows:>14,}{wide_ms:>10.2f}")
    print(f"  {'department':<14}{norm_rows:>14,}{norm_ms:>10.2f}")
    print(f"\n  One fact changed. The wide table rewrote {wide_rows:,} rows "
          "to change it.")


def part3(cx: sqlite3.Connection) -> None:
    print("\n=== 3. A half-finished update ===")
    # Simulate what a program that crashes part-way, or a clerk who filters
    # on the wrong column, leaves behind: only FALL2025 rows are renamed.
    cx.execute("UPDATE report SET dept_name = 'Information Tech.' "
               "WHERE dept_code = 'IT' AND semester = 'FALL2025'")
    cx.commit()
    names = cx.execute("SELECT dept_name, COUNT(*) FROM report "
                       "WHERE dept_code = 'IT' GROUP BY dept_name").fetchall()
    print("  report now holds, for dept_code = 'IT':")
    for nm, n in names:
        print(f"      {nm!r:<28}{n:>8,} rows")
    print(f"  -> {len(names)} different names for one department. Which is "
          "right?")
    print("\n  department holds one row per dept_code, so the same mistake")
    print("  there is impossible: there is only one place to write it.")


def part4(cx: sqlite3.Connection) -> None:
    print("\n=== 4. The insertion and deletion anomalies ===")
    # insertion: a new course with no enrolments yet
    try:
        cx.execute("INSERT INTO course VALUES ('CS-999','Quantum Databases',"
                   "3,'CS')")
        cx.commit()
        ok_norm = "yes"
    except sqlite3.Error as e:
        ok_norm = f"no ({e})"
    in_report = cx.execute("SELECT COUNT(*) FROM report "
                           "WHERE course_code = 'CS-999'").fetchone()[0]
    print("  INSERTION: a new course, nobody enrolled yet")
    print(f"    normalised -- one row in course:          {ok_norm}")
    print(f"    report     -- rows that could hold it:    {in_report}")
    print("    report has no row to put it in without inventing an enrolment")
    print("    -- a roll number, a section, a grade -- that does not exist.")

    # deletion: remove the enrolments of the course with the fewest rows
    code, n = cx.execute("SELECT course_code, COUNT(*) FROM report "
                         "GROUP BY course_code ORDER BY 2 LIMIT 1").fetchone()
    title = cx.execute("SELECT title FROM report WHERE course_code = ? "
                       "LIMIT 1", (code,)).fetchone()[0]
    cx.execute("DELETE FROM report WHERE course_code = ?", (code,))
    cx.commit()
    left = cx.execute("SELECT COUNT(*) FROM report WHERE course_code = ?",
                      (code,)).fetchone()[0]
    still = cx.execute("SELECT COUNT(*) FROM course WHERE course_code = ?",
                       (code,)).fetchone()[0]
    print(f"\n  DELETION: remove the {n:,} enrolments of {code} "
          f"({title!r})")
    print(f"    report     -- rows still describing the course:  {left}")
    print(f"    normalised -- course row still present:          "
          f"{'yes' if still else 'no'}")
    print("    In the wide table the course's title, credit hours and")
    print("    department went with its last enrolment. Nothing else held "
          "them.")


TRANSCRIPT_WIDE = ("SELECT course_code, title, credit_hours, semester, grade "
                   "FROM report WHERE roll_no = ?")
TRANSCRIPT_NORM = ("SELECT c.course_code, c.title, c.credit_hours, "
                   "e.semester, e.grade FROM enrolment e JOIN course c "
                   "ON c.course_code = e.course_code WHERE e.roll_no = ?")
DEPT_WIDE = ("SELECT dept_name, COUNT(*), AVG(credit_hours) FROM report "
             "GROUP BY dept_code")
DEPT_NORM = ("SELECT d.name, COUNT(*), AVG(c.credit_hours) FROM enrolment e "
             "JOIN course c ON c.course_code = e.course_code "
             "JOIN department d ON d.dept_code = c.dept_code "
             "GROUP BY d.dept_code")


def part5(path: str) -> None:
    # The heading of this part used to say "the one place the wide table
    # should win". It lost the first test by two orders of magnitude, and
    # the reason is the same one part 1 found: CREATE TABLE AS made a table
    # with no index, while the normalised enrolment table's primary key
    # begins with roll_no. "No joins" is not the advantage people assume;
    # the advantage, where there is one, is not reading other tables at all.
    print("\n=== 5. Read cost: where the wide table is supposed to win ===")
    cx = sqlite3.connect(path)
    roll = cx.execute("SELECT MIN(roll_no) FROM student").fetchone()[0]
    print("  (a) one student's transcript -- a lookup")
    a, ra = timed(cx, TRANSCRIPT_WIDE, (roll,))
    b, rb = timed(cx, TRANSCRIPT_NORM, (roll,))
    print(f"      report, no index on roll_no        {a:8.2f} ms  "
          f"{len(ra)} rows")
    print(f"      normalised, 1 join                 {b:8.2f} ms  "
          f"{len(rb)} rows")
    cx.execute("CREATE INDEX idx_report_roll ON report(roll_no)")
    c, rc = timed(cx, TRANSCRIPT_WIDE, (roll,))
    print(f"      report, WITH an index on roll_no   {c:8.2f} ms  "
          f"{len(rc)} rows")
    print(f"      -> without an index the wide table was {a / b:.0f}x "
          "SLOWER, despite having no join;")
    print(f"         with one, the two are within {max(b, c) / min(b, c):.1f}x "
          "of each other -- a tie.")
    print("  (b) enrolments and average credit hours per department "
          "-- a full scan")
    d, rd = timed(cx, DEPT_WIDE)
    e, re_ = timed(cx, DEPT_NORM)
    print(f"      report, no joins                   {d:8.2f} ms")
    print(f"      normalised, 2 joins                {e:8.2f} ms")
    same = sorted(round(x[2], 6) for x in rd) == \
        sorted(round(x[2], 6) for x in re_)
    print(f"      same averages: {'yes' if same else 'NO'}")
    print(f"      -> here the wide table IS faster, by {e / d:.1f}x: a "
          "report over every row")
    print("         reads one table instead of joining three.")
    cx.close()


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}, {STUDENTS:,} students"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "u.db")
        cx = build(path, students=STUDENTS)
        cx.executescript(REPORT)
        cx.commit()
        cx.close()
        # read costs measured on a clean copy, before any anomaly is staged
        clean = os.path.join(tmp, "clean.db")
        shutil.copy(path, clean)
        cx = sqlite3.connect(path)
        cx.execute("PRAGMA foreign_keys = ON")
        part1(path, cx)
        part2(cx)
        part3(cx)
        part4(cx)
        cx.close()
        part5(clean)
    print()


if __name__ == "__main__":
    main()
