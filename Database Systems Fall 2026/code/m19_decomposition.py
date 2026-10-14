"""Chapter 19 -- decompositions that lose things, and denormalising on purpose.

Three parts:
  1. A LOSSY decomposition: split enrolment into (roll_no, course_code) and
     (course_code, grade), join them back, and count the tuples that were
     never in the original. Then the lossless split, and the one-line test
     that tells them apart.
  2. A decomposition that is lossless but loses a DEPENDENCY: once
     dept_code -> dept_name lives in no single table, nothing refuses data
     that breaks it.
  3. Denormalisation as a measured decision. A per-department summary,
     computed on demand by joins, against a summary TABLE kept current by
     triggers: what each read saves, what each write costs, and the number
     of reads per write at which the summary table pays for itself.

    python m19_decomposition.py
    python m19_decomposition.py --quick
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
STUDENTS = 2000 if QUICK else 4000


def timed(cx, sql, params=(), reps=3):
    best, out = float("inf"), None
    for _ in range(reps):
        t0 = time.perf_counter()
        out = cx.execute(sql, params).fetchall()
        best = min(best, (time.perf_counter() - t0) * 1000)
    return best, out


def part1(cx: sqlite3.Connection) -> None:
    print("\n=== 1. Lossy and lossless decompositions of enrolment ===")
    cx.executescript("""
      CREATE TEMP TABLE r1 AS SELECT DISTINCT roll_no, course_code
                                FROM enrolment WHERE grade IS NOT NULL;
      CREATE TEMP TABLE r2 AS SELECT DISTINCT course_code, grade
                                FROM enrolment WHERE grade IS NOT NULL;
    """)
    n1 = cx.execute("SELECT COUNT(*) FROM r1").fetchone()[0]
    n2 = cx.execute("SELECT COUNT(*) FROM r2").fetchone()[0]
    back = cx.execute("SELECT COUNT(*) FROM r1 JOIN r2 "
                      "USING (course_code)").fetchone()[0]
    real = cx.execute("SELECT COUNT(*) FROM (SELECT DISTINCT roll_no, "
                      "course_code, grade FROM enrolment "
                      "WHERE grade IS NOT NULL)").fetchone()[0]
    print("  LOSSY: R1(roll_no, course_code) and R2(course_code, grade)")
    print(f"    graded (roll_no, course_code, grade) facts     {real:>9,}")
    print(f"    R1 rows {n1:,},  R2 rows {n2:,}")
    print(f"    R1 JOIN R2 on course_code                       {back:>9,}")
    print(f"    spurious tuples -- facts that were never true   "
          f"{back - real:>9,}")
    print(f"    the join returned {back / real:.1f}x the real facts. Every "
          "student now")
    print("    appears to hold every grade anybody got in the course.")

    cx.executescript("""
      CREATE TEMP TABLE s1 AS SELECT roll_no, name, programme FROM student;
      CREATE TEMP TABLE s2 AS SELECT roll_no, semester, cnic FROM student;
    """)
    sn = cx.execute("SELECT COUNT(*) FROM student").fetchone()[0]
    sb = cx.execute("SELECT COUNT(*) FROM s1 JOIN s2 "
                    "USING (roll_no)").fetchone()[0]
    print("\n  LOSSLESS: S1(roll_no, name, programme), "
          "S2(roll_no, semester, cnic)")
    print(f"    student rows {sn:,};  S1 JOIN S2 {sb:,};  spurious "
          f"{sb - sn}")
    print("\n  The test: the shared attributes must be a superkey of one side.")
    print("    lossy:    R1 and R2 share {course_code} -- a key of neither")
    print("    lossless: S1 and S2 share {roll_no}     -- the key of both")


def part2(tmp: str) -> None:
    print("\n=== 2. Lossless, but a dependency is lost ===")
    cx = sqlite3.connect(os.path.join(tmp, "dp.db"))
    # course_dept(course_code, dept_code) and course_dname(course_code,
    # dept_name): both keyed on course_code, so the join is lossless -- but
    # dept_code -> dept_name is in neither table.
    cx.executescript("""
      CREATE TABLE course_dept  (course_code TEXT PRIMARY KEY,
                                 dept_code   TEXT NOT NULL) STRICT;
      CREATE TABLE course_dname (course_code TEXT PRIMARY KEY,
                                 dept_name   TEXT NOT NULL) STRICT;
      INSERT INTO course_dept  VALUES ('CS-101','CS'), ('CS-150','CS');
      INSERT INTO course_dname VALUES ('CS-101','Computer Science');
    """)
    try:
        cx.execute("INSERT INTO course_dname VALUES ('CS-150','Computing')")
        cx.commit()
        verdict = "ACCEPTED"
    except sqlite3.Error as e:
        verdict = f"refused ({e})"
    rows = cx.execute("SELECT d.dept_code, n.dept_name FROM course_dept d "
                      "JOIN course_dname n USING (course_code) "
                      "ORDER BY 1, 2").fetchall()
    print("  two courses, both in CS; insert a second name for CS:  "
          f"{verdict}")
    print("  the join now says:")
    for r in rows:
        print(f"      dept_code {r[0]}  ->  {r[1]!r}")
    names = len({r[1] for r in rows})
    print(f"  dept_code -> dept_name is violated ({names} names for one "
          "code), and no")
    print("  constraint in either table could have refused it: the rule")
    print("  spans two tables, and a CHECK sees one row of one table.")
    cx.close()


ONDEMAND = ("SELECT d.dept_code, COUNT(*), SUM(c.credit_hours) "
            "FROM enrolment e JOIN course c ON c.course_code = e.course_code "
            "JOIN department d ON d.dept_code = c.dept_code "
            "GROUP BY d.dept_code")
SUMMARY = "SELECT dept_code, n_enrol, sum_credits FROM dept_stats"

TRIGGERS = """
CREATE TABLE dept_stats (
  dept_code   TEXT PRIMARY KEY NOT NULL REFERENCES department(dept_code),
  n_enrol     INTEGER NOT NULL,
  sum_credits INTEGER NOT NULL
) STRICT;
INSERT INTO dept_stats
  SELECT d.dept_code, COUNT(e.roll_no), COALESCE(SUM(c.credit_hours), 0)
    FROM department d
    LEFT JOIN course c    ON c.dept_code = d.dept_code
    LEFT JOIN enrolment e ON e.course_code = c.course_code
   GROUP BY d.dept_code;
CREATE TRIGGER stats_ins AFTER INSERT ON enrolment BEGIN
  UPDATE dept_stats
     SET n_enrol = n_enrol + 1,
         sum_credits = sum_credits +
           (SELECT credit_hours FROM course WHERE course_code = NEW.course_code)
   WHERE dept_code =
           (SELECT dept_code FROM course WHERE course_code = NEW.course_code);
END;
CREATE TRIGGER stats_del AFTER DELETE ON enrolment BEGIN
  UPDATE dept_stats
     SET n_enrol = n_enrol - 1,
         sum_credits = sum_credits -
           (SELECT credit_hours FROM course WHERE course_code = OLD.course_code)
   WHERE dept_code =
           (SELECT dept_code FROM course WHERE course_code = OLD.course_code);
END;
"""


def new_enrolments(cx: sqlite3.Connection, k: int) -> list[tuple]:
    """k enrolments that do not exist yet: real students, real sections,
    in a new term, so the composite key is never violated."""
    studs = [r[0] for r in cx.execute(
        "SELECT roll_no FROM student ORDER BY roll_no LIMIT ?", (k,))]
    secs = cx.execute("SELECT course_code, section_no FROM section "
                      "WHERE semester = 'FALL2026' ORDER BY 1, 2").fetchall()
    cx.executemany("INSERT OR IGNORE INTO section VALUES (?,?,?,?,?)",
                   [(c, "SPRING2027", n, "R1", None) for c, n in secs])
    return [(s, secs[i % len(secs)][0], "SPRING2027",
             secs[i % len(secs)][1], None) for i, s in enumerate(studs)]


def insert_cost(cx: sqlite3.Connection, rows: list[tuple]) -> float:
    """ms per row, one row per transaction -- as an enrolment screen
    inserts them, not as a bulk load does."""
    t0 = time.perf_counter()
    for r in rows:
        cx.execute("INSERT INTO enrolment VALUES (?,?,?,?,?)", r)
        cx.commit()
    return (time.perf_counter() - t0) * 1000 / len(rows)


def part3(tmp: str) -> None:
    print("\n=== 3. Denormalisation as a measured decision ===")
    print("  'enrolments and credit hours per department'")
    k = 200 if QUICK else 400

    cx = build(os.path.join(tmp, "plain.db"), students=STUDENTS)
    read_join, ra = timed(cx, ONDEMAND)
    rows = new_enrolments(cx, k)
    cx.commit()
    w_plain = insert_cost(cx, rows)
    cx.close()

    cx = build(os.path.join(tmp, "trig.db"), students=STUDENTS)
    cx.executescript(TRIGGERS)
    cx.commit()
    read_tab, rb = timed(cx, SUMMARY)
    same = sorted(ra) == sorted(rb)
    rows = new_enrolments(cx, k)
    cx.commit()
    w_trig = insert_cost(cx, rows)
    # does the summary still agree with the truth after the writes?
    truth = sorted(cx.execute(ONDEMAND).fetchall())
    kept = sorted(cx.execute(SUMMARY).fetchall())
    cx.close()

    saved = read_join - read_tab
    extra = w_trig - w_plain
    print(f"  {'':<40}{'ms':>10}")
    print(f"  {'read: computed by two joins':<40}{read_join:>10.2f}")
    print(f"  {'read: from the summary table':<40}{read_tab:>10.3f}")
    print(f"  {'  saved per read':<40}{saved:>10.2f}")
    print(f"  {'write: one enrolment, no trigger':<40}{w_plain:>10.3f}")
    print(f"  {'write: one enrolment, with trigger':<40}{w_trig:>10.3f}")
    print(f"  {'  added per write':<40}{extra:>10.3f}")
    print(f"\n  summary agreed with the joins before the writes: "
          f"{'yes' if same else 'NO'}")
    print(f"  ...and after {k} more enrolments:              "
          f"{'yes' if truth == kept else 'NO'}")
    if extra > 0:
        print(f"\n  break-even: the summary table pays for itself once there")
        print(f"  are more than {extra / saved:.3f} reads per write -- i.e. "
              f"one read every {saved / extra:,.0f} writes.")
    else:
        print("\n  the trigger added no measurable cost per write; the summary")
        print("  table is cheaper at any ratio of reads to writes.")
    print("  Note what the per-write cost is mostly made of: the COMMIT.")
    print("  The trigger's two UPDATEs ride inside a transaction that was")
    print("  going to be made durable anyway.")


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}, {STUDENTS:,} students"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        cx = build(os.path.join(tmp, "u.db"), students=STUDENTS)
        part1(cx)
        cx.close()
        part2(tmp)
        part3(tmp)
    print()


if __name__ == "__main__":
    main()
