"""Chapter 20 -- views and triggers: what they cost, and what they keep honest.

Four parts:
  1. A view against the query it wraps: compiled programs compared, as
     Chapter 16 compared WHERE and HAVING. Does a view cost anything?
  2. A filter applied to the view: is it pushed inside, or applied after
     the whole view has been computed?
  3. Trigger overhead per row: an audit trigger recording every grade
     change, against the same updates with no trigger -- one row per
     transaction, and all rows in one.
  4. The view and Chapter 19's summary table, side by side after the same
     UPDATE the summary table's triggers do not handle: which one is still
     right, and what each costs to read.

    python m20_views.py
    python m20_views.py --quick
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


def program(cx, sql):
    return [(r[1], r[2], r[3], r[4], str(r[5]))
            for r in cx.execute("EXPLAIN " + sql)]


def plan(cx, sql):
    return [r[3] for r in cx.execute("EXPLAIN QUERY PLAN " + sql)]


RAW = ("SELECT e.roll_no, c.course_code, c.title, e.semester, e.grade "
       "FROM enrolment e JOIN course c ON c.course_code = e.course_code")
VIEW_DDL = "CREATE VIEW transcript AS " + RAW


def part1(cx: sqlite3.Connection) -> None:
    print("\n=== 1. A view against the query it wraps ===")
    cx.execute(VIEW_DDL)
    roll = cx.execute("SELECT MIN(roll_no) FROM student").fetchone()[0]
    q_raw = RAW + " WHERE e.roll_no = ?"
    q_view = "SELECT * FROM transcript WHERE roll_no = ?"
    a, ra = timed(cx, q_raw, (roll,))
    b, rb = timed(cx, q_view, (roll,))
    pr = program(cx, q_raw.replace("?", f"'{roll}'"))
    pv = program(cx, q_view.replace("?", f"'{roll}'"))
    # The first draft compared every field of every instruction and printed
    # "identical: no" directly under the claim that a view costs nothing.
    # Both were true. The instructions are the same, in the same order; the
    # view's version numbers its table cursors one higher, and that is the
    # whole of the difference. So compare what the program DOES -- the
    # opcode sequence -- and then report separately what else differs.
    same_ops = [o[0] for o in pr] == [o[0] for o in pv]
    diffs = [(x, y) for x, y in zip(pr, pv) if x != y]
    # (opcode, p1, p2, p3, p4). Cursor numbers live in p1, and in p3 for
    # DeferredSeek; p2 (jump targets, page numbers) and p4 (key info) must
    # match exactly if the two programs do the same thing.
    only_cursor = all(x[0] == y[0] and x[2] == y[2] and x[4] == y[4]
                      for x, y in diffs)
    print(f"  the query written out     {a:7.3f} ms   {len(ra)} rows   "
          f"{len(pr)} ops")
    print(f"  the same through a view   {b:7.3f} ms   {len(rb)} rows   "
          f"{len(pv)} ops")
    print(f"  same rows: {'yes' if sorted(ra) == sorted(rb) else 'NO'};  "
          f"same instructions in the same order: "
          f"{'YES' if same_ops else 'no'}")
    print(f"  instructions that differ at all: {len(diffs)}, and they differ "
          f"only in which cursor number they use: "
          f"{'yes' if only_cursor else 'NO'}")
    print("\n  A view is stored as its SELECT text, and SQLite substitutes")
    print("  that text into the query that uses it before compiling. The")
    print("  program that results is the hand-written one with its cursors")
    print("  renumbered. There is nothing left at run time for the view to")
    print("  cost.")


def part2(cx: sqlite3.Connection) -> None:
    print("\n=== 2. Is a filter on the view pushed inside it? ===")
    roll = cx.execute("SELECT MIN(roll_no) FROM student").fetchone()[0]
    q = f"SELECT * FROM transcript WHERE roll_no = '{roll}'"
    for p in plan(cx, q):
        print(f"      {p}")
    print("  SEARCH e ... (roll_no=?): the WHERE reached the enrolment table")
    print("  inside the view and used its key. The view was never computed")
    print("  in full and then filtered.")

    # A view SQLite cannot flatten: one with an aggregate. It is run as a
    # co-routine, and the question is whether the outer WHERE still reaches
    # inside it. With no index on course_code the plan cannot tell you --
    # both forms must scan every enrolment anyway, so the first draft of
    # this part read a SCAN as "not pushed down" and was wrong. Add the
    # index and the plan answers the question.
    cx.execute("CREATE VIEW per_course AS SELECT course_code, COUNT(*) AS n "
               "FROM enrolment GROUP BY course_code")
    q2 = "SELECT * FROM per_course WHERE course_code = 'CS-101'"
    w2 = ("SELECT course_code, COUNT(*) FROM enrolment "
          "WHERE course_code = 'CS-101' GROUP BY course_code")
    print("\n  a view with GROUP BY, filtered on its grouping column")
    print("  -- no index on course_code:")
    for p in plan(cx, q2):
        print(f"      {p}")
    a, _ = timed(cx, q2)
    b, _ = timed(cx, w2)
    print(f"      through the view {a:.2f} ms; WHERE written inside "
          f"{b:.2f} ms -- both scan everything")
    cx.execute("CREATE INDEX idx_enr_course ON enrolment(course_code)")
    print("  -- WITH an index on course_code:")
    for p in plan(cx, q2):
        print(f"      {p}")
    a, _ = timed(cx, q2)
    b, _ = timed(cx, w2)
    print(f"      through the view {a:.3f} ms; WHERE written inside "
          f"{b:.3f} ms")
    print("  SEARCH ... (course_code=?) INSIDE the co-routine: the filter was")
    print("  pushed into the grouped view, and only one course's rows were")
    print("  ever grouped. Without the index it had nothing to search WITH.")
    cx.execute("DROP INDEX idx_enr_course")


AUDIT = """
CREATE TABLE grade_audit (
  roll_no     TEXT NOT NULL,
  course_code TEXT NOT NULL,
  semester    TEXT NOT NULL,
  section_no  INTEGER NOT NULL,
  old_grade   TEXT,
  new_grade   TEXT,
  changed_at  TEXT NOT NULL DEFAULT (datetime('now'))
) STRICT;
CREATE TRIGGER audit_grade AFTER UPDATE OF grade ON enrolment
WHEN OLD.grade IS NOT NEW.grade BEGIN
  INSERT INTO grade_audit (roll_no, course_code, semester, section_no,
                           old_grade, new_grade)
  VALUES (OLD.roll_no, OLD.course_code, OLD.semester, OLD.section_no,
          OLD.grade, NEW.grade);
END;
"""


def keys(cx, k):
    return cx.execute("SELECT roll_no, course_code, semester, section_no "
                      "FROM enrolment WHERE grade IS NOT NULL "
                      "ORDER BY 1, 2, 3, 4 LIMIT ?", (k,)).fetchall()


UPD = ("UPDATE enrolment SET grade = CASE grade WHEN 'A' THEN 'B' ELSE 'A' "
       "END WHERE roll_no = ? AND course_code = ? AND semester = ? "
       "AND section_no = ?")


def run_updates(cx, rows, one_txn: bool) -> float:
    t0 = time.perf_counter()
    if one_txn:
        cx.executemany(UPD, rows)
        cx.commit()
    else:
        for r in rows:
            cx.execute(UPD, r)
            cx.commit()
    return (time.perf_counter() - t0) * 1000 / len(rows)


MANUAL = ("INSERT INTO grade_audit (roll_no, course_code, semester, "
          "section_no, old_grade, new_grade) VALUES (?,?,?,?,NULL,NULL)")


def run_manual(cx, rows, one_txn: bool) -> float:
    """The control: the same UPDATE, plus an INSERT into a second table
    written by hand -- no trigger. If this costs what the trigger costs, the
    price is the second table, not the trigger mechanism."""
    t0 = time.perf_counter()
    if one_txn:
        for r in rows:
            cx.execute(UPD, r)
            cx.execute(MANUAL, r)
        cx.commit()
    else:
        for r in rows:
            cx.execute(UPD, r)
            cx.execute(MANUAL, r)
            cx.commit()
    return (time.perf_counter() - t0) * 1000 / len(rows)


def part3(tmp: str) -> None:
    # The first draft concluded the trigger's RELATIVE cost was "far higher"
    # in bulk. Measured, it was 5% one-per-transaction and 8% in bulk --
    # similar, not far apart. What WAS far apart is the absolute cost: about
    # 1.4 ms a row with a commit per row against 0.02 ms in bulk. A guess at
    # the reason is not a measurement, so this part now runs a control: the
    # same work done by hand, with no trigger at all.
    # Second draft. The control was added to test the explanation "the cost
    # is the second table, not the trigger", and it did not support it: the
    # same INSERT written by hand cost MORE than the trigger did. And a single
    # run with a commit per row is mostly disk-sync time on this machine's
    # mechanical disk -- about 29 ms a commit, varying by several ms -- so a
    # one-run difference of 1 ms means nothing. Hence: medians of REPS runs,
    # and a conclusion printed from the numbers rather than written in
    # advance.
    # Third draft. run_all.py then caught the second draft's conclusion
    # failing on a re-run: with 150-200 rows, the bulk case takes ~30 ms in
    # total, the trigger's added cost per row came out at -0.02, +0.05 and
    # -0.08 ms on three runs -- NEGATIVE twice -- and "the trigger was
    # cheaper than by hand" was one run that happened to separate. The two
    # modes need different sizes: a commit per row is slow, so keep it small;
    # one commit for everything is cheap, so measure it on enough rows that
    # the effect is larger than the noise. And report ranges for both.
    print("\n=== 3. Trigger overhead per row ===")
    k_single = 100 if QUICK else 200
    k_bulk = 10_000 if QUICK else 30_000
    reps = 3 if QUICK else 5

    def med(v):
        return sorted(v)[len(v) // 2]

    res = {}
    for mode in ("plain", "trigger", "manual"):
        singles, bulks = [], []
        for r in range(reps):
            cx = build(os.path.join(tmp, f"t_{mode}_{r}.db"),
                       students=STUDENTS)
            cx.executescript(AUDIT)
            if mode != "trigger":
                cx.execute("DROP TRIGGER audit_grade")
            cx.commit()
            run = run_manual if mode == "manual" else run_updates
            if r < 3:                       # the slow mode: three runs
                singles.append(run(cx, keys(cx, k_single), one_txn=False))
            bulks.append(run(cx, keys(cx, k_bulk), one_txn=True))
            cx.close()
            os.remove(os.path.join(tmp, f"t_{mode}_{r}.db"))
        res[mode] = (med(singles), med(bulks), min(singles), max(singles),
                     min(bulks), max(bulks))
    print(f"  commit per row: {k_single} changes, median of 3;  "
          f"one commit: {k_bulk:,} changes, median of {reps}")
    print(f"  {'':<38}{'one per txn':>12}{'(range)':>14}{'all in one':>12}"
          f"{'(range)':>18}")
    labels = {"plain": "UPDATE only",
              "trigger": "UPDATE + audit trigger",
              "manual": "UPDATE + same INSERT by hand"}
    for m in ("plain", "trigger", "manual"):
        s, b, lo, hi, blo, bhi = res[m]
        print(f"  {labels[m]:<38}{s:>9.3f} ms {lo:6.1f}-{hi:<6.1f}"
              f"{b:>9.4f} ms {blo:7.4f}-{bhi:<7.4f}")
    p, t, h = res["plain"], res["trigger"], res["manual"]
    print(f"\n  {'':<26}{'one per txn':>14}{'all in one':>16}")
    for lab, x in (("trigger adds per row", t), ("by hand adds per row", h)):
        print(f"  {lab:<26}{x[0] - p[0]:>9.3f} ms"
              f" ({(x[0] / p[0] - 1) * 100:+3.0f}%)"
              f"{x[1] - p[1]:>9.4f} ms ({(x[1] / p[1] - 1) * 100:+3.0f}%)")

    def apart(x, y):           # do the two bulk ranges fail to overlap?
        return x[5] < y[4] or y[5] < x[4]

    print()
    if apart(t, p):
        print(f"  In bulk the trigger's cost is a result: its range clears the")
        print(f"  plain update's. It adds {(t[1] / p[1] - 1) * 100:.0f}% to an "
              "update.")
    else:
        print("  In bulk the trigger's cost is inside the run-to-run range of")
        print("  a plain update: too small to separate from noise.")
    if apart(t, h):
        cheaper = "the trigger" if t[1] < h[1] else "the hand-written INSERT"
        print(f"  Against the same INSERT written by hand, {cheaper} is")
        print("  cheaper, and the ranges do not overlap.")
        if t[1] < h[1]:
            print("  It runs inside the UPDATE statement, in the engine; the")
            print("  hand-written version is a second statement the program")
            print("  must send for every row.")
    else:
        print("  Against the same INSERT written by hand, the ranges OVERLAP:")
        print("  neither is reliably cheaper. A trigger costs about what the")
        print("  work inside it costs -- the mechanism adds nothing measurable.")
    overlap = not (t[2] > p[3] or p[2] > t[3])
    if overlap:
        print("  One commit per row, the plain and trigger ranges OVERLAP: the")
        print("  disk sync dominates, and the trigger's share is inside its")
        print("  noise.")
    else:
        print("  One commit per row, the trigger's cost stands clear of the")
        print("  run-to-run range of the commit.")


SUMMARY_DDL = """
CREATE TABLE dept_stats (dept_code TEXT PRIMARY KEY NOT NULL,
                         n_enrol INTEGER NOT NULL) STRICT;
INSERT INTO dept_stats SELECT c.dept_code, COUNT(*) FROM enrolment e
  JOIN course c ON c.course_code = e.course_code GROUP BY c.dept_code;
CREATE TRIGGER s_ins AFTER INSERT ON enrolment BEGIN
  UPDATE dept_stats SET n_enrol = n_enrol + 1 WHERE dept_code =
    (SELECT dept_code FROM course WHERE course_code = NEW.course_code);
END;
CREATE TRIGGER s_del AFTER DELETE ON enrolment BEGIN
  UPDATE dept_stats SET n_enrol = n_enrol - 1 WHERE dept_code =
    (SELECT dept_code FROM course WHERE course_code = OLD.course_code);
END;
CREATE VIEW dept_view AS SELECT c.dept_code, COUNT(*) AS n_enrol
  FROM enrolment e JOIN course c ON c.course_code = e.course_code
 GROUP BY c.dept_code;
"""


def part4(tmp: str) -> None:
    print("\n=== 4. The view and the summary table, after an UPDATE ===")
    cx = build(os.path.join(tmp, "sv.db"), students=STUDENTS)
    cx.executescript(SUMMARY_DDL)
    cx.commit()
    before_t = dict(cx.execute("SELECT * FROM dept_stats"))
    before_v = dict(cx.execute("SELECT * FROM dept_view"))
    # move one enrolment from a CS course to an MA section in the same
    # term -- the change Chapter 19's triggers have no UPDATE trigger for
    src = cx.execute("SELECT e.rowid, e.semester FROM enrolment e JOIN "
                     "course c ON c.course_code = e.course_code "
                     "WHERE c.dept_code = 'CS' ORDER BY e.rowid "
                     "LIMIT 1").fetchone()
    dst = cx.execute("SELECT s.course_code, s.section_no FROM section s "
                     "JOIN course c ON c.course_code = s.course_code "
                     "WHERE c.dept_code = 'MA' AND s.semester = ? "
                     "ORDER BY 1, 2 LIMIT 1", (src[1],)).fetchone()
    cx.execute("UPDATE enrolment SET course_code = ?, section_no = ? "
               "WHERE rowid = ?", (dst[0], dst[1], src[0]))
    cx.commit()
    after_t = dict(cx.execute("SELECT * FROM dept_stats"))
    after_v = dict(cx.execute("SELECT * FROM dept_view"))
    truth = dict(cx.execute("SELECT c.dept_code, COUNT(*) FROM enrolment e "
                            "JOIN course c ON c.course_code = e.course_code "
                            "GROUP BY c.dept_code"))
    print("  one enrolment moved from a CS course to an MA course:")
    print(f"  {'':<6}{'truth':>8}{'view':>8}{'table':>8}")
    for d in ("CS", "MA"):
        print(f"  {d:<6}{truth[d]:>8,}{after_v[d]:>8,}{after_t[d]:>8,}"
              f"   (was {before_t[d]:,})")
    print(f"  view correct: {'yes' if after_v == truth else 'NO'};  "
          f"summary table correct: {'yes' if after_t == truth else 'NO'}")
    a, _ = timed(cx, "SELECT * FROM dept_view")
    b, _ = timed(cx, "SELECT * FROM dept_stats")
    print(f"\n  read the view {a:.2f} ms;  read the summary table {b:.3f} ms")
    print("  The view cannot be stale: it is the query, run now. The table")
    print("  is fast and is exactly as correct as its triggers are complete.")
    cx.close()


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}, {STUDENTS:,} students"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        cx = build(os.path.join(tmp, "u.db"), students=STUDENTS)
        part1(cx)
        part2(cx)
        cx.close()
        part3(tmp)
        part4(tmp)
    print()


if __name__ == "__main__":
    main()
