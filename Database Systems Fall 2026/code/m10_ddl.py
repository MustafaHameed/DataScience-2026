"""Chapter 10 -- what a declared column type actually promises.

Chapter 8 measured a constraint that was declared and not enforced. This is
the same shape one level down: a column declared INTEGER that stores text.

SQLite's type system is "type affinity": a declared type is a PREFERENCE.
Given a value it cannot convert, the column keeps it as it is. The same table
declared STRICT refuses it. One keyword apart.

Four parts:
  1. What each declared type does with the same five values.
  2. The same table STRICT.
  3. Why it matters: a mixed-type column sorts and aggregates by a rule
     nobody expects, so MAX() returns the wrong row and no error is raised.
  4. What STRICT costs.

    python m10_ddl.py
"""
from __future__ import annotations

import os
import sqlite3
import tempfile
import time

VALUES = [42, "42", "hello", 3.7, None]
LABELS = ["42 (int)", "'42' (text)", "'hello'", "3.7", "NULL"]


def part1(cx: sqlite3.Connection) -> None:
    print("\n=== 1. One value, five declared types: what is STORED ===")
    types = ["INTEGER", "TEXT", "REAL", "NUMERIC", ""]
    cols = ", ".join(f"c{i} {t}" for i, t in enumerate(types))
    cx.executescript(f"CREATE TABLE affinity ({cols})")
    print(f"  {'inserted':<14}" + "".join(
        f"{t or '(no type)':<11}" for t in types))
    print("  " + "-" * 70)
    for label, v in zip(LABELS, VALUES):
        cx.execute("DELETE FROM affinity")
        cx.execute(f"INSERT INTO affinity VALUES ({','.join('?' * 5)})",
                   [v] * 5)
        got = cx.execute(
            "SELECT " + ", ".join(f"typeof(c{i})" for i in range(5))
            + " FROM affinity").fetchone()
        print(f"  {label:<14}" + "".join(f"{g:<11}" for g in got))
    print("\n  Read the 'hello' row. Every column was asked to store text it")
    print("  could not convert, and every one of them kept it.")


def part2(cx: sqlite3.Connection) -> None:
    print("\n=== 2. The same table, one keyword different ===")
    cx.executescript("CREATE TABLE strict_t (c INTEGER) STRICT")
    cx.executescript("CREATE TABLE loose_t  (c INTEGER)")
    print(f"  {'inserted into c INTEGER':<28}{'plain':<22}STRICT")
    print("  " + "-" * 70)
    for label, v in zip(LABELS, VALUES):
        out = []
        for tbl in ("loose_t", "strict_t"):
            cx.execute(f"DELETE FROM {tbl}")
            try:
                cx.execute(f"INSERT INTO {tbl} VALUES (?)", (v,))
                t = cx.execute(
                    f"SELECT typeof(c) FROM {tbl}").fetchone()[0]
                out.append(f"stored as {t}")
            except sqlite3.IntegrityError as e:
                out.append(f"refused: {str(e)[:30]}")
        print(f"  {label:<28}{out[0]:<22}{out[1]}")


def part3(cx: sqlite3.Connection) -> None:
    print("\n=== 3. Why it matters: ORDER BY and MAX on a mixed column ===")
    cx.executescript("CREATE TABLE marks (roll TEXT, score INTEGER)")
    rows = [("A", 95), ("B", 72), ("C", "absent"), ("D", 88), ("E", "100")]
    cx.executemany("INSERT INTO marks VALUES (?,?)", rows)
    print("  inserted:  " + ",  ".join(f"{r}={s!r}" for r, s in rows))
    print("\n  SELECT roll, score, typeof(score) FROM marks ORDER BY score:")
    for roll, score, t in cx.execute(
            "SELECT roll, score, typeof(score) FROM marks ORDER BY score"):
        print(f"      {roll}   {str(score):<8}{t}")
    top = cx.execute("SELECT roll, score FROM marks "
                     "WHERE score = (SELECT MAX(score) FROM marks)").fetchone()
    n_all = cx.execute("SELECT COUNT(*) FROM marks").fetchone()[0]
    n_num = cx.execute("SELECT COUNT(*) FROM marks "
                       "WHERE typeof(score) IN ('integer','real')").fetchone()[0]
    avg = cx.execute("SELECT AVG(score) FROM marks").fetchone()[0]
    # The average is the worst of the three, and it is worth computing the
    # honest one beside it. AVG does not skip the text row the way it skips a
    # NULL -- it casts 'absent' to 0 and still divides by 5, so the class
    # average is reported about 20% low and nothing says so.
    real = cx.execute("SELECT AVG(score) FROM marks WHERE "
                      "typeof(score) IN ('integer','real')").fetchone()[0]
    print(f"\n  highest mark, per MAX(score):   {top[0]} with {top[1]!r}")
    print(f"  rows: {n_all} total, {n_num} actually numeric")
    print(f"  AVG(score) as written:          {avg:.2f}")
    print(f"  AVG over the numeric rows only: {real:.2f}")
    print(f"  the reported average is         {(avg / real - 1) * 100:.0f}% out")
    print("\n  Three things went wrong and none of them raised an error:")
    print("    ORDER BY  sorts by storage class first (NULL < number < text),")
    print("              so every text value sorts ABOVE every number;")
    print("    MAX       therefore returns 'absent', and the student with 95")
    print("              is not the top of the class;")
    print("    AVG       does not skip the text row the way it skips a NULL.")
    print("              It reads 'absent' as 0 and still divides by 5.")
    print("\n  Note also that '100' was converted to the integer 100 on the")
    print("  way in. Affinity silently succeeded there and silently failed")
    print("  on 'absent', and the column looks the same either way.")


def part4(tmp: str) -> None:
    print("\n=== 4. What STRICT costs ===")
    n, reps = 50000, 5
    timing: dict[bool, list[float]] = {False: [], True: []}
    for r in range(reps):
        for strict in (False, True):
            path = os.path.join(tmp, f"s{int(strict)}_{r}.db")
            cx = sqlite3.connect(path)
            cx.execute("CREATE TABLE t (a INTEGER, b TEXT)"
                       + (" STRICT" if strict else ""))
            rows = [(i, f"row {i}") for i in range(n)]
            t0 = time.perf_counter()
            cx.executemany("INSERT INTO t VALUES (?,?)", rows)
            cx.commit()
            timing[strict].append((time.perf_counter() - t0) * 1000)
            cx.close()
            os.remove(path)

    def med(v: list[float]) -> float:
        return sorted(v)[len(v) // 2]

    print(f"  {n:,} rows, {reps} runs")
    print(f"  {'':<22}{'median':>9}{'min':>9}{'max':>9}")
    for strict, label in ((False, "plain"), (True, "STRICT")):
        v = timing[strict]
        print(f"  {label:<22}{med(v):8.1f}{min(v):9.1f}{max(v):9.1f}")
    d = med(timing[True]) - med(timing[False])
    print(f"  {'difference':<22}{d:8.1f} ms"
          f"   ({d / med(timing[False]) * 100:+.0f}%)")
    # State the comparison rather than leaving the reader to make it. If the
    # two ranges overlap, the median difference is not a result -- which is
    # the opposite of Chapter 8's foreign-key figure, where they did not.
    overlap = (min(timing[True]) <= max(timing[False])
               and min(timing[False]) <= max(timing[True]))
    print("\n  The two ranges " + ("OVERLAP" if overlap else "do not overlap")
          + ", so the median difference is "
          + ("noise, not a cost." if overlap else "a real cost."))
    if overlap:
        print("  STRICT is free. It is checking a type it already knows.")


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}")
    with tempfile.TemporaryDirectory() as tmp:
        cx = sqlite3.connect(os.path.join(tmp, "affinity.db"))
        part1(cx)
        part2(cx)
        part3(cx)
        cx.close()
        part4(tmp)
    print()


if __name__ == "__main__":
    main()
