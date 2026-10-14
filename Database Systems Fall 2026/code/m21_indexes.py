"""Chapter 21 -- indexes, and where they stop helping.

Four parts:
  1. A point lookup, scanned and indexed, at three table sizes. The scan
     grows with the table; the index barely moves.
  2. THE CROSSOVER. A range query that selects a growing fraction of the
     table, forced to use the index and forced to scan. Past some fraction
     the index is SLOWER than reading everything -- this finds where.
  3. An ORDER BY the index cannot help, and one it can.
  4. What each index costs: bytes on disk, and time per insert.

The table is built directly rather than seeded, because the crossover needs
a million rows and a column whose order is unrelated to the table's own --
otherwise every "random" index lookup is really a sequential read.

    python m21_indexes.py
    python m21_indexes.py --quick       200,000 rows instead of 1,000,000
"""
from __future__ import annotations

import os
import random
import sqlite3
import sys
import tempfile
import time

QUICK = "--quick" in sys.argv
BIG = 200_000 if QUICK else 1_000_000


def make(path: str, n: int, seed: int = 21) -> sqlite3.Connection:
    if os.path.exists(path):
        os.remove(path)
    rng = random.Random(seed)
    cx = sqlite3.connect(path)
    cx.execute("CREATE TABLE t (id INTEGER PRIMARY KEY, k INTEGER NOT NULL, "
               "v TEXT NOT NULL) STRICT")
    # k is a random permutation of 0..n-1: unique, and in no relation to id,
    # so index order and table order are unrelated.
    ks = list(range(n))
    rng.shuffle(ks)
    pad = "x" * 60                    # a realistic row width, ~80 bytes
    cx.executemany("INSERT INTO t (k, v) VALUES (?, ?)",
                   ((k, f"{k:08d}{pad}") for k in ks))
    cx.commit()
    return cx


def timed(cx, sql, params=(), reps=3):
    best = float("inf")
    for _ in range(reps):
        t0 = time.perf_counter()
        cx.execute(sql, params).fetchall()
        best = min(best, (time.perf_counter() - t0) * 1000)
    return best


def plan(cx, sql, params=()):
    return " / ".join(r[3] for r in
                      cx.execute("EXPLAIN QUERY PLAN " + sql, params))


def part1(tmp: str) -> None:
    print("\n=== 1. A point lookup, scanned and indexed ===")
    sizes = [10_000, 100_000, BIG]
    print(f"  {'rows':>11}{'scan':>11}{'index':>11}{'speed-up':>11}")
    print("  " + "-" * 44)
    for n in sizes:
        cx = make(os.path.join(tmp, f"p{n}.db"), n)
        q = "SELECT v FROM t WHERE k = ?"
        key = (n // 2,)
        a = timed(cx, q, key)
        cx.execute("CREATE INDEX ik ON t(k)")
        b = timed(cx, q, key)
        cx.close()
        os.remove(os.path.join(tmp, f"p{n}.db"))
        print(f"  {n:>11,}{a:>9.2f}ms{b:>9.3f}ms{a / b:>10,.0f}x")
    print("\n  The scan reads every row, so it grows with the table. The")
    print("  index descends a B-tree whose depth grows with the LOG of the")
    print("  table, so a hundred times the rows costs almost nothing more.")


def part2(tmp: str) -> list[tuple]:
    print(f"\n=== 2. The crossover: a range over a growing share of "
          f"{BIG:,} rows ===")
    cx = make(os.path.join(tmp, "x.db"), BIG)
    cx.execute("CREATE INDEX ik ON t(k)")
    cx.execute("ANALYZE")
    # Force each plan, so both are measured at every fraction rather than
    # only the one the optimiser would have chosen.
    via_index = ("SELECT SUM(length(v)) FROM t INDEXED BY ik "
                 "WHERE k BETWEEN ? AND ?")
    via_scan = ("SELECT SUM(length(v)) FROM t NOT INDEXED "
                "WHERE k BETWEEN ? AND ?")
    chosen = ("SELECT SUM(length(v)) FROM t WHERE k BETWEEN ? AND ?")
    fracs = [0.001, 0.01, 0.02, 0.05, 0.10, 0.20, 0.35, 0.50, 1.00]
    rows = []
    print(f"  {'share':>7}{'rows':>11}{'via index':>12}{'full scan':>12}"
          f"{'winner':>9}   optimiser chose")
    print("  " + "-" * 70)
    for f in fracs:
        hi = int(BIG * f) - 1
        p = (0, hi)
        a = timed(cx, via_index, p)
        b = timed(cx, via_scan, p)
        ch = "index" if "USING INDEX" in plan(cx, chosen, p) else "scan"
        win = "index" if a < b else "scan"
        rows.append((f, hi + 1, a, b))
        print(f"  {f * 100:>6.1f}%{hi + 1:>11,}{a:>10.1f}ms{b:>10.1f}ms"
              f"{win:>9}   {ch}")
    cx.close()
    # interpolate the crossover share
    cross = None
    for (f0, _, a0, b0), (f1, _, a1, b1) in zip(rows, rows[1:]):
        if a0 < b0 and a1 >= b1:
            d0, d1 = b0 - a0, a1 - b1
            cross = f0 + (f1 - f0) * d0 / (d0 + d1)
            break
    if cross is not None:
        print(f"\n  The index stops winning at about {cross * 100:.0f}% of the "
              "table.")
    else:
        print("\n  No crossover within the fractions measured.")
    print("  Below that, it reads only the rows it needs. Above it, each of")
    print("  those reads is a separate jump to a different page of the table,")
    print("  while a scan reads the pages in order, once each -- and a jump")
    print("  per row loses to a straight read once enough rows are wanted.")

    # The optimiser chose the index at EVERY fraction, including 100%, where
    # it was twenty times slower -- and this is after ANALYZE. The reason is
    # in what ANALYZE could record. Without SQLITE_ENABLE_STAT4 (not in the
    # build Python ships) it stores only a row count and the average number
    # of rows per index key: it knows k is unique, and nothing about how the
    # values are spread. So it cannot tell BETWEEN 0 AND 999 from BETWEEN 0
    # AND 999999, and treats every range on an indexed column as selective.
    stat4 = any("STAT4" in r[0] for r in cx_opts())
    print("\n  But look at the last column: the optimiser chose the index at")
    print("  EVERY share, including 100%, after ANALYZE.")
    print(f"    STAT4 histograms compiled into this SQLite: "
          f"{'yes' if stat4 else 'NO'}")
    cx = sqlite3.connect(os.path.join(tmp, "x.db"))
    print(f"    what ANALYZE recorded (sqlite_stat1): "
          f"{cx.execute('SELECT stat FROM sqlite_stat1').fetchone()[0]!r}"
          "  -- rows, rows per k")
    print("    It knows how MANY rows and that k is unique. It does not know")
    print("    how the values are spread, so every range looks selective.")
    every = (0, BIG - 1)
    plus = ("SELECT SUM(length(v)) FROM t WHERE +k BETWEEN ? AND ?")
    a = timed(cx, chosen, every)
    b = timed(cx, plus, every)
    print(f"\n  the whole table, as written ({plan(cx, chosen, every)[:28]})"
          f"  {a:8.1f} ms")
    print(f"  the same with +k             ({plan(cx, plus, every)[:28]})"
          f"  {b:8.1f} ms")
    print("  A unary + makes the column an expression the index cannot")
    print("  serve, and so forces the scan for that one query.")
    cx.close()
    return rows


def cx_opts():
    return sqlite3.connect(":memory:").execute("PRAGMA compile_options")


def part3(tmp: str) -> None:
    print("\n=== 3. ORDER BY: one the index helps, one it cannot ===")
    cx = make(os.path.join(tmp, "o.db"), 100_000)
    cx.execute("CREATE INDEX ik ON t(k)")
    cases = [
        ("ORDER BY k LIMIT 10   (the indexed column)",
         "SELECT v FROM t ORDER BY k LIMIT 10"),
        ("ORDER BY v LIMIT 10   (not indexed)",
         "SELECT v FROM t ORDER BY v LIMIT 10"),
        ("ORDER BY k % 7, k LIMIT 10   (an expression)",
         "SELECT v FROM t ORDER BY k % 7, k LIMIT 10"),
    ]
    for label, sql in cases:
        ms = timed(cx, sql)
        p = plan(cx, sql)
        tmpb = "TEMP B-TREE" in p
        print(f"  {label:<46}{ms:>8.2f} ms   "
              f"{'sorts (TEMP B-TREE)' if tmpb else 'reads in index order'}")
    cx.close()
    print("\n  An index is a sorted copy of ONE ordering. A query wanting that")
    print("  order gets it free -- and stops after 10 rows. Any other order,")
    print("  including an expression OF the indexed column, means sorting")
    print("  every row first to find the ten.")


def part4(tmp: str) -> None:
    # Medians of three. A single run put three indexes FASTER to insert into
    # than two, which is noise -- the third index cannot make inserting
    # cheaper -- and a reader would have been right to distrust the rest.
    print("\n=== 4. What each index costs (insert: median of 3) ===")
    n, reps = 100_000, 3
    print(f"  {n:,} rows   {'file bytes':>13}{'insert 10,000 more':>22}"
          f"{'(range)':>16}")
    defs = ["CREATE INDEX i1 ON t(k)", "CREATE INDEX i2 ON t(v)",
            "CREATE INDEX i3 ON t(k, v)"]
    for k_idx in range(0, 4):
        times, size = [], 0
        for r in range(reps):
            path = os.path.join(tmp, f"c{k_idx}_{r}.db")
            cx = make(path, n)
            for d in defs[:k_idx]:
                cx.execute(d)
            cx.commit()
            cx.execute("VACUUM")
            size = os.path.getsize(path)
            extra = [(n + i, f"{n + i:08d}" + "y" * 60)
                     for i in range(10_000)]
            t0 = time.perf_counter()
            cx.executemany("INSERT INTO t (k, v) VALUES (?, ?)", extra)
            cx.commit()
            times.append((time.perf_counter() - t0) * 1000)
            cx.close()
            os.remove(path)
        med = sorted(times)[reps // 2]
        print(f"  {k_idx} index{'es' if k_idx != 1 else '  '}"
              f"{size:>16,}{med:>18.1f} ms"
              f"{min(times):>9.1f}-{max(times):<6.1f}")
    print("\n  Read the ranges before the medians. 0 and 1 index overlap: the")
    print("  new rows arrive in key order, so each new entry goes on the")
    print("  right-hand end of the index -- the cheapest case there is, as")
    print("  for roll numbers issued in sequence. The index on v, a 68-byte")
    print("  text column, is the one that costs: it doubled the insert time")
    print("  and added more bytes than the integer index and the table's")
    print("  growth together. Every index is another B-tree kept in step")
    print("  with the table on every write, whether or not any query uses it;")
    print("  what it costs depends on how wide its keys are and on the order")
    print("  they arrive in.")


def main() -> None:
    print(f"\n  SQLite {sqlite3.sqlite_version}"
          + ("   (--quick)" if QUICK else ""))
    with tempfile.TemporaryDirectory() as tmp:
        part1(tmp)
        part2(tmp)
        part3(tmp)
        part4(tmp)
    print()


if __name__ == "__main__":
    main()
