"""Chapter 22 -- transactions, concurrency and recovery, made to happen.

Every part uses two real connections to one file, or a real failure, rather
than describing one. Python's sqlite3 module is put in true autocommit mode
(isolation_level=None) so that every BEGIN and COMMIT below is one we wrote.

  1. ATOMICITY. Move a student between sections -- a DELETE and an INSERT --
     and fail between the two. Without a transaction and with one.
  2. THE LOST UPDATE. Two clerks each add one to a section's seat count,
     read-then-write, interleaved. Without a transaction, with BEGIN
     DEFERRED, and with BEGIN IMMEDIATE.
  3. SQLITE_BUSY. One connection holds the write lock for 300 ms; a second
     tries to write with busy_timeout 0, and with busy_timeout 1000.
  4. DURABILITY, PRICED. 200 inserts, one commit each and one commit for
     all; and what synchronous = OFF buys, and what it gives up.

    python m22_transactions.py
"""
from __future__ import annotations

import os
import sqlite3
import tempfile
import threading
import time


def connect(path: str, timeout_ms: int = 0) -> sqlite3.Connection:
    cx = sqlite3.connect(path, isolation_level=None, timeout=timeout_ms / 1000,
                         check_same_thread=False)
    cx.execute(f"PRAGMA busy_timeout = {timeout_ms}")
    return cx


def fresh(path: str) -> None:
    for suffix in ("", "-journal", "-wal", "-shm"):
        if os.path.exists(path + suffix):
            os.remove(path + suffix)
    cx = connect(path)
    cx.executescript("""
      CREATE TABLE section (id TEXT PRIMARY KEY, seats_taken INTEGER NOT NULL)
        STRICT;
      CREATE TABLE enrolment (roll_no TEXT NOT NULL, section TEXT NOT NULL,
                              PRIMARY KEY (roll_no, section)) STRICT;
      INSERT INTO section VALUES ('CS-103/1', 40), ('CS-103/2', 38);
      INSERT INTO enrolment VALUES ('2021-IT-001', 'CS-103/1');
    """)
    cx.close()


class Crash(Exception):
    """Stands in for a power cut, a killed process, a dropped connection."""


def move(cx: sqlite3.Connection, use_txn: bool) -> None:
    if use_txn:
        cx.execute("BEGIN")
    try:
        cx.execute("DELETE FROM enrolment WHERE roll_no = '2021-IT-001' "
                   "AND section = 'CS-103/1'")
        raise Crash("the program dies here")
        cx.execute("INSERT INTO enrolment VALUES "      # never reached
                   "('2021-IT-001', 'CS-103/2')")
    except Crash:
        if use_txn:
            cx.execute("ROLLBACK")


def part1(tmp: str) -> None:
    print("\n=== 1. Atomicity: fail between the DELETE and the INSERT ===")
    print(f"  {'':<22}{'enrolments of 2021-IT-001 afterwards':<40}")
    for use_txn in (False, True):
        path = os.path.join(tmp, f"a{int(use_txn)}.db")
        fresh(path)
        cx = connect(path)
        move(cx, use_txn)
        rows = cx.execute("SELECT section FROM enrolment "
                          "WHERE roll_no = '2021-IT-001'").fetchall()
        cx.close()
        label = "BEGIN ... ROLLBACK" if use_txn else "autocommit"
        print(f"  {label:<22}{[r[0] for r in rows]!s:<16}"
              f"{'still in section 1' if rows else 'IN NOTHING -- lost'}")
    print("\n  In autocommit every statement is its own transaction, so the")
    print("  DELETE was permanent the moment it finished. Inside one")
    print("  transaction the two statements are one unit: all, or none.")


def read_seats(cx) -> int:
    return cx.execute("SELECT seats_taken FROM section "
                      "WHERE id = 'CS-103/1'").fetchone()[0]


def write_seats(cx, n) -> None:
    cx.execute("UPDATE section SET seats_taken = ? WHERE id = 'CS-103/1'",
               (n,))


def part2(tmp: str) -> None:
    print("\n=== 2. The lost update: two clerks, one seat count ===")
    print("  Both read the count, add one, write it back. Interleaved:")
    print("    A reads, B reads, A writes, B writes.   Start: 40.")
    print(f"\n  {'':<26}{'what happened':<46}{'final':>6}")

    # (a) no transaction: read and write are separate statements
    path = os.path.join(tmp, "l0.db")
    fresh(path)
    a, b = connect(path), connect(path)
    na = read_seats(a)
    nb = read_seats(b)
    write_seats(a, na + 1)
    write_seats(b, nb + 1)
    final = read_seats(a)
    a.close()
    b.close()
    print(f"  {'no transaction':<26}{'both read 40, both wrote 41':<46}"
          f"{final:>6}   <- one enrolment LOST")

    # (b) BEGIN DEFERRED: both take a read lock; the first to write wins
    #     the RESERVED lock, the second is refused -- and the first then
    #     cannot COMMIT while the second still holds its read lock.
    path = os.path.join(tmp, "l1.db")
    fresh(path)
    a, b = connect(path), connect(path)
    a.execute("BEGIN DEFERRED")
    b.execute("BEGIN DEFERRED")
    na = read_seats(a)
    nb = read_seats(b)
    write_seats(a, na + 1)
    events = []
    try:
        write_seats(b, nb + 1)
        events.append("B wrote")
    except sqlite3.OperationalError as e:
        events.append(f"B's write refused ({e})")
    try:
        a.execute("COMMIT")
        events.append("A committed")
    except sqlite3.OperationalError as e:
        events.append(f"A's COMMIT refused ({e})")
    b.execute("ROLLBACK")
    events.append("B rolled back")
    if a.in_transaction:
        a.execute("COMMIT")
        events.append("then A committed")
    final = read_seats(a)
    a.close()
    b.close()
    print(f"  {'BEGIN DEFERRED':<26}{events[0]:<46}{final:>6}")
    for ev in events[1:]:
        print(f"  {'':<26}{ev}")
    print(f"  {'':<26}-> a DEADLOCK: B held a read lock and wanted to write;")
    print(f"  {'':<26}   A held the write lock and needed B's read lock gone")
    print(f"  {'':<26}   to commit. Neither could proceed until B gave up,")
    print(f"  {'':<26}   so B's +1 must be retried -- but B was TOLD.")

    # (c) BEGIN IMMEDIATE: the write lock is taken at BEGIN, so B cannot
    #     even start until A has finished -- and then reads A's result.
    path = os.path.join(tmp, "l2.db")
    fresh(path)
    a, b = connect(path), connect(path, timeout_ms=0)
    a.execute("BEGIN IMMEDIATE")
    na = read_seats(a)
    try:
        b.execute("BEGIN IMMEDIATE")
        started = "B started"
    except sqlite3.OperationalError as e:
        started = f"B could not begin ({e})"
    write_seats(a, na + 1)
    a.execute("COMMIT")
    b.execute("BEGIN IMMEDIATE")
    nb = read_seats(b)
    write_seats(b, nb + 1)
    b.execute("COMMIT")
    final = read_seats(a)
    a.close()
    b.close()
    print(f"  {'BEGIN IMMEDIATE':<26}{started[:46]:<46}{final:>6}   "
          "<- both counted")
    print(f"  {'':<26}then B began again, read {nb}, wrote {nb + 1}")
    print("\n  DEFERRED takes a lock only when it first needs one, so two")
    print("  clerks can both read the old value and only discover the clash")
    print("  at the write -- when one of them must give up its work. IMMEDIATE")
    print("  takes the write lock at BEGIN: the second clerk waits before")
    print("  reading anything, and so reads the first clerk's result.")


def part3(tmp: str) -> None:
    print("\n=== 3. SQLITE_BUSY, and busy_timeout ===")
    print("  A holds the write lock for 300 ms; B tries to write meanwhile.")
    print(f"\n  {'B with busy_timeout':<24}{'outcome':<40}{'B waited':>10}")
    for timeout in (0, 1000):
        path = os.path.join(tmp, f"b{timeout}.db")
        fresh(path)
        a = connect(path)
        b = connect(path, timeout_ms=timeout)
        a.execute("BEGIN IMMEDIATE")
        write_seats(a, 99)

        def release():
            time.sleep(0.3)
            a.execute("COMMIT")
        th = threading.Thread(target=release)
        th.start()
        time.sleep(0.02)                      # make sure A holds the lock
        t0 = time.perf_counter()
        try:
            b.execute("UPDATE section SET seats_taken = seats_taken + 1 "
                      "WHERE id = 'CS-103/2'")
            out = "succeeded, after A committed"
        except sqlite3.OperationalError as e:
            out = f"failed at once: {e}"
        waited = (time.perf_counter() - t0) * 1000
        th.join()
        a.close()
        b.close()
        print(f"  {str(timeout) + ' ms':<24}{out:<40}{waited:>8.0f} ms")
    print("\n  SQLite allows one writer at a time for the whole file. With a")
    print("  timeout of 0 the second writer is told 'database is locked' and")
    print("  its work is simply refused; with a timeout it waits, and goes")
    print("  ahead the moment the lock is free. SET IT on every connection --")
    print("  the default is 0.")


def part4(tmp: str) -> None:
    print("\n=== 4. Durability, priced ===")
    n = 200
    rows = [(f"R-{i:05d}", "CS-103/1") for i in range(n)]
    print(f"  {n} inserts{'':<34}{'total':>10}{'per row':>12}")
    results = {}
    for label, sync, one in (("commit each, synchronous=FULL", "FULL", False),
                             ("one commit,  synchronous=FULL", "FULL", True),
                             ("commit each, synchronous=OFF ", "OFF", False)):
        path = os.path.join(tmp, f"d_{sync}_{int(one)}.db")
        fresh(path)
        cx = connect(path)
        cx.execute(f"PRAGMA synchronous = {sync}")
        t0 = time.perf_counter()
        if one:
            cx.execute("BEGIN")
        for r in rows:
            cx.execute("INSERT INTO enrolment VALUES (?, ?)", r)
        if one:
            cx.execute("COMMIT")
        ms = (time.perf_counter() - t0) * 1000
        cx.close()
        results[label] = ms
        print(f"  {label:<43}{ms:>8.1f} ms{ms / n:>10.3f} ms")
    each = results["commit each, synchronous=FULL"]
    once = results["one commit,  synchronous=FULL"]
    off = results["commit each, synchronous=OFF "]
    print(f"\n  One commit for all {n} was {each / once:,.0f}x faster than "
          f"one each.")
    print(f"  synchronous=OFF made one-commit-each {each / off:.0f}x faster -- "
          f"and was still {off / once:.0f}x")
    print("  slower than batching the inserts into one transaction.")
    print("\n  A COMMIT under synchronous=FULL waits until the disk confirms")
    print("  the data is really on it. That wait IS the durability, and on a")
    print("  mechanical disk it is most of the time. synchronous=OFF skips it,")
    print("  and a power cut can then lose transactions that were reported")
    print("  as committed, or corrupt the file. Batching keeps the guarantee")
    print("  AND is faster: group the writes that belong together.")


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
