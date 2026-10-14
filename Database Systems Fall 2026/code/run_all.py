#!/usr/bin/env python3
"""Run every measurement harness and check it still says what the handout says.

The handout quotes figures, and figures rot: SQLite is upgraded, the machine
changes, a harness is edited and nobody re-reads the chapter that cites it.
This runs each harness and matches its output against patterns describing the
SHAPE of each result the handout claims -- not the exact millisecond, which is
a fact about one machine, but the structure that would have to survive for the
chapter's argument to hold.

For example, Chapter 14's claim is not "NOT IN returned 0 in 0.76 ms"; it is
that NOT IN returned zero while NOT EXISTS returned three on the same data. The
pattern requires exactly that, and nothing about the timing.

Several claims ARE about exact counts -- 0 bytes for an empty database, 8,192
for one table -- because those are facts about SQLite's file format, not about
the machine, and the chapter says so. Those patterns are exact.

    python run_all.py              everything (a few minutes)
    python run_all.py --quick      skip the slow ones (20, 21)
    python run_all.py --ch 13 15   just those
    python run_all.py --list       what it would run

Exit status is the number of claims that failed, so this can gate a build.
Every harness works in a temporary directory and deletes what it made.
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))

# (chapter, argv, slow?, [(what the handout claims, pattern that shows it)])
CHECKS: list[tuple[int, list[str], bool, list[tuple[str, str]]]] = [
    (4, ["m04_firstdatabase.py"], False, [
        ("an empty database is 0 bytes",
         r"an empty database, just connected\s+0\b"),
        ("one CREATE TABLE costs two pages",
         r"after one CREATE TABLE, no rows\s+8,192"),
        ("the file grows by whole 4,096-byte pages",
         r"page size 4,096 bytes"),
        ("VACUUM recovers nothing with no deletions",
         r"10,000 rows\s+(\S+)\s[\s\S]*after VACUUM\s+\1"),
    ]),
    (8, ["m08_constraints.py"], False, [
        ("an orphan is accepted with foreign_keys untouched",
         r"default \(untouched\)\s+1\s+accepted"),
        ("and refused with it ON", r"\bON\s+0\s+refused"),
        ("the setting does not survive the connection",
         r"connection B, orphan insert\s+->\s+accepted"),
        ("TEXT PRIMARY KEY stores a null",
         r"TEXT PRIMARY KEY\s+accepted\s+null"),
        ("INTEGER PRIMARY KEY substitutes an integer",
         r"INTEGER PRIMARY KEY\s+accepted\s+integer"),
        ("STRICT refuses the null", r"STRICT\s+refused"),
    ]),
    (10, ["m10_ddl.py"], False, [
        ("'hello' is stored as text in all five declared types",
         r"'hello'\s+text\s+text\s+text\s+text\s+text"),
        ("STRICT refuses it", r"refused: cannot store TEXT value in INT"),
        ("MAX returns the text value", r"MAX\(score\):\s+C with 'absent'"),
        ("AVG reads the text as zero and still divides by all rows",
         r"AVG over the numeric rows only:\s+88\.75"),
        ("the cost of STRICT is compared against the spread",
         r"OVERLAP|do not overlap"),
    ]),
    (13, ["m13_joins.py", "--quick"], False, [
        ("three spellings agree", r"The first three agree"),
        ("NATURAL JOIN returns nothing", r"NATURAL JOIN returned 0\."),
        ("because two columns share a name",
         r"share BY NAME: roll_no, semester"),
        ("an inner join drops the unassigned sections",
         r"\d+ sections are missing from the inner-join"),
        ("SQLite builds an automatic index",
         r"AUTOMATIC COVERING INDEX \(name=\?\)"),
    ]),
    (15, ["m15_subqueries.py", "--quick"], False, [
        ("NOT IN returns 0", r"NOT IN\s+0\s[\s\S]*?one null empties it"),
        ("the three other forms return 3",
         r"NOT EXISTS\s+3\s[\s\S]*IS NULL\s+3\s[\s\S]*EXCEPT\s+3\s"),
        ("all four correlation variants agree",
         r"all four agree on every count: yes"),
        ("the index turns SCAN into SEARCH inside the sub-query",
         r"SEARCH e USING COVERING INDEX idx_enr_course"),
    ]),
    (16, ["m16_grouping.py", "--quick"], False, [
        ("COALESCE to zero lowers the average",
         r"the second is \d+% lower"),
        ("null forms its own group", r"None\s+[\d,]+"),
        ("an index removes the grouping sort",
         r"USE TEMP B-TREE FOR GROUP BY[\s\S]*COVERING INDEX idx_enr_course"),
        ("WHERE and HAVING compile to the same program",
         r"identical: YES"),
        ("an aggregate condition is tested after grouping",
         r"COUNT\(\*\) > 1000 .*AFTER grouping"),
        ("window and correlated ranks agree", r"x\s+yes"),
    ]),
    (18, ["m18_normalforms.py", "--quick"], False, [
        ("like for like, the wide table is larger",
         r"like for like, the wide table is [\d.]+x"),
        ("a half-finished update leaves two names",
         r"2 different names for one department"),
        ("a new course has nowhere to go in the wide table",
         r"rows that could hold it:\s+0"),
        ("deleting enrolments deletes the course",
         r"rows still describing the course:\s+0"),
        ("the wide table is slower on a lookup without an index",
         r"wide table was \d+x SLOWER"),
        ("and faster on a full scan", r"the wide table IS faster"),
    ]),
    (19, ["m19_decomposition.py", "--quick"], False, [
        ("a lossy split invents tuples",
         r"spurious tuples -- facts that were never true\s+[\d,]*[1-9]"),
        ("a lossless split invents none", r"spurious 0\b"),
        ("a lost dependency accepts a violating insert",
         r"second name for CS:\s+ACCEPTED"),
        ("the summary table stays correct through the writes",
         r"after \d+ more enrolments:\s+yes"),
        ("a break-even is computed", r"break-even|cheaper at any ratio"),
    ]),
    (20, ["m20_views.py", "--quick"], True, [
        ("a view compiles to the same instructions",
         r"same instructions in the same order: YES"),
        ("differing only in cursor numbers",
         r"only in which cursor number they use: yes"),
        ("a filter reaches inside a grouped view",
         r"CO-ROUTINE per_course\s+SEARCH enrolment USING COVERING INDEX"),
        # The handout's claim, after this check caught an earlier one failing
        # on a re-run: a trigger costs what its body costs -- neither it nor
        # the same INSERT written by hand is reliably cheaper.
        ("a trigger costs what the same work by hand costs",
         r"ranges OVERLAP:\s+neither is reliably cheaper"),
        ("after an UPDATE the view is right and the table is not",
         r"view correct: yes;\s+summary table correct: NO"),
    ]),
    (21, ["m21_indexes.py", "--quick"], True, [
        ("a point lookup barely grows with the table",
         r"200,000\s+[\d.]+ms\s+0\.\d+ms"),
        ("a crossover share is found",
         r"The index stops winning at about \d+%"),
        ("the optimiser chooses the index even for the whole table",
         r"100\.0%.*scan\s+index"),
        ("+k forces the scan", r"with \+k\s+\(SCAN t\)"),
        ("an index the query cannot use for ORDER BY sorts",
         r"ORDER BY v LIMIT 10.*TEMP B-TREE"),
    ]),
    (22, ["m22_transactions.py"], False, [
        ("autocommit loses the student",
         r"autocommit\s+\[\]\s+IN NOTHING"),
        ("a transaction keeps them",
         r"ROLLBACK\s+\['CS-103/1'\]\s+still in section 1"),
        ("no transaction loses an update",
         r"no transaction\s+both read 40, both wrote 41\s+41"),
        ("DEFERRED deadlocks", r"A's COMMIT refused \(database is locked\)"),
        ("IMMEDIATE counts both", r"BEGIN IMMEDIATE\s+.*\s42\s"),
        ("busy_timeout 0 fails at once, 1000 waits and succeeds",
         r"failed at once: database is locked[\s\S]*succeeded, after A"),
        ("one commit is far faster than one each",
         r"One commit for all \d+ was [\d,]+x faster"),
    ]),
]


def run(argv: list[str], timeout: int = 1200) -> tuple[int, str, float]:
    t0 = time.perf_counter()
    try:
        r = subprocess.run([sys.executable] + argv, cwd=HERE, text=True,
                           capture_output=True, timeout=timeout)
        return r.returncode, r.stdout + r.stderr, time.perf_counter() - t0
    except subprocess.TimeoutExpired:
        return 99, "(timed out)", time.perf_counter() - t0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    ap.add_argument("--ch", type=int, nargs="*")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    picked = [c for c in CHECKS
              if (not a.ch or c[0] in a.ch) and not (a.quick and c[2])]

    if a.list:
        for ch, argv, slow, claims in picked:
            print(f"  Ch {ch:>2}  {'slow ' if slow else '     '}"
                  f"{' '.join(argv):<36} {len(claims)} claim(s)")
        return 0

    missing = [argv[0] for _ch, argv, _s, _c in picked
               if not os.path.exists(os.path.join(HERE, argv[0]))]
    if missing:
        print(f"missing harness(es): {', '.join(missing)}", file=sys.stderr)
        return 1

    print(f"\n  Running {len(picked)} harness(es). "
          f"{'Quick mode: slow ones skipped.' if a.quick else ''}\n")
    head = f"  {'ch':>3}  {'harness':<26}{'secs':>7}  checks"
    print(head)
    print("  " + "-" * (len(head) + 10))

    total = passed = 0
    failures: list[str] = []
    for ch, argv, _slow, claims in picked:
        code, out, secs = run(argv)
        hits = []
        for claim, pattern in claims:
            total += 1
            ok = bool(re.search(pattern, out))
            passed += ok
            hits.append("." if ok else "X")
            if not ok:
                failures.append(f"Ch {ch}: {claim}")
        flag = "" if code == 0 else f"  (exit {code})"
        print(f"  {ch:>3}  {argv[0]:<26}{secs:>7.0f}  "
              f"{''.join(hits)}{flag}")
        if a.verbose or "X" in hits or code:
            for line in out.strip().splitlines()[-6:]:
                print(f"         {line[:92]}")

    print(f"\n  {passed}/{total} claims still hold")
    if failures:
        print("\n  not reproduced:")
        for f in failures:
            print(f"    {f}")
        print("\n  A failure here means the handout and SQLite now disagree.")
        print("  Find out which one moved before editing either.")
    return len(failures)


if __name__ == "__main__":
    sys.exit(main())
