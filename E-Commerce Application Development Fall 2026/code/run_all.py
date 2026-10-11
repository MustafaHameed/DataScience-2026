#!/usr/bin/env python3
"""Run every measurement harness and check it still says what the handout says.

The handout quotes figures. Figures rot: the platform changes, the machine
changes, a harness is edited and nobody re-reads the chapter that cites it.
This runs each harness and matches its output against a pattern describing
the SHAPE of the result the handout claims -- not the exact number, which is
a fact about one laptop, but the structure that would have to survive for the
chapter's argument to hold.

For example, Chapter 9 claims the session record stays at 18 bytes while the
basket grows. The pattern does not require "18"; it requires that the session
column holds the same value on every row. If that stops being true, the
chapter is wrong whatever the number is.

    python run_all.py                  everything (fifteen minutes or so)
    python run_all.py --quick          skip the slow ones (6, 14, 20)
    python run_all.py --ch 9 10        just those
    python run_all.py --list           what it would run

Exit status is the number of failures, so this can gate a build.

NOTE: several harnesses change the shop -- they seed catalogues, place orders,
install extensions and toggle settings. Each cleans up after itself, and the
whole thing is safe to run against the teaching container because the
teaching container is disposable (Chapter 5). Do not point it at a shop you
care about.
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
    (5, ["m05_coldstart.py", "--warm"], False, [
        ("a warm start reaches HTTP 200", r"200|ok\s"),
        ("the stack's footprint is reported", r"(?i)MB|GB"),
    ]),
    (6, ["m06_catalogue.py", "--quick"], True, [
        ("page time grows with the catalogue", r"\b100\b.*\b1000\b|1,000"),
        ("queries per page are counted", r"(?i)quer"),
    ]),
    (7, ["m07_variations.py"], False, [
        ("options and variants are compared", r"(?i)option.*variant|variant"),
    ]),
    (8, ["m08_pageweight.py", "--compare"], False, [
        ("three pages are weighed", r"(?i)home[\s\S]*category[\s\S]*product"),
        ("CSS and JS dominate", r"(?i)\bcss\b[\s\S]*\bjs\b"),
        ("a total is reported", r"(?i)TOTAL"),
    ]),
    (9, ["m09_basket.py", "--where"], False, [
        ("the cookie does not grow with the basket",
         r"\(nothing\)\s+(\d+)[\s\S]*?x .*?\s+\1\s"),
        ("the session record is the same size throughout",
         r"\(nothing\)\s+\d+\s+(\d+)[\s\S]*?x .*?\s+\d+\s+\1\s"),
        ("the basket is rows in oc_cart", r"(?i)oc_cart|cart rows"),
    ]),
    (10, ["m10_checkout.py", "--repeat", "2"], False, [
        ("every checkout step is timed", r"(?i)fill basket[\s\S]*confirm"),
        ("confirm does the writes", r"(?i)confirm"),
        ("orders are created", r"orders in the table: \d+ -> \d+"),
    ]),
    (12, ["m12_discounts.py", "--sortorder"], False, [
        ("the two totals differ", r"(?i)NO -- [+-]"),
        ("both orders are itemised", r"(?i)Sub-Total[\s\S]*Coupon"),
    ]),
    (14, ["m14_gateway.py", "--latency", "60"], True, [
        ("the tail is far above the median", r"(?i)median[\s\S]*p99"),
        ("some requests never answer", r"(?i)never answered"),
    ]),
    (15, ["m15_accounts.py", "--cost"], False, [
        ("each cost factor roughly doubles the work", r"\b8\b[\s\S]*\b14\b"),
        ("the shipped default is bcrypt", r"(?i)bcrypt"),
    ]),
    (16, ["m16_security.py", "--check"], False, [
        ("the battery reports a score", r"---- \d+ of \d+ ----"),
        ("the error log is checked", r"(?i)Error log"),
    ]),
    (17, ["m17_dashboard.py", "--counts"], False, [
        ("the dashboard's parts disagree",
         r"(?i)Orders widget[\s\S]*The chart[\s\S]*The table itself"),
    ]),
    (18, ["m18_orders.py", "--refund"], False, [
        ("a return writes two rows and nothing else",
         r"(?i)oc_return[\s\S]*oc_return_history"),
    ]),
    (19, ["m19_modules.py", "--measure"], False, [
        ("the cost per module instance is reported", r"(?i)queries each"),
        ("it is measured at 0, 1, 3 and 5", r"\n\s+0\s[\s\S]*\n\s+5\s"),
    ]),
    (20, ["m20_deploy.py", "--backup"], True, [
        ("a restore is timed", r"(?i)database restore"),
        ("the restore is the slower of the two", r"(?i)restore"),
    ]),
    (21, ["m21_seo.py", "--crawl"], False, [
        ("there is no sitemap", r"(?i)sitemap\.xml\s+HTTP 404"),
        ("first byte is reported per page", r"(?i)first byte"),
    ]),
]


def run(argv: list[str], timeout: int = 900) -> tuple[int, str, float]:
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
    head = f"  {'ch':>3}  {'harness':<34}{'secs':>7}  checks"
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
        print(f"  {ch:>3}  {argv[0]:<34}{secs:>7.0f}  "
              f"{''.join(hits)}{flag}")
        if a.verbose or "X" in hits:
            for line in out.strip().splitlines()[-6:]:
                print(f"         {line[:92]}")

    print(f"\n  {passed}/{total} claims still hold")
    if failures:
        print("\n  not reproduced:")
        for f in failures:
            print(f"    {f}")
        print("\n  A failure here means the handout and the software now")
        print("  disagree. Find out which one moved before editing either.")
    return len(failures)


if __name__ == "__main__":
    sys.exit(main())
