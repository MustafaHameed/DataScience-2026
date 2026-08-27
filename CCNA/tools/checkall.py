#!/usr/bin/env python3
"""Run every gate and report one verdict.

    python tools/checkall.py

Assumes the PDF has already been built -- logcheck reads the log, it does not
compile. Run build.ps1 first, or use it as:

    .\\build.ps1; python tools\\checkall.py

Each gate exists because the class of error it catches actually happened and
was not caught by anything else. Exit status is 0 only when all pass.
"""
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

GATES = [
    ("lintsource.py",       "source traps LaTeX reports badly"),
    ("logcheck.py",         "the seven build-log gates"),
    ("check_coverage.py",   "all 29 blueprint topics covered"),
    ("check_furniture.py",  "every chapter has its required blocks"),
    ("check_answers.py",    "every checkpoint question is answered"),
    ("check_sectionrefs.py", "literal section references resolve"),
]


def main():
    width = max(len(g) for g, _ in GATES)
    failed = []

    for script, what in GATES:
        r = subprocess.run([sys.executable, os.path.join(HERE, script)],
                           capture_output=True, text=True)
        ok = r.returncode == 0
        print("%-*s  %-46s %s" % (width, script, what,
                                  "PASS" if ok else "FAIL"))
        if not ok:
            failed.append((script, r.stdout.strip() or r.stderr.strip()))

    print()
    if failed:
        for script, out in failed:
            print("--- %s ---" % script)
            print(out)
            print()
        print("%d gate(s) failed." % len(failed))
        return 1

    print("All %d gates pass." % len(GATES))
    return 0


if __name__ == "__main__":
    sys.exit(main())
