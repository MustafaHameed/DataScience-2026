#!/usr/bin/env python3
"""Prove every chapter carries the pedagogy blocks the book claims it does.

The front matter tells the reader that a Break & Fix is "practised in every
chapter rather than revised once at the end". That was untrue for chapter 1
for most of this book's life, and nothing caught it -- a missing environment
is not a LaTeX error, it is just an absence.

Core furniture is required in every chapter. config, verify, hostcmd,
playbook and lab are NOT required: chapters covering describe-only topics
legitimately have no configuration, and inventing some would teach the wrong
exam. Those are reported as a census, not as failures.

Exit status is 0 when every chapter has all the core blocks.
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PARTS = os.path.join(ROOT, "parts")
MASTER = os.path.join(ROOT, "CCNA_v2_Handout.tex")

REQUIRED = ["outcomes", "prereq", "keyterms", "pitfall", "breakfix",
            "checkpoint", "chaptersummary", "reviewq"]
OPTIONAL = ["examnote", "config", "verify", "hostcmd", "playbook", "lab"]


def chapters():
    text = io.open(MASTER, encoding="utf-8").read()
    n = 0
    for stem in re.findall(r"\\input\{parts/([^}]+)\}", text):
        path = os.path.join(PARTS, stem + ".tex")
        if not os.path.exists(path):
            continue
        body = io.open(path, encoding="utf-8").read()
        if not re.search(r"^\\chapter\{", body, re.M):
            continue
        n += 1
        if n > 34:            # appendices carry different furniture
            return
        title = re.search(r"\\chapter\{(.+?)\}", body).group(1)
        yield n, title, body


def main():
    missing = []
    totals = {b: 0 for b in REQUIRED + OPTIONAL}

    for n, title, body in chapters():
        for b in REQUIRED + OPTIONAL:
            c = len(re.findall(r"\\begin\{" + b + r"\}", body))
            totals[b] += c
            if b in REQUIRED and c == 0:
                missing.append((n, title, b))

    print("Block census across 34 chapters")
    for b in REQUIRED:
        print("  %-16s %3d   (required)" % (b, totals[b]))
    for b in OPTIONAL:
        print("  %-16s %3d" % (b, totals[b]))
    print()

    if missing:
        for n, title, b in missing:
            print("ch %-2d %-42s has no %s" % (n, title[:42], b))
        print("\n%d chapter(s) missing required furniture." % len(missing))
        return 1

    print("Every chapter carries all required furniture.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
