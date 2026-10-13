#!/usr/bin/env python3
"""Check every lab listing in parts/ against the two rules that bite.

Both faults are silent until the build, and one of them is silent even then:

* **Non-ASCII.** listings has no Unicode support under pdfLaTeX, so a typographic
  dash or a non-breaking space pasted from a terminal is a compile error well
  away from the character that caused it.
* **Over-wide lines.** A listing line past 76 columns overflows the box, which
  the build reports only as an anonymous "Overfull \\hbox in paragraph", and on
  the slides it pushes the code panel below the 10.5 pt floor.

It also reports the step markers, since the slide builder cuts a lab into one
slide per ``# --- n.`` comment and a lab with none becomes a single wall of code.

    python checklabs.py           every chapter
    python checklabs.py 02 03     only these
"""
from __future__ import annotations

import glob
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PARTS = os.path.join(HERE, "parts")
MAXCOL = 76

LISTING = re.compile(r"\\begin\{lstlisting\}(\[[^\]]*\])?\n(.*?)\\end\{lstlisting\}",
                     re.S)
# A step marker opens with whatever starts a comment in the lab's language:
# # for shell, YAML, Dockerfile and Terraform; /* or // for C.  The slide
# builder splits a lab at these, so it must recognise the same set.
STEP = re.compile(r"^\s*(?:#|//|/\*|--)\s*---\s*(\d+)\.")


LAB = re.compile(r"\\begin\{lab\}.*?\\end\{lab\}", re.S)


def check(path: str) -> list[str]:
    name = os.path.basename(path)
    src = io.open(path, encoding="utf-8").read()
    probs: list[str] = []
    listings = LISTING.findall(src)
    if not listings:
        return probs
    # The step-marker rule applies only to listings inside a lab box: the
    # slide builder splits those into one slide per step.  A reference
    # listing in an appendix is a lookup table and has no steps to split.
    in_labs = set()
    for m in LAB.finditer(src):
        for _, body in LISTING.findall(m.group(0)):
            in_labs.add(body)
    for opts, body in listings:
        is_lab = body in in_labs
        lines = body.splitlines()
        steps = [m.group(1) for m in (STEP.match(x) for x in lines) if m]
        for i, line in enumerate(lines, 1):
            if len(line) > MAXCOL:
                probs.append(f"{name} listing line {i}: {len(line)} columns "
                             f"-- {line.strip()[:58]!r}")
            bad = [c for c in line if ord(c) > 126]
            if bad:
                probs.append(f"{name} listing line {i}: non-ASCII "
                             f"{''.join(sorted(set(bad)))!r} -- "
                             f"{line.strip()[:48]!r}")
        if not steps:
            if is_lab:
                probs.append(f"{name}: a lab listing has no '# --- n.' step "
                             f"markers; the slide builder cannot split it")
        else:
            want = [str(n) for n in range(int(steps[0]),
                                          int(steps[0]) + len(steps))]
            if steps != want:
                probs.append(f"{name}: step markers out of order: "
                             f"{', '.join(steps)}")
        if opts and "language=" not in opts:
            probs.append(f"{name}: a listing has no language= option")
    return probs


def main() -> int:
    want = sys.argv[1:]
    paths = sorted(glob.glob(os.path.join(PARTS, "*.tex")))
    if want:
        paths = [p for p in paths
                 if any(os.path.basename(p).startswith(w) for w in want)]
    probs: list[str] = []
    for p in paths:
        probs += check(p)
    if probs:
        for q in probs:
            print("  " + q)
        print(f"\n  {len(probs)} problem(s)")
        return 1
    print("  lab listings clean")
    return 0


if __name__ == "__main__":
    sys.exit(main())
