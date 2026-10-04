#!/usr/bin/env python3
"""Check that the code shown in the handout really is the code in code/.

The sibling courses print whole programs in their lab boxes, so their
sync_labs.py can extract the listing and diff it against the file. This course
cannot: its programs are 150-300 lines and the lab boxes show an ABRIDGED
excerpt -- the parts worth reading, with the printf boilerplate elided.

So the contract here is weaker and still worth enforcing:

    every non-comment, non-blank line of every handout listing must appear in
    that chapter's .cpp file, compared as TOKENS rather than characters.

Comments are exempt because they are reflowed to the handout's narrower
measure. Code is compared with whitespace removed and the "std::" prefix made
optional, for the same reason: the handout's listings are set in a 76-column
measure, so a multi-line lambda is often compressed onto one line and "i - 1"
written "i-1". Those are typographical, and a checker that rejected them would
be rejecting the house style rather than finding defects.

What this still catches, and what it is for: a line of C++ printed in the
handout that is not in the program at all. A student types what is on the
page, so an invented line, a renamed variable or a flipped operator is a real
defect whatever the spacing.

    python sync_labs.py            check every chapter
    python sync_labs.py --ch 9     just that one
    python sync_labs.py --verbose  list what was checked, not only failures
"""
from __future__ import annotations

import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PARTS = os.path.join(os.path.dirname(HERE), "parts")

# A line that is only a comment, or only the remains of one. Also the
# "/* ... */" placeholders the excerpts use where a block was cut out.
COMMENT = re.compile(r"^\s*(//|/\*|\*/|\*)")


def norm(s: str) -> str:
    """Code reduced to tokens: no whitespace, no optional std:: prefix.

    This is what makes the check tolerant of the handout's narrower measure
    and intolerant of everything else.
    """
    return re.sub(r"\s+", "", s).replace("std::", "")


def listings(tex: str) -> list[list[str]]:
    """Every lstlisting body in a tex file, as lists of lines."""
    out = []
    for m in re.finditer(r"\\begin\{lstlisting\}(?:\[[^\]]*\])?\n(.*?)"
                         r"\\end\{lstlisting\}", tex, re.S):
        out.append(m.group(1).split("\n"))
    return out


def code_lines(lines: list[str]) -> list[str]:
    """Drop blanks, comment-only lines, and lines inside a /* */ block."""
    out, in_block = [], False
    for ln in lines:
        s = ln.strip()
        if in_block:
            if "*/" in s:
                in_block = False
            continue
        if s.startswith("/*") and "*/" not in s:
            in_block = True
            continue
        if not s or COMMENT.match(ln):
            continue
        # strip a trailing // comment, which is reflowed like any other
        s = re.sub(r"\s*//.*$", "", s).strip()
        if s:
            out.append(s)
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", type=int, nargs="*")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    chapters = []
    for f in sorted(os.listdir(PARTS)):
        m = re.match(r"(\d\d)_.*\.tex$", f)
        if m:
            chapters.append((int(m.group(1)), f))
    if a.ch:
        chapters = [(n, f) for n, f in chapters if n in a.ch]

    problems, checked_lines, checked_ch = [], 0, 0
    for n, f in chapters:
        tex = open(os.path.join(PARTS, f), encoding="utf-8").read()
        ls = listings(tex)
        if not ls:
            continue
        src = [g for g in sorted(os.listdir(HERE))
               if g.startswith(f"ch{n:02d}_") and g.endswith(".cpp")]
        if not src:
            problems.append(f"ch {n:2d}: {len(ls)} listing(s) but no .cpp")
            continue
        cpp = open(os.path.join(HERE, src[0]), encoding="utf-8").read()
        # Strip the source's own comments before normalising, so a handout
        # line can never be "found" inside a comment that quotes it.
        bare = re.sub(r"//[^\n]*", "", re.sub(r"/\*.*?\*/", "", cpp, flags=re.S))
        hay = norm(bare)

        missing = []
        for body in ls:
            for ln in code_lines(body):
                checked_lines += 1
                needle = norm(ln)
                if len(needle) < 6:          # too short to match meaningfully
                    continue
                if needle in hay:
                    continue
                missing.append(ln)
        checked_ch += 1
        if missing:
            problems.append(f"ch {n:2d} ({src[0]}): {len(missing)} line(s) "
                            f"not found in the source")
            for ln in missing[:6]:
                problems.append(f"        {ln[:72]}")
            if len(missing) > 6:
                problems.append(f"        ... and {len(missing)-6} more")
        elif a.verbose:
            print(f"  ch {n:2d}  {src[0]:24s} "
                  f"{len(code_lines(sum(ls, []))):3d} code line(s)  ok")

    print(f"\n  {checked_ch} chapter(s), {checked_lines} code line(s) checked")
    if problems:
        print()
        for p in problems:
            print("  " + p)
        print(f"\n  {len([p for p in problems if p.startswith('ch')])} "
              f"problem(s)")
        return 1
    print("  every line of C++ in the handout is in the source")
    return 0


if __name__ == "__main__":
    sys.exit(main())
