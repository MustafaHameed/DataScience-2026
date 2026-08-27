#!/usr/bin/env python3
"""Verify hand-written 'Section N.M' cross-references point where they claim.

\\ref{} is checked by LaTeX; a section number typed as prose is not. Insert a
section into a chapter and every literal 'Section 29.4' after it silently
becomes wrong, with no warning from anything.

This walks each chapter, numbers its \\section commands in order, and reports
any literal reference whose target does not exist -- printing the section
that number actually names so the fix is obvious.

Exit status is 0 when every reference resolves.
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PARTS = os.path.join(ROOT, "parts")
MASTER = os.path.join(ROOT, "CCNA_v2_Handout.tex")

REF = re.compile(r"Section~?([0-9A-G]+)\.(\d+)")


def numbered_files():
    """(label, path) for each chapter: '1'..'34' then 'A'..'G'."""
    text = io.open(MASTER, encoding="utf-8").read()
    out, n, appendix = [], 0, False
    for m in re.finditer(r"\\appendix|\\input\{parts/([^}]+)\}", text):
        if m.group(0) == r"\appendix":
            appendix = True
            continue
        path = os.path.join(PARTS, m.group(1) + ".tex")
        if not os.path.exists(path):
            continue
        body = io.open(path, encoding="utf-8").read()
        if not re.search(r"^\\chapter\{", body, re.M):
            continue
        if appendix:
            label = os.path.basename(path)[0]
        else:
            n += 1
            label = str(n)
        out.append((label, path, body))
    return out


def main():
    problems = 0
    for label, path, body in numbered_files():
        # Section numbering restarts at each chapter; starred sections do not
        # take a number.
        sections = re.findall(r"^\\section\{(.+?)\}", body, re.M)
        name = os.path.basename(path)

        for m in REF.finditer(body):
            chap, num = m.group(1), int(m.group(2))
            if chap != label:
                continue          # a reference to another chapter; skip
            if num < 1 or num > len(sections):
                problems += 1
                print("%s: 'Section %s.%d' -- chapter has only %d sections"
                      % (name, chap, num, len(sections)))
                for i, s in enumerate(sections, 1):
                    print("      %s.%d  %s" % (chap, i, s))
            else:
                target = sections[num - 1]
                print("ok  %-26s Section %s.%d -> %s"
                      % (name, chap, num, target))

    print()
    if problems:
        print("%d bad section reference(s)." % problems)
        return 1
    print("All literal section references resolve.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
