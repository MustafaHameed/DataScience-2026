#!/usr/bin/env python3
"""Fail on a page that carries almost nothing.

A page with one review question on it and eight inches of white below is not
an error TeX will report -- the box is not overfull, no reference is missing,
nothing is undefined -- and it survived a page-by-page read of the whole
volume, five times over. It is only obvious when the pages are counted rather
than looked at.

Part openers are meant to be nearly empty; everything else is not.

    python tools/check_pages.py

Exit status is 0 when every page carries either real content or is a part
opener.
"""
from __future__ import annotations

import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PDF = os.path.join(ROOT, "CCNA_v2_Handout.pdf")

# Below this many lines of text a page is not carrying its weight. The running
# head and the folio are two of them, so the real floor is about seven lines.
FLOOR = 9

# A part opener is a page whose whole content is "Part <roman>" and a title.
OPENER = re.compile(r"\bPart\s+(I|II|III|IV|V|VI)\b")


def main() -> int:
    if not os.path.exists(PDF):
        print("no PDF -- run build.ps1 first")
        return 1
    r = subprocess.run(["pdftotext", "-layout", PDF, "-"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print("pdftotext failed:", r.stderr.strip()[:200])
        return 1

    thin = []
    for n, page in enumerate(r.stdout.split("\f"), 1):
        lines = [ln for ln in page.split("\n") if ln.strip()]
        if not lines or len(lines) >= FLOOR:
            continue
        body = " ".join(lines)
        if OPENER.search(body) or "Appendices" in body:
            continue                       # deliberate
        thin.append((n, len(lines), body[:88]))

    if thin:
        print(f"{len(thin)} page(s) with almost nothing on them:")
        for n, k, body in thin:
            print(f"  p{n}: {k} lines -- {body}")
        print("\nUsually a block that should have been kept together. See the"
              "\n`reviewq` environment in ccna.sty for how the last one was"
              "\nfixed: \\needspace, plus enumitem's midpenalty.")
        return 1
    print("No orphaned pages.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
