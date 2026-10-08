#!/usr/bin/env python3
"""Report on the last build, reading only ECAD_Fall2026_Handout.log.

build.ps1 prints the same figures, but it writes them straight to the console
with Out-Host, so they cannot be captured or filtered by a caller.  This reads
the log instead and prints a few lines, which is what you want when the build
runs inside a tool that shows you everything it emits.

    python checklog.py          the report
    python checklog.py -v       the report, then the first errors in full
"""
from __future__ import annotations

import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
LOG = os.path.join(HERE, "ECAD_Fall2026_Handout.log")

# A longtable that breaks across a page reports infinite glue shrinkage once per
# break.  It is upstream behaviour, not a fault in this document; build.ps1
# filters it and so does this.
BENIGN = re.compile(r"Infinite glue shrinkage")


def main() -> int:
    if not os.path.exists(LOG):
        print("no log -- run .\\build.ps1 first")
        return 1
    txt = io.open(LOG, encoding="utf-8", errors="replace").read()
    lines = txt.splitlines()

    errors = [l for l in lines
              if (l.startswith("!") or re.match(r"^.*:\d+: ", l))
              and not BENIGN.search(l)]
    undef = [l for l in lines
             if re.search(r"undefined references|Reference .* undefined", l)]
    ovh = [l for l in lines if "Overfull \\hbox" in l]
    ovv = [l for l in lines if "Overfull \\vbox" in l]
    under = [l for l in lines if l.startswith("Underfull")]
    fonts = [l for l in lines if "LaTeX Font Warning" in l]
    pdfstr = [l for l in lines if "Token not allowed" in l]
    m = re.search(r"Output written on .*?\((\d+) pages", txt)
    pages = m.group(1) if m else "?"

    rows = [("Pages", pages), ("Errors", len(errors)),
            ("Undefined refs", len(undef)), ("Overfull hboxes", len(ovh)),
            ("Overfull vboxes", len(ovv)), ("Underfull boxes", len(under)),
            ("Font warnings", len(fonts)), ("PDF-string warns", len(pdfstr))]
    print("--- Build report ---")
    for k, v in rows:
        print(f"{k:<22}{v}")

    bad = errors + undef + ovh + ovv + under + fonts + pdfstr
    if bad:
        print()
        for l in bad[:25]:
            print("  " + l[:150])
        if len(bad) > 25:
            print(f"  ... and {len(bad) - 25} more")
    if "-v" in sys.argv and errors:
        i = lines.index(errors[0])
        print("\n--- first error in context ---")
        print("\n".join(lines[i:i + 14]))
    return 1 if (errors or undef) else 0


if __name__ == "__main__":
    sys.exit(main())
