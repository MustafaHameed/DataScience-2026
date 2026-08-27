#!/usr/bin/env python3
"""Read CCNA_v2_Handout.log and report the seven quality gates.

build.ps1 already does this, but its output is buried in latexmk's own
narration, and a hand-rolled `grep '^!'` misses everything: latexmk is run
with -file-line-error, which reformats errors as `file:line: message` and
leaves no leading `!` to find. That mistake once let a real error sit in a
build reported as clean, so the check lives in a file now rather than in
whatever regex seemed right at the time.

Exit status is 0 only when every gate reads zero.
"""
import re
import sys
from pathlib import Path

LOG = Path(__file__).resolve().parent.parent / "CCNA_v2_Handout.log"

# longtable emits this whenever a table crosses a page break. Upstream
# behaviour, not ours, and harmless.
IGNORE = "Infinite glue shrinkage"

GATES = [
    ("Errors",          re.compile(r"^!|^.*:\d+: ")),
    ("Undefined refs",  re.compile(r"undefined references|Reference .* undefined")),
    ("Overfull hboxes", re.compile(r"Overfull \\hbox")),
    ("Overfull vboxes", re.compile(r"Overfull \\vbox")),
    ("Underfull boxes", re.compile(r"^Underfull")),
    ("Font warnings",   re.compile(r"LaTeX Font Warning")),
    ("PDF-string warns", re.compile(r"Token not allowed")),
]


def main() -> int:
    if not LOG.exists():
        print(f"No log at {LOG}. Run build.ps1 first.")
        return 1

    lines = LOG.read_text(encoding="utf-8", errors="replace").split("\n")

    pages = "(no PDF produced)"
    for line in lines:
        m = re.search(r"Output written on .*\((\d+) pages", line)
        if m:
            pages = m.group(1)

    print(f"{'Pages:':<20}{pages}")

    failed = 0
    for label, pattern in GATES:
        hits = [ln for ln in lines if pattern.search(ln) and IGNORE not in ln]
        mark = "" if not hits else "  <-- "
        print(f"{label + ':':<20}{len(hits)}{mark}")
        if hits:
            failed += len(hits)
            for ln in hits[:6]:
                print(f"    {ln.strip()[:110]}")
            if len(hits) > 6:
                print(f"    ... and {len(hits) - 6} more")

    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
