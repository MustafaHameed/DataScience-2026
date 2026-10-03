#!/usr/bin/env python3
"""Extract every lab listing from parts/ into runnable files here.

The handout is the single source: a lab is written once, inside its
``lab`` box, and this script writes it out. Editing a file in this
directory by hand is therefore pointless --- the next sync overwrites it.
Fix the chapter instead.

Each lab's first line names its own file, which is how the mapping is
derived:

    # ch06_pressure.sh -- ready time and the reclamation ladder, observed.
    /* ch03_popf.c -- show that POPF fails silently in user mode.

    python sync_labs.py            write the files
    python sync_labs.py --check    fail if anything is out of date
"""
from __future__ import annotations

import glob
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PARTS = os.path.join(os.path.dirname(HERE), "parts")

LAB = re.compile(r"\\begin\{lab\}.*?\\end\{lab\}", re.S)
LISTING = re.compile(
    r"\\begin\{lstlisting\}(?:\[[^\]]*\])?\n(.*?)\\end\{lstlisting\}", re.S)
# the first line of every lab names the file it becomes
NAMED = re.compile(r"^\s*(?:#|/\*)\s*(\w+\.\w+)\s*--")

SHEBANG = {".sh": "#!/usr/bin/env bash\n"}


def labs() -> dict[str, str]:
    """Map filename -> body, for every lab listing in parts/."""
    out: dict[str, str] = {}
    for path in sorted(glob.glob(os.path.join(PARTS, "*.tex"))):
        src = io.open(path, encoding="utf-8").read()
        for lab in LAB.findall(src):
            for body in LISTING.findall(lab):
                lines = body.splitlines()
                name = None
                for line in lines[:3]:
                    m = NAMED.match(line)
                    if m:
                        name = m.group(1)
                        break
                if not name:
                    print(f"  no filename in a lab of "
                          f"{os.path.basename(path)}", file=sys.stderr)
                    continue
                if name in out:
                    print(f"  duplicate filename {name}", file=sys.stderr)
                out[name] = body.rstrip() + "\n"
    return out


def main() -> int:
    check = "--check" in sys.argv
    found = labs()
    stale, written = [], 0
    for name, body in found.items():
        # The shebang is in the handout for .sh labs already; do not double it.
        if name.endswith(".sh") and not body.startswith("#!"):
            body = SHEBANG[".sh"] + body
        path = os.path.join(HERE, name)
        old = io.open(path, encoding="utf-8").read() if os.path.exists(path) \
            else None
        if old == body:
            continue
        if check:
            stale.append(name)
            continue
        io.open(path, "w", encoding="utf-8", newline="\n").write(body)
        if name.endswith(".sh"):
            os.chmod(path, 0o755)
        written += 1
    if check:
        if stale:
            print("  out of date: " + ", ".join(sorted(stale)))
            print(f"\n  {len(stale)} file(s) differ from the handout")
            return 1
        print(f"  {len(found)} lab file(s) match the handout")
        return 0
    print(f"  {len(found)} lab(s) found, {written} written")
    return 0


if __name__ == "__main__":
    sys.exit(main())
