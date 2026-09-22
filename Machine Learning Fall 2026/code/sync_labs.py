#!/usr/bin/env python3
"""Extract every chapter's lab listing into code/, and optionally run them.

The handout is the single source of truth: each chapter's "Try It Yourself"
box holds one self-contained Python listing, and this script copies it,
unchanged, into ``code/chNN_<slug>.py``. Editing a file here by hand is
pointless -- the next sync overwrites it. Edit the chapter instead.

    python sync_labs.py            extract all labs
    python sync_labs.py --run      extract, then run each one and report
    python sync_labs.py --run 9    extract all, run chapter 9 only

Checks made on every extract:

* the listing is pure ASCII (listings under pdfLaTeX cannot typeset UTF-8);
* no line is longer than 78 characters (wider lines wrap in the handout).
"""
from __future__ import annotations

import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PARTS = os.path.join(ROOT, "parts")
MAX_COLS = 78


STOP = {"a", "an", "the", "and", "then", "by", "of", "in", "to", "for",
        "with", "from", "at", "on"}


def slug(title: str, limit: int = 22) -> str:
    """A short file name from the lab title: its first content words, at most
    *limit* characters, so the name fits on the lab slide's 'runnable copy'
    card (the slide pipeline shows code/<file>)."""
    t = re.sub(r"\\[a-zA-Z]+\*?", "", title).lower()
    words = [w for w in re.split(r"[^a-z0-9]+", t) if w and w not in STOP]
    out = ""
    for w in words:
        cand = f"{out}_{w}" if out else w
        if len(cand) > limit:
            break
        out = cand
    return out or words[0][:limit]


def labs() -> list[tuple[int, str, str]]:
    """[(chapter, slug, code)] for every chapter with a lab listing."""
    out = []
    for f in sorted(os.listdir(PARTS)):
        m = re.match(r"(\d\d)_.*\.tex$", f)
        if not m or f.startswith("00_"):
            continue
        src = open(os.path.join(PARTS, f), encoding="utf-8").read()
        lab = re.search(r"\\begin\{lab\}\[([^\]]*)\](.*?)\\end\{lab\}", src,
                        re.S)
        if not lab:
            continue
        code = re.search(r"\\begin\{lstlisting\}(?:\[[^\]]*\])?\n(.*?)"
                         r"\\end\{lstlisting\}", lab.group(2), re.S)
        if code:
            out.append((int(m.group(1)), slug(lab.group(1)), code.group(1)))
    return out


def check(n: int, code: str) -> list[str]:
    problems = []
    for i, line in enumerate(code.splitlines(), 1):
        bad = [c for c in line if ord(c) > 127]
        if bad:
            problems.append(f"ch{n:02d} line {i}: non-ASCII {''.join(bad)!r}")
        if len(line) > MAX_COLS:
            problems.append(f"ch{n:02d} line {i}: {len(line)} columns")
    return problems


def main() -> int:
    run = "--run" in sys.argv
    only = [int(a) for a in sys.argv[1:] if a.isdigit()]
    found = labs()
    problems = []
    written = []
    for n, s, code in found:
        # drop stale extracts of the same chapter under an older slug
        for f in os.listdir(HERE):
            if f.startswith(f"ch{n:02d}_") and f.endswith(".py") \
                    and f != f"ch{n:02d}_{s}.py":
                os.remove(os.path.join(HERE, f))
        path = os.path.join(HERE, f"ch{n:02d}_{s}.py")
        header = (f"# Chapter {n} lab -- extracted from parts/ by "
                  f"sync_labs.py. Edit the chapter, not this file.\n")
        with open(path, "w", encoding="ascii", errors="replace",
                  newline="\n") as fh:
            fh.write(header + code)
        written.append((n, path))
        problems += check(n, code)
    print(f"extracted {len(written)} lab(s)")
    for p in problems:
        print("  WARNING", p)

    failed = 0
    if run:
        for n, path in written:
            if only and n not in only:
                continue
            t0 = time.time()
            env = dict(os.environ, MPLBACKEND="Agg", PYTHONWARNINGS="ignore")
            r = subprocess.run([sys.executable, path], capture_output=True,
                               text=True, cwd=HERE, env=env, timeout=600)
            ok = r.returncode == 0
            failed += not ok
            print(f"\n=== ch{n:02d} {'ok' if ok else 'FAILED'} "
                  f"({time.time() - t0:.1f}s) {os.path.basename(path)}")
            print(r.stdout.rstrip())
            if not ok:
                print(r.stderr.rstrip()[-2000:])
    return 1 if failed or problems else 0


if __name__ == "__main__":
    sys.exit(main())
