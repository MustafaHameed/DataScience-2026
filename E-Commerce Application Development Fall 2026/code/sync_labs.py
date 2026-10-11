#!/usr/bin/env python3
"""Check that what the handout tells a student to run actually exists.

The sibling courses print whole programs in their lab boxes, so their
sync_labs.py extracts a listing and diffs it against a source file. This
course's labs are not programs. They are COMMANDS -- a sequence of things to
type at a shell, against a running container -- so the contract has to be a
different one, and it is a more useful one for this subject:

  1. Every `python mNN_xxx.py` a lab names must EXIST in code/, and every
     --flag it is given must be one that file actually accepts.
  2. Every `docker compose <service>` must name a service that is in
     docker-compose.yml.
  3. Every PHP line printed in the handout must appear in the module source
     under modules/, compared as tokens.
  4. Every file path the handout tells a student to look at must exist.

What this is for: a handout rots in exactly one way. Somebody renames a flag,
or splits a harness, and twenty lab steps across fourteen chapters quietly
start referring to something that is not there. A student types what is on the
page and gets an error that is not their fault. Nothing else in the build
catches that, because LaTeX is perfectly happy to typeset a command that does
not exist.

    python sync_labs.py            check everything
    python sync_labs.py --ch 16    one chapter
    python sync_labs.py --verbose  list what was checked, not only failures
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PARTS = os.path.join(ROOT, "parts")

LST = re.compile(r"\\begin\{lstlisting\}(\[[^\]]*\])?(.*?)\\end\{lstlisting\}",
                 re.S)
LANG = re.compile(r"language=([A-Za-z+]+)")
# "python m09_basket.py --where" anywhere in a listing
PYRUN = re.compile(r"python\s+(m\d\d_[a-z_]+\.py|[a-z_]+/[a-z_]+\.py|"
                   r"[a-z_]+\.py)((?:\s+--?[a-z-]+(?:\s+[^\s\\|;]+)?)*)")
FLAG = re.compile(r"--([a-z][a-z-]*)")
# Spaces-but-not-newlines. Plain \s crosses a line break, so
# "docker compose up -d" followed by "python waitfor.py" on the NEXT line
# reported a service called "python". A checker that invents findings is a
# checker somebody switches off.
SP = r"[ \t]"
COMPOSE = re.compile(rf"docker{SP}+compose{SP}+"
                     rf"(?:exec|up|logs|restart|cp)"
                     rf"(?:{SP}+-\w+)*{SP}+([a-z]+)")
# Paths the handout points at, inside \texttt{...} in a listing or in prose
PATHISH = re.compile(r"(?:^|\s)((?:code/|modules/|shop/|seed/|gateway/)"
                     r"[A-Za-z0-9_./-]+)")


def read(path: str) -> str:
    with open(path, encoding="utf-8") as fh:
        return fh.read()


def chapter_files(only: int | None) -> list[tuple[str, str]]:
    out = []
    for name in sorted(os.listdir(PARTS)):
        m = re.fullmatch(r"(\d\d)_([a-z]+)\.tex", name)
        if not m:
            continue
        n = int(m.group(1))
        if only and n != only:
            continue
        out.append((f"Ch {n:>2}", os.path.join(PARTS, name)))
    return out


def listings(text: str) -> list[tuple[str, str]]:
    return [(LANG.search(opt or "").group(1) if LANG.search(opt or "") else "",
             body) for opt, body in LST.findall(text)]


def script_flags(path: str) -> set[str] | None:
    """The --flags a harness accepts, read from its source rather than run."""
    if not os.path.exists(path):
        return None
    src = read(path)
    flags = set(re.findall(r'add_argument\(\s*[`"\']--([a-z][a-z-]*)', src))
    flags |= set(re.findall(r'f"--\{(\w+)\}"', src))
    # Not every harness uses argparse. m05_coldstart.py reads sys.argv
    # directly, and a checker that only knows argparse reports its real
    # flags as nonexistent.
    flags |= set(re.findall(r'"--([a-z][a-z-]*)" in sys\.argv', src))
    # the loop form: for flag in ("where", "sessions", ...)
    for group in re.findall(r'for flag in \(([^)]*)\)', src):
        flags |= set(re.findall(r'"([a-z][a-z-]*)"', group))
    return flags


def compose_services() -> set[str]:
    path = os.path.join(HERE, "docker-compose.yml")
    if not os.path.exists(path):
        return set()
    out, in_services = set(), False
    for line in read(path).splitlines():
        if re.match(r"^services:", line):
            in_services = True
            continue
        if in_services and re.match(r"^[a-z]", line):
            break
        m = re.match(r"^  ([a-z][a-z0-9_-]*):\s*$", line)
        if in_services and m:
            out.add(m.group(1))
    return out


def php_lines(body: str) -> list[str]:
    keep = []
    for line in body.splitlines():
        s = line.strip()
        if not s or s.startswith(("//", "/*", "*", "#")):
            continue
        keep.append(re.sub(r"\s+", "", s))
    return keep


def module_source() -> str:
    """Every PHP file under modules/, whitespace removed, as one string."""
    blob = []
    base = os.path.join(HERE, "modules")
    for root, _dirs, files in os.walk(base):
        for f in files:
            if f.endswith(".php"):
                blob.append(re.sub(r"\s+", "", read(os.path.join(root, f))))
    return "".join(blob)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", type=int)
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    services = compose_services()
    php_blob = module_source()
    problems, checked = [], 0

    for label, path in chapter_files(a.ch):
        text = read(path)
        for lang, body in listings(text):
            # 1 and 2: the commands
            for script, tail in PYRUN.findall(body):
                checked += 1
                full = os.path.join(HERE, script)
                if not os.path.exists(full):
                    problems.append(f"{label}: no such file  code/{script}")
                    continue
                accepted = script_flags(full)
                for flag in FLAG.findall(tail):
                    if accepted is not None and flag not in accepted:
                        problems.append(
                            f"{label}: {script} does not accept --{flag}")
                if a.verbose:
                    print(f"  ok  {label}  python {script}{tail}")

            for svc in COMPOSE.findall(body):
                checked += 1
                if svc not in services and svc not in {"-d", "up", "down"}:
                    problems.append(
                        f"{label}: docker compose names service '{svc}', "
                        f"which is not in docker-compose.yml")

            # 3: PHP shown in the handout must be in a module source
            if lang == "php" and php_blob:
                for line in php_lines(body):
                    checked += 1
                    probe = line.rstrip(";")
                    if len(probe) < 12 or "..." in probe:
                        continue
                    if probe not in php_blob:
                        problems.append(
                            f"{label}: PHP line not found in modules/: "
                            f"{line[:62]}")

            # 4: paths the lab tells you to look at
            for p in PATHISH.findall(body):
                p = p.rstrip(".,;'\"")
                if "*" in p or p.endswith("/"):
                    continue
                checked += 1
                if not os.path.exists(os.path.join(HERE, p[len("code/"):]
                                                   if p.startswith("code/")
                                                   else p)):
                    problems.append(f"{label}: path does not exist: {p}")

    print(f"\n  {checked} things checked in the handout's laboratories")
    if problems:
        print(f"  {len(problems)} problem(s):\n")
        for p in problems:
            print(f"    {p}")
        return 1
    print("  every command, service, path and PHP line in the labs exists")
    return 0


if __name__ == "__main__":
    sys.exit(main())
