#!/usr/bin/env python3
"""Check that what the handout tells a student to type would actually work.

A handout rots in one way above all: somebody renames a flag, adds a column,
or moves a file, and lab steps across several chapters quietly start pointing
at something that is not there. A student types what is on the page and gets
an error that is not their fault. LaTeX is perfectly happy to typeset a command
that cannot run, so nothing else in the build catches it.

This course's labs are SQL and short commands, so the contracts are:

  1. Every `python <path>.py` a lab names exists, and every --flag given to it
     is one that file accepts.
  2. Every `.read <path>` names a file that exists.
  3. Every PRAGMA a lab uses is one this SQLite knows.
  4. Every `INSERT INTO t VALUES (...)` supplies as many values as t has
     columns -- where t is whatever the chapter's own lab created, else the
     table in code/schema/v1_create.sql, with ALTER TABLE ... ADD COLUMN
     applied in order. (A NOT NULL column added with no DEFAULT is refused by
     SQLite on a table with rows, and Chapter 11 does that on purpose, so it
     is not counted.)
  5. Every code/ path mentioned anywhere in the handout exists, except the
     files a lab tells the student to write.

Contract 4 was added because it found a real fault on its first run: Chapter
9's laboratory inserted four values into a student table that has eight.

    python sync_labs.py            check everything
    python sync_labs.py --ch 9     one chapter
    python sync_labs.py --verbose  list what was checked, not only failures
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sqlite3
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PARTS = os.path.join(ROOT, "parts")
V1 = os.path.join(HERE, "schema", "v1_create.sql")

LAB = re.compile(r"\\begin\{lab\}(.*?)\\end\{lab\}", re.S)
LST = re.compile(r"\\begin\{lstlisting\}(?:\[[^\]]*\])?(.*?)"
                 r"\\end\{lstlisting\}", re.S)
PYRUN = re.compile(r"python\s+((?:code/)?[\w/]+\.py)((?:[ \t]+--?[a-z][\w-]*"
                   r"(?:[ \t]+[^\s\\|;#-][^\s\\|;#]*)?)*)")
FLAG = re.compile(r"--([a-z][a-z-]*)")
READ = re.compile(r"\.read\s+([\w./-]+)")
PRAGMA = re.compile(r"(?i)\bPRAGMA\s+([a-z_]+)")
CODEPATH = re.compile(r"code/[\w/]+\.(?:py|sql)")
# files the student is told to write, so they cannot exist in the repo
STUDENT_WRITES = {"code/schema/v2_rollback.sql", "code/schema/v3_final.sql",
                  "code/schema/v3_normalised.sql"}


def strip_comments(sql: str) -> str:
    return re.sub(r"--[^\n]*", "", sql)


def split_top(s: str) -> list[str]:
    """Split on commas at bracket depth 0, outside quotes."""
    out, cur, depth, q = [], [], 0, None
    for ch in s:
        if q:
            cur.append(ch)
            if ch == q:
                q = None
            continue
        if ch in "'\"":
            q = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            out.append("".join(cur))
            cur = []
            continue
        cur.append(ch)
    if "".join(cur).strip():
        out.append("".join(cur))
    return out


def group(s: str, i: int) -> tuple[str, int]:
    """The text inside the bracket opening at s[i], and the index after it."""
    depth, q, j = 0, None, i
    while j < len(s):
        ch = s[j]
        if q:
            if ch == q:
                q = None
        elif ch in "'\"":
            q = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return s[i + 1:], len(s)


CONSTRAINT = re.compile(r"(?i)^\s*(PRIMARY\s+KEY|FOREIGN\s+KEY|UNIQUE|CHECK|"
                        r"CONSTRAINT)\b")


def column_count(body: str) -> int:
    return sum(1 for part in split_top(body) if part.strip()
               and not CONSTRAINT.match(part))


CREATE = re.compile(r"(?i)CREATE\s+(?:TEMP\s+)?TABLE\s+(\w+)\s*\(")
ALTER = re.compile(r"(?i)ALTER\s+TABLE\s+(\w+)\s+ADD\s+COLUMN\s+([^;]*);")
INSERT = re.compile(r"(?i)INSERT\s+INTO\s+(\w+)\s+VALUES\s*\(")


def schema_from(sql: str, known: dict[str, int]) -> None:
    for m in CREATE.finditer(sql):
        body, _ = group(sql, m.end() - 1)
        known[m.group(1).lower()] = column_count(body)


def check_inserts(sql: str, known: dict[str, int], where: str,
                  errs: list[str], seen: list[str]) -> None:
    """Walk statements in order, so CREATE and ALTER apply before INSERT."""
    events = []
    for m in CREATE.finditer(sql):
        events.append((m.start(), "create", m))
    for m in ALTER.finditer(sql):
        events.append((m.start(), "alter", m))
    for m in INSERT.finditer(sql):
        events.append((m.start(), "insert", m))
    for _pos, kind, m in sorted(events, key=lambda e: e[0]):
        t = m.group(1).lower()
        if kind == "create":
            body, _ = group(sql, m.end() - 1)
            known[t] = column_count(body)
        elif kind == "alter":
            spec = m.group(2)
            refused = (re.search(r"(?i)NOT\s+NULL", spec)
                       and not re.search(r"(?i)DEFAULT", spec))
            if not refused and t in known:
                known[t] += 1
        else:
            vals, _ = group(sql, m.end() - 1)
            n = len(split_top(vals))
            seen.append(f"{where}: INSERT INTO {t} ({n} values)")
            if t in known and known[t] != n:
                errs.append(f"{where}: INSERT INTO {t} gives {n} values; "
                            f"{t} has {known[t]} columns")


def accepted_flags(path: str) -> set[str]:
    src = io.open(path, encoding="utf-8").read()
    flags = set(re.findall(r"add_argument\(\s*[\"']--([a-z][a-z-]*)", src))
    flags |= set(re.findall(r"[\"']--([a-z][a-z-]*)[\"']\s+in\s+sys\.argv",
                            src))
    return flags


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", type=int, nargs="*")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    base: dict[str, int] = {}
    schema_from(strip_comments(io.open(V1, encoding="utf-8").read()), base)
    pragmas = {r[0] for r in
               sqlite3.connect(":memory:").execute("PRAGMA pragma_list")}

    errs: list[str] = []
    seen: list[str] = []
    for name in sorted(os.listdir(PARTS)):
        m = re.match(r"(\d\d)_", name)
        if not name.endswith(".tex"):
            continue
        src = io.open(os.path.join(PARTS, name), encoding="utf-8").read()
        ch = int(m.group(1)) if m else None

        # contract 5: every code/ path, anywhere in the file
        for p in sorted(set(CODEPATH.findall(src.replace("\\_", "_")))):
            if p in STUDENT_WRITES:
                continue
            seen.append(f"{name}: path {p}")
            if not os.path.exists(os.path.join(ROOT, p)):
                errs.append(f"{name}: names {p}, which does not exist")

        if a.ch and ch not in a.ch:
            continue
        if not m or name.startswith("00_"):
            continue
        known = dict(base)
        for lab in LAB.findall(src):
            for body in LST.findall(lab):
                sql = strip_comments(body)
                where = f"Ch {ch}"
                # contract 1
                for py, rest in PYRUN.findall(body):
                    path = os.path.join(ROOT, py if py.startswith("code/")
                                        else os.path.join("code", py))
                    seen.append(f"{where}: python {py}{rest}")
                    if not os.path.exists(path):
                        errs.append(f"{where}: python {py} -- no such file")
                        continue
                    ok = accepted_flags(path)
                    for f in FLAG.findall(rest):
                        if f not in ok:
                            errs.append(f"{where}: {py} given --{f}, which "
                                        f"it does not accept")
                # contract 2
                for p in READ.findall(body):
                    seen.append(f"{where}: .read {p}")
                    if not os.path.exists(os.path.join(ROOT, p)):
                        errs.append(f"{where}: .read {p} -- no such file")
                # contract 3
                for pr in PRAGMA.findall(sql):
                    seen.append(f"{where}: PRAGMA {pr}")
                    if pr.lower() not in pragmas:
                        errs.append(f"{where}: PRAGMA {pr} is not a pragma "
                                    f"this SQLite knows")
                # contract 4
                check_inserts(sql, known, where, errs, seen)

    if a.verbose:
        for s in seen:
            print(f"  ok?  {s}")
    print(f"\n  {len(seen)} things checked, {len(errs)} problem(s)")
    for e in errs:
        print(f"    {e}")
    return len(errs)


if __name__ == "__main__":
    sys.exit(main())
