#!/usr/bin/env python3
"""Catch the two TikZ faults in this book that cost the most to diagnose.

Both produce a cascade of errors anchored at the *end* of the figure, dozens of
lines away from the node that caused them, and both begin with the unhelpful
``LaTeX Error: Not allowed in LR mode``.

1. **A ``\\\\`` in a node that is not aligned.** TikZ only gives ``\\\\`` a
   meaning inside a node whose options set ``align`` or ``text width``; without
   one, ``\\\\`` is the ordinary LaTeX line break and is illegal in the
   restricted horizontal mode a node is built in.

2. **A ``\\\\`` inside a nested group of an aligned node** -- typically
   ``{Docker\\\\\\emph{the image,\\\\not the VM}}``. The alignment is only in
   force at the node's own brace level, so the inner ``\\\\`` is again the
   LaTeX line break. This one is worse than the first: the node looks correct,
   and the failure is three figures further down the file.

Run it before every build; it is instant.

    python checkfigs.py            every chapter
    python checkfigs.py 04 05      only these
"""
from __future__ import annotations

import glob
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PARTS = os.path.join(HERE, "parts")

PICTURE = re.compile(r"\\begin\{tikzpicture\}(.*?)\\end\{tikzpicture\}", re.S)
STYLEDEF = re.compile(r"([A-Za-z@][\w@]*)/\.style\s*=\s*\{")
ALIGNY = re.compile(r"\b(align|text width)\s*=")

# Defining a style with one of these names shadows a built-in TikZ key, and the
# error names the key rather than the style, which sends you looking in the
# wrong place.  "cap" cost an hour in this book; the rest are from the sibling
# courses' READMEs and from pgf's own key list.
RESERVED = {
    "cap", "step", "pos", "in", "out", "at", "to", "node", "draw", "fill",
    "text", "above", "below", "left", "right", "shift", "scale", "rotate",
    "anchor", "shape", "color", "name", "label", "sloped", "swap", "grid",
    "help lines", "domain", "samples", "join", "line cap", "line join",
    "midway", "sharp corners", "rounded corners", "clip", "overlay", "opacity",
}


def _balanced(s: str, start: int) -> tuple[str, int]:
    """Return the group beginning at s[start] == '{', and the index after it."""
    depth, i = 0, start
    while i < len(s):
        if s[i] == "{" and (i == 0 or s[i - 1] != "\\"):
            depth += 1
        elif s[i] == "}" and s[i - 1] != "\\":
            depth -= 1
            if depth == 0:
                return s[start + 1:i], i + 1
        i += 1
    return s[start + 1:], len(s)


def aligned_styles(pic: str) -> set[str]:
    """Style names in this picture that end up setting align or text width.

    A style counts if its own body sets one, or if it names another style that
    does -- ``g/.style={st, draw=...}`` is aligned whenever ``st`` is. Resolved
    to a fixed point, since the chains are arbitrary.
    """
    bodies: dict[str, str] = {}
    for m in STYLEDEF.finditer(pic):
        bodies[m.group(1)], _ = _balanced(pic, m.end() - 1)
    out = {n for n, b in bodies.items() if ALIGNY.search(b)}
    changed = True
    while changed:
        changed = False
        for n, b in bodies.items():
            if n in out:
                continue
            if set(re.findall(r"[A-Za-z@][\w@]*", b)) & out:
                out.add(n)
                changed = True
    return out


def check_picture(pic: str, name: str, base_line: int) -> list[str]:
    probs: list[str] = []
    for m in STYLEDEF.finditer(pic):
        if m.group(1) in RESERVED:
            line = base_line + pic[:m.start()].count("\n") + 1
            probs.append(f"{name}:{line}: style named {m.group(1)!r} shadows a "
                         f"built-in TikZ key -- rename it")
    ok_styles = aligned_styles(pic)
    for m in re.finditer(r"\\node\s*\[", pic):
        opts, after = _balanced("{" + pic[m.end():], 0)
        # re-find the real option text: the bracket group, not a brace group
        depth, i = 1, m.end()
        while i < len(pic) and depth:
            if pic[i] == "[":
                depth += 1
            elif pic[i] == "]":
                depth -= 1
            i += 1
        opts = pic[m.end():i - 1]
        j = pic.find("{", i - 1)
        if j < 0:
            continue
        text, _ = _balanced(pic, j)
        if "\\\\" not in text:
            continue
        line = base_line + pic[:m.start()].count("\n") + 1
        named = set(re.findall(r"[A-Za-z@][\w@]*", opts))
        has_align = bool(ALIGNY.search(opts)) or bool(named & ok_styles)
        if not has_align:
            probs.append(f"{name}:{line}: node has a line break but no "
                         f"align/text width -- {text.strip()[:52]!r}")
            continue
        # fault 2: a line break inside a nested group of an aligned node
        k = 0
        while k < len(text):
            if text[k] == "{":
                inner, nxt = _balanced(text, k)
                if "\\\\" in inner:
                    probs.append(f"{name}:{line}: line break inside a nested "
                                 f"group -- {{{inner.strip()[:46]}}}")
                k = nxt
            else:
                k += 1
    return probs


def check(path: str) -> list[str]:
    name = os.path.basename(path)
    src = io.open(path, encoding="utf-8").read()
    probs: list[str] = []
    for m in PICTURE.finditer(src):
        probs += check_picture(m.group(1), name, src[:m.start()].count("\n"))
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
    print("  figures clean")
    return 0


if __name__ == "__main__":
    sys.exit(main())
