#!/usr/bin/env python3
"""Catch the source-level traps in parts/*.tex that LaTeX reports badly.

Each check exists because the error it prevents was made more than once,
and because the message TeX gives for it points somewhere unhelpful.

1. Special characters in a box title. A raw _ in a `verify` or `config`
   title is a subscript, and TeX reports the error at the \\end{} of the
   environment rather than at the title -- so the line number in the log is
   actively misleading. Same for # $ % & ~ ^.

2. A tinted colour passed into a TikZ style argument. Styles here are
   written `s/.style={draw=#1, fill=#1!10, text=#1!40!black}`; invoking one
   as `s=dom4!60` expands `#1!40!black` to `dom4!60!40!black`, and xcolor
   reads the `40` as a colour name. The error is "Undefined color `40'",
   which names neither the style nor the node. Pass the base colour and
   override the tint on the node: `[s=dom4, fill=dom4!28]`.

3. A `.style` that redefines a key TikZ already owns. `cap` is TikZ's own
   line-cap key, so `cap/.style={...}` shadows it and TeX reports "The key
   '/tikz/cap' requires a value" at the first node that uses the style --
   naming a key the source never mentions. Made three times here before it
   was checked; the check then found a fourth, latent one in chapter 30,
   which worked only because that style happened to take an argument.

Run before building. Exit status is 0 when nothing is found.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PARTS = ROOT / "parts"
# The hand-annotated slide figures are the same kind of TikZ and hit
# the same traps, so they are linted from here rather than separately.
FIGSRC = ROOT / "slides" / "figures" / "src"

TITLE_ENVS = ("verify", "config", "hostcmd", "lab", "breakfix")
TITLE_RE = re.compile(
    r"\\begin\{(" + "|".join(TITLE_ENVS) + r")\}"
    r"((?:\[[^\]\n]*\]|\{[^{}\n]*\})+)"
)
BAD_CHARS = set("_#$%&^~")

# Key names TikZ/PGF already own. Redefining one with `name/.style={...}`
# shadows it, and the error names the key rather than the style or the node.
RESERVED = {
    "cap", "join", "step", "at", "to", "in", "out", "pos", "shift",
    "scale", "rotate", "anchor", "name", "label", "text", "font",
    "fill", "draw",
    "color", "opacity", "sloped", "midway", "above", "below", "left",
    "right", "distance", "looseness", "densely", "loosely", "help lines",
    "every node", "line width", "dash pattern", "rounded corners",
    "inner sep", "outer sep", "minimum width", "minimum height", "align",
}

# A style key taking a colour argument, invoked with a tint: s=dom4!60.
# Styles that merely pass #1 straight through as an option are safe, so
# this reports candidates and says why rather than failing blindly.
TINT_RE = re.compile(r"\[([a-zA-Z@]+)=((?:dom[1-5]|primary|accent|neutral)![0-9]+)")


def main() -> int:
    problems = 0
    sources = sorted(PARTS.glob("*.tex"))
    if FIGSRC.is_dir():
        sources += sorted(FIGSRC.glob("*.tex"))
    for path in sources:
        lines = path.read_text(encoding="utf-8").split("\n")

        # Which style keys in this file build a derived colour from #1?
        derived = set()
        for line in lines:
            for m in re.finditer(r"([a-zA-Z@]+)/\.style\s*=\s*\{([^\n]*)", line):
                if re.search(r"#1![0-9]+", m.group(2)):
                    derived.add(m.group(1))

        for n, line in enumerate(lines, 1):
            for m in TITLE_RE.finditer(line):
                args = re.sub(r"\\.", "", m.group(2))   # drop \_ \# \% etc
                hits = sorted(set(args) & BAD_CHARS)
                if hits:
                    problems += 1
                    print(f"{path.name}:{n}: raw {''.join(hits)} in "
                          f"{m.group(1)} title -- escape it")
                    print(f"    {m.group(0)[:96]}")

            for m in re.finditer(r"([a-zA-Z@ ]+)/\.style\s*=", line):
                key = m.group(1).strip().split(",")[-1].strip()
                if key in RESERVED:
                    problems += 1
                    print(f"{path.name}:{n}: .style redefines the "
                          f"reserved TikZ key '{key}' -- rename it "
                          f"(the error TeX gives names the key, not "
                          f"this line)")

            for m in TINT_RE.finditer(line):
                if m.group(1) in derived:
                    problems += 1
                    print(f"{path.name}:{n}: tint passed to style "
                          f"'{m.group(1)}' which derives a colour from it "
                          f"-- xcolor will fail")
                    print(f"    {m.group(0)[:96]}")

    if problems:
        print(f"\n{problems} problem(s) found.")
        return 1
    print("Source lint clean.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
