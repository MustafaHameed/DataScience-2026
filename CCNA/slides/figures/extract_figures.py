#!/usr/bin/env python3
"""Render the book's TikZ figures as slide-ready PNGs.

The volume already contains 32 figures drawn in TikZ. Redrawing them for the
slides would guarantee they drift, so they are compiled from the book's own
sources instead.

Pipeline per figure::

    parts/NN_x.tex  ->  build/<label>.tex  (standalone)  ->  PDF  ->  PNG

The standalone preamble is *generated from ``ccna.sty`` at build time* -- the
colour definitions, the tikz libraries and the device macros are lifted out of
it rather than copied by hand, so a palette change in the book reaches the
slides on the next run. ccna.sty itself cannot simply be loaded: it sets up
geometry, fancyhdr and titlesec against a ``report`` class that standalone
does not provide.

Progressive builds
------------------
A figure with an annotated copy at ``figures/src/<label>.tex`` is compiled once
per build level. Annotate by wrapping elements in ``\\stepvis{n}{...}``: they
appear only when ``n <= \\buildlevel``. The result is ``<label>-1.png`` ...
``<label>-N.png``, which ``build_slides.py`` lays onto consecutive slides for
PowerPoint's Morph to animate between.

Usage::

    python extract_figures.py              # all figures
    python extract_figures.py --ch 1 3     # only these chapters
    python extract_figures.py --list       # what would be built
"""
from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SLIDES = os.path.dirname(HERE)
ROOT = os.path.dirname(SLIDES)                 # ...\CCNA
PARTS = os.path.join(ROOT, "parts")
STY = os.path.join(ROOT, "ccna.sty")
SRC = os.path.join(HERE, "src")                # hand-annotated step versions
BUILD = os.path.join(HERE, "build")
OUT = os.path.join(HERE, "out")

DPI = 300


# ---------------------------------------------------------------------------
# Preamble, generated from ccna.sty
# ---------------------------------------------------------------------------

def _grab(sty: str, start: str, end: str) -> str:
    i = sty.find(start)
    if i == -1:
        return ""
    j = sty.find(end, i + len(start))
    return sty[i:j if j != -1 else len(sty)]


def preamble() -> str:
    """Assemble the smallest preamble that renders the book's figures."""
    sty = open(STY, encoding="utf-8").read()

    libs = re.search(r"\\usetikzlibrary\{[^}]*\}", sty, re.S)
    colours = "\n".join(re.findall(r"\\definecolor\{[^}]*\}\{[^}]*\}\{[^}]*\}",
                                   sty))
    # device styles and macros: from the tikzset that defines dev/ to the end
    # of the topology environment definition
    devs = _grab(sty, "\\tikzset{\n  dev/.style=", "\\newenvironment{topology}")
    topo = _grab(sty, "\\newenvironment{topology}", "%---------")

    return "\n".join([
        r"\documentclass[border=3pt]{standalone}",
        r"\usepackage[T1]{fontenc}",
        r"\usepackage[utf8]{inputenc}",
        r"\usepackage{lmodern}",
        r"\usepackage{amsmath,amssymb}",
        r"\usepackage{xcolor}",
        r"\usepackage{tikz}",
        r"\usepackage{fontawesome5}",
        r"\usepackage{ragged2e}",
        libs.group(0) if libs else r"\usetikzlibrary{arrows.meta,positioning}",
        r"\usetikzlibrary{decorations.pathreplacing,decorations.markings}",
        colours,
        r"% --- carried over from ccna.sty so figures match the book exactly",
        devs,
        topo,
        r"\newcommand{\cmd}[1]{\texttt{\small #1}}",
        r"\newcommand{\term}[1]{\textbf{\textcolor{primary}{#1}}}",
        r"\newcommand{\chref}[1]{Chapter~\ref{ch:#1}}",
        r"\newcommand{\ccnafitw}[1]{#1}",
        r"\providecommand{\thefigure}{}",
        r"% --- progressive builds",
        r"\newcount\buildlevel \buildlevel=99",
        r"% \stepvis{n}{..} appears from level n onward; \steponly{n}{..} at",
        r"% level n alone, which is how a caption is swapped rather than",
        r"% accumulated as the build advances. \stepdim{n}{..} is the",
        r"% complement of \stepvis: it shows only BEFORE level n, so a",
        r"% pair of them draws an element faint and then solid. That keeps",
        r"% the canvas full from step 1, which matters for a figure whose",
        r"% shape is itself the lesson -- a five-rung ladder revealed one",
        r"% rung at a time leaves a void where the method should be.",
        r"\newcommand{\stepvis}[2]{\ifnum#1>\buildlevel\relax\else#2\fi}",
        r"\newcommand{\steponly}[2]{\ifnum#1=\buildlevel #2\fi}",
        r"\newcommand{\stepdim}[2]{\ifnum#1>\buildlevel #2\fi}",
        r"% Slide text does not hyphenate. A projected caption broken as",
        r"% ``uni-cast'' costs the reader a beat that print can afford and a",
        r"% lecture cannot, and the ragged edge it avoids is invisible at",
        r"% four metres anyway.",
        r"\hyphenpenalty=10000 \exhyphenpenalty=10000",
        r"\begin{document}",
    ])


# ---------------------------------------------------------------------------
# Finding figures in the book
# ---------------------------------------------------------------------------

def _balanced(s: str, i: int) -> tuple[str, int]:
    depth, j = 0, i
    while j < len(s):
        if s[j] == "{" and s[j - 1] != "\\":
            depth += 1
        elif s[j] == "}" and s[j - 1] != "\\":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return s[i + 1:], len(s)


def find_figures(stem: str) -> list[dict]:
    """Every \\ccnafig / \\ccnafigh in one chapter file."""
    path = os.path.join(PARTS, stem + ".tex")
    src = open(path, encoding="utf-8").read()
    figs = []
    for m in re.finditer(r"\\ccnafigh?\{", src):
        i = m.end() - 1
        body, j = _balanced(src, i)
        while j < len(src) and src[j] in " \n%":
            j += 1
        if j >= len(src) or src[j] != "{":
            continue
        caption, j = _balanced(src, j)
        while j < len(src) and src[j] in " \n%":
            j += 1
        if j >= len(src) or src[j] != "{":
            continue
        label, _ = _balanced(src, j)
        figs.append({"label": label.strip(),
                     "body": body.strip().lstrip("%").strip(),
                     "caption": re.sub(r"\s+", " ", caption).strip()})
    return figs


def chapter_stems() -> list[tuple[int, str]]:
    master = open(os.path.join(ROOT, "CCNA_v2_Handout.tex"),
                  encoding="utf-8").read()
    out, n = [], 0
    for stem in re.findall(r"\\input\{parts/([^}]+)\}", master):
        p = os.path.join(PARTS, stem + ".tex")
        if not os.path.exists(p):
            continue
        body = open(p, encoding="utf-8").read()
        if re.search(r"^\\chapter\{", body, re.M):
            n += 1
            out.append((n, stem))
        elif "opener" in stem or "frontmatter" in stem:
            out.append((0, stem))
    return out


# ---------------------------------------------------------------------------
# Compiling
# ---------------------------------------------------------------------------

def steps_for(label: str) -> int:
    """How many build levels an annotated figure declares, else 1."""
    p = os.path.join(SRC, label + ".tex")
    if not os.path.exists(p):
        return 1
    body = open(p, encoding="utf-8").read()
    # Both markers count. A figure whose last level is introduced by
    # \steponly -- a caption swap with nothing new appearing -- was being
    # built one level short, silently dropping its final state.
    levels = [int(x) for x in
              re.findall(r"\\step(?:vis|only|dim)\{(\d+)\}", body)]
    return max(levels) if levels else 1


def body_for(label: str, fallback: str) -> str:
    p = os.path.join(SRC, label + ".tex")
    if os.path.exists(p):
        return open(p, encoding="utf-8").read()
    return fallback


def render(label: str, body: str, level: int | None, pre: str) -> str | None:
    """Compile one figure (at one build level) to PNG. Returns the path."""
    os.makedirs(BUILD, exist_ok=True)
    os.makedirs(OUT, exist_ok=True)
    name = label if level is None else f"{label}-{level}"
    tex = os.path.join(BUILD, name + ".tex")
    setlvl = "" if level is None else f"\\buildlevel={level}\n"
    with open(tex, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(pre + "\n" + setlvl + body + "\n\\end{document}\n")

    r = subprocess.run(
        ["pdflatex", "-interaction=nonstopmode", "-halt-on-error",
         "-output-directory", BUILD, tex],
        capture_output=True, text=True, cwd=BUILD)
    pdf = os.path.join(BUILD, name + ".pdf")
    if r.returncode != 0 or not os.path.exists(pdf):
        log = os.path.join(BUILD, name + ".log")
        why = ""
        if os.path.exists(log):
            txt = open(log, encoding="utf-8", errors="replace").read()
            m = re.search(r"^! (.+)$", txt, re.M)
            why = m.group(1) if m else ""
        print(f"    FAILED {name}  {why}")
        return None

    png = os.path.join(OUT, name)
    subprocess.run(["pdftocairo", "-png", "-r", str(DPI), "-transp",
                    "-singlefile", pdf, png],
                   capture_output=True, text=True)
    final = png + ".png"
    return final if os.path.exists(final) else None


def crop_check(label: str, body: str, level: int | None,
               pre: str) -> str:
    """Report content that a pinned bounding box will silently crop.

    ``\\useasboundingbox`` fixes the canvas so Morph fades instead of sliding,
    but it also *crops*: a caption without an explicit ``text width`` runs off
    the edge and is cut mid-word, with no warning from anywhere. So the figure
    is compiled a second time with the pin turned into a ``\\path`` -- which
    only ever grows the box -- and the two page sizes are compared.

    Returns "" when nothing is cropped, else a description.
    """
    if "useasboundingbox" not in body:
        return ""
    loose = body.replace("\\useasboundingbox", "\\path")
    a = _page_size(label + "-pin", body, level, pre)
    b = _page_size(label + "-loose", loose, level, pre)
    if not a or not b:
        return ""
    dw, dh = b[0] - a[0], b[1] - a[1]
    if dw > 2.0 or dh > 2.0:            # points; ignore rounding
        return f"content extends {dw:.0f}x{dh:.0f}pt beyond the pinned box"
    return ""


def _page_size(name: str, body: str, level: int | None,
               pre: str) -> tuple[float, float] | None:
    os.makedirs(BUILD, exist_ok=True)
    tag = name if level is None else f"{name}-{level}"
    tex = os.path.join(BUILD, "chk-" + tag + ".tex")
    setlvl = "" if level is None else f"\\buildlevel={level}\n"
    with open(tex, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(pre + "\n" + setlvl + body + "\n\\end{document}\n")
    r = subprocess.run(
        ["pdflatex", "-interaction=nonstopmode", "-halt-on-error",
         "-output-directory", BUILD, tex],
        capture_output=True, text=True, cwd=BUILD)
    pdf = os.path.join(BUILD, "chk-" + tag + ".pdf")
    if r.returncode != 0 or not os.path.exists(pdf):
        return None
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True)
    m = re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", out.stdout)
    return (float(m.group(1)), float(m.group(2))) if m else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", nargs="*", type=int, help="chapters to build")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--clean", action="store_true")
    ap.add_argument("--check-crop", action="store_true",
                    help="also verify nothing falls outside a pinned "
                         "bounding box (slow: one extra compile per level)")
    a = ap.parse_args()

    if a.clean and os.path.isdir(BUILD):
        shutil.rmtree(BUILD)

    pre = preamble()
    os.makedirs(SRC, exist_ok=True)

    total, made, failed, cropped = 0, 0, [], []
    for n, stem in chapter_stems():
        if a.ch and n not in a.ch:
            continue
        figs = find_figures(stem)
        if not figs:
            continue
        print(f"  ch {n:>2}  {stem}")
        for f in figs:
            total += 1
            nsteps = steps_for(f["label"])
            body = body_for(f["label"], f["body"])
            annotated = os.path.exists(os.path.join(SRC, f["label"] + ".tex"))
            tag = f" [{nsteps} steps]" if nsteps > 1 else ""
            src = " (annotated)" if annotated else ""
            print(f"        {f['label']}{tag}{src}")
            if a.list:
                continue
            if nsteps > 1:
                for lv in range(1, nsteps + 1):
                    p = render(f["label"], body, lv, pre)
                    made += 1 if p else 0
                    if not p:
                        failed.append(f"{f['label']}-{lv}")
                    elif a.check_crop:
                        why = crop_check(f["label"], body, lv, pre)
                        if why:
                            cropped.append(f"{f['label']}-{lv}: {why}")
            else:
                p = render(f["label"], body, None, pre)
                made += 1 if p else 0
                if not p:
                    failed.append(f["label"])

    if a.list:
        print(f"\n  {total} figures")
        return 0
    print(f"\n  {made} PNG written to figures/out/  ({total} figures)")
    if cropped:
        print("  CROPPED (widen the box, or give the node a text width):")
        for c in cropped:
            print("    " + c)
    if failed:
        print("  failed: " + ", ".join(failed))
        return 1
    return 1 if cropped else 0


if __name__ == "__main__":
    sys.exit(main())
