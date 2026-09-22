#!/usr/bin/env python3
"""The handout's Font Awesome icons, as recolourable vector pictures.

Every pedagogy box in the handout carries an icon (a book for a definition, a
lightbulb for a concept, a flask for a lab ...). The slides use the *same*
glyphs, so a student who has learnt that the calculator means "worked
example" in print sees it again on the screen.

Each icon is typeset once by pdflatex + fontawesome5 in a sentinel colour,
converted to SVG, and recoloured on demand by rewriting the fill -- one
compile per icon, any number of colours. A PNG fallback is produced from the
alpha channel in the same colour.
"""
from __future__ import annotations

import os
import re
import subprocess

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
DIR = os.path.join(HERE, "figures", "icons")
BUILD = os.path.join(HERE, "figures", "build", "icons")
SENTINEL = "FF00FF"

NAMES = ["book", "lightbulb", "exclamation-triangle", "code", "bullseye",
         "link", "tags", "calculator", "exclamation-circle", "flask",
         "question-circle", "clipboard-check", "compass", "graduation-cap",
         "tasks", "microchip", "shield-alt", "pen-fancy", "arrow-right",
         "check", "times", "list-ul", "map-signs", "book-open", "file-code",
         "calendar-alt", "clipboard-list", "user-tie", "comments",
         "chart-line", "project-diagram", "layer-group", "play-circle",
         "flag-checkered", "sitemap", "table", "image", "quote-left"]


def _compile_all() -> None:
    os.makedirs(BUILD, exist_ok=True)
    os.makedirs(DIR, exist_ok=True)
    missing = [n for n in NAMES
               if not os.path.exists(os.path.join(DIR, n + ".svg"))]
    if not missing:
        return
    body = "\n".join(r"\begin{ic}\faIcon{%s}\end{ic}" % n for n in missing)
    src = (r"\documentclass[border=0.6pt,multi=ic]{standalone}"
           "\n\\usepackage{fontawesome5}\\usepackage{xcolor}"
           "\n\\definecolor{s}{HTML}{%s}\n\\begin{document}\\color{s}"
           "\\fontsize{40}{40}\\selectfont\n%s\n\\end{document}\n"
           % (SENTINEL, body))
    tex = os.path.join(BUILD, "icons.tex")
    with open(tex, "w", encoding="utf-8") as fh:
        fh.write(src)
    r = subprocess.run(["pdflatex", "-interaction=nonstopmode",
                        "icons.tex"], cwd=BUILD, capture_output=True)
    pdf = os.path.join(BUILD, "icons.pdf")
    if not os.path.exists(pdf):
        raise RuntimeError("icon compile failed: " + pdf)
    import fitz
    doc = fitz.open(pdf)
    for i, n in enumerate(missing):
        subprocess.run(["pdftocairo", "-svg", "-f", str(i + 1), "-l",
                        str(i + 1), pdf, os.path.join(DIR, n + ".svg")],
                       capture_output=True)
        pix = doc[i].get_pixmap(dpi=600, alpha=True)
        pix.save(os.path.join(DIR, n + ".png"))
    doc.close()


def _rgb_pct(hexv: str) -> str:
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return r, g, b


def icon(name: str, colour: str) -> tuple[str, str, float]:
    """(svg, png, aspect w/h) for *name* drawn in *colour* (hex)."""
    _compile_all()
    colour = colour.upper().lstrip("#")
    svg_out = os.path.join(DIR, f"{name}-{colour}.svg")
    png_out = os.path.join(DIR, f"{name}-{colour}.png")
    base_png = os.path.join(DIR, name + ".png")
    if not (os.path.exists(svg_out) and os.path.exists(png_out)):
        svg = open(os.path.join(DIR, name + ".svg"), encoding="utf-8").read()
        r, g, b = _rgb_pct(colour)
        # pdftocairo writes fills as rgb(100%, 0%, 100%) for the sentinel
        svg = re.sub(r"rgb\(\s*100%\s*,\s*0%\s*,\s*100%\s*\)",
                     f"rgb({r}, {g}, {b})", svg)
        svg = re.sub(r"#ff00ff", "#" + colour.lower(), svg, flags=re.I)
        with open(svg_out, "w", encoding="utf-8") as fh:
            fh.write(svg)
        im = Image.open(base_png).convert("RGBA")
        alpha = im.getchannel("A")
        solid = Image.new("RGBA", im.size, (r, g, b, 255))
        solid.putalpha(alpha)
        solid.save(png_out)
    with Image.open(base_png) as im:
        w, h = im.size
    return svg_out, png_out, w / h
