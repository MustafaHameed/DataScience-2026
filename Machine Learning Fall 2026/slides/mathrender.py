#!/usr/bin/env python3
"""Typeset maths for the slides with LuaLaTeX, in the slides' own faces.

Inline maths that Unicode can carry (P(A ∣ B), F₁, σ/√n) stays as live text.
What it cannot -- display equations, \\dfrac, \\underbrace, a sum with limits
-- is typeset here instead, and placed on the slide as a vector picture.

The faces are chosen to disappear into the slide: text in **Segoe UI** (the
deck's body face) and maths in **Cambria Math** (PowerPoint's own equation
face), both loaded from Windows by fontspec/unicode-math. A worked-example line
such as "**Precision** = 80/(80+110) = 42.1%" therefore reads as one line of
slide text that happens to contain a real fraction.

Each snippet is typeset at the exact width and point size of the box it will
sit in, so TeX makes the line breaks and the picture needs no scaling.
Rendering is batched: the deck builder *requests* snippets during a dry run,
``flush()`` compiles every missing one in a single LuaLaTeX pass (one snippet
per page), and the real build finds them all cached.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(HERE, "figures", "math")
BUILD = os.path.join(HERE, "figures", "build", "math")
INDEX = os.path.join(CACHE, "index.json")
PNG_DPI = 300

# Colours the snippets may use, by name -- same values as mlhandout.sty.
COLOURS = {
    "ink": "1C252B", "primary": "1A5276", "secondary": "2E86C1",
    "accent": "E74C3C", "greenhl": "27AE60", "purplehl": "8E44AD",
    "goldhl": "B7950B", "tealhl": "117A65", "neutral": "707B7C",
    "white": "FFFFFF", "muted": "5D6D7E", "lensLA": "6C3483",
    "lensPM": "1F618D", "lensIOT": "117864", "lensSEC": "922B21",
    "golddark": "7D6608", "defink": "7E3F12",
}

PREAMBLE = r"""\documentclass[border=0pt,multi=snip,varwidth=false]{standalone}
\usepackage{fontspec}
\usepackage{unicode-math}
\setmainfont{Segoe UI}[BoldFont={Segoe UI Bold},ItalicFont={Segoe UI Italic},
  BoldItalicFont={Segoe UI Bold Italic}]
\setmonofont{Consolas}
\setmathfont{Cambria Math}
\usepackage{xcolor}
\usepackage{amsmath}
\usepackage{ragged2e}
%(colours)s
\newcommand{\term}[1]{\textbf{\textcolor{primary}{#1}}}
\newcommand{\dsref}[1]{Figure}
\newcommand{\chref}[1]{Chapter}
\newcommand{\faIcon}[1]{}
\newcommand{\lensitem}[3]{#3}
\renewcommand{\_}{\textunderscore}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}
\begin{document}
"""


def _colours() -> str:
    return "\n".join(r"\definecolor{%s}{HTML}{%s}" % kv
                     for kv in COLOURS.items())


def _key(tex: str, width_in: float, size_pt: float, colour: str,
         align: str, display: bool) -> str:
    # "fit" versions display snippets: they shrink to fit instead of wrapping
    h = hashlib.sha1(json.dumps([tex, round(width_in, 3), round(size_pt, 2),
                                 colour, align, display, PREAMBLE] +
                                (["fit"] if display else []))
                     .encode("utf-8")).hexdigest()
    return h[:20]


def _clean(tex: str) -> str:
    """Make handout prose safe for a LuaLaTeX snippet."""
    t = tex
    # \mathbf{42.1\%} -> bold text: unicode-math would set the digits bold
    # but leave the point and the percent sign in the regular weight.
    t = re.sub(r"\\mathbf\{([0-9.,{}%\\ ]+)\}",
               lambda m: r"\text{\textbf{%s}}" % m.group(1), t)
    t = t.replace("\\smallskip", "").replace("\\medskip", "")
    t = re.sub(r"\\(Needspace|needspace)\*?\{[^}]*\}", "", t)
    return t.strip()


class MathRenderer:
    """Request -> (flush) -> fetch. Cached on disk across builds."""

    def __init__(self):
        os.makedirs(CACHE, exist_ok=True)
        self.index = {}
        if os.path.exists(INDEX):
            try:
                self.index = json.load(open(INDEX, encoding="utf-8"))
            except Exception:
                self.index = {}
        self.pending: dict[str, dict] = {}

    # -- public -----------------------------------------------------------

    def get(self, tex: str, width_in: float, size_pt: float,
            colour: str = "ink", align: str = "left",
            display: bool = False):
        """Return (svg, png, w_in, h_in), or a pending estimate.

        When the snippet is not yet cached, it is queued and an *estimate*
        is returned with svg=None, so a dry-run layout can proceed.
        """
        k = _key(tex, width_in, size_pt, colour, align, display)
        e = self.index.get(k)
        svg = os.path.join(CACHE, k + ".svg")
        png = os.path.join(CACHE, k + ".png")
        if e and os.path.exists(svg) and os.path.exists(png):
            return svg, png, e["w"], e["h"]
        self.pending[k] = dict(tex=tex, width=width_in, size=size_pt,
                               colour=colour, align=align, display=display)
        return None, None, width_in, self._estimate(tex, width_in, size_pt,
                                                    display)

    def flush(self) -> int:
        """Compile every pending snippet in one LuaLaTeX run."""
        if not self.pending:
            return 0
        os.makedirs(BUILD, exist_ok=True)
        keys = list(self.pending)
        body = []
        for k in keys:
            body.append(self._snippet(self.pending[k]))
        src = PREAMBLE.replace("%(colours)s", _colours()) + \
            "\n".join(body) + "\n\\end{document}\n"
        name = "batch"
        tex = os.path.join(BUILD, name + ".tex")
        with open(tex, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(src)
        r = subprocess.run(["lualatex", "-interaction=nonstopmode",
                            "-halt-on-error", name + ".tex"],
                           cwd=BUILD, capture_output=True, text=True,
                           encoding="utf-8", errors="replace")
        pdf = os.path.join(BUILD, name + ".pdf")
        if r.returncode != 0 or not os.path.exists(pdf):
            log = open(os.path.join(BUILD, name + ".log"), encoding="utf-8",
                       errors="replace").read()
            m = re.search(r"^! (.+?)$(.*?)^l\.(\d+)", log, re.M | re.S)
            raise RuntimeError("maths batch failed: " +
                               (m.group(1) + " at line " + m.group(3)
                                if m else "see " + pdf))
        import fitz
        doc = fitz.open(pdf)
        if len(doc) != len(keys):
            raise RuntimeError(f"maths batch: {len(doc)} pages for "
                               f"{len(keys)} snippets")
        for i, k in enumerate(keys):
            page = doc[i]
            w_in, h_in = page.rect.width / 72.0, page.rect.height / 72.0
            pix = page.get_pixmap(dpi=PNG_DPI, alpha=True)
            pix.save(os.path.join(CACHE, k + ".png"))
            # MuPDF, not pdftocairo, writes the SVG here: cairo's SVG surface
            # misplaces glyphs of LuaLaTeX's CID TrueType fonts (Segoe UI),
            # and PowerPoint then draws the line up to 0.14 in to the right
            # of where the PNG fallback puts it. Glyphs become paths.
            with open(os.path.join(CACHE, k + ".svg"), "w",
                      encoding="utf-8") as fh:
                fh.write(page.get_svg_image(text_as_path=True))
            self.index[k] = {"w": round(w_in, 4), "h": round(h_in, 4),
                             "tex": self.pending[k]["tex"][:120]}
        doc.close()
        n = len(keys)
        self.pending.clear()
        with open(INDEX, "w", encoding="utf-8") as fh:
            json.dump(self.index, fh, indent=0)
        return n

    # -- internals ----------------------------------------------------------

    @staticmethod
    def _estimate(tex, width_in, size_pt, display) -> float:
        lines = max(1, int(len(re.sub(r"\\[a-z]+|[{}$]", "", tex)) * 0.5 *
                           size_pt / 72.0 / max(width_in, 0.5)) + 1)
        tall = 2.1 if ("frac" in tex or "\\sum" in tex or display) else 1.3
        return lines * tall * size_pt / 72.0

    @staticmethod
    def _snippet(p: dict) -> str:
        size = p["size"]
        lead = size * 1.25
        w = p["width"]
        col = p["colour"]
        tex = _clean(p["tex"])
        if p["display"]:
            # A formula up to 15% too wide is set at a smaller size rather
            # than broken across lines (TeX would break it after a relation
            # and strand the result); a wider one still wraps.
            fit = (r"\setbox0=\hbox{$\displaystyle %s$}"
                   r"\ifdim\wd0>%.4fin\relax"
                   r"\ifdim\wd0<1.15\dimexpr %.4fin\relax"
                   r"\dimen0=\dimexpr %.2fpt*\number\dimexpr 0.985\dimexpr"
                   r" %.4fin\relax\relax/\number\wd0\relax"
                   r"\fontsize{\the\dimen0}{\the\dimen0}\selectfont"
                   r"\fi\fi" % (tex, w, w, size, w))
            inner = fit + (r"\centering$\displaystyle %s$" % tex
                           if p["align"] == "center"
                           else r"$\displaystyle %s$" % tex)
        else:
            inner = (r"\centering " if p["align"] == "center" else
                     r"\RaggedRight ") + tex
        # \lineskip keeps a fraction on one line from touching a root or a
        # fraction on the next: at the default 1pt they collide.
        return (r"\begin{snip}\fontsize{%.2f}{%.2f}\selectfont"
                r"\lineskiplimit=%.2fpt\lineskip=%.2fpt"
                r"\color{%s}\parbox[t]{%.4fin}{\strut %s\strut}\end{snip}"
                % (size, lead, size * 0.25, size * 0.3, col, w, inner))


# A process-wide renderer the deck builder shares.
_R: MathRenderer | None = None


def renderer() -> MathRenderer:
    global _R
    if _R is None:
        _R = MathRenderer()
    return _R
