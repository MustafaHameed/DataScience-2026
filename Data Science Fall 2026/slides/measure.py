#!/usr/bin/env python3
"""Measure text the way PowerPoint will set it.

python-pptx has no working autofit: PowerPoint recalculates only when a human
opens and edits the file. The CCNA decks therefore guessed sizes from an
average character width, and paid for the guess in slack (short text in tall
boxes) and in the odd overflow.

Here the guess is replaced by measurement. Glyph advances come from the actual
Windows font files through Pillow, and two facts were calibrated against
PowerPoint's own renderer (slide export at 150 px/in):

* a line is exactly ``1.2 x size x line_spacing`` tall, for Segoe UI,
  Consolas and Cambria alike;
* Pillow's advance widths match PowerPoint's to within 0.3%.

Word wrapping is greedy at spaces, as PowerPoint's is, so a paragraph's line
count -- and therefore its height -- is known before the slide is drawn.
"""
from __future__ import annotations

import os
from functools import lru_cache

from PIL import ImageFont

FONTS = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
LINE = 1.2                  # PowerPoint single-spacing factor (calibrated)
SAFETY = 1.02               # widths: absorb kerning / renderer differences

_FILES = {
    ("Segoe UI", False, False): "segoeui.ttf",
    ("Segoe UI", True, False): "segoeuib.ttf",
    ("Segoe UI", False, True): "segoeuii.ttf",
    ("Segoe UI", True, True): "segoeuiz.ttf",
    ("Segoe UI Semibold", False, False): "seguisb.ttf",
    ("Segoe UI Semibold", True, False): "seguisb.ttf",
    ("Segoe UI Semibold", False, True): "seguisbi.ttf",
    ("Segoe UI Semibold", True, True): "seguisbi.ttf",
    ("Segoe UI Light", False, False): "segoeuil.ttf",
    ("Segoe UI Black", False, False): "seguibl.ttf",
    ("Cambria", False, False): "cambria.ttc",
    ("Cambria", True, False): "cambriab.ttf",
    ("Cambria", False, True): "cambriai.ttf",
    ("Cambria", True, True): "cambriaz.ttf",
    ("Consolas", False, False): "consola.ttf",
    ("Consolas", True, False): "consolab.ttf",
    ("Consolas", False, True): "consolai.ttf",
    ("Consolas", True, True): "consolaz.ttf",
}
_UNITS = 1000               # fonts are loaded at 1000 px; widths scale


@lru_cache(maxsize=64)
def _font(family: str, bold: bool, italic: bool):
    f = _FILES.get((family, bold, italic)) or \
        _FILES.get((family, bold, False)) or \
        _FILES.get((family, False, False)) or "segoeui.ttf"
    return ImageFont.truetype(os.path.join(FONTS, f), _UNITS)


@lru_cache(maxsize=200000)
def width(text: str, size: float, family: str = "Segoe UI",
          bold: bool = False, italic: bool = False) -> float:
    """Advance width of *text* in inches at *size* points."""
    if not text:
        return 0.0
    return _font(family, bold, italic).getlength(text) / _UNITS * \
        size / 72.0 * SAFETY


# A "piece" is (text, size, family, bold, italic): one run's worth of style.

def _tokens(pieces):
    """Split styled pieces into wrap tokens: (word_width, space_width)."""
    toks = []          # list of [width_without_trailing_space, space_width]
    cur_w = 0.0
    for text, size, fam, b, it in pieces:
        parts = text.split(" ")
        for k, part in enumerate(parts):
            if k > 0:
                # a space ends the current word
                toks.append([cur_w, width(" ", size, fam, b, it)])
                cur_w = 0.0
            cur_w += width(part, size, fam, b, it)
    toks.append([cur_w, 0.0])
    return [t for t in toks if t[0] > 0 or t[1] > 0]


def count_lines(pieces, box_w: float, first_indent: float = 0.0) -> int:
    """Lines a paragraph of styled pieces occupies in a box *box_w* wide."""
    toks = _tokens(pieces)
    if not toks:
        return 1
    box_w += 1e-6        # a box sized to its text must not wrap it
    lines, x = 1, first_indent
    for w, sp in toks:
        if x > first_indent and x + w > box_w:
            lines += 1
            x = first_indent
        if w > box_w - first_indent:
            # a token longer than a line breaks by character
            over = int(w // max(box_w - first_indent, 0.1))
            lines += over
            x = first_indent + (w - over * (box_w - first_indent))
        else:
            x += w
        x += sp
    return lines


def line_widths(pieces, box_w: float) -> list[float]:
    """Width of each wrapped line -- for sizing chips and tight cards."""
    toks = _tokens(pieces)
    out, x = [], 0.0
    for w, sp in toks:
        if x > 0 and x + w > box_w:
            out.append(x)
            x = 0.0
        x += w + sp
    out.append(x)
    return out


def natural_width(pieces) -> float:
    return sum(w + sp for w, sp in _tokens(pieces))


def para_height(pieces, box_w: float, size: float, spacing: float = 1.0,
                first_indent: float = 0.0) -> float:
    return count_lines(pieces, box_w, first_indent) * LINE * size * \
        spacing / 72.0
