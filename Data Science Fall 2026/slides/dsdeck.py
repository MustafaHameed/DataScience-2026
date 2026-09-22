#!/usr/bin/env python3
"""The design system for the Data Science Fall 2026 lecture decks.

One module decides what a slide *looks* like, so 33 decks read as one course.
``build_slides.py`` decides *what* goes on each slide; this module draws it.

Principles, most of them learnt on the CCNA and ARM decks first:

* **Measure, do not guess.** Every text box is sized from real glyph widths
  (``measure.py``), so cards fit their content, text never runs off a slide,
  and long content is split *before* it overflows.
* **Hang content from the title.** Body content starts at one fixed height on
  every slide; the eye does not have to re-find it after each click.
* **Say it once.** A figure's caption becomes its title and takeaway; the
  handout reference lives in the footer, not in the body.
* **Real structure.** Titles are title placeholders (outline view, the
  accessibility checker and "go to slide" all work), pictures carry alt text,
  ornaments are marked decorative, the theme carries the course palette and
  faces, and the deck has PowerPoint sections that mirror the handout.
* **Motion with a purpose.** Lists, worked-example steps and cards build one
  click at a time; Morph carries code walk-throughs, lens cards, stepped
  figures and question-to-answer. Nothing moves for decoration.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

import icons
import measure as M
import pptxfx as FX
import texparse as T
from mathrender import renderer

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")
FIGS = os.path.join(HERE, "figures", "out")

# ---------------------------------------------------------------------------
# Palette -- dshandout.sty, verbatim, plus the few neutrals a screen needs
# ---------------------------------------------------------------------------

PRIMARY = "1A5276"
SECONDARY = "2E86C1"
ACCENT = "E74C3C"
LIGHTBG = "EBF5FB"
DEFBG = "FDF2E9"
GREENHL = "27AE60"
PURPLEHL = "8E44AD"
GOLDHL = "B7950B"
TEALHL = "117A65"
NEUTRAL = "707B7C"
INK = "1C252B"
MUTED = "5F6F7B"
FAINT = "98A4AE"
RULE = "D9E1E7"
TRACK = "EDF1F4"
WHITE = "FFFFFF"
CODEBG = "1E2833"
CODEBAR = "28343F"

LENS = {"LA": ("6C3483", "Learning Analytics", "graduation-cap"),
        "PM": ("1F618D", "Project Management", "tasks"),
        "IOT": ("117864", "Internet of Things", "microchip"),
        "SEC": ("922B21", "Cybersecurity", "shield-alt")}
LENS_ORDER = ("LA", "PM", "IOT", "SEC")

# The handout's dependency map colours the six parts; the decks follow it.
PART_HUE = {1: PRIMARY, 2: SECONDARY, 3: GREENHL, 4: PURPLEHL, 5: TEALHL,
            6: GOLDHL}

NAMED = {"primary": PRIMARY, "secondary": SECONDARY, "accent": ACCENT,
         "greenhl": "1E8449", "purplehl": PURPLEHL, "goldhl": "7D6608",
         "tealhl": TEALHL, "gray": MUTED, "grey": MUTED, "neutral": NEUTRAL,
         "black": INK, "white": WHITE, "red": "C0392B", "blue": SECONDARY,
         "lensLA": "6C3483", "lensPM": "1F618D", "lensIOT": "117864",
         "lensSEC": "922B21", "orange": "B35900"}


@dataclass(frozen=True)
class Style:
    label: str
    icon: str
    rule: str          # left rule / kicker / accent
    bg: str            # card fill
    ink: str = INK     # bold-lead colour inside the card


STYLES = {
    "definitionbox": Style("DEFINITION", "book", "B35900", "FDF2E9",
                           "8A4500"),
    "conceptbox": Style("CONCEPT", "lightbulb", SECONDARY, "EBF5FB",
                        "1F618D"),
    "alertbox": Style("WATCH OUT", "exclamation-triangle", ACCENT, "FDEDEC",
                      "A93226"),
    "examplebox": Style("EXAMPLE", "code", GREENHL, "EAF7EF", "1E8449"),
    "worked": Style("WORKED EXAMPLE", "calculator", GOLDHL, "FBF8EC",
                    "7D6608"),
    "pitfall": Style("COMMON PITFALLS", "exclamation-circle", ACCENT,
                     "FEF5F4", "A93226"),
    "lab": Style("TRY IT YOURSELF", "flask", TEALHL, "EEF6F4", "0E6655"),
    "checkpoint": Style("CHECKPOINT", "question-circle", "256B9A", "EAF2F8",
                        "256B9A"),
    "chaptersummary": Style("CHAPTER SUMMARY", "clipboard-check", "15425E",
                            "EEF3F7", "15425E"),
    "outcomes": Style("LEARNING OUTCOMES", "bullseye", PRIMARY, "F2F6F9",
                      PRIMARY),
    "prereq": Style("BEFORE YOU START", "link", "4D5656", "F4F6F6",
                    "4D5656"),
    "keyterms": Style("KEY TERMS", "tags", PURPLEHL, "F5EEF8", "6C3483"),
    "fourlenses": Style("FOUR LENSES", "compass", "123953", "F4F6F8",
                        "123953"),
    "reviewq": Style("REVIEW QUESTIONS", "pen-fancy", PRIMARY, "F2F6F9",
                     PRIMARY),
    "table": Style("AT A GLANCE", "table", PRIMARY, "F2F6F9", PRIMARY),
    "statement": Style("KEY IDEA", "quote-left", PRIMARY, "F2F6F9", PRIMARY),
    "points": Style("IN BRIEF", "list-ul", PRIMARY, "F2F6F9", PRIMARY),
    "course": Style("THE COURSE", "map-signs", PRIMARY, "F2F6F9", PRIMARY),
    "agenda": Style("IN THIS CHAPTER", "list-ul", PRIMARY, "F2F6F9",
                    PRIMARY),
}

BODY = "Segoe UI"
HEAD = "Segoe UI Semibold"
MONO = "Consolas"
MATHF = "Cambria"

# ---------------------------------------------------------------------------
# Grid (inches). One grid, used by every layout.
# ---------------------------------------------------------------------------

W, H = 13.333, 7.5
ML = MR = 0.72
CW = W - ML - MR                  # 11.893
KICK_Y, KICK_H = 0.44, 0.28
TITLE_Y, TITLE_H = 0.74, 1.02
BODY_Y = 1.90
BODY_B = 6.90
BODY_H = BODY_B - BODY_Y          # 5.00
FOOT_Y = 7.04


def rgb(h: str) -> RGBColor:
    return RGBColor.from_string(h.upper())


def mix(a: str, b: str, t: float) -> str:
    """a*(1-t) + b*t, as hex."""
    ca = [int(a[i:i + 2], 16) for i in (0, 2, 4)]
    cb = [int(b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join("%02X" % round(x * (1 - t) + y * t) for x, y in zip(ca, cb))


def ground(h: str) -> str:
    """The hue darkened until it carries white type (never lightened).

    Four part hues are fine as they are; the gold of Part VI and the green of
    Part III are not, and white type on them washes out. Scaling to a common
    luminance fixes those without inventing a colour.
    """
    c = [int(h[i:i + 2], 16) for i in (0, 2, 4)]
    lum = (c[0] * 299 + c[1] * 587 + c[2] * 114) / 1000
    if lum <= 88:
        return h.upper()
    k = 88 / lum
    return "".join("%02X" % min(255, int(v * k)) for v in c)


# ---------------------------------------------------------------------------
# Paragraph specs and the flow engine
# ---------------------------------------------------------------------------

@dataclass
class PSpec:
    """One paragraph to draw: styled runs plus how to set them."""
    runs: list
    size: float
    color: str = INK
    bold: bool = False
    italic: bool = False
    family: str = BODY
    align: str = "l"
    marker: str = ""           # bullet / number, hung in the left margin
    marker_color: str = PRIMARY
    marker_w: float = 0.0      # hanging indent width (in)
    space_after: float = 0.0   # pt
    spacing: float = 1.0       # line spacing multiple
    term_color: str = PRIMARY
    lead_color: str | None = None   # colour for a bold lead phrase

    def pieces(self):
        out = []
        for r in self.runs:
            size = self.size * (0.66 if r.script else 1.0)
            fam = MONO if r.mono else (MATHF if r.math else self.family)
            if r.mono:
                size = size * 0.94
            out.append((r.text, size, fam if fam in (BODY, MONO, MATHF)
                        else fam, self.bold or r.bold, self.italic or r.italic))
        return out

    def height(self, w: float) -> float:
        return M.para_height(self.pieces(), max(0.3, w - self.marker_w),
                             self.size, self.spacing)


def paras_height(ps: list[PSpec], w: float) -> float:
    h = 0.0
    for i, p in enumerate(ps):
        h += p.height(w)
        if i < len(ps) - 1:
            h += p.space_after / 72.0
    return h


@dataclass
class Blk:
    """A vertical slab of content inside a flow."""
    kind: str                  # text | image | table | code | gap
    h: float = 0.0
    w: float = 0.0
    paras: list = field(default_factory=list)       # text
    svg: str | None = None                            # image
    png: str | None = None
    align: str = "l"
    table: object = None                              # T.TableE
    code: object = None                               # T.CodeE
    size: float = 16
    step: bool = True          # revealed on its own click (in a build)
    per_para: bool = False     # a list: reveal item by item
    callout: str | None = None  # draw a coloured bar at the left
    alt: str = ""
    extra: dict = field(default_factory=dict)


def style_runs(p, runs, size, color=INK, bold=False, italic=False,
               family=BODY, term_color=PRIMARY, lead_color=None):
    """Write T.Run fragments into a python-pptx paragraph."""
    first_bold_done = False
    for i, r in enumerate(runs):
        run = p.add_run()
        run.text = r.text
        f = run.font
        sz = size * (0.94 if r.mono else 1.0)
        f.size = Pt(round(sz * 2) / 2)
        fam = MONO if r.mono else (MATHF if r.math else family)
        f.name = fam
        f.bold = bool(bold or r.bold)
        f.italic = bool(italic or r.italic)
        c = color
        if r.color and r.color in NAMED:
            c = NAMED[r.color]
        if r.term:
            c = term_color
        elif r.mono and not r.color:
            c = "1F4E79" if color == INK else color
        elif lead_color and r.bold and not first_bold_done and i == 0:
            c = lead_color
        f.color.rgb = rgb(c)
        if r.bold and i == 0:
            first_bold_done = True
        FX.baseline(run, r.script)


def _hang(p, w_in: float) -> None:
    pPr = p._p.get_or_add_pPr()
    mar = int(w_in * 914400)
    pPr.set("marL", str(mar))
    pPr.set("indent", str(-mar))


# ---------------------------------------------------------------------------
# The deck
# ---------------------------------------------------------------------------

@dataclass
class Rec:
    slide: object
    kind: str
    group: str
    progress: bool = True


class Deck:
    """One chapter's deck (or the course overview)."""

    def __init__(self, ch: T.Chapter | None, course="Data Science",
                 edition="BSIT · Fall 2026", author="Dr. Mustafa Hameed",
                 dept="Department of Information Technology",
                 uni="The Islamia University of Bahawalpur", part=None):
        self.prs = Presentation()
        self.prs.slide_width = Inches(W)
        self.prs.slide_height = Inches(H)
        self.ch = ch
        self.course, self.edition = course, edition
        self.author, self.dept, self.uni = author, dept, uni
        self.part = part if part is not None else (ch.part if ch else 1)
        self.hue = PART_HUE.get(self.part, PRIMARY)
        self.ground = ground(self.hue)
        self.deep = mix(self.ground, "000000", 0.28)
        self.accent_txt = self.ground         # hue readable on white
        self.tint = mix(self.hue, WHITE, 0.90)
        self._layout = self.prs.slide_layouts[5]      # "Title Only"
        self.recs: list[Rec] = []
        self.group = "Opening"
        self.links: list[tuple] = []         # (shape, target_key)
        self.anchors: dict = {}              # key -> slide
        self.math = renderer()
        FX.set_theme(self.prs, {
            "dk1": INK, "lt1": WHITE, "dk2": PRIMARY, "lt2": LIGHTBG,
            "accent1": PRIMARY, "accent2": SECONDARY, "accent3": ACCENT,
            "accent4": GREENHL, "accent5": PURPLEHL, "accent6": GOLDHL,
            "hlink": SECONDARY, "folHlink": PURPLEHL})
        FX.set_title_style(self.prs, PRIMARY, 30)
        cp = self.prs.core_properties
        cp.author = author
        cp.title = (f"Chapter {ch.number}: {ch.title}" if ch
                    else "Data Science — Course Overview")
        cp.subject = "Data Science — From Foundations to Intelligent " \
                     "Systems · BSIT · Fall 2026"
        cp.keywords = "data science; lecture slides; BSIT; Fall 2026"
        cp.category = "Lecture slides"

    # -- plumbing ------------------------------------------------------------

    def new(self, kind: str, bg: str | None = None, progress=True):
        s = self.prs.slides.add_slide(self._layout)
        if bg:
            f = s.background.fill
            f.solid()
            f.fore_color.rgb = rgb(bg)
        self.recs.append(Rec(s, kind, self.group, progress))
        return s

    # What to do with each kind of slide, when the handout gives no prose to
    # put in the notes. Short, and about the room, not the content.
    CUES = {
        "card": "Read the box aloud, then ask for an example from the room "
                "before moving on.",
        "table": "Do not read the table row by row: pick the row the room "
                 "is most likely to get wrong and ask why.",
        "statement": "Say this in your own words first; the slide is the "
                     "handout's wording.",
        "pitfall": "For each pitfall, ask who has seen it happen. One real "
                   "story beats five bullets.",
        "summary": "Ask the room to say each takeaway before you reveal it.",
        "divider": "Pause. Name the question this section answers.",
        "figure": "Give the room ten seconds with the figure before "
                  "explaining it.",
        "code": "Read the code aloud; ask what each line will print before "
                "running it.",
        "worked": "Each click adds one step. Ask for the next number before "
                  "revealing it.",
        "lens": "Ask which domain the room works in, and start there.",
        "lab": "Say what the room will have at the end of the lab, then let "
               "them start. Circulate rather than lecture.",
    }

    def notes(self, slide, text: str) -> None:
        text = (text or "").strip()
        rec = self.recs[-1] if self.recs and self.recs[-1].slide is slide \
            else None
        if not text and rec is not None:
            text = self.CUES.get(rec.kind, "")
        if text:
            slide.notes_slide.notes_text_frame.text = text[:8000]

    def rect(self, s, x, y, w, h, fill=None, line=None, name=None, lw=1.0,
             shape=MSO_SHAPE.RECTANGLE, radius=None, deco=True):
        sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w),
                                Inches(h))
        if fill is None:
            sh.fill.background()
        else:
            sh.fill.solid()
            sh.fill.fore_color.rgb = rgb(fill)
        if line is None:
            sh.line.fill.background()
        else:
            sh.line.color.rgb = rgb(line)
            sh.line.width = Pt(lw)
        sh.shadow.inherit = False
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = radius
        if name:
            sh.name = name
        tf = sh.text_frame
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        if deco:
            FX.decorative(sh)
        return sh

    def rrect(self, s, x, y, w, h, fill=None, line=None, name=None, r_in=0.08,
              lw=1.0, deco=True):
        rad = min(0.5, r_in / max(0.01, min(w, h)))
        return self.rect(s, x, y, w, h, fill, line, name, lw,
                         MSO_SHAPE.ROUNDED_RECTANGLE, rad, deco)

    def box(self, s, x, y, w, h, name=None, anchor=MSO_ANCHOR.TOP,
            wrap=True):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = wrap
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        FX.no_autofit(tf)
        if name:
            tb.name = name
        return tb

    def write(self, s, x, y, w, h, specs: list[PSpec], name=None,
              anchor=MSO_ANCHOR.TOP, wrap=True):
        """A text box holding *specs*, one paragraph each."""
        tb = self.box(s, x, y, w, h, name, anchor, wrap)
        tf = tb.text_frame
        for i, sp in enumerate(specs):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER,
                           "r": PP_ALIGN.RIGHT}[sp.align]
            p.line_spacing = sp.spacing
            p.space_after = Pt(sp.space_after if i < len(specs) - 1 else 0)
            p.space_before = Pt(0)
            if sp.marker:
                _hang(p, sp.marker_w)
                mr = p.add_run()
                # the tab jumps to the hanging indent, so wrapped lines and
                # the first line's text share one left edge
                mr.text = sp.marker + "\t"
                mr.font.size = Pt(sp.size)
                mr.font.bold = True
                mr.font.name = sp.family
                mr.font.color.rgb = rgb(sp.marker_color)
            style_runs(p, sp.runs, sp.size, sp.color, sp.bold, sp.italic,
                       sp.family, sp.term_color, sp.lead_color)
        return tb

    def text(self, s, x, y, w, h, txt, size=18, color=INK, bold=False,
             family=BODY, align="l", name=None, anchor=MSO_ANCHOR.TOP,
             spacing=1.0, italic=False, spc=None, wrap=True):
        sp = PSpec([T.Run(txt)], size, color, bold, italic, family, align,
                   spacing=spacing)
        tb = self.write(s, x, y, w, h, [sp], name, anchor, wrap)
        if spc:
            for r in tb.text_frame.paragraphs[0].runs:
                FX.letter_spacing(r, spc)
        return tb

    def picture(self, s, png, x, y, w=None, h=None, svg=None, name=None,
                alt="", deco=False):
        kw = {}
        if w is not None:
            kw["width"] = Inches(w)
        if h is not None:
            kw["height"] = Inches(h)
        pic = s.shapes.add_picture(png, Inches(x), Inches(y), **kw)
        if svg and os.path.exists(svg):
            FX.add_svg(s, pic, svg)
        if name:
            pic.name = name
        if deco:
            FX.decorative(pic)
        elif alt:
            FX.alt_text(pic, alt)
        return pic

    def icon(self, s, name, colour, x, y, h, nm=None):
        svg, png, ar = icons.icon(name, colour)
        return self.picture(s, png, x, y, h * ar, h, svg, nm, deco=True)

    # -- chrome ----------------------------------------------------------------

    def title(self, s, text, size=None, color=PRIMARY, x=ML, y=TITLE_Y,
              w=CW, h=TITLE_H, align="l", name="!!title", anchor=None,
              family=HEAD):
        """Fill the slide's real title placeholder."""
        t = s.shapes.title
        t.left, t.top = Inches(x), Inches(y)
        t.width, t.height = Inches(w), Inches(h)
        t.name = name
        tf = t.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        tf.vertical_anchor = anchor or MSO_ANCHOR.TOP
        FX.no_autofit(tf)
        if size is None:
            size = self.fit_size([T.Run(text)], w, h, 30, 22, HEAD)
        p = tf.paragraphs[0]
        p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER}[align]
        p.line_spacing = 0.96
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.name = family
        r.font.bold = False
        r.font.color.rgb = rgb(color)
        return t

    def fit_size(self, runs, w, h, hi, lo, family=BODY, spacing=0.96,
                 bold=False):
        for sz in range(int(hi), int(lo) - 1, -1):
            sp = PSpec(runs, sz, family=family, bold=bold, spacing=spacing)
            if sp.height(w) <= h + 0.01:
                return sz
        return lo

    def kicker(self, s, label, colour=None, icon_name=None, x=ML, y=KICK_Y,
               name="!!kicker"):
        colour = colour or self.accent_txt
        cx = x
        if icon_name:
            self.icon(s, icon_name, colour, x, y + 0.01, 0.24,
                      nm=name + "-icon")
            cx = x + 0.36
        tb = self.text(s, cx, y, CW - (cx - x), KICK_H, label.upper(), 12.5,
                       colour, True, name=name, spc=160)
        return tb

    def footer(self, s, ref: str = "", dark=False):
        col = "B9C6D0" if dark else FAINT
        left = f"{self.course}  ·  " + (
            f"Chapter {self.ch.number}  ·  {self.ch.title}" if self.ch
            else "Course overview")
        self.text(s, ML, FOOT_Y, CW * 0.62, 0.26, left, 10, col,
                  name="!!footL")
        tb = self.box(s, ML + CW * 0.55, FOOT_Y, CW * 0.45, 0.26,
                      name="!!footR")
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        if ref:
            r = p.add_run()
            r.text = ref + "     "
            r.font.size = Pt(10)
            r.font.name = BODY
            r.font.color.rgb = rgb(col)
        FX.slide_number_field(p, 10, "8795A1" if not dark else "D5DEE5",
                              BODY, str(len(self.recs)))

    def chrome(self, s, title, kicker, style: Style | None = None, ref="",
               title_color=PRIMARY, kicker_color=None):
        if kicker:
            self.kicker(s, kicker, kicker_color or
                        (style.rule if style else None),
                        style.icon if style else None)
        self.title(s, title, color=title_color)
        self.footer(s, ref)

    # -- flow engine -------------------------------------------------------

    def elements_to_blocks(self, elements, w, size, style: Style | None,
                           reveal_lists=False, color=INK):
        """Convert parsed handout elements into measured blocks."""
        blocks: list[Blk] = []
        lead = style.ink if style else PRIMARY
        gap = size * 0.55 / 72.0
        for e in elements:
            if isinstance(e, T.Para):
                if e.heavy:
                    blocks.append(self._math_block(e.tex, w, size, False,
                                                   color, "left", e.text))
                else:
                    sp = PSpec(e.runs, size, color, lead_color=lead,
                               spacing=1.04)
                    blocks.append(Blk("text", sp.height(w), w, [sp]))
            elif isinstance(e, T.ListE):
                items = []
                for k, it in enumerate(e.items):
                    mk = f"{k + 1}." if e.numbered else "•"
                    mw = M.width("00." if e.numbered else "•", size, BODY,
                                 True) + size * 0.45 / 72.0
                    if it.heavy:
                        if items:
                            blocks.append(self._list_blk(items, w,
                                                         reveal_lists))
                            items = []
                        b = self._math_block(it.tex, w - mw, size, False,
                                             color, "left", it.text)
                        b.extra["marker"] = (mk, mw, lead)
                        blocks.append(b)
                        continue
                    items.append(PSpec(it.runs, size, color, marker=mk,
                                       marker_color=lead, marker_w=mw,
                                       space_after=size * 0.35,
                                       spacing=1.04, lead_color=lead))
                if items:
                    blocks.append(self._list_blk(items, w, reveal_lists))
            elif isinstance(e, T.MathE):
                blocks.append(self._math_block(e.tex, w, size * 1.08, True,
                                               color, "center", e.tex))
            elif isinstance(e, T.TableE):
                tsize = max(12, size - 3)
                blocks.append(self._table_blk(e, w, tsize))
            elif isinstance(e, T.CodeE):
                blocks.append(self._code_blk(e, w, max(12, size - 5)))
            elif isinstance(e, T.Box):
                # a nested box (rare): flatten its body
                blocks += self.elements_to_blocks(e.body, w, size, style,
                                                  reveal_lists, color)
        for b in blocks:
            b.extra.setdefault("gap", gap)
        return blocks

    def _list_blk(self, items, w, per_para):
        h = paras_height(items, w)
        return Blk("text", h, w, items, per_para=per_para)

    def _math_block(self, tex, w, size, display, color, align, alt):
        colname = {INK: "ink", WHITE: "white"}.get(color, "ink")
        svg, png, iw, ih = self.math.get(tex, w, size, colname, align,
                                         display)
        return Blk("image", ih, iw, svg=svg, png=png, align=align,
                   size=size, alt=alt)

    def _table_blk(self, t, w, size):
        widths, heights, sz = self.table_geometry(t, w, size)
        return Blk("table", sum(heights), sum(widths), table=t, size=sz,
                   extra={"widths": widths, "heights": heights})

    def _code_blk(self, c, w, size):
        lines = c.code.split("\n")
        size = self.code_size(lines, w - 0.9, 30, size, 11)
        h = 0.36 + 0.24 + len(lines) * M.LINE * size / 72.0
        return Blk("code", h, w, code=c, size=size)

    def flow_height(self, blocks) -> float:
        h = 0.0
        for i, b in enumerate(blocks):
            h += b.h
            if i < len(blocks) - 1:
                h += b.extra.get("gap", 0.12)
        return h

    def draw_flow(self, s, blocks, x, y, w, name="body"):
        """Draw blocks top-down from (x, y). Returns [(block, [shapes])]."""
        placed = []
        for i, b in enumerate(blocks):
            shapes = []
            if b.kind == "text":
                tb = self.write(s, x, y, w, b.h + 0.05, b.paras,
                                name=f"{name}-{i}")
                shapes.append(tb)
            elif b.kind == "image":
                if "marker" in b.extra:
                    mk, mw, col = b.extra["marker"]
                    msz = b.size
                    shapes.append(self.write(
                        s, x, y, mw, 0.4,
                        [PSpec([T.Run(mk)], msz, col, bold=True)],
                        name=f"{name}-{i}m"))
                    ix = x + mw
                elif b.align == "center":
                    ix = x + (w - b.w) / 2
                else:
                    ix = x
                if b.png:
                    shapes.append(self.picture(s, b.png, ix, y, b.w, b.h,
                                               b.svg, f"{name}-{i}",
                                               alt=b.alt))
                else:      # dry run: a placeholder of the estimated size
                    shapes.append(self.rect(s, ix, y, b.w, b.h, None,
                                            name=f"{name}-{i}"))
            elif b.kind == "table":
                shapes.append(self.draw_table(s, b.table, x, y,
                                              b.extra["widths"],
                                              b.extra["heights"], b.size,
                                              name=f"{name}-{i}"))
            elif b.kind == "code":
                shapes += self.draw_code(s, b.code, x, y, w, b.h, b.size,
                                         name=f"{name}-{i}")
            if b.callout:
                bar = self.rect(s, x - 0.22, y + 0.02, 0.06, b.h - 0.02,
                                b.callout, name=f"{name}-{i}bar")
                shapes.insert(0, bar)
            placed.append((b, shapes))
            y += b.h + b.extra.get("gap", 0.12)
        return placed

    def paginate(self, make_blocks, room, sizes):
        """Pick the largest size at which the content fits; else split.

        *make_blocks(size)* builds the blocks at a size. Returns
        (size, [pages of blocks]).
        """
        for sz in sizes:
            bl = make_blocks(sz)
            if self.flow_height(bl) <= room:
                return sz, [bl]
        sz = sizes[-1] if len(sizes) < 3 else sizes[min(2, len(sizes) - 1)]
        bl = make_blocks(sz)
        pages, cur, h = [], [], 0.0
        for b in bl:
            bh = b.h + (b.extra.get("gap", 0.12) if cur else 0)
            if b.kind == "text" and len(b.paras) > 1 and h + bh > room:
                # split a list across pages, item by item
                for p in b.paras:
                    ph = p.height(b.w) + p.space_after / 72.0
                    if cur and h + ph > room:
                        pages.append(cur)
                        cur, h = [], 0.0
                    if cur and cur[-1].kind == "text" and \
                            cur[-1].extra.get("split") is b:
                        cur[-1].paras.append(p)
                        cur[-1].h = paras_height(cur[-1].paras, b.w)
                    else:
                        nb = Blk("text", p.height(b.w), b.w, [p],
                                 per_para=b.per_para, step=b.step,
                                 extra={"split": b,
                                        "gap": b.extra.get("gap", 0.12)})
                        cur.append(nb)
                    h = self.flow_height(cur)
                continue
            if cur and h + bh > room:
                pages.append(cur)
                cur, h = [], 0.0
                bh = b.h
            cur.append(b)
            h += bh
        if cur:
            pages.append(cur)
        return sz, pages

    # -- tables ------------------------------------------------------------

    def table_geometry(self, t, w_max, size):
        ncol = len(t.widths)
        for sz in range(int(size), 10, -1):
            nat = []
            for c in range(ncol):
                cells = [r[c] for r in t.rows if c < len(r)]
                if t.header:
                    cells.append(t.header[c])
                mx = max([M.natural_width(PSpec(cc, sz, bold=(cc is
                          (t.header[c] if t.header else None))).pieces())
                          for cc in cells] + [0.3])
                nat.append(mx + 0.28)
            spec = t.widths
            if all(x == 0 for x in spec):
                widths = nat
            else:
                # paragraph columns share what the natural ones leave
                fixed = sum(n for n, s_ in zip(nat, spec) if s_ == 0)
                rel = sum(s_ for s_ in spec if s_ > 0)
                room = max(1.0, w_max - fixed)
                widths = [n if s_ == 0 else room * s_ / rel
                          for n, s_ in zip(nat, spec)]
                # paragraph columns never need more than their content
                total_para = sum(wd for wd, s_ in zip(widths, spec) if s_)
                nat_para = sum(n for n, s_ in zip(nat, spec) if s_)
                if nat_para < total_para:
                    widths = [n if s_ == 0 else max(n, wd * nat_para /
                                                    total_para)
                              for n, wd, s_ in zip(nat, widths, spec)]
            if sum(widths) > w_max:
                k = w_max / sum(widths)
                widths = [x * k for x in widths]
            heights = []
            rows = ([t.header] if t.header else []) + t.rows
            for ri, r in enumerate(rows):
                hh = 0.0
                for c in range(ncol):
                    cell = r[c] if c < len(r) else []
                    sp = PSpec(cell, sz, bold=(ri == 0 and bool(t.header)))
                    hh = max(hh, sp.height(widths[c] - 0.24))
                heights.append(hh + 0.16)
            if sum(heights) <= BODY_H - 0.2 or sz <= 12:
                return widths, heights, sz
        return widths, heights, sz

    def draw_table(self, s, t, x, y, widths, heights, size, name="table",
                   rows_slice=None):
        rows = t.rows if rows_slice is None else t.rows[rows_slice]
        all_rows = ([t.header] if t.header else []) + rows
        hs = heights if rows_slice is None else \
            ([heights[0]] if t.header else []) + \
            heights[(1 if t.header else 0):][rows_slice]
        gf = s.shapes.add_table(len(all_rows), len(widths), Inches(x),
                                Inches(y), Inches(sum(widths)),
                                Inches(sum(hs)))
        gf.name = name
        tbl = gf.table
        tblPr = tbl._tbl.tblPr
        for a in ("firstRow", "bandRow", "firstCol", "lastRow", "lastCol",
                  "bandCol"):
            tblPr.set(a, "0")
        sid = tblPr.find(qn("a:tableStyleId"))
        if sid is None:
            sid = tblPr.makeelement(qn("a:tableStyleId"), {})
            tblPr.append(sid)
        sid.text = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"
        for c, wd in enumerate(widths):
            tbl.columns[c].width = Inches(wd)
        band = {"lightbg": "F2F8FC", "defbg": "FEF7F1"}.get(t.band, "F2F8FC")
        head_fill = mix(PRIMARY, WHITE, 0.85)
        for ri, r in enumerate(all_rows):
            tbl.rows[ri].height = Inches(hs[ri])
            is_head = ri == 0 and bool(t.header)
            for c in range(len(widths)):
                cell = tbl.cell(ri, c)
                cell.margin_left = cell.margin_right = Inches(0.12)
                cell.margin_top = Inches(0.07)
                cell.margin_bottom = Inches(0.05)
                cell.vertical_anchor = MSO_ANCHOR.TOP
                fill = head_fill if is_head else (
                    band if (ri - (1 if t.header else 0)) % 2 == 1 else WHITE)
                cell.fill.solid()
                cell.fill.fore_color.rgb = rgb(fill)
                tf = cell.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                al = t.align[c] if c < len(t.align) else "l"
                p.alignment = {"l": PP_ALIGN.LEFT, "r": PP_ALIGN.RIGHT,
                               "c": PP_ALIGN.CENTER}[al]
                runs = r[c] if c < len(r) else []
                style_runs(p, runs, size, PRIMARY if is_head else INK,
                           bold=is_head)
                # an empty cell still has a paragraph mark, at the
                # template's 18pt unless told otherwise -- which pads every
                # row of the table to that height
                end = p._p.get_or_add_endParaRPr()
                end.set("sz", str(int(size * 100)))
                self._cell_borders(cell, top=(ri == 0),
                                   bottom=(is_head or ri == len(all_rows) - 1),
                                   heavy=(ri == 0 or ri == len(all_rows) - 1))
        return gf

    def _cell_borders(self, cell, top=False, bottom=False, heavy=False):
        """Booktabs rules: a line above and below the header and at the end,
        nothing vertical -- the handout's table style."""
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for old in tcPr.findall(qn(tag)):
                tcPr.remove(old)

        def ln(tag, on, w_pt, col):
            el = tcPr.makeelement(qn(tag), {"w": str(int(w_pt * 12700)),
                                            "cmpd": "sng"})
            if on:
                sf = el.makeelement(qn("a:solidFill"), {})
                c = sf.makeelement(qn("a:srgbClr"), {"val": col})
                sf.append(c)
                el.append(sf)
            else:
                el.append(el.makeelement(qn("a:noFill"), {}))
            return el
        # schema order: lnL, lnR, lnT, lnB, ..., solidFill
        els = [ln("a:lnL", False, 0, PRIMARY), ln("a:lnR", False, 0, PRIMARY),
               ln("a:lnT", top, 1.5 if heavy else 0.75, PRIMARY),
               ln("a:lnB", bottom, 1.5 if heavy else 0.75, PRIMARY)]
        for i, el in enumerate(els):
            tcPr.insert(i, el)

    # -- code ----------------------------------------------------------------

    CODE_COL = {"kw": "5DADE2", "str": "7DCEA0", "num": "F7DC6F",
                "com": "8795A1", "fn": "C39BD3", "cls": "76D7C4",
                "op": "AEB6BF", "txt": "E8EDF2", "dec": "F0B27A"}

    def code_size(self, lines, w, max_lines, hi, lo):
        longest = max((len(l) for l in lines), default=1)
        for sz in range(int(hi), lo - 1, -1):
            if M.width("M" * longest, sz, MONO) <= w and \
                    len(lines) * M.LINE * sz / 72.0 <= BODY_H - 0.8:
                return sz
        return lo

    def _tokens(self, code, lang):
        from pygments import lex
        from pygments.lexers import (BashLexer, PythonLexer, SLexer,
                                     TextLexer)
        from pygments.token import Token
        lexer = {"python": PythonLexer(), "r": SLexer(),
                 "bash": BashLexer(), "sh": BashLexer()}.get(lang,
                                                             TextLexer())
        out = []
        for tok, val in lex(code, lexer):
            if tok in Token.Comment:
                k = "com"
            elif tok in Token.Keyword:
                k = "kw"
            elif tok in Token.Literal.String:
                k = "str"
            elif tok in Token.Literal.Number:
                k = "num"
            elif tok in Token.Name.Function or tok in Token.Name.Builtin:
                k = "fn"
            elif tok in Token.Name.Class or tok in Token.Name.Namespace:
                k = "cls"
            elif tok in Token.Name.Decorator:
                k = "dec"
            elif tok in Token.Operator or tok in Token.Punctuation:
                k = "op"
            else:
                k = "txt"
            out.append((k, val))
        return out

    def draw_code(self, s, c, x, y, w, h, size, name="code", label=None,
                  first_line=1, dim_before=0):
        """A dark editor panel with highlighted code and line numbers."""
        shapes = []
        bar = 0.36
        shapes.append(self.rrect(s, x, y, w, h, CODEBG, name=f"!!{name}-bg",
                                 r_in=0.10))
        shapes.append(self.rect(s, x + 0.02, y + bar - 0.01, w - 0.04, 0.012,
                                CODEBAR, name=f"!!{name}-sep"))
        for k, col in enumerate(("E74C3C", "F4D03F", "58D68D")):
            shapes.append(self.rect(s, x + 0.22 + k * 0.2, y + 0.125, 0.11,
                                    0.11, col, shape=MSO_SHAPE.OVAL,
                                    name=f"!!{name}-dot{k}"))
        lang = {"python": "Python", "r": "R", "bash": "Shell",
                "sh": "Shell"}.get(c.lang, "Text")
        shapes.append(self.text(s, x + 0.95, y + 0.07, w - 1.2, 0.24,
                                label or lang, 11, "9FB0BF", True,
                                name=f"!!{name}-label"))
        lines = c.code.split("\n")
        toks = self._tokens(c.code, c.lang)
        # split tokens into lines
        per_line = [[]]
        for k, val in toks:
            parts = val.split("\n")
            for j, part in enumerate(parts):
                if j > 0:
                    per_line.append([])
                if part:
                    per_line[-1].append((k, part))
        per_line = per_line[:len(lines)]
        num_w = 0.55
        body_y = y + bar + 0.12
        tb = self.box(s, x + 0.22 + num_w, body_y, w - num_w - 0.42,
                      h - bar - 0.18, name=f"!!{name}-text", wrap=False)
        tf = tb.text_frame
        nb = self.box(s, x + 0.12, body_y, num_w - 0.12, h - bar - 0.18,
                      name=f"!!{name}-num", wrap=False)
        for i, toks_l in enumerate(per_line):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.line_spacing = 1.0
            p.space_after = Pt(0)
            dim = i < dim_before
            if not toks_l:
                r = p.add_run()
                r.text = " "
                r.font.size = Pt(size)
                r.font.name = MONO
            for k, val in toks_l:
                r = p.add_run()
                r.text = val
                r.font.size = Pt(size)
                r.font.name = MONO
                col = self.CODE_COL[k]
                r.font.color.rgb = rgb(mix(col, CODEBG, 0.62) if dim
                                       else col)
                r.font.italic = k == "com"
            q = nb.text_frame.paragraphs[0] if i == 0 else \
                nb.text_frame.add_paragraph()
            q.alignment = PP_ALIGN.RIGHT
            q.line_spacing = 1.0
            rr = q.add_run()
            rr.text = str(first_line + i)
            rr.font.size = Pt(size)
            rr.font.name = MONO
            rr.font.color.rgb = rgb("4F5D6A")
        FX.alt_text(tb, c.code[:1400])
        shapes += [tb, nb]
        return shapes

    # =======================================================================
    # Slide types
    # =======================================================================

    def title_slide(self, part_title: str, week=None):
        ch = self.ch
        s = self.new("title", self.ground, progress=False)
        # the chapter number, set very large as a watermark -- behind the
        # title placeholder, which the layout put at the bottom of the stack
        wm = self.text(s, W - 6.3, -0.55, 6.0, 5.0, f"{ch.number:02d}", 280,
                       mix(self.ground, WHITE, 0.10),
                       family="Segoe UI Black", align="r", name="!!wm")
        FX.decorative(wm)
        FX.to_back(wm)
        self.rect(s, ML, 1.42, 0.9, 0.07, WHITE, name="!!tbar")
        self.text(s, ML, 1.66, 9.5, 0.32,
                  f"PART {_roman(ch.part)}  ·  {part_title.upper()}",
                  13, mix(self.ground, WHITE, 0.72), True, spc=180,
                  name="!!tpart")
        self.text(s, ML, 2.04, 9.5, 0.32,
                  f"CHAPTER {ch.number}" + (f"  ·  WEEK {week.week}"
                                            if week else ""),
                  13, mix(self.ground, WHITE, 0.72), True, spc=180,
                  name="!!tch")
        self.title(s, ch.title, color=WHITE, x=ML, y=2.52, w=9.3, h=2.1,
                   size=self.fit_size([T.Run(ch.title)], 9.3, 2.1, 46, 32,
                                      HEAD))
        self.text(s, ML, 4.72, 9.3, 0.36,
                  "Data Science — From Foundations to Intelligent Systems",
                  16, mix(self.ground, WHITE, 0.80), name="!!tsub")
        # bottom band: the person on the left, the institution on the right
        self.rect(s, 0, 5.78, W, H - 5.78, self.deep, name="!!band")
        self.text(s, ML, 6.06, 5.0, 0.26, "LECTURER", 10.5,
                  mix(self.deep, WHITE, 0.62), True, spc=160,
                  name="!!lab")
        self.text(s, ML, 6.32, 6.0, 0.4, self.author, 19, WHITE, True,
                  family=HEAD, name="!!author")
        self.text(s, ML, 6.76, 6.4, 0.3, f"{self.dept}  ·  {self.edition}",
                  12, mix(self.deep, WHITE, 0.72), name="!!dept")
        self._logos(s, W - MR, 6.0, 0.98)
        FX.transition(s, "fade")
        self.notes(s, f"Chapter {ch.number}: {ch.title}. "
                      f"Part {ch.part} — {part_title}."
                      + (f" Week {week.week}: {week.topic}." if week else "")
                      + f" Handout pages {ch.page}–{ch.last_page}.")
        return s

    def _logos(self, s, right, y, h):
        names = [("iub-logo-white.png", "The Islamia University of "
                  "Bahawalpur"),
                 ("dit-logo.png", "Department of Information Technology")]
        found = [(os.path.join(ASSETS, n), a) for n, a in names
                 if os.path.exists(os.path.join(ASSETS, n))]
        if not found:
            self.text(s, W - MR - 5.0, y + 0.3, 5.0, 0.3, self.uni.upper(),
                      12, WHITE, True, align="r", name="!!uni")
            return
        from PIL import Image
        ws = []
        for p, _ in found:
            with Image.open(p) as im:
                ws.append(h * im.size[0] / im.size[1])
        gap = 0.28
        x = right - sum(ws) - gap * (len(ws) - 1)
        uni_w = 4.4
        self.text(s, x - uni_w - 0.3, y + 0.22, uni_w, 0.3,
                  self.uni.upper(), 11.5, WHITE, True, align="r", spc=60,
                  name="!!uni")
        self.text(s, x - uni_w - 0.3, y + 0.52, uni_w, 0.3,
                  "Bahawalpur, Pakistan", 10.5, mix(self.deep, WHITE, 0.62),
                  align="r", name="!!uni2")
        for (p, alt), wd in zip(found, ws):
            pic = self.picture(s, p, x, y, wd, h, name="!!logo-" +
                               os.path.basename(p)[:-4], alt=alt)
            x += wd + gap

    # -- where this chapter sits ---------------------------------------------

    def roadmap_slide(self, chapters, builds_on, leads_to, prereq_elems):
        """All 32 chapters as a track, this one lit; then its dependencies.

        *chapters*: [(num, title, part)]. *builds_on* / *leads_to*: lists of
        chapter numbers, both derived from the handout's prerequisite boxes.
        """
        ch = self.ch
        s = self.new("roadmap")
        self.chrome(s, "Where this chapter sits", "The course",
                    STYLES["course"], ref=f"Handout p. {ch.page}")
        # the track
        n = len(chapters)
        parts = sorted({p for _, _, p in chapters})
        gap_part = 0.30
        # the current chapter's dot is twice the size of the others, so it
        # gets room of its own rather than sitting on its neighbours
        halo = 0.17
        step = (CW - 0.2 - gap_part * (len(parts) - 1) - 2 * halo) / n
        y0 = BODY_Y + 0.72
        x = ML + 0.1
        xs = {}
        last_part = None
        for num, t, p in chapters:
            if last_part is not None and p != last_part:
                x += gap_part
            if num == ch.number:
                x += halo
            xs[num] = (x + step / 2, p)
            last_part = p
            x += step
            if num == ch.number:
                x += halo
        # part labels and underlines
        for p in parts:
            px = [xs[c][0] for c, _, pp in chapters if pp == p]
            lo, hi = min(px) - step / 2 + 0.03, max(px) + step / 2 - 0.03
            col = ground(PART_HUE[p])
            live = p == ch.part
            self.text(s, lo, BODY_Y + 0.02, hi - lo, 0.26,
                      f"PART {_roman(p)}", 11, col if live else FAINT, True,
                      align="c", spc=120, name=f"!!rp{p}")
            self.rect(s, lo, BODY_Y + 0.32, hi - lo, 0.05 if live else 0.03,
                      col if live else RULE, name=f"!!rpl{p}")
        # dots
        for num, t, p in chapters:
            cx, _ = xs[num]
            col = PART_HUE[p]
            if num == ch.number:
                d = 0.62
                self.rect(s, cx - d / 2, y0 - d / 2, d, d, ground(col),
                          shape=MSO_SHAPE.OVAL, name=f"!!dot{num}")
                self.text(s, cx - d / 2, y0 - 0.17, d, 0.34, str(num), 15,
                          WHITE, True, align="c", family=HEAD,
                          name=f"!!dotn{num}")
            else:
                d = 0.30
                past = num < ch.number
                self.rect(s, cx - d / 2, y0 - d / 2, d, d,
                          mix(col, WHITE, 0.15) if past else WHITE,
                          None if past else mix(col, WHITE, 0.45),
                          shape=MSO_SHAPE.OVAL, lw=1.25,
                          name=f"!!dot{num}")
                self.text(s, cx - 0.25, y0 - 0.11, 0.5, 0.22, str(num), 9.5,
                          WHITE if past else mix(col, INK, 0.2),
                          True, align="c", name=f"!!dotn{num}")
        # the chapter, named under its dot
        cx = xs[ch.number][0]
        lw = 5.0
        lx = min(max(ML, cx - lw / 2), ML + CW - lw)
        self.text(s, lx, y0 + 0.44, lw, 0.34, ch.title, 15,
                  self.accent_txt, True, align="c", name="!!dotlabel")
        # dependencies
        yy = y0 + 1.12
        col_w = (CW - 0.5) / 2

        def chips(label, nums, x0):
            self.text(s, x0, yy, col_w, 0.28, label, 12, MUTED, True,
                      spc=150, name=f"!!dep-{label[:5]}")
            cy = yy + 0.40
            if not nums:
                self.text(s, x0, cy, col_w, 0.3,
                          "—  nothing earlier in the course" if
                          label.startswith("BUILDS") else
                          "—  the course ends here", 14, FAINT,
                          name=f"!!dep-none-{label[:5]}")
                return
            titles = {c: (t, p) for c, t, p in chapters}
            for c in nums[:5]:
                t, p = titles.get(c, ("", 1))
                colr = ground(PART_HUE[p])
                self.rrect(s, x0, cy, col_w, 0.46, mix(colr, WHITE, 0.90),
                           name=f"!!chip{label[:2]}{c}", r_in=0.08)
                self.rect(s, x0, cy, 0.07, 0.46, colr,
                          name=f"!!chipb{label[:2]}{c}")
                self.write(s, x0 + 0.22, cy + 0.1, col_w - 0.3, 0.3,
                           [PSpec([T.Run(f"Ch {c}  ", bold=True,
                                         color=None),
                                   T.Run(t)], 14, INK,
                                  lead_color=colr)],
                           name=f"!!chipt{label[:2]}{c}")
                cy += 0.56
        chips("BUILDS ON", builds_on, ML)
        chips("USED LATER IN", leads_to, ML + col_w + 0.5)
        FX.transition(s, "fade")
        txt = " ".join(e.text for e in prereq_elems
                       if isinstance(e, T.Para))
        self.notes(s, "Place the chapter. " + (f"Before you start: {txt}"
                                               if txt else ""))
        return s

    # -- outcomes ------------------------------------------------------------

    def outcomes_slide(self, lead, items, style=STYLES["outcomes"]):
        s = self.new("outcomes")
        self.chrome(s, "What you will be able to do", style.label, style,
                    ref=f"Handout p. {self.ch.page}" if self.ch else "")
        y = BODY_Y
        if lead:
            self.write(s, ML, y, CW, 0.36, [PSpec(lead, 16, MUTED)],
                       name="!!olead")
            y += 0.52
        room = BODY_B - y
        badge = 0.46
        tx = ML + badge + 0.26
        tw = CW - (tx - ML)
        for sz in (21, 20, 19, 18, 17, 16, 15):
            hs = [max(badge, PSpec(it.runs, sz, spacing=1.04).height(tw))
                  for it in items]
            gap = 0.20
            if sum(hs) + gap * (len(hs) - 1) <= room:
                break
        steps = []
        for k, (it, hh) in enumerate(zip(items, hs)):
            b = self.rect(s, ML, y + 0.01, badge, badge, self.ground,
                          shape=MSO_SHAPE.OVAL, name=f"!!ob{k}")
            n = self.text(s, ML, y + 0.075, badge, 0.3, str(k + 1), 15,
                          WHITE, True, align="c", family=HEAD,
                          name=f"!!obn{k}")
            t = self.write(s, tx, y + max(0, (badge - M.LINE * sz / 72) / 2)
                           if hh <= badge + 0.01 else y, tw, hh,
                           [PSpec(it.runs, sz, INK, spacing=1.04,
                                  lead_color=self.accent_txt)],
                           name=f"!!ot{k}")
            steps.append([FX.fx(b, "zoom", 300), FX.fx(n, "fade", 300),
                          FX.fx(t, "wipeRight", 420)])
            y += hh + gap
        FX.transition(s, "fade")
        FX.animate(s, steps)
        self.notes(s, "Read these out: they are the chapter's contract with "
                      "the room. Each click adds one.")
        return s

    # -- key terms -----------------------------------------------------------

    def terms_slide(self, terms, style=STYLES["keyterms"]):
        s = self.new("terms")
        self.chrome(s, "Vocabulary for this chapter", style.label, style,
                    ref=f"Handout p. {self.ch.page}" if self.ch else "")
        size = 17 if len(terms) <= 12 else (16 if len(terms) <= 16 else 15)
        pad_x, ch_h, gap = 0.2, 0.52, 0.16
        rows = [[]]
        x = 0.0
        for t in terms:
            w = M.width(t, size) + 2 * pad_x + 0.08
            if x + w > CW and rows[-1]:
                rows.append([])
                x = 0.0
            rows[-1].append((t, w))
            x += w + gap
        y = BODY_Y + 0.12
        shapes = []
        for row in rows:
            x = ML
            for t, w in row:
                c = self.rrect(s, x, y, w, ch_h, mix(style.rule, WHITE, 0.92),
                               name=f"!!term-{t[:14]}", r_in=0.26)
                c2 = self.rect(s, x + 0.14, y + ch_h / 2 - 0.05, 0.1, 0.1,
                               style.rule, shape=MSO_SHAPE.OVAL,
                               name=f"!!termd-{t[:14]}")
                tt = self.text(s, x + pad_x + 0.12, y + (ch_h - M.LINE *
                               size / 72) / 2, w - pad_x, 0.34, t, size,
                               INK, name=f"!!termt-{t[:14]}", wrap=False)
                shapes += [c, c2, tt]
                x += w + gap
            y += ch_h + gap
        self.text(s, ML, min(BODY_B - 0.3, y + 0.22), CW, 0.3,
                  "Each term is defined where it first matters — not here.",
                  13.5, MUTED, italic=True, name="!!termnote")
        FX.transition(s, "fade")
        self.notes(s, "Do not define these now; they arrive one at a time. "
                      "This slide is for the students' notes. Terms: "
                      + "; ".join(terms) + ".")
        return s

    # -- agenda --------------------------------------------------------------

    def agenda_slide(self, sections):
        """*sections*: [(key, num, title, page)] -> clickable rows."""
        s = self.new("agenda")
        self.chrome(s, "In this chapter", STYLES["agenda"].label,
                    STYLES["agenda"],
                    ref=f"Handout pp. {self.ch.page}–{self.ch.last_page}")
        n = len(sections)
        row = min(0.62, (BODY_H - 0.1) / max(n, 1) - 0.08)
        size = 18 if row >= 0.55 else 16
        y = BODY_Y + 0.04
        for i, (key, num, t, page) in enumerate(sections):
            r = self.rrect(s, ML, y, CW, row, "F4F7F9", name=f"!!ag{i}",
                           r_in=0.06, deco=False)
            FX.alt_text(r, f"Go to section {num} {t}")
            self.rect(s, ML, y, 0.07, row, self.hue, name=f"!!agb{i}")
            ty = y + (row - M.LINE * size / 72) / 2
            self.text(s, ML + 0.3, ty, 1.0, 0.36, num, size,
                      self.accent_txt, True, family=HEAD, name=f"!!agn{i}")
            self.text(s, ML + 1.3, ty, CW - 3.0, 0.36, t, size, INK,
                      name=f"!!agt{i}")
            self.text(s, ML + CW - 1.9, ty + 0.03, 1.7, 0.36,
                      f"p. {page}  ›", size - 3, MUTED, align="r",
                      name=f"!!agp{i}")
            self.links.append((r, key))
            y += row + 0.08
        FX.transition(s, "fade")
        self.notes(s, "The route through the chapter. Each row is a link: "
                      "click it to jump to that section in the slide show.")
        return s

    # -- section divider -----------------------------------------------------

    def divider_slide(self, num, title, idx, sections, key):
        s = self.new("divider")
        panel = 4.55
        self.rect(s, 0, 0, panel, H, self.ground, name="!!dpanel")
        self.text(s, 0.72, 2.10, panel - 1.0, 0.3,
                  f"SECTION {idx} OF {len(sections)}", 12.5,
                  mix(self.ground, WHITE, 0.70), True, spc=180,
                  name="!!dsec")
        self.text(s, 0.68, 2.46, panel - 0.9, 1.6, num, 76, WHITE,
                  family=HEAD, name="!!dnum")
        self.title(s, title, x=panel + 0.75, y=2.05, w=W - panel - 1.5,
                   h=2.0, size=self.fit_size([T.Run(title)],
                                             W - panel - 1.5, 2.0, 38, 26,
                                             HEAD),
                   color=INK, anchor=MSO_ANCHOR.MIDDLE)
        # the other sections, small, so the room sees the route
        y = 4.55
        for i, (k, n, t, p) in enumerate(sections):
            live = n == num
            self.text(s, panel + 0.75, y, W - panel - 1.5, 0.28,
                      f"{n}   {t}", 12.5 if not live else 13,
                      self.accent_txt if live else FAINT, live,
                      name=f"!!dl{i}")
            y += 0.30
            if y > 6.85:
                break
        # footer on the white side only; the panel carries the number
        self.text(s, panel + 0.75, FOOT_Y, 5.0, 0.26,
                  f"{self.course}  ·  Chapter {self.ch.number}", 10, FAINT,
                  name="!!footL")
        tb = self.box(s, ML + CW * 0.55, FOOT_Y, CW * 0.45, 0.26,
                      name="!!footR")
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        r.text = f"Handout §{num}     "
        r.font.size = Pt(10)
        r.font.name = BODY
        r.font.color.rgb = rgb(FAINT)
        FX.slide_number_field(p, 10, "8795A1", BODY, str(len(self.recs)))
        FX.transition(s, "push")
        self.notes(s, "")
        self.anchors[key] = s
        return s

    # -- cards ---------------------------------------------------------------

    def card_slides(self, style: Style, title: str, elements, notes="",
                    ref="", reveal=False, kicker=None):
        """A tinted card holding a box's content. Splits when it must."""
        pad = 0.34
        iw = CW - 2 * pad - 0.1

        def mk(sz):
            return self.elements_to_blocks(elements, iw, sz, style,
                                           reveal_lists=reveal)
        size, pages = self.paginate(mk, BODY_H - 2 * pad - 0.05,
                                    [23, 22, 21, 20, 19, 18, 17, 16])
        out = []
        for pi, blocks in enumerate(pages):
            s = self.new("card")
            ttl = title if pi == 0 else f"{title} (continued)"
            self.chrome(s, ttl, kicker or style.label, style, ref)
            h = self.flow_height(blocks) + 2 * pad
            card = self.rrect(s, ML, BODY_Y, CW, h, style.bg,
                              name="!!card", r_in=0.10)
            FX.shadow(card, 10, 2, 10)
            self.rect(s, ML, BODY_Y, 0.09, h, style.rule, name="!!cardrule")
            placed = self.draw_flow(s, blocks, ML + pad + 0.1, BODY_Y + pad,
                                    iw)
            steps = []
            if reveal:
                for b, shapes in placed:
                    if b.per_para and shapes:
                        for k in range(len(b.paras)):
                            steps.append([FX.fx(shapes[0], "fade", 380,
                                                para=k)])
            FX.transition(s, "fade")
            FX.animate(s, steps)
            self.notes(s, notes if pi == 0 else "")
            out.append(s)
        return out

    # -- worked example ------------------------------------------------------

    def worked_slides(self, title, elements, notes="", ref=""):
        """The problem, then each step on its own click."""
        style = STYLES["worked"]
        pad = 0.30
        iw = CW - 2 * pad - 0.1

        def mk(sz):
            bl = self.elements_to_blocks(elements, iw, sz, style,
                                         reveal_lists=False)
            # the conclusion paragraphs: bold-led prose after the working
            seen_math = False
            for b in bl:
                if b.kind in ("image", "table") or (
                        b.kind == "text" and len(b.paras) > 1):
                    seen_math = True
                elif b.kind == "text" and seen_math and b.paras and \
                        b.paras[0].runs and b.paras[0].runs[0].bold:
                    b.callout = style.rule
            return bl
        size, pages = self.paginate(mk, BODY_H - 2 * pad,
                                    [21, 20, 19, 18, 17, 16])
        out = []
        for pi, blocks in enumerate(pages):
            s = self.new("worked")
            ttl = title if pi == 0 else f"{title} (continued)"
            self.chrome(s, ttl, style.label, style, ref)
            h = self.flow_height(blocks) + 2 * pad
            card = self.rrect(s, ML, BODY_Y, CW, h, style.bg,
                              name="!!card", r_in=0.10)
            FX.shadow(card, 10, 2, 10)
            self.rect(s, ML, BODY_Y, 0.09, h, style.rule, name="!!cardrule")
            placed = self.draw_flow(s, blocks, ML + pad + 0.12,
                                    BODY_Y + pad, iw)
            steps = []
            # the set-up (first block on the first page) is on screen at once
            start = 1 if pi == 0 else 0
            if pi == 0 and len(placed) > 1 and placed[1][0].kind == "text" \
                    and len(placed[1][0].paras) > 1 and \
                    placed[1][0].paras[0].marker:
                start = 2          # set-up plus its list of givens
            for b, shapes in placed[start:]:
                if not shapes:
                    continue
                if b.per_para or (b.kind == "text" and len(b.paras) > 1
                                  and b.paras[0].marker):
                    for k in range(len(b.paras)):
                        steps.append([FX.fx(shapes[-1] if b.callout
                                            else shapes[0], "fade", 350,
                                            para=k)])
                    if b.callout:
                        steps[-len(b.paras)].insert(
                            0, FX.fx(shapes[0], "wipeDown", 300))
                else:
                    steps.append([FX.fx(sh, "fade" if k else "wipeRight",
                                        420) for k, sh in enumerate(shapes)])
            FX.transition(s, "fade")
            FX.animate(s, steps)
            self.notes(s, (notes + "\n\n" if notes and pi == 0 else "") +
                       "Each click adds one step of the working. Ask for "
                       "the next number before you reveal it.")
            out.append(s)
        return out

    # -- figures -------------------------------------------------------------

    def figure_slide(self, fig: T.FigureE, number: str, page: int,
                     section: str, notes="", ref="", step=None, morph=False):
        cap = T.plain(fig.caption)
        title, rest = split_caption(cap, section)
        png = os.path.join(FIGS, (step or fig.label) + ".png")
        svg = os.path.join(FIGS, (step or fig.label) + ".svg")
        s = self.new("figure")
        self.chrome(s, title, f"Figure {number}" if number else "Figure",
                    None, ref, kicker_color=self.accent_txt)
        take_h = 0.0
        take_sz = 17
        if rest:
            take_sz = self.fit_size([T.Run(rest)], CW - 0.3, 0.9, 17, 14)
            take_h = PSpec([T.Run(rest)], take_sz, spacing=1.04).height(
                CW - 0.3) + 0.24
        area_h = BODY_H - take_h - 0.08
        if os.path.exists(png):
            from PIL import Image
            with Image.open(png) as im:
                ar = im.size[0] / im.size[1]
            w = min(CW, area_h * ar)
            h = w / ar
            x = ML + (CW - w) / 2
            self.picture(s, png, x, BODY_Y + (area_h - h) / 2 * 0.35, w, h,
                         svg, name="!!figure", alt=f"Figure {number}. {cap}")
            fig_bottom = BODY_Y + (area_h - h) / 2 * 0.35 + h
        else:
            self.rect(s, ML, BODY_Y, CW, area_h, LIGHTBG, name="!!figure")
            self.text(s, ML, BODY_Y + area_h / 2 - 0.2, CW, 0.4,
                      f"[figure not built: {fig.label}]", 16, ACCENT,
                      align="c", name="!!figmiss")
            fig_bottom = BODY_Y + area_h
        if rest:
            ty = max(fig_bottom + 0.14, BODY_B - take_h + 0.1)
            ty = min(ty, BODY_B - take_h + 0.14)
            self.rect(s, ML, ty + 0.02, 0.06, take_h - 0.2, self.hue,
                      name="!!takebar")
            self.write(s, ML + 0.24, ty, CW - 0.3, take_h,
                       [PSpec([T.Run(rest)], take_sz, "3E4C57",
                              spacing=1.04)], name="!!take")
        FX.transition(s, "morph" if morph else "fade")
        self.notes(s, "\n\n".join(x for x in (
            notes, f"Figure {number} (handout p. {page}): {cap}") if x))
        return s

    # -- tables --------------------------------------------------------------

    def table_slides(self, t: T.TableE, title, notes="", ref="",
                     kicker=None):
        widths, heights, size = self.table_geometry(t, CW, 17)
        room = BODY_H - 0.1
        head_h = heights[0] if t.header else 0.0
        body_h = heights[1:] if t.header else heights
        pages, cur, h = [], [], head_h
        for i, rh in enumerate(body_h):
            if cur and h + rh > room:
                pages.append(cur)
                cur, h = [], head_h
            cur.append(i)
            h += rh
        if cur:
            pages.append(cur)
        out = []
        for pi, idx in enumerate(pages):
            s = self.new("table")
            self.chrome(s, title if pi == 0 else f"{title} (continued)",
                        kicker or STYLES["table"].label, STYLES["table"], ref)
            sl = slice(idx[0], idx[-1] + 1)
            x = ML
            self.draw_table(s, t, x, BODY_Y, widths, heights, size,
                            name="!!table", rows_slice=sl)
            FX.transition(s, "fade")
            self.notes(s, notes if pi == 0 else "")
            out.append(s)
        return out

    # -- statement and points --------------------------------------------

    def statement_slide(self, title, elements, notes="", ref=""):
        st = STYLES["statement"]
        iw = CW - 0.5

        def mk(sz):
            return self.elements_to_blocks(elements, iw, sz, st)
        size, pages = self.paginate(mk, BODY_H - 0.2,
                                    [26, 25, 24, 23, 22, 21, 20, 19, 18])
        out = []
        for pi, blocks in enumerate(pages):
            s = self.new("statement")
            self.chrome(s, title if pi == 0 else f"{title} (continued)",
                        st.label, st, ref)
            h = self.flow_height(blocks)
            self.rect(s, ML, BODY_Y + 0.08, 0.08, h - 0.04, self.hue,
                      name="!!stbar")
            placed = self.draw_flow(s, blocks, ML + 0.42, BODY_Y + 0.06, iw)
            steps = [[FX.fx(sh, "fade", 450) for sh in shapes]
                     for b, shapes in placed[1:]]
            FX.transition(s, "fade")
            FX.animate(s, steps)
            self.notes(s, notes)
            out.append(s)
        return out

    def points_slides(self, title, elements, notes="", ref="", style=None):
        return self.card_slides(style or STYLES["points"], title, elements,
                                notes, ref, reveal=True)

    # -- pitfalls ------------------------------------------------------------

    def pitfall_slides(self, title, items, lead=None, notes="", ref=""):
        st = STYLES["pitfall"]
        badge = 0.40
        tx = ML + badge + 0.26
        tw = CW - (tx - ML)
        y0 = BODY_Y + (0.5 if lead else 0.0)
        room = BODY_B - y0
        for sz in (21, 20, 19, 18, 17):
            hs = [max(badge, PSpec(it.runs, sz, spacing=1.04).height(tw))
                  for it in items]
            if sum(hs) + 0.22 * (len(hs) - 1) <= room:
                break
        pages, cur, h = [], [], 0.0
        for it, hh in zip(items, hs):
            if cur and h + hh > room:
                pages.append(cur)
                cur, h = [], 0.0
            cur.append((it, hh))
            h += hh + 0.22
        if cur:
            pages.append(cur)
        out = []
        for pi, page in enumerate(pages):
            s = self.new("pitfall")
            self.chrome(s, title if pi == 0 else f"{title} (continued)",
                        st.label, st, ref)
            y = BODY_Y
            if lead and pi == 0:
                self.write(s, ML, y, CW, 0.4, [PSpec(lead.runs, 16, MUTED)],
                           name="!!plead")
                y += 0.5
            steps = []
            for k, (it, hh) in enumerate(page):
                b = self.rect(s, ML, y + 0.02, badge, badge,
                              mix(ACCENT, WHITE, 0.86), shape=MSO_SHAPE.OVAL,
                              name=f"!!pb{k}")
                x_ = self.text(s, ML, y + 0.045, badge, 0.3, "✕", 15,
                               ACCENT, True, align="c", name=f"!!px{k}")
                t = self.write(s, tx, y, tw, hh,
                               [PSpec(it.runs, sz, INK, spacing=1.04,
                                      lead_color="A93226")],
                               name=f"!!pt{k}")
                steps.append([FX.fx(b, "zoom", 260), FX.fx(x_, "fade", 260),
                              FX.fx(t, "wipeRight", 400)])
                y += hh + 0.22
            FX.transition(s, "fade")
            FX.animate(s, steps)
            self.notes(s, notes if pi == 0 else "")
            out.append(s)
        return out

    # -- four lenses ---------------------------------------------------------

    def lens_slides(self, topic, lenses: dict, notes=""):
        """Overview grid, then each lens Morph-expanded to its full text."""
        st = STYLES["fourlenses"]
        out = []
        gx, gy = 0.26, 0.24
        cw2 = (CW - gx) / 2
        ch2 = (BODY_H - gy) / 2
        head = 0.54

        def lead_of(elems):
            """The lens's first sentence, styling kept."""
            runs = []
            for e in elems:
                if isinstance(e, T.Para):
                    runs = e.runs
                    break
            txt = T.plain(runs)
            m = re.search(r"[.!?](\s|$)", txt)
            end = m.end() if m else len(txt)
            # a sentence that ends inside "e.g." is not a sentence end
            while m and txt[:m.start() + 1].endswith(("e.g.", "i.e.")):
                m = re.search(r"[.!?](\s|$)", txt[end:])
                end = end + m.end() if m else len(txt)
            return _slice_runs(runs, 0, end)

        def card(s, key, x, y, w, h, body, full, active):
            col, name, ic = LENS[key]
            c = self.rrect(s, x, y, w, h, mix(col, WHITE, 0.94),
                           name=f"!!lens-{key}", r_in=0.10)
            if active:
                FX.shadow(c, 14, 3, 14)
            self.rrect(s, x, y, w, head, col, name=f"!!lensh-{key}",
                       r_in=0.10)
            self.rect(s, x, y + head - 0.12, w, 0.12, col,
                      name=f"!!lensh2-{key}")
            ipic = self.icon(s, ic, WHITE, x + 0.22, y + 0.13, 0.28,
                             nm=f"!!lensi-{key}")
            tx = x + 0.22 + ipic.width / 914400 + 0.16
            label = name.upper() if w > 4 else name
            self.text(s, tx, y + 0.11, x + w - tx - 0.1, 0.32, label,
                      13.5 if w > 4 else 13, WHITE, True,
                      spc=100 if w > 4 else 0, name=f"!!lensn-{key}",
                      wrap=False)
            if body is None:
                return
            bw = w - 0.5
            if full:
                sz = self.fit_size_elems(body, bw, h - head - 0.4, 22, 14)
                blocks = self.elements_to_blocks(body, bw, sz, None,
                                                 color=INK)
                for b in blocks:
                    for p in b.paras:
                        p.lead_color = col
                self.draw_flow(s, blocks, x + 0.26, y + head + 0.22, bw,
                               name=f"!!lensb-{key}")
            else:
                runs = lead_of(body)
                sz = self.fit_size(runs, bw, h - head - 0.4, 23, 14)
                self.write(s, x + 0.26, y + head + 0.2, bw, h - head - 0.3,
                           [PSpec(runs, sz, INK, spacing=1.04,
                                  lead_color=col)],
                           name=f"!!lensb-{key}-0")

        # overview
        s = self.new("lenses")
        self.chrome(s, f"Four lenses: {topic}", st.label, st,
                    ref=self._ref_for_lenses())
        for i, key in enumerate(LENS_ORDER):
            x = ML + (i % 2) * (cw2 + gx)
            y = BODY_Y + (i // 2) * (ch2 + gy)
            card(s, key, x, y, cw2, ch2, lenses.get(key, []), False, False)
        FX.transition(s, "fade")
        self.notes(s, "The same method, four domains. Each of the next four "
                      "slides opens one lens in full. " + notes)
        out.append(s)
        # expansions
        big_w = CW - 3.1
        for key in LENS_ORDER:
            s = self.new("lens")
            col, name, _ = LENS[key]
            self.chrome(s, f"{name}: {topic}", st.label, st,
                        ref=self._ref_for_lenses(), kicker_color=col)
            card(s, key, ML, BODY_Y, big_w, BODY_H, lenses.get(key, []),
                 True, True)
            yy = BODY_Y
            others = [k for k in LENS_ORDER if k != key]
            for k in others:
                card(s, k, ML + big_w + 0.26, yy, CW - big_w - 0.26, head,
                     None, False, False)
                yy += head + 0.18
            FX.transition(s, "morph", 700)
            txt = " ".join(e.text for e in lenses.get(key, [])
                           if isinstance(e, T.Para))
            self.notes(s, f"{name}. {txt}")
            out.append(s)
        return out

    def _ref_for_lenses(self):
        return f"Handout p. {self.ch.last_page}" if self.ch else ""

    def fit_size_elems(self, elems, w, h, hi, lo):
        for sz in range(hi, lo - 1, -1):
            if self.flow_height(self.elements_to_blocks(elems, w, sz,
                                                        None)) <= h:
                return sz
        return lo

    # -- lab -----------------------------------------------------------------

    def lab_slides(self, title, elements, code_file="", week=None, ref=""):
        st = STYLES["lab"]
        out = []
        codes = [e for e in elements if isinstance(e, T.CodeE)]
        first_code = next((i for i, e in enumerate(elements)
                           if isinstance(e, T.CodeE)), len(elements))
        before = [e for e in elements[:first_code]
                  if not isinstance(e, T.CodeE)]
        after = [e for e in elements[first_code:]
                 if not isinstance(e, T.CodeE)]
        # briefing
        s = self.new("lab")
        self.chrome(s, f"Try it yourself: {title}", st.label, st, ref)
        y = BODY_Y
        if before:
            blocks = self.elements_to_blocks(before, CW, 20, st)
            self.draw_flow(s, blocks, ML, y, CW)
            y += self.flow_height(blocks) + 0.34
        # the route through the code, from the listing's own step markers
        heads = []
        for c in codes:
            for h_, _ in split_code(c.code):
                h_ = h_.replace(" (continued)", "")
                if h_ and h_ not in heads:
                    heads.append(h_)
        if len(heads) >= 2:
            self.text(s, ML, y, CW, 0.3, "WHAT YOU WILL DO", 12, st.ink,
                      True, spc=150, name="!!labsteps-h")
            y += 0.42
            room = BODY_B - 1.5 - y
            sz = 19
            items = []
            while sz >= 14:
                mw = M.width("00.", sz, BODY, True) + 0.12
                items = [PSpec([T.Run(h_)], sz, INK, marker=f"{k + 1}.",
                               marker_color=st.rule, marker_w=mw,
                               space_after=sz * 0.3)
                         for k, h_ in enumerate(heads)]
                if paras_height(items, CW) <= room:
                    break
                sz -= 1
            tb = self.write(s, ML, y, CW, paras_height(items, CW) + 0.1,
                            items, name="!!labsteps")
            FX.animate(s, [[FX.fx(tb, "fade", 300, para=k)]
                           for k in range(len(items))])
            y += paras_height(items, CW) + 0.34
        facts = []
        if week:
            facts.append(("THIS WEEK'S LAB", f"Week {week.week}: {week.lab}"))
        facts.append(("WHAT YOU WILL RUN",
                      f"{len(codes)} listing{'s' if len(codes) != 1 else ''}"
                      f", {sum(len(c.code.splitlines()) for c in codes)} "
                      f"lines of "
                      f"{', '.join(sorted({_lang_name(c.lang) for c in codes}))}"))
        if code_file:
            facts.append(("RUNNABLE COPY", f"code/{code_file}"))
        fw = (CW - 0.3 * (len(facts) - 1)) / len(facts)
        fx_ = ML
        fy = max(y, BODY_B - 1.2) if len(heads) >= 2 else \
            max(y, BODY_Y + 0.2)
        for lab_, val in facts:
            self.rrect(s, fx_, fy, fw, 1.18, st.bg, name=f"!!lf{lab_[:4]}",
                       r_in=0.1)
            self.rect(s, fx_, fy, 0.07, 1.18, st.rule,
                      name=f"!!lfb{lab_[:4]}")
            self.text(s, fx_ + 0.3, fy + 0.18, fw - 0.45, 0.26, lab_, 11.5,
                      st.ink, True, spc=140, name=f"!!lfl{lab_[:4]}")
            vsz = self.fit_size([T.Run(val)], fw - 0.45, 0.62, 18, 13)
            self.text(s, fx_ + 0.3, fy + 0.5, fw - 0.45, 0.62, val, vsz,
                      INK, family=MONO if lab_.startswith("RUNNABLE")
                      else BODY, name=f"!!lfv{lab_[:4]}")
            fx_ += fw + 0.3
        FX.transition(s, "fade")
        self.notes(s, " ".join(e.text for e in before
                               if isinstance(e, T.Para)))
        out.append(s)
        # the code, one step per slide
        for ci, c in enumerate(codes):
            chunks = split_code(c.code)
            total = len(chunks)
            line_no = 1
            for k, (head, body) in enumerate(chunks):
                s = self.new("code")
                self.kicker(s, f"LAB  ·  STEP {k + 1} OF {total}"
                            + (f"  ·  LISTING {ci + 1}" if len(codes) > 1
                               else ""), st.rule, st.icon)
                self.title(s, title, size=26)
                self.footer(s, ref)
                sub = head or ("Set-up" if k == 0 else "")
                y = BODY_Y - 0.12
                if sub:
                    self.text(s, ML, y, CW, 0.36, sub, 19, st.ink, True,
                              family=HEAD, name="!!codehead")
                    y += 0.5
                lines = body.split("\n")
                size = self.code_size(lines, CW - 1.1, 40, 20, 11)
                h = min(BODY_B - y, 0.36 + 0.3 + len(lines) * M.LINE *
                        size / 72.0)
                self.draw_code(s, T.CodeE(body, c.lang), ML, y, CW, h, size,
                               name="code", first_line=line_no,
                               label=(f"{_lang_name(c.lang)}  ·  "
                                      f"code/{code_file}" if code_file
                                      else _lang_name(c.lang)))
                line_no += len(lines)
                FX.transition(s, "morph" if k > 0 else "fade", 600)
                self.notes(s, (f"Step {k + 1}: {sub}. " if sub else "") +
                           "Read the code aloud; ask what each line would "
                           "print before running it.")
                out.append(s)
        if after:
            out += self.card_slides(st, "What the lab shows", after,
                                    ref=ref)
        return out

    # -- checkpoint ----------------------------------------------------------

    def checkpoint_slides(self, questions, answers, ref=""):
        st = STYLES["checkpoint"]
        out = []
        total = len(questions)
        for i, q in enumerate(questions):
            gq = ground(st.rule)
            s = self.new("question", gq, progress=False)
            self.icon(s, st.icon, WHITE, ML, 1.42, 0.42, nm="!!qicon")
            self.text(s, ML + 0.62, 1.47, 8, 0.34,
                      f"CHECKPOINT  {i + 1} OF {total}", 14,
                      mix(gq, WHITE, 0.70), True, spc=200, name="!!qlab")
            qs = self.fit_size(q.runs, CW, 3.2, 34, 22, HEAD)
            t = s.shapes.title
            self._title_runs(t, q.runs, qs, WHITE, ML, 2.2, CW, 3.3,
                             name="!!q")
            self.text(s, ML, 6.3, CW, 0.34,
                      "Discuss with the person next to you, then advance.",
                      14, mix(gq, WHITE, 0.55), italic=True, name="!!qhint")
            self.footer(s, ref, dark=True)
            FX.transition(s, "fade")
            self.notes(s, "Ask the room. Do not advance until someone "
                          "commits to an answer.")
            out.append(s)
            if i < len(answers):
                a = answers[i]
                s = self.new("answer")
                self.kicker(s, f"Checkpoint {i + 1} of {total}  ·  answer",
                            st.rule, "check")
                qsz = self.fit_size(q.runs, CW, 1.25, 24, 18, HEAD)
                t = s.shapes.title
                qh = PSpec(q.runs, qsz, family=HEAD, spacing=0.96).height(CW)
                self._title_runs(t, q.runs, qsz, PRIMARY, ML, TITLE_Y, CW,
                                 max(0.5, qh + 0.05), name="!!q")
                ay = TITLE_Y + qh + 0.42
                room = BODY_B - ay
                elems = [a]
                sz = self.fit_size_elems(elems, CW - 0.4, room, 22, 15)
                blocks = self.elements_to_blocks(elems, CW - 0.4, sz, st)
                ah = self.flow_height(blocks)
                self.rect(s, ML, ay, 0.07, ah, GREENHL, name="!!abar")
                self.draw_flow(s, blocks, ML + 0.36, ay, CW - 0.4,
                               name="!!ans")
                self.footer(s, ref)
                FX.transition(s, "morph", 650)
                self.notes(s, "Answer (Appendix D): " + a.text)
                out.append(s)
        return out

    def _title_runs(self, t, runs, size, color, x, y, w, h, name="!!q"):
        t.left, t.top, t.width, t.height = (Inches(x), Inches(y), Inches(w),
                                            Inches(h))
        t.name = name
        tf = t.text_frame
        tf.clear()
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Emu(0)
        tf.margin_top = tf.margin_bottom = Emu(0)
        tf.vertical_anchor = MSO_ANCHOR.TOP
        FX.no_autofit(tf)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 0.96
        style_runs(p, runs, size, color, family=HEAD, term_color=color)

    # -- summary -------------------------------------------------------------

    def summary_slides(self, items, ref=""):
        st = STYLES["chaptersummary"]
        badge = 0.44
        tx = ML + badge + 0.26
        tw = CW - (tx - ML)
        room = BODY_H
        for sz in (21, 20, 19, 18, 17):
            hs = [max(badge, PSpec(it.runs, sz, spacing=1.04).height(tw))
                  for it in items]
            if sum(hs) + 0.2 * (len(hs) - 1) <= room:
                break
        pages, cur, h = [], [], 0.0
        for k, (it, hh) in enumerate(zip(items, hs)):
            if cur and h + hh > room:
                pages.append(cur)
                cur, h = [], 0.0
            cur.append((k, it, hh))
            h += hh + 0.2
        if cur:
            pages.append(cur)
        out = []
        for pi, page in enumerate(pages):
            s = self.new("summary")
            self.chrome(s, "What to take away" if pi == 0 else
                        "What to take away (continued)", st.label, st, ref)
            y = BODY_Y
            steps = []
            for k, it, hh in page:
                b = self.rrect(s, ML, y + 0.01, badge, badge, st.rule,
                               name=f"!!sb{k}", r_in=0.1)
                n = self.text(s, ML, y + 0.07, badge, 0.3, str(k + 1), 15,
                              WHITE, True, align="c", family=HEAD,
                              name=f"!!sbn{k}")
                t = self.write(s, tx, y, tw, hh,
                               [PSpec(it.runs, sz, INK, spacing=1.04)],
                               name=f"!!st{k}")
                steps.append([FX.fx(b, "zoom", 260), FX.fx(n, "fade", 260),
                              FX.fx(t, "wipeRight", 400)])
                y += hh + 0.2
            FX.transition(s, "fade")
            FX.animate(s, steps)
            self.notes(s, "")
            out.append(s)
        return out

    # -- review questions ----------------------------------------------------

    def review_slides(self, items, ref=""):
        st = STYLES["reviewq"]
        out = []
        mcq, rest = [], []
        for it in items:
            tag, body = split_tag(it.runs)
            (mcq if tag == "MCQ" else rest).append((tag, body))
        for n, (tag, body) in enumerate(mcq, 1):
            stem, opts = split_options(body)
            s = self.new("mcq")
            self.chrome(s, f"Quick check {n} of {len(mcq)}",
                        "Review  ·  multiple choice", st, ref)
            sz = self.fit_size(stem, CW, 1.3, 26, 19)
            sh = PSpec(stem, sz, spacing=1.04).height(CW)
            self.write(s, ML, BODY_Y, CW, sh + 0.05,
                       [PSpec(stem, sz, INK, spacing=1.04)], name="!!mstem")
            y = BODY_Y + sh + 0.4
            if opts:
                gx = 0.26
                ow = (CW - gx) / 2
                oh = min(1.1, (BODY_B - y - 0.2) / 2)
                for k, (letter, runs) in enumerate(opts[:4]):
                    x = ML + (k % 2) * (ow + gx)
                    yy = y + (k // 2) * (oh + 0.2)
                    self.rrect(s, x, yy, ow, oh, "F4F7F9", RULE,
                               name=f"!!mo{k}", r_in=0.1, lw=1.0)
                    self.rrect(s, x + 0.2, yy + (oh - 0.5) / 2, 0.5, 0.5,
                               self.ground, name=f"!!mol{k}", r_in=0.25)
                    self.text(s, x + 0.2, yy + (oh - 0.5) / 2 + 0.09, 0.5,
                              0.3, letter.upper(), 15, WHITE, True,
                              align="c", family=HEAD, name=f"!!molt{k}")
                    osz = self.fit_size(runs, ow - 1.1, oh - 0.2, 20, 14)
                    self.write(s, x + 0.9, yy + (oh - PSpec(runs, osz)
                                                 .height(ow - 1.1)) / 2,
                               ow - 1.1, oh, [PSpec(runs, osz, INK)],
                               name=f"!!mot{k}")
            FX.transition(s, "fade")
            self.notes(s, "Take a show of hands for each option before "
                          "discussing. (The handout does not print MCQ "
                          "answers.)")
            out.append(s)
        if rest:
            specs = []
            for tag, body in rest:
                specs.append((tag, body))
            badge_w = 1.25
            tw = CW - badge_w - 0.2
            for sz in (19, 18, 17, 16, 15):
                hs = [PSpec(b, sz, spacing=1.04).height(tw) for _, b in specs]
                if sum(hs) + 0.24 * len(hs) <= BODY_H:
                    break
            pages, cur, h = [], [], 0.0
            for (tag, body), hh in zip(specs, hs):
                if cur and h + hh > BODY_H:
                    pages.append(cur)
                    cur, h = [], 0.0
                cur.append((tag, body, hh))
                h += hh + 0.24
            if cur:
                pages.append(cur)
            for pi, page in enumerate(pages):
                s = self.new("review")
                self.chrome(s, "For your notebook" if pi == 0 else
                            "For your notebook (continued)",
                            "Review questions", st, ref)
                y = BODY_Y
                for k, (tag, body, hh) in enumerate(page):
                    col = {"SHORT": SECONDARY, "APPLIED": TEALHL}.get(
                        tag, PRIMARY)
                    self.rrect(s, ML, y + 0.02, badge_w - 0.1, 0.34,
                               mix(col, WHITE, 0.88), name=f"!!rt{k}",
                               r_in=0.17)
                    self.text(s, ML, y + 0.065, badge_w - 0.1, 0.26,
                              tag or "Q", 11, ground(col), True, align="c",
                              spc=120, name=f"!!rtt{k}")
                    self.write(s, ML + badge_w + 0.1, y, tw, hh,
                               [PSpec(body, sz, INK, spacing=1.04)],
                               name=f"!!rq{k}")
                    y += hh + 0.24
                FX.transition(s, "fade")
                self.notes(s, "Set these for independent study. Model "
                              "answers are not printed in the handout.")
                out.append(s)
        return out

    # -- closing -------------------------------------------------------------

    def closing_slide(self, next_ch=None, week=None, takeaways=None,
                      last=False):
        ch = self.ch
        s = self.new("closing", self.ground, progress=False)
        wm = self.text(s, W - 6.3, -0.55, 6.0, 5.0, f"{ch.number:02d}", 280,
                       mix(self.ground, WHITE, 0.08),
                       family="Segoe UI Black", align="r", name="!!wm")
        FX.decorative(wm)
        FX.to_back(wm)
        self.rect(s, ML, 1.12, 0.9, 0.07, WHITE, name="!!tbar")
        self.title(s, "End of the course" if last else
                   f"End of Chapter {ch.number}", color=WHITE, x=ML, y=1.36,
                   w=9.0, h=0.9, size=40)
        y = 2.5
        cards = []
        if next_ch:
            cards.append(("NEXT", f"Chapter {next_ch[0]} — {next_ch[1]}"))
        elif last:
            cards.append(("THE COURSE", "32 chapters  ·  six parts  ·  four "
                          "lenses  ·  one capstone"))
        cards.append(("READ", f"Handout Chapter {ch.number}, pages "
                              f"{ch.page}–{ch.last_page}"))
        if week:
            cards.append(("THIS WEEK'S LAB", week.lab))
            if week.assessment:
                cards.append(("ASSESSMENT", week.assessment))
        cw_ = (CW - 0.3) / 2
        for i, (lab, val) in enumerate(cards):
            x = ML + (i % 2) * (cw_ + 0.3)
            yy = y + (i // 2) * 1.42
            self.rrect(s, x, yy, cw_, 1.2, mix(self.ground, "000000", 0.22),
                       name=f"!!cl{i}", r_in=0.1)
            self.text(s, x + 0.3, yy + 0.2, cw_ - 0.6, 0.26, lab, 11.5,
                      mix(self.ground, WHITE, 0.62), True, spc=160,
                      name=f"!!cll{i}")
            vsz = self.fit_size([T.Run(val)], cw_ - 0.6, 0.62, 21, 14, HEAD)
            self.text(s, x + 0.3, yy + 0.52, cw_ - 0.6, 0.62, val, vsz,
                      WHITE, family=HEAD, name=f"!!clv{i}")
        self.footer(s, "", dark=True)
        FX.transition(s, "fade")
        self.notes(s, "Close. " + " ".join(f"{a}: {b}." for a, b in cards))
        return s

    # -- finishing ------------------------------------------------------------

    def finish(self, path: str) -> str:
        n = len(self.recs)
        idx = 0
        for i, rec in enumerate(self.recs):
            if not rec.progress:
                continue
            s = rec.slide
            frac = (i + 1) / n
            self.rect(s, 0, 0, W, 0.06, TRACK, name="!!ptrack")
            self.rect(s, 0, 0, max(0.05, W * frac), 0.06, self.hue,
                      name="!!pfill")
            idx += 1
        for shape, key in self.links:
            tgt = self.anchors.get(key)
            if tgt is not None:
                FX.link_to_slide(shape, tgt)
        groups = []
        for rec in self.recs:
            if not groups or groups[-1][0] != rec.group:
                groups.append((rec.group, []))
            groups[-1][1].append(rec.slide)
        FX.sections(self.prs, groups)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.prs.save(path)
        return path


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _roman(n: int) -> str:
    return ["", "I", "II", "III", "IV", "V", "VI", "VII"][n]


def _lang_name(l: str) -> str:
    return {"python": "Python", "r": "R", "bash": "Shell",
            "sh": "Shell"}.get(l, "Text")


def split_caption(cap: str, section: str) -> tuple[str, str]:
    """A caption's first sentence becomes the slide title, the rest its
    takeaway -- unless that sentence would make a poor title."""
    m = re.match(r"(.+?[.!?])(\s+|$)(.*)", cap, re.S)
    first, rest = (m.group(1), m.group(3).strip()) if m else (cap, "")
    first = first.rstrip(".")
    bad = (len(first) > 92 or len(first) < 12 or
           re.match(r"(Left|Right|Top|Bottom|Above|Below)\b", first) or
           first.count("(") != first.count(")"))
    if bad:
        return section or "Figure", cap
    return first, rest


def split_tag(runs):
    """'(MCQ) Precision is ...' -> ('MCQ', runs without the tag)."""
    txt = T.plain(runs)
    m = re.match(r"\s*\((MCQ|Short|Applied|Essay|Long)\)\s*", txt, re.I)
    if not m:
        return "", runs
    tag = m.group(1).upper()
    n = m.end()
    out, cut = [], 0
    for r in runs:
        if cut >= n:
            out.append(r)
            continue
        if cut + len(r.text) <= n:
            cut += len(r.text)
            continue
        k = n - cut
        out.append(T.Run(r.text[k:], r.bold, r.italic, r.mono, r.term,
                         r.math, r.script, r.color))
        cut = n
    if out:
        out[0].text = out[0].text.lstrip()
    return tag, out


def split_options(runs):
    """Split an MCQ stem from its '(a) ... (b) ...' options."""
    txt = T.plain(runs)
    marks = list(re.finditer(r"\(([a-e])\)\s*", txt))
    if len(marks) < 2:
        return runs, []
    pieces = []
    bounds = [m.start() for m in marks] + [len(txt)]
    stem = _slice_runs(runs, 0, marks[0].start())
    for i, m in enumerate(marks):
        seg = _slice_runs(runs, m.end(), bounds[i + 1])
        if seg:
            seg[-1].text = seg[-1].text.rstrip(" .;,")
        pieces.append((m.group(1), seg))
    return stem, pieces


def _slice_runs(runs, a, b):
    out, pos = [], 0
    for r in runs:
        s, e = pos, pos + len(r.text)
        pos = e
        lo, hi = max(a, s), min(b, e)
        if lo < hi:
            out.append(T.Run(r.text[lo - s:hi - s], r.bold, r.italic, r.mono,
                             r.term, r.math, r.script, r.color))
    if out:
        out[0].text = out[0].text.lstrip()
        out[-1].text = out[-1].text.rstrip()
    return [r for r in out if r.text]


_STEP = re.compile(r"^\s*(#|//|--)\s*(-{3,}|={3,})\s*(.*?)\s*[-=]*\s*$")


def split_code(code: str, max_lines: int = 20):
    """Split a lab listing at the handout's '# --- n. step ---' markers.

    Returns [(heading, chunk)]. The marker line stays in the chunk (it is
    part of the program); its words become the slide's step heading.
    Chunks longer than *max_lines* are split again at blank lines.
    """
    lines = code.split("\n")
    marks = [i for i, l in enumerate(lines) if _STEP.match(l)]
    if not marks:
        marks = []
    bounds = ([0] if not marks or marks[0] != 0 else []) + marks + \
        [len(lines)]
    chunks = []
    for a, b in zip(bounds, bounds[1:]):
        seg = lines[a:b]
        while seg and not seg[-1].strip():
            seg.pop()
        if not seg:
            continue
        m = _STEP.match(seg[0])
        head = ""
        if m:
            head = re.sub(r"^\d+[.)]\s*", "", m.group(3)).strip(" -=")
            if head.isupper() and len(head) > 4:
                # "BASELINE FIRST" is emphasis in a comment; as a heading
                # among sentence-case ones it reads as shouting
                head = head.capitalize()
            head = head[:1].upper() + head[1:]
        # long chunks: split at blank lines
        while len(seg) > max_lines:
            cut = max((i for i, l in enumerate(seg[:max_lines])
                       if not l.strip() and i > 4), default=max_lines)
            chunks.append((head, "\n".join(seg[:cut]).rstrip()))
            seg = seg[cut:]
            while seg and not seg[0].strip():
                seg.pop(0)
            head = (head + " (continued)") if head and \
                not head.endswith("(continued)") else head
        if seg:
            chunks.append((head, "\n".join(seg).rstrip()))
    return chunks or [("", code)]
