#!/usr/bin/env python3
"""The course-overview deck: the handout's front matter and Appendix G.

One deck for the first session of the semester, and for anyone who needs to see
the shape of the course without reading 263 pages. Everything on it is parsed
from the handout: the course map and the six part openers, the sixteen-week
schedule, the assessment weights and the dependency map.

    python overview.py          # or: python build_slides.py --overview
"""
from __future__ import annotations

import os
import sys

import dsdeck as D
import texparse as T
from dsdeck import ML, CW, BODY_Y, BODY_H, BODY_B, PSpec, STYLES

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")


def _figure(stem: str, label: str):
    for e in T.walk(T.parse_file(stem)):
        if isinstance(e, T.FigureE) and e.label == label:
            return e
    return None


def _first_box(stem: str, env: str, title_has: str = ""):
    """The first box of this kind; with *title_has*, the first whose title
    contains it. Appendix G has two alertboxes and only one of them is for
    students."""
    for e in T.walk(T.parse_file(stem)):
        if isinstance(e, T.Box) and e.env == env:
            if not title_has or title_has.lower() in (e.title_text or
                                                      "").lower():
                return e
    return None


def _tables(stem: str):
    return [e for e in T.walk(T.parse_file(stem))
            if isinstance(e, T.TableE)]


def build(chapters, deps, figman, verbose=True):
    deck = D.Deck(None, part=1)
    deck.prs.core_properties.title = (
        "Analysis of Algorithms — Course Overview")

    # -- title ------------------------------------------------------------
    deck.group = "Opening"
    s = deck.new("title", deck.ground, progress=False)
    wm = deck.text(s, D.W - 7.4, -0.5, 7.0, 5.2, "23", 280,
                   D.mix(deck.ground, D.WHITE, 0.10),
                   family="Segoe UI Black", align="r", name="!!wm")
    D.FX.decorative(wm)
    D.FX.to_back(wm)
    deck.rect(s, ML, 1.42, 0.9, 0.07, D.WHITE, name="!!tbar")
    deck.text(s, ML, 1.66, 9.5, 0.32, "BSIT  ·  FALL 2026", 13,
              D.mix(deck.ground, D.WHITE, 0.72), True, spc=180, name="!!tk")
    deck.title(s, "Analysis of Algorithms", color=D.WHITE, x=ML,
               y=2.1, w=11.9, h=1.3, size=46)
    deck.text(s, ML, 3.5, 9.3, 0.5,
              "Design, Proof and Measurement", 26,
              D.mix(deck.ground, D.WHITE, 0.86), family=D.HEAD,
              name="!!tsub")
    deck.text(s, ML, 4.55, 9.3, 0.36,
              "Six parts  ·  23 chapters  ·  16 weeks  ·  one project",
              16, D.mix(deck.ground, D.WHITE, 0.74),
              name="!!tsub2")
    deck.rect(s, 0, 5.78, D.W, D.H - 5.78, deck.deep, name="!!band")
    deck.text(s, ML, 6.06, 5.0, 0.26, "LECTURER", 10.5,
              D.mix(deck.deep, D.WHITE, 0.62), True, spc=160, name="!!lab")
    deck.text(s, ML, 6.32, 6.0, 0.4, deck.author, 19, D.WHITE, True,
              family=D.HEAD, name="!!author")
    deck.text(s, ML, 6.76, 6.4, 0.3,
              f"{deck.dept}  ·  {deck.edition}", 12,
              D.mix(deck.deep, D.WHITE, 0.72), name="!!dept")
    deck._logos(s, D.W - D.MR, 6.0, 0.98)
    D.FX.transition(s, "fade")
    deck.notes(s, "Open the semester with this deck. It answers: what is this "
                  "course, how is it sequenced, and how will I be assessed?")

    # -- the course on one page -------------------------------------------
    fig = _figure("00_frontmatter", "coursemap")
    if fig:
        deck.figure_slide(fig, T.ref_num("fig:coursemap"),
                          T.ref_page("fig:coursemap"), "The whole course",
                          "Thirty-two chapters in six parts. The two "
                          "application chapters the course is built towards "
                          "are 25 and 26.", "Handout front matter")

    # -- what the course virtualizes ---------------------------------------
    # This course has no four-lens device (see ../README.md). Its organising
    # idea is the one Chapter 1 states: the same move, applied six times to
    # six different resources.
    s = deck.new("cards")
    deck.chrome(s, "One idea, six times over", "The spine of the course",
                STYLES["conceptbox"], "Handout §1.5")
    layers = [
        (D.PRIMARY,   "THE WHOLE MACHINE",
         "An unmodified operating system boots on it.\nParts I and II."),
        (D.SECONDARY, "THE OPERATING SYSTEM",
         "Processes see a private system, sharing one kernel.\nPart III."),
        (D.GREENHL,   "STORAGE",
         "A file behaves as a disk; many disks behave as one.\nChapter 12."),
        (D.PURPLEHL,  "THE NETWORK",
         "A switch exists in software; a LAN spans sites.\nChapter 13."),
        (D.TEALHL,    "THE DESKTOP",
         "The screen is here and the computer is elsewhere.\nChapter 15."),
        (D.GOLDHL,    "AN ACCELERATOR",
         "One GPU serves several tenants at once.\nChapter 20."),
    ]
    gx, gy = 0.24, 0.20
    cw2 = (CW - 2 * gx) / 3
    ch2 = (BODY_H - 0.62 - gy) / 2
    for i, (col, name, blurb) in enumerate(layers):
        x = ML + (i % 3) * (cw2 + gx)
        y = BODY_Y + (i // 3) * (ch2 + gy)
        deck.rrect(s, x, y, cw2, ch2, D.mix(col, D.WHITE, 0.94),
                   name=f"!!lay{i}", r_in=0.10)
        deck.rect(s, x, y, 0.07, ch2, col, name=f"!!layb{i}")
        deck.text(s, x + 0.28, y + 0.16, cw2 - 0.5, 0.28, name, 11.5,
                  col, True, spc=120, name=f"!!layn{i}")
        deck.text(s, x + 0.28, y + 0.52, cw2 - 0.5, ch2 - 0.7, blurb, 14,
                  D.INK, name=f"!!layv{i}")
    deck.text(s, ML, BODY_B - 0.5, CW, 0.42,
              "In every case the question is the same: what is the lie, who "
              "maintains it, and what does maintaining it cost?", 15,
              D.MUTED, name="!!laynote")
    D.FX.transition(s, "fade")
    deck.notes(s, "The organising idea of the whole course. Each row of "
                  "Chapter 1's table becomes a part or a chapter, and the "
                  "same three questions are asked of each.")

    # -- the six parts ------------------------------------------------------
    for p in range(1, 7):
        stem = f"00_part{p}_opener"
        deck.group = f"Part {D._roman(p)}"
        hue = D.PART_HUE[p]
        ground = D.ground(hue)
        quote = _first_box(stem, "tcolorbox")
        elems = T.parse_file(stem)
        paras = [e for e in elems if isinstance(e, T.Para)]
        s = deck.new("divider")
        panel = 4.55
        deck.rect(s, 0, 0, panel, D.H, ground, name="!!dpanel")
        deck.text(s, 0.72, 2.10, panel - 1.0, 0.3, f"PART {D._roman(p)}", 12.5,
                  D.mix(ground, D.WHITE, 0.70), True, spc=180, name="!!dsec")
        lo, hi = T.PART_RANGES[p]
        deck.text(s, 0.72, 2.55, panel - 1.0, 1.2,
                  f"{D._roman(p)}", 76, D.WHITE, family=D.HEAD,
                  name="!!dnum")
        deck.text(s, 0.74, 4.0, panel - 1.0, 0.3, f"CHAPTERS {lo}–{hi}", 12.5,
                  D.mix(ground, D.WHITE, 0.70), True, spc=140, name="!!dch")
        deck.title(s, T.PART_TITLES[p], x=panel + 0.75, y=1.9,
                   w=D.W - panel - 1.5, h=1.1, size=34, color=D.INK)
        if paras:
            txt = paras[0].runs
            sz = deck.fit_size(txt, D.W - panel - 1.5, 2.6, 20, 15)
            deck.write(s, panel + 0.75, 3.15, D.W - panel - 1.5, 2.6,
                       [PSpec(txt, sz, D.MUTED, spacing=1.06)],
                       name="!!dquote")
        deck.footer(s, f"Handout part opener")
        D.FX.transition(s, "push")
        deck.notes(s, T.plain(paras[0].runs) if paras else "")
        f = _figure(stem, f"part{p}map")
        if f:
            deck.figure_slide(f, T.ref_num(f"fig:part{p}map"),
                              T.ref_page(f"fig:part{p}map"),
                              f"Part {D._roman(p)}: {T.PART_TITLES[p]}",
                              "", "Handout part opener")

    # -- schedule, assessment, dependencies ---------------------------------
    deck.group = "Schedule and assessment"
    tabs = _tables("G_schedule")
    if tabs:
        deck.table_slides(tabs[0], "Sixteen weeks", "", "Handout Appendix G",
                          kicker="The teaching schedule")
    if len(tabs) > 1:
        deck.table_slides(tabs[1], "How the marks are earned", "",
                          "Handout Appendix G", kicker="Assessment")
    dep = _figure("G_schedule", "dependency-map")
    if dep:
        deck.figure_slide(dep, T.ref_num("fig:dependency-map"),
                          T.ref_page("fig:dependency-map"),
                          "What depends on what", "", "Handout Appendix G")
    # The students' alertbox, not the lecturer's "if you must cut something".
    exam = _first_box("G_schedule", "alertbox", "examinations")
    if exam:
        deck.card_slides(STYLES["alertbox"], exam.title_text or
                         "What the examinations actually ask", exam.body, "",
                         "Handout Appendix G")

    # -- close --------------------------------------------------------------
    deck.group = "Wrap-up"
    s = deck.new("closing", deck.ground, progress=False)
    wm = deck.text(s, D.W - 7.4, -0.5, 7.0, 5.2, "01", 280,
                   D.mix(deck.ground, D.WHITE, 0.08),
                   family="Segoe UI Black", align="r", name="!!wm")
    D.FX.decorative(wm)
    D.FX.to_back(wm)
    deck.rect(s, ML, 1.12, 0.9, 0.07, D.WHITE, name="!!tbar")
    deck.title(s, "We start with Chapter 1", color=D.WHITE, x=ML, y=1.36,
               w=9.0, h=0.9, size=40)
    cards = [("FIRST SESSION", "Chapter 1 — What Virtualization Is"),
             ("BRING", "A laptop with virtualization enabled"),
             ("FIRST LAB", "Install a hypervisor; a first guest"),
             ("READ BEFORE WEEK 2", "Chapters 1 and 2")]
    cw_ = (CW - 0.3) / 2
    for i, (lab, val) in enumerate(cards):
        x = ML + (i % 2) * (cw_ + 0.3)
        y = 2.5 + (i // 2) * 1.42
        deck.rrect(s, x, y, cw_, 1.2, D.mix(deck.ground, "000000", 0.22),
                   name=f"!!cl{i}", r_in=0.1)
        deck.text(s, x + 0.3, y + 0.2, cw_ - 0.6, 0.26, lab, 11.5,
                  D.mix(deck.ground, D.WHITE, 0.62), True, spc=160,
                  name=f"!!cll{i}")
        vsz = deck.fit_size([T.Run(val)], cw_ - 0.6, 0.62, 21, 14, D.HEAD)
        deck.text(s, x + 0.3, y + 0.52, cw_ - 0.6, 0.62, val, vsz, D.WHITE,
                  family=D.HEAD, name=f"!!clv{i}")
    deck.footer(s, "", dark=True)
    D.FX.transition(s, "fade")
    deck.notes(s, "Close the first session here.")

    path = os.path.join(OUT, "AoA_00_Course_Overview.pptx")
    for attempt in range(4):
        if not deck.math.pending:
            break
        deck.math.flush()
    deck.finish(path)
    if verbose:
        print(f"  ov    {os.path.basename(path):60} "
              f"{len(deck.recs):>3} slides")
    return path


if __name__ == "__main__":
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    chapters = [(n, T.parse_chapter(n).title, T.parse_chapter(n).part)
                for n, _ in T.chapter_order()]
    build(chapters, {}, {})
