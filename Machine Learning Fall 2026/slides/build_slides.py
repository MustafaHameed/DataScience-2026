#!/usr/bin/env python3
"""Build the Machine Learning Fall 2026 lecture decks from the handout.

    python build_slides.py --ch 14         one chapter
    python build_slides.py --part 3        a whole part
    python build_slides.py --all           all 22 chapters

Every slide is derived from ``../parts/*.tex``. Nothing here invents content:
this file decides which handout block becomes which kind of slide, where
prose becomes speaker notes, and how the deck is grouped. ``dsdeck.py``
draws; ``texparse.py`` reads; ``mathrender.py`` typesets maths.

Figures must be built first (``figures/extract_figures.py``). Maths is
rendered on demand: each deck is laid out once as a dry run, every missing
formula is typeset in one LuaLaTeX pass, and the deck is laid out again.
"""
from __future__ import annotations

import argparse
import json
import functools
import os
import re
import sys
import time

import dsdeck as D
import texparse as T

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
MANIFEST = os.path.join(HERE, "figures", "out", "manifest.json")

PARTS = [(1, 4), (5, 8), (9, 14), (15, 17), (18, 20), (21, 22)]
FURNITURE = {"pitfall": "Pitfalls", "fourlenses": "Four lenses",
             "lab": "Lab", "checkpoint": "Checkpoint",
             "chaptersummary": "Wrap-up", "reviewq": "Wrap-up"}
CARDS = ("definitionbox", "conceptbox", "alertbox", "examplebox")


def safe(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")


def manifest() -> dict:
    try:
        return json.load(open(MANIFEST, encoding="utf-8"))
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# Dependencies between chapters, from the prerequisite boxes
# ---------------------------------------------------------------------------

def prereq_refs() -> dict[int, list[int]]:
    """{chapter: [chapters its 'Before You Start' box cites]}"""
    out = {}
    labels = {}
    for n, stem in T.chapter_order():
        src = open(os.path.join(T.PARTS, stem + ".tex"), encoding="utf-8").read()
        m = re.search(r"\\label\{ch:([^}]+)\}", src)
        if m:
            labels[m.group(1)] = n
    for n, stem in T.chapter_order():
        src = open(os.path.join(T.PARTS, stem + ".tex"), encoding="utf-8").read()
        m = re.search(r"\\begin\{prereq\}(.*?)\\end\{prereq\}", src, re.S)
        refs = []
        if m:
            for k in re.findall(r"\\chref\{([^}]+)\}", m.group(1)):
                if k in labels and labels[k] not in refs and labels[k] != n:
                    refs.append(labels[k])
        out[n] = sorted(refs)
    return out


@functools.lru_cache(None)
def body_refs() -> dict[int, set[int]]:
    """{chapter: chapters cited anywhere in it} -- for 'used later in',
    where a later chapter may rely on this one without listing it among
    its prerequisites (Chapter 20's function approximation, Chapter 13)."""
    labels, srcs = {}, {}
    for n, stem in T.chapter_order():
        srcs[n] = open(os.path.join(T.PARTS, stem + ".tex"),
                       encoding="utf-8").read()
        m = re.search(r"\\label\{ch:([^}]+)\}", srcs[n])
        if m:
            labels[m.group(1)] = n
    return {n: {labels[k] for k in re.findall(r"\\chref\{([^}]+)\}", src)
                if k in labels and labels[k] != n}
            for n, src in srcs.items()}


# ---------------------------------------------------------------------------
# One chapter -> one deck
# ---------------------------------------------------------------------------

def _para_text(elems) -> str:
    return " ".join(e.text for e in elems if isinstance(e, T.Para))


def _items(box):
    for e in box.body:
        if isinstance(e, T.ListE):
            return e.items
    return []


def _slideful(el) -> bool:
    if isinstance(el, T.Box):
        return el.env not in ("outcomes", "prereq", "keyterms") and \
            el.env not in FURNITURE
    return isinstance(el, (T.FigureE, T.TableE, T.ListE, T.CodeE))


def draw_chapter(deck: D.Deck, ch: T.Chapter, chapters, deps, figman):
    n = ch.number
    week = T.week_of(n)
    part_title = T.PART_TITLES.get(ch.part, "")
    deck.group = "Opening"
    deck.title_slide(part_title, week)

    # A prerequisite box may also point *forward* ("the foundation of ...
    # in Chapter 14"), so direction comes from chapter order, not the box.
    builds = [c for c in deps.get(n, []) if c < n]
    leads = sorted({c for c, refs in deps.items() if n in refs and c > n} |
                   {c for c in deps.get(n, []) if c > n})
    # the prerequisite boxes are the curated list; when it is short, add
    # later chapters that cite this one in their text (the slide shows five)
    extra = sorted(c for c, refs in body_refs().items()
                   if n in refs and c > n and c not in leads)
    leads = sorted((leads + extra)[:5])
    prereq = ch.first("prereq")
    deck.roadmap_slide(chapters, builds, leads,
                       prereq.body if prereq else [])

    oc = ch.first("outcomes")
    if oc:
        lead = [e for e in oc.body if isinstance(e, T.Para)]
        deck.outcomes_slide(lead[0].runs if lead else None, _items(oc))
    kt = ch.first("keyterms")
    if kt:
        txt = _para_text(kt.body)
        terms = [t.strip() for t in txt.split("•") if t.strip()]
        deck.terms_slide(terms)

    # sections: which get a divider, which are prose-only
    secs = [s for s in ch.sections if s.level == 1]
    weight, prose_only = {}, {}
    for it in ch.items:
        if not it.secnum or "." not in it.secnum:
            continue
        top = ".".join(it.secnum.split(".")[:2])
        weight.setdefault(top, 0)
        prose_only.setdefault(top, True)
        if isinstance(it.el, T.Box) and it.el.env in FURNITURE:
            continue
        if _slideful(it.el):
            weight[top] += 1
            prose_only[top] = False
        elif isinstance(it.el, T.Box):
            weight[top] += 1
            prose_only[top] = False
    agenda = [(s.num, s.num, s.title, s.page) for s in secs]
    if agenda:
        deck.agenda_slide(agenda)

    pending: list = []            # prose waiting to become speaker notes
    cur_top = None
    statement: list = []          # a prose-only section's paragraphs
    stmt_title = ""

    def ref_for(it):
        if it.secnum:
            return f"§{it.secnum}  ·  Handout p. {it.page}"
        return f"Handout p. {ch.page}"

    def notes():
        t = _para_text(pending)
        pending.clear()
        return t

    def flush_statement(it_ref):
        nonlocal statement
        if statement:
            first = len(deck.recs)
            deck.statement_slide(stmt_title, statement, "", it_ref)
            statement = []
            if cur_top and deck.anchors.get(cur_top) is None:
                deck.anchors[cur_top] = deck.recs[first].slide
            return first
        return None

    last_ref = f"Handout p. {ch.page}"
    for it in ch.items:
        el = it.el
        if isinstance(el, T.Box) and el.env in ("outcomes", "prereq",
                                                "keyterms"):
            continue
        top = ".".join(it.secnum.split(".")[:2]) if it.secnum else None
        furniture = isinstance(el, T.Box) and el.env in FURNITURE

        # entering a new top-level section
        if top and top != cur_top and not furniture:
            flush_statement(last_ref)
            cur_top = top
            sec = next((s for s in secs if s.num == top), None)
            deck.group = f"{top} {sec.title}" if sec else top
            if sec and weight.get(top, 0) > 1:
                idx = [s.num for s in secs].index(top) + 1
                deck.divider_slide(top, sec.title, idx, agenda, top)
            first_of_section = len(deck.recs)
            deck.anchors.setdefault(top, None)
            if prose_only.get(top):
                stmt_title = sec.title if sec else it.top
        last_ref = ref_for(it)
        start = len(deck.recs)

        if furniture:
            deck.group = FURNITURE[el.env]

        if isinstance(el, T.Para) or (isinstance(el, T.MathE)):
            if top and prose_only.get(top):
                statement.append(el)
            elif isinstance(el, T.MathE) or el.heavy:
                deck.points_slides(it.section or ch.title,
                                   pending[-1:] + [el], notes(), last_ref)
            else:
                pending.append(el)
            continue

        if isinstance(el, T.ListE):
            lead = []
            if pending and pending[-1].text.rstrip().endswith(":"):
                lead = [pending.pop()]
            deck.points_slides(it.section or ch.title, lead + [el], notes(),
                               last_ref)
        elif isinstance(el, T.FigureE):
            num = T.ref_num("fig:" + el.label)
            page = T.ref_page("fig:" + el.label)
            steps = sorted((k for k, v in figman.items()
                            if v.get("label") == el.label and v.get("level")),
                           key=lambda k: figman[k]["level"])
            if steps:
                for k, st in enumerate(steps):
                    deck.figure_slide(el, num, page, it.section,
                                      notes() if k == 0 else "", last_ref,
                                      step=st, morph=k > 0)
            else:
                deck.figure_slide(el, num, page, it.section, notes(),
                                  last_ref)
        elif isinstance(el, T.TableE):
            deck.table_slides(el, it.section or ch.title, notes(), last_ref)
        elif isinstance(el, T.CodeE):
            deck.card_slides(D.STYLES["examplebox"], it.section, [el],
                             notes(), last_ref)
        elif isinstance(el, T.Box):
            env = el.env
            title = el.title_text or it.section or ch.title
            if env in CARDS:
                deck.card_slides(D.STYLES[env], title, el.body, notes(),
                                 last_ref, reveal=True)
            elif env == "worked":
                deck.worked_slides(title, el.body, notes(), last_ref)
            elif env == "pitfall":
                lead = [e for e in el.body if isinstance(e, T.Para)]
                deck.pitfall_slides(el.title_text or "Common pitfalls",
                                    _items(el), lead[0] if lead else None,
                                    notes(), last_ref)
            elif env == "fourlenses":
                lenses = {c.arg: c.body for c in el.children}
                deck.lens_slides(el.arg, lenses, notes())
            elif env == "lab":
                deck.lab_slides(el.title_text or "Lab", el.body,
                                T.lab_code_file(n), week, last_ref)
            elif env == "checkpoint":
                deck.checkpoint_slides(_items(el), T.checkpoint_answers(n),
                                       f"Handout p. {ch.last_page}  ·  "
                                       f"answers in Appendix D")
            elif env == "chaptersummary":
                deck.summary_slides(_items(el), f"Handout p. {ch.last_page}")
            elif env == "reviewq":
                deck.review_slides(_items(el), f"Handout p. {ch.last_page}")
        if top and deck.anchors.get(top) is None and len(deck.recs) > start \
                and not furniture:
            deck.anchors[top] = deck.recs[start].slide
    flush_statement(last_ref)
    # anchor any section whose first slide was a divider or statement
    for k, v in list(deck.anchors.items()):
        if v is None:
            deck.anchors.pop(k)

    deck.group = "Wrap-up"
    nxt = next(((c, t) for c, t, _ in chapters if c == n + 1), None)
    deck.closing_slide(nxt, week, last=n == len(chapters))


def build_chapter(n: int, chapters, deps, figman, verbose=True) -> str:
    ch = T.parse_chapter(n)
    t0 = time.time()
    for attempt in range(5):
        deck = D.Deck(ch)
        draw_chapter(deck, ch, chapters, deps, figman)
        pend = len(deck.math.pending)
        if not pend:
            break
        k = deck.math.flush()
        if verbose:
            print(f"    typeset {k} formula(s)")
    path = os.path.join(OUT, f"ML_Ch{n:02d}_{safe(ch.title)}.pptx")
    # the anchors dict above points into the last deck built
    deck.finish(path)
    if verbose:
        print(f"  ch {n:>2}  {os.path.basename(path):60} "
              f"{len(deck.recs):>3} slides  {time.time() - t0:5.1f}s")
    return path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", nargs="*", type=int)
    ap.add_argument("--part", type=int)
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--overview", action="store_true")
    a = ap.parse_args()

    chapters = [(n, T.parse_chapter(n).title, T.parse_chapter(n).part)
                for n, _ in T.chapter_order()]
    deps = prereq_refs()
    figman = manifest()
    if a.part:
        lo, hi = PARTS[a.part - 1]
        want = list(range(lo, hi + 1))
    elif a.all:
        want = [c for c, _, _ in chapters]
    elif a.ch:
        want = a.ch
    else:
        want = [] if a.overview else [1]
    os.makedirs(OUT, exist_ok=True)
    for n in want:
        build_chapter(n, chapters, deps, figman)
    if a.overview:
        # The Data Science pipeline has an overview deck; this course does not
        # (yet). Fail clearly rather than with an ImportError.
        print("no overview deck for this course")
    return 0


if __name__ == "__main__":
    sys.exit(main())
