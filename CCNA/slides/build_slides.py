#!/usr/bin/env python3
"""Build the CCNA lecture decks from the book.

    python build_slides.py --ch 8          one chapter
    python build_slides.py --part 1        a whole part
    python build_slides.py --all           all 34

Every slide is derived from ``parts/NN_*.tex``. Nothing here invents content;
it decides which block becomes which kind of slide, how long content is split,
and where the guided-diagnosis beats fall. ``ccnadeck.py`` draws.
"""
from __future__ import annotations

import argparse
import os
import re
import sys

from ccnadeck import (ACCENT, DEFBG, Deck, GOLDHL, LIGHTBG, PRIMARY, _est_h,
                      _flat, CW, BODY_H)
from texparse import (Block, Run, checkpoint_answers, chapter_index,
                      parse_chapter, plain)

HERE = os.path.dirname(os.path.abspath(__file__))
FIGS = os.path.join(HERE, "figures", "out")
OUT = os.path.join(HERE, "out")

PARTS = [
    (1, 7, "Part I — Foundations"),
    (8, 14, "Part II — Switching and Network Access"),
    (15, 19, "Part III — IP Routing"),
    (20, 26, "Part IV — Network Services and Security"),
    (27, 32, "Part V — Wireless, Virtualisation and Operations"),
    (33, 34, "Part VI — Sitting the Exam"),
]

# How each pedagogy block presents on a slide.
STYLE = {
    "definitionbox": ("DEFINITION", DEFBG),
    "conceptbox": ("CONCEPT", LIGHTBG),
    "alertbox": ("WATCH OUT", None),
    "examnote": ("ON THE EXAM", None),
    "worked": ("WORKED EXAMPLE", None),
    "pitfall": ("COMMON PITFALLS", None),
}


def part_of(n: int) -> str:
    for a, b, name in PARTS:
        if a <= n <= b:
            return name
    return ""


def slug(stem: str) -> str:
    return stem.replace("_", "-")


def sentence(s: str) -> str:
    """Capitalise the first letter and touch nothing else.

    The book's block titles read as the tail of a sentence -- ``\\begin{config}
    {a router that leases its own WAN address}`` prints as "Configure: a router
    ...".  On a slide the title stands alone, so it needs a capital.  ``str
    .capitalize()`` also lower-cases everything after it, which turned WAN into
    "wan" on 31 of the book's 58 configuration slides.
    """
    return s[:1].upper() + s[1:] if s else s


# ---------------------------------------------------------------------------
# Splitting content that will not fit one slide
# ---------------------------------------------------------------------------

def split_bullets(bullets: list, budget: float = BODY_H - 1.15,
                  size: int = 18) -> list[list]:
    """Chunk bullets so no slide overflows. Keeps order."""
    if not bullets:
        return []
    out, cur, h = [], [], 0.0
    for b in bullets:
        bh = _est_h([b], CW - 0.6, size) + 0.10
        if cur and h + bh > budget:
            out.append(cur)
            cur, h = [], 0.0
        cur.append(b)
        h += bh
    if cur:
        out.append(cur)
    return out


def split_code(code: str, max_lines: int = 21) -> list[str]:
    """Chunk a long listing across slides, breaking at comment boundaries.

    A 46-line configuration on one slide is unreadable from the back of a
    room whatever the point size, so long listings are split. IOS comment
    lines (``!``) mark the author's own logical steps, so breaking there
    keeps each slide a coherent piece of work rather than an arbitrary
    window onto a wall of text.
    """
    lines = code.split("\n")
    if len(lines) <= max_lines:
        return [code]

    # candidate break points: a bare "!" or a "! comment" starting a step
    breaks = [i for i, ln in enumerate(lines) if ln.strip().startswith("!")]
    chunks, start = [], 0
    while start < len(lines):
        limit = start + max_lines
        if limit >= len(lines):
            chunks.append("\n".join(lines[start:]))
            break
        cut = max((b for b in breaks if start + 6 < b <= limit), default=limit)
        chunks.append("\n".join(lines[start:cut]).rstrip())
        start = cut
    # Strip blank lines at a chunk boundary: a leading one prints as an
    # empty first row of the panel, which reads as output that scrolled off.
    return [c.strip("\n") for c in chunks if c.strip()]


def _fignote(note: str, caption) -> str:
    """Speaker notes for a stepped figure, keeping the book's own caption."""
    if not caption:
        cap = ""
    elif isinstance(caption, str):
        cap = caption
    else:
        cap = plain(caption)
    return "\n\n".join(t for t in (note, cap) if t)


def _notes(pending: list[Block], extra: str = "") -> str:
    """Speaker notes: the book's own prose around this point."""
    txt = " ".join(plain(p) for b in pending for p in b.paras)
    txt = (extra + "\n\n" + txt).strip() if extra else txt
    pending.clear()
    return txt


# ---------------------------------------------------------------------------
# Break & Fix -> guided diagnosis
# ---------------------------------------------------------------------------

_D_MARK = re.compile(r"^(diagnosis|the tell|what has happened)", re.I)
_F_MARK = re.compile(r"^(fix|the fix|immediate|then check|the lasting|"
                     r"the process lesson|the wider lesson|the lesson|"
                     r"and (change|expect)|the verification lesson|"
                     r"what should have happened|the rule)", re.I)


def breakfix_slides(deck: Deck, b: Block, note: str) -> None:
    """Six beats: symptom, pause, evidence, pause, diagnosis, fix."""
    setup, diagnosis, fix = [], [], []
    bucket = setup
    for p in b.paras:
        head = plain(p)[:40]
        if _D_MARK.match(head):
            bucket = diagnosis
        elif _F_MARK.match(head):
            bucket = fix
        bucket.append(p)

    evidence = [c for c in b.children if c.kind in ("verify", "hostcmd")]
    remedy = [c for c in b.children if c.kind in ("config", "playbook")]

    # 1 — the symptom
    deck.breakfix_slide("Symptom", b.title or "Something is broken",
                        paras=setup[:3], note=note)
    # 2 — think
    deck.breakfix_slide(
        "Your move", b.title or "",
        prompt="What would you check first — and what do you expect to see?",
        note="Pause here. Take two or three answers before advancing. The "
             "point is the prediction, not the command.")
    # 3 — the evidence
    for c in evidence:
        dev = c.args[0] if c.args else "IOS"
        cmd = c.args[1] if len(c.args) > 1 else ""
        deck.breakfix_slide("Evidence", cmd or "The output",
                            code=c.verbatim,
                            device=f"{dev}#" if cmd else dev,
                            note="Read it aloud. Ask what single field decides "
                                 "the answer.")
    # 4 — think again
    if evidence:
        deck.breakfix_slide(
            "Your move", "",
            prompt="What is the fault — and what does this output rule out?",
            note="The elimination matters as much as the answer.")
    # 5 — the diagnosis
    if diagnosis:
        for chunk in _chunk_paras(diagnosis, 4):
            deck.breakfix_slide("Diagnosis", "Why it is doing that",
                                paras=chunk, note="")
    # 6 — the fix
    for c in remedy:
        deck.breakfix_slide("The fix",
                            sentence(c.args[0]) if c.args else "The fix",
                            code=c.verbatim,
                            device=("Ansible  ·  YAML" if c.kind == "playbook"
                                    else "IOS  ·  configuration mode"),
                            note="")
    if fix:
        for chunk in _chunk_paras(fix, 4):
            deck.breakfix_slide("The fix", "And the lesson", paras=chunk,
                                note="")


def _chunk_paras(paras: list, per: int) -> list[list]:
    out, cur, h = [], [], 0.0
    for p in paras:
        ph = _est_h([p], CW, 18)
        if cur and (len(cur) >= per or h + ph > BODY_H - 1.9):
            out.append(cur)
            cur, h = [], 0.0
        cur.append(p)
        h += ph
    if cur:
        out.append(cur)
    return out


# ---------------------------------------------------------------------------
# One chapter -> one deck
# ---------------------------------------------------------------------------

def build_chapter(n: int, next_title: str = "") -> str:
    ch = parse_chapter(n)
    deck = Deck(n, ch.title, ch.topics)
    deck.title_slide(part_of(n))

    pending: list[Block] = []
    summary_items: list = []

    for b in ch.blocks:
        if b.kind == "prose":
            pending.append(b)
            continue
        if b.kind == "reviewq":
            continue                      # assessment, not lecture material

        note = _notes(pending)

        if b.kind == "outcomes":
            deck.content_slide("What you will be able to do",
                               bullets=b.items, kicker="LEARNING OUTCOMES",
                               note=note or "Read these out. They are the "
                                            "chapter's contract with the room.")

        elif b.kind == "prereq":
            deck.content_slide("Before we start", paras=b.paras,
                               tail=b.tail,
                               bullets=b.items, kicker="ASSUMED KNOWLEDGE",
                               tint=LIGHTBG, note=note)

        elif b.kind == "keyterms":
            terms = [t.strip() for t in
                     re.split(r"•", plain(b.paras[0]) if b.paras else "")
                     if t.strip()]
            deck.terms_slide(
                terms,
                note="Do not define these now - they are defined as they "
                     "arrive. This slide is for the students' notes.")

        elif b.kind == "figure":
            label = b.args[0] if b.args else ""
            steps = sorted(f for f in os.listdir(FIGS)
                           if re.fullmatch(re.escape(label) + r"-\d+\.png", f)) \
                if os.path.isdir(FIGS) else []
            title = b.section or ch.title
            if steps:
                # No deck caption on a stepped figure: an annotated figure
                # carries a caption of its own on every beat, drawn inside the
                # image so it cannot drift from what is on screen. Printing
                # the book's caption underneath as well says the same thing
                # twice on the last step. The book's wording is not lost --
                # it is in the speaker notes.
                for i, f in enumerate(steps):
                    deck.figure_slide(title, os.path.join(FIGS, f),
                                      caption="", kicker="",
                                      note=(_fignote(note, b.title)
                                            if i == 0 else ""),
                                      morph=True)
            else:
                deck.figure_slide(title, os.path.join(FIGS, label + ".png"),
                                  caption=b.title, note=note)

        elif b.kind in STYLE:
            kicker, tint = STYLE[b.kind]
            # A pitfall box never carries a title in the book -- it is a
            # standard piece of chapter furniture -- so falling back to the
            # enclosing section named the slide after whatever topic the box
            # happened to sit under: six wireless pitfalls arrived titled
            # "Everything else", and chapter 34's under "On the day". The
            # section is still useful context, so it moves to the kicker.
            if b.kind == "pitfall" and not b.title:
                title = "Common pitfalls"
                if b.section:
                    kicker = f"COMMON PITFALLS  ·  {b.section}"
            else:
                title = b.title or b.section or kicker.title()
            # Reserve height for the closing prose, or the bullets fill the
            # body and the tail lands in a box too small for it.
            reserve = _est_h(b.tail, CW, 17) + 0.25 if b.tail else 0.0
            chunks = split_bullets(b.items,
                                   budget=BODY_H - 1.15 - reserve) or [None]
            for i, part in enumerate(chunks):
                deck.content_slide(
                    title if i == 0 else f"{title} (continued)",
                    paras=b.paras if i == 0 else None,
                    tail=b.tail if i == len(chunks) - 1 else None,
                    bullets=part, kicker=kicker,
                    tint=tint if b.kind != "examnote" else None,
                    note=note if i == 0 else "",
                    icon_text="" if b.kind != "examnote"
                    else "This is what the blueprint actually asks for")

        elif b.kind in ("config", "playbook"):
            title = sentence(b.args[0]) if b.args else "Configuration"
            # The bar names the thing being typed into; "Configure" merely
            # repeated the kicker two inches above it.
            bar = ("Ansible  ·  YAML" if b.kind == "playbook"
                   else "IOS  ·  configuration mode")
            for i, part in enumerate(split_code(b.verbatim)):
                deck.terminal_slide(
                    title if i == 0 else f"{title} (continued)",
                    part, kicker="CONFIGURE", device=bar,
                    note=note if i == 0 else "")

        elif b.kind in ("verify", "hostcmd"):
            dev = b.args[0] if b.args else "IOS"
            cmd = b.args[1] if len(b.args) > 1 else ""
            # The command is the slide's title, so the panel bar carries only
            # the prompt. Printing the command in both places wastes the one
            # line of the panel that could say where the output came from.
            bar = f"{dev}#" if b.kind == "verify" else dev
            for i, part in enumerate(split_code(b.verbatim)):
                deck.terminal_slide(
                    (cmd or "Verification") if i == 0
                    else f"{cmd or 'Verification'} (continued)",
                    part, kicker="VERIFY", device=bar,
                    note=note if i == 0 else "")

        elif b.kind == "breakfix":
            breakfix_slides(deck, b, note)

        elif b.kind == "lab":
            title = b.args[0] if b.args else "Lab"
            plat = b.args[1] if len(b.args) > 1 else ""
            reserve = _est_h(b.tail, CW, 17) + 0.25 if b.tail else 0.0
            chunks = split_bullets(b.items,
                                   budget=BODY_H - 1.15 - reserve) or [None]
            for i, part in enumerate(chunks):
                deck.content_slide(
                    title if i == 0 else f"{title} (continued)",
                    paras=b.paras if i == 0 else None, bullets=part,
                    tail=b.tail if i == len(chunks) - 1 else None,
                    numbered=True, kicker=f"LAB · {plat}".upper(),
                    note=note if i == 0 else "")

        elif b.kind == "checkpoint":
            answers = checkpoint_answers(n)
            total = len(b.items)
            for i, q in enumerate(b.items):
                deck.question_slide(i + 1, total, q, None,
                                    note="Ask the room. Do not advance until "
                                         "somebody commits to an answer.")
                if i < len(answers):
                    deck.question_slide(i + 1, total, q, answers[i])

        elif b.kind == "chaptersummary":
            summary_items = b.items
            chunks = split_bullets(b.items)
            for i, part in enumerate(chunks):
                deck.content_slide(
                    "What to take away" if i == 0
                    else "What to take away (continued)",
                    bullets=part, kicker="CHAPTER SUMMARY", note=note if i == 0
                    else "")

    deck.closing_slide(next_title, summary_items[:5])

    path = os.path.join(OUT, f"ch{n:02d}-{slug(ch.stem)[3:]}.pptx")
    deck.save(path)
    return path


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", nargs="*", type=int)
    ap.add_argument("--part", type=int)
    ap.add_argument("--all", action="store_true")
    a = ap.parse_args()

    if a.part:
        lo, hi, _ = PARTS[a.part - 1]
        want = list(range(lo, hi + 1))
    elif a.all:
        want = list(range(1, 35))
    elif a.ch:
        want = a.ch
    else:
        want = [1]

    _, order = chapter_index()
    titles = {n: parse_chapter(n).title for n, _ in order if n <= 34}

    os.makedirs(OUT, exist_ok=True)
    total = 0
    for n in want:
        nxt = f"Chapter {n + 1} — {titles[n + 1]}" if n + 1 in titles else ""
        p = build_chapter(n, nxt)
        from pptx import Presentation
        cnt = len(Presentation(p).slides._sldIdLst)
        total += cnt
        print(f"  ch {n:>2}  {os.path.basename(p):34} {cnt:>3} slides")
    print(f"\n  {len(want)} deck(s), {total} slides -> slides/out/")
    return 0


if __name__ == "__main__":
    sys.exit(main())
