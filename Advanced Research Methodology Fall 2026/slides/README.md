# Lecture slides

PowerPoint decks for the Advanced Research Methodology course, **generated
from the handout source** rather than written separately.

```powershell
cd slides
python build_slides.py              # 33 chapter decks + course overview
python build_slides.py --ch 10      # rebuild one chapter
python build_slides.py --part 3     # rebuild one part's chapters
python build_slides.py --per-part   # one deck per part instead of per chapter
```

Output lands in `out/`. Requires `python-pptx` and `lxml`
(`pip install python-pptx`).

**Current build: 34 decks, 613 slides.**

---

## Why it is generated

The decks parse `../parts/*.tex` directly. Correct a definition in the
handout, rerun the script, and the slides follow. Nothing is retyped, so the
two cannot drift — which matters for a volume whose whole argument is that
claims should be traceable to a source.

This works because the chapter furniture is fixed and documented in
`../README.md`: every chapter carries `\locator`, `outcomes`, `prereq`,
`keyterms`, the teaching boxes, `worked`, `pitfall`, `traditions`,
`checkpoint`, `chaptersummary` and `reviewq`, in that order. **If you change
that structure, update `texparse.py` with it.**

Chapter numbers are read from the compiled `.aux`, so the slides always agree
with the printed volume. Build the handout before the slides; if the `.aux`
is missing the script warns and numbers come out as 0.

| File | Role |
|---|---|
| `texparse.py` | LaTeX → structured content. All the parsing and text conversion. |
| `pptxfx.py` | Slide transitions and entrance animations (raw OOXML). |
| `build_slides.py` | Theme, layouts, and the deck assembly. |

---

## Deck structure

Each chapter deck follows the chapter:

```
Title  ·  Learning outcomes  ·  Key terms
  Section divider → content slides (definitions, concepts, warnings)   ×N
Worked example  ·  Pitfalls  ·  Regulations (where the chapter quotes any)
Four traditions  ·  Summary  ·  Review questions  ·  Close
```

Colours are lifted from `../armhandout.sty` — the same teal `#0E4D64`,
`#137177` and accent `#C84B31`, and the same four-tradition colours — so a
slide and the page it came from read as one course.

---

## Animation

`python-pptx` has no animation API, so `pptxfx.py` builds the `<p:timing>`
and `<p:transition>` trees and grafts them onto the slide XML.

Entrance effects use PowerPoint's own preset ids (fade, wipe, fly-in, zoom),
so the animation pane shows familiar names rather than "Custom" and a
lecturer can edit them normally. Bullet lists advance one item per click;
cards and quadrants animate as groups.

**The nesting is load-bearing.** PowerPoint expects one `tmRoot` → one
`mainSeq` → one `<p:par>` per click, with effects sharing a click as siblings
inside it. Get it wrong and PowerPoint *silently drops every animation on the
slide* rather than reporting an error — so if you edit `build_timing()`,
verify in PowerPoint, not by reading the XML.

---

## Verification

Because a broken deck fails silently, the checks run PowerPoint itself:

1. **Open every deck.** Malformed timing XML makes PowerPoint refuse the file
   or offer to repair it. All 34 must open clean.
2. **Count animated slides.** Query `Slide.TimeLine.MainSequence.Count` — it
   must be non-zero on every slide (613/613).
3. **Check for overflow.** Compare `TextRange.BoundHeight` against the shape
   height. Text spilling past its box is invisible in the XML and obvious to
   a room.
4. **Look at the slides.** Export to PNG and actually read a sample.

Last run: 34/34 decks open, 613/613 slides animated, 0 overflows.

---

## Known limits

- **Tables and TikZ figures do not transfer.** The handout's figures are
  drawn in TikZ and its tables in `longtable`; neither has a PowerPoint
  equivalent the script can produce. A section built round a table falls back
  to its prose. Where a figure carries the argument — the alignment chain in
  Ch. 3, the PRISMA diagram in Ch. 5 — **project the handout page**.
- **Prose is condensed to whole sentences.** Long passages are cut at a
  sentence boundary, never mid-sentence. A slide is a pointer to the chapter,
  not a replacement for it.
- Speaker notes are not generated. The handout is the notes.
