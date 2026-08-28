# CCNA 200-301 v2.0 — Lecture Decks

PowerPoint decks generated from the book in `../parts/`. One deck per chapter.
Nothing here is authored twice: every slide is derived from the chapter source,
so the lecture, the notes and the labs cannot drift apart.

**Current state:** Parts I–III built — chapters 1–19, **19 decks, 603 slides**,
229 carrying speaker notes. Static checks clean; every slide rendered and
reviewed.

## Build

```powershell
python figures\extract_figures.py            # TikZ -> PNG (all figures)
python build_slides.py --part 1               # -> out\*.pptx
python qa.py --render                                  # check, then look
```

`--all` does all 34. `--ch 8` does one. `qa.py` without `--render` runs the
static checks only, which is fast enough to run on every edit.

## Why PowerPoint, and what it costs

Chosen so colleagues can edit the decks. python-pptx **cannot create entrance
animations** and supports no SVG or EMF, and there is no Inkscape or
LibreOffice here to convert with. Three techniques recover most of what that
would otherwise cost:

| Want | How |
| --- | --- |
| Animated builds | **Stepped slides + Morph.** Consecutive slides carry the same shapes, in the same places, under the same *names*; PowerPoint's Morph transition tweens between them. Real motion, no fragile XML, every step still editable. |
| Crisp figures | The book's TikZ compiled to 300 dpi transparent PNG — ~4000 px at slide width, sharper than any projector. |
| Transitions | `<p:transition>` injected into the slide XML, which python-pptx does not expose. |

Morph needs PowerPoint 2016+. Older versions fall back to a cut, not a break.

## The pieces

| File | Does |
| --- | --- |
| `texparse.py` | Chapter `.tex` → typed blocks and styled runs. `--all` audits that nothing is dropped. |
| `ccnadeck.py` | The design system: palette, geometry, every slide layout. |
| `build_slides.py` | Decides which block becomes which slide, and where content splits. |
| `figures/extract_figures.py` | TikZ → standalone → PDF → PNG, with build steps. |
| `qa.py` | Static checks, then PowerPoint → PDF → PNG so slides can be read. |

## Deck grammar

Title · roadmap · outcomes · before-you-start · key terms · on-the-exam ·
concept and definition slides · figures · configure · verify · pitfalls ·
**Break & Fix** · lab briefing · checkpoints · summary · next.

The Break & Fix is the distinctive one. Each of the book's 34 becomes a paced
sequence — symptom, *what would you check?*, evidence, *what is it?*,
diagnosis, fix — so the class thinks before being told.

## Design

Palette is lifted from `../ccna.sty`, which already defines everything as HTML
hex, so slides and book match exactly. **Each deck is tinted by its chapter's
primary blueprint domain** — switching teal, routing purple, security red — so
students absorb the exam's five-domain structure over a semester.

Segoe UI for prose, Consolas for all IOS text, nothing below 13 pt.

## Logos

Drop `assets/iub-logo.png` and `assets/dit-logo.png` in and they are used on
the title slide and section dividers. Absent, a typographic institutional
lockup renders instead. No code change either way.

## Things learned the hard way

- **Sizing is the whole game.** python-pptx has no working autofit — PowerPoint
  only recalculates when it opens the file. Every layout sizes text up front,
  and `qa.py` measures boxes afterwards. When adding a layout, size the panel
  to its content: a six-line `show` output in a five-inch black box reads as a
  rendering fault, not as restraint.
- **PowerPoint's float line spacing multiplies a ~1.2 baseline**, so
  `line_spacing = 1.14` renders at about 1.37× the font size. Panel maths and
  the checker must use the same factor or one of them lies.
- **Pin figure bounding boxes with `\useasboundingbox`, not `\path`.** A path
  only ever grows the box, so the build step with the widest caption came out a
  different size and Morph slid the whole diagram sideways instead of fading
  parts into it. Anything outside the pinned box is cropped, so anchor captions
  at a fixed left edge rather than centring them.
- **Slide figures are not print figures.** Chapter 3's two side-by-side panels
  are right in the book, where a reader compares them at a glance. On a slide
  that wasted two thirds of the canvas at step 1. The slide version is one
  diagram that evolves in place — see `figures/src/switchflood.tex`.
- **A morph needs one shared shape name, not three.** A checkpoint question
  deliberately morphs from a full-bleed slide into the answer slide's heading,
  sharing only the question text. Zero shared names is the real fault.
- **TikZ draws no arrowhead on a closed path.** `\draw[->] ... circle` renders
  a plain ring with no sense of direction. Use an arc — `arc (0:308:r)`. This
  was wrong in the book's Figure 12.1 too, and is now fixed in both.
- **Track section headings positionally.** Reading all of a gap's `\section`
  commands and applying the last one gave every figure the heading of the
  section *after* it — Figure 12.1, which shows why a loop is fatal, arrived
  on a slide titled "How the tree is chosen".
- **Keep prose before and after a list apart.** Merging them printed a box's
  closing sentence as if it introduced the box. `Block.tail` holds the
  after-part, and the layout reserves height for it before chunking bullets.
- **A box can hold more than one list.** Taking only the first left chapter
  4's second list flattened into prose, with the literal word "tight" printed
  on the slide because `\begin{tight}` was not being dropped.
