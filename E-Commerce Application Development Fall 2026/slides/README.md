# `slides/` — the deck pipeline

Twenty-four decks, built from the handout rather than written beside it. The
handout is the single source: change `parts/`, rebuild, and the slides follow.
Nothing here is edited by hand, and a deck that has been edited by hand is a
deck that will silently diverge from the chapter it claims to present.

```bash
python figures/extract_figures.py     # every TikZ figure to a cropped PDF
python build_slides.py --all          # 23 chapter decks
python overview.py                    # the course overview deck
python qa.py --render                 # static checks, then LOOK at the result
```

## The chain

| Stage | File | What it does |
|---|---|---|
| Parse | `texparse.py` | Reads `parts/*.tex` into chapters, sections, boxes, tables, listings, figures. Resolves `\ref` and page numbers from the master's `.aux` |
| Draw | `dsdeck.py` | The house deck: palette, card styles, chips, code frames, the per-box-type styling |
| Build | `build_slides.py` | One deck per chapter, in the house grammar |
| Fix up | `pptxfx.py` | OOXML the library cannot reach: Morph transitions, alt text, decorative flags |
| Figures | `figures/extract_figures.py` | Compiles each `\dsfig` body standalone against `ecadhandout.sty` and crops it |
| Maths | `mathrender.py` | LuaLaTeX for the inline formulae, as images |
| Check | `qa.py` | Static checks, then renders every slide to PNG and a contact sheet |
| Prose | `scanprose.py` | **The retarget gate.** See below |

## What differs from the Analysis of Algorithms copy

This pipeline was copied from the Algorithms course. The removals matter as
much as the additions, because a pipeline that still knows about theorems will
silently mis-count what it dropped.

**Removed**

- `RESULT_ENVS` and `_fold_proofs()` — this course states no theorems and
  proves nothing. Left in, they corrupt `texparse.py --all`'s dropped-content
  accounting, and an audit that miscounts is an audit nobody reads.
- `theorem`, `lemma`, `corollary` from `BOX_OPT` and from `STYLES`.
- `_MATH_WRAP`'s `\bigO`, `\bigOm`, `\bigTh`.

**Added**

- `measured` as a box type, in **three** places — `BOX_OPT` in `texparse.py`,
  `STYLES` in `dsdeck.py`, and `CARDS` in `build_slides.py`. The plan named
  two. The third was found only because Chapters 5 and 6 produced no measured
  slides at all while every other gate reported success, so there is now an
  import-time assertion:

  ```python
  _ORPHANS = (T.BOX_OPT | T.BOX_PLAIN) - _HANDLED
  assert not _ORPHANS, (...)
  ```

  A new box type now fails the build instead of disappearing from it.

- `\adminpath{System, Settings, Edit, Option}` → `System › Settings › Edit ›
  Option`, and `\uiel{...}` → bold. A GUI course names a click path on nearly
  every page; without these the trail arrives as a comma list and stops looking
  like navigation.
- Lexers for PHP, SQL, Twig, YAML, JSON, ini, bash and Python, and a widened
  `_lang_name`. The Algorithms course lost a day to `language=C++` parsing as
  `C`; this course has seven languages.
- `declared_chapters()` — the master's chapter count, not the count of
  chapters written so far. Without it, whichever chapter was finished last
  closed its deck with **"End of the course"**.
- `_best_aux()` — picks whichever master `.aux` knows about the most chapters,
  so the parse gate works while the book is being written and needs no edit
  when it is finished.

## `scanprose.py`, and why it exists

The retarget replaces names. It does not replace *descriptions*. Three blocks
of Virtual Systems prose survived into the Algorithms `overview.py` because
they never used a word the retarget searched for, and only a rendered contact
sheet found them.

The same thing happened here. Every chapter deck carried the tagline **"Design,
Proof and Measurement"** — the Algorithms subtitle — on its title slide, past
every text check, because it was a string literal nobody had reason to grep
for. `overview.py` still said *"A loop invariant, not a passing test"*.

```bash
python scanprose.py        # 0 suspicious literal(s)
```

It walks the Python with `ast` and inspects real string literals only — an
earlier regex version matched comments and produced findings nobody could act
on. Its `OTHER` vocabulary covers the Algorithms, VSS, Data Science, ML, AI and
ARM word lists; `container` and `docker` are in `ALLOW`, because here they are
this course's own vocabulary.

Sanity-checked against the un-retargeted Algorithms copy, where it reports 9
hits. A gate that cannot fail is not a gate.

## Rendering is not optional

`qa.py` without `--render` checks geometry, alt text, notes and overflow from
the XML. That is necessary and it is not sufficient. **Every defect listed
below was invisible to it and obvious in a contact sheet:**

- the Algorithms tagline on all 23 title slides;
- Chapter 6's title overflowing Chapter 5's "used later in" chip by 156%;
- a 15-row table running off the bottom of a slide;
- a 30-line PHP listing doing the same.

Look at the sheets. It takes four minutes for twenty-four decks and it is the
only gate that has ever caught a whole class of defect.

## Branding

Dr. Mustafa Hameed · IUB crest and DIT shield · Department of Information
Technology · BSIT · Fall 2026. Identical to the six sibling courses, so the
series reads as one.
