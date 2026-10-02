# Data Science — Fall 2026 · Lecture Decks

PowerPoint decks generated from the handout in `../parts/`. One deck per
chapter, plus a course overview. Nothing is authored twice: every slide is
derived from the chapter source, so the lecture, the book and the labs cannot
drift apart — and a correction made in the handout reaches the slides on the
next build.

**Current state:** 33 decks, 1,406 slides, every slide carrying speaker notes.
All 103 handout figures are embedded as vector art. Static checks clean; every
slide rendered and looked at.

## Build

```powershell
python figures\extract_figures.py     # TikZ -> SVG + PNG (only what changed)
python build_slides.py --all          # -> out\*.pptx
python qa.py --render                 # check, then look at render\*_sheet.png
```

`--ch 14` does one chapter, `--part 3` a part, `--overview` the overview deck.
`qa.py` without `--render` runs the static checks only, which is fast enough
for every edit. `extract_figures.py --check-crop` additionally verifies that
nothing falls outside a figure's pinned bounding box.

## The pieces

| File | Does |
| --- | --- |
| `texparse.py` | Chapter `.tex` → typed elements and styled runs. `--all` audits that nothing is dropped. |
| `measure.py` | Text measurement from the real font files, calibrated against PowerPoint. |
| `mathrender.py` | Display and worked-example maths, typeset by LuaLaTeX in the deck's own faces. |
| `icons.py` | The handout's Font Awesome icons, recolourable. |
| `dsdeck.py` | The design system: palette, grid, flow engine, every slide layout. |
| `pptxfx.py` | What python-pptx cannot do: animation, Morph, SVG, alt text, sections, links. |
| `build_slides.py` | Decides which handout block becomes which slide, and where content splits. |
| `overview.py` | The course-overview deck, from the front matter and Appendix G. |
| `qa.py` | Static checks, then PowerPoint → PNG → contact sheet. |

## Deck grammar

Title · **where this chapter sits** · outcomes · key terms · agenda · then, per
handout section: a divider, definition/concept/watch-out cards, figures, worked
examples, tables · pitfalls · **Four Lenses** · lab · checkpoints · summary ·
review questions · close.

A divider is a pause, so it has to be worth pausing for: only sections that put
more than one slide on screen get one.

The distinctive slides are the **Four Lenses** (a 2×2 of domain cards, then each
lens Morph-expands to its full text) and the **lab walk-through**, which splits
the handout's listing at its own `# --- n.` step markers, one step per slide,
with Morph tweening between them.

## Design

Palette and part colours are lifted from `../dshandout.sty` and the handout's
own dependency map, so the decks and the book match: Part I primary, II
secondary, III green, IV purple, V teal, VI gold. `ground()` darkens a hue until
it carries white type, which is what makes the gold of Part VI usable full-bleed.

Segoe UI for prose, Cambria for inline maths, Consolas for code, nothing below
10.5 pt.

## Things worth knowing before you edit

- **Measure, do not guess.** `measure.py` reads the actual Windows font files;
  two facts were calibrated against PowerPoint's own renderer: a line is exactly
  `1.2 × size × line_spacing` tall, and Pillow's advance widths match
  PowerPoint's to within 0.3%. Sizing from measurement is what lets cards fit
  their content instead of being padded to a guess.
- **A box sized to its text must not wrap it.** Floating-point round-off made a
  table column exactly as wide as the word "420" wrap it onto a second line, and
  every row inherited the extra height. `count_lines` carries a 1e-6 tolerance.
- **An empty table cell still has a paragraph mark**, at the template's 18 pt
  unless told otherwise — which silently padded every row of the cost table in
  Chapter 14.
- **cairo's SVG output misplaces LuaLaTeX's glyphs.** Converting the typeset
  maths with `pdftocairo -svg` shifted whole lines up to 0.14 in from where the
  PNG fallback put them, because the CID TrueType fonts (Segoe UI) are
  positioned differently in cairo's SVG surface than in its raster one. MuPDF
  writes those SVGs instead, with glyphs as paths. The TikZ figures, which are
  pdflatex with Type 1 fonts, convert correctly either way.
- **The title placeholder is at the bottom of the z-order.** Anything drawn
  full-bleed covers it, which is why the title slides lost their titles until
  the watermark was sent to the back. Use the slide's background fill, not a
  rectangle.
- **A tab, not spaces, after a bullet.** The marker is an inline run, so the
  hanging indent (`marL` plus a negative `indent`) only lines up the wrapped
  lines; the tab is what puts the first line's text at the same edge.
- **Morph pairs shapes by name when the name starts with `!!`.** Every shape
  meant to travel between slides is named that way, and `qa.py` checks that
  consecutive Morph slides share at least one.
- **`%` eats the newline in TeX, and the comment stripper must too.** Leaving it
  behind turned a comment-only line into a blank one, which ends the paragraph —
  and a blank line inside a `pgfplots` option list ends the compile.
- **A near-square figure wastes most of the slide.** Scaled to the body
  height, a figure with an aspect ratio near 1 fills barely a third of the
  column width. Below 1.55 the takeaway stands beside the figure instead of
  under it, which roughly doubles the drawing. Four figures qualify, and two
  of them are the diagrams the course opens with.
- **Rendering each figure alone at 300 dpi found faults the page-by-page review
  of the PDF had missed.** Hidden labels, clipped annotations, an arrow drawn
  through the boxes it was meant to skip, and two ROC curves whose stated AUCs
  were not the AUCs they had. All are fixed in the handout, which is the point of
  generating from it: see `../README.md` for the list.
