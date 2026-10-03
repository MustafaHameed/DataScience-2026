# Virtual Systems and Services — Fall 2026 · Lecture Decks

PowerPoint decks generated from the handout in `../parts/`. One deck per
chapter, plus a course overview. Nothing is authored twice: every slide is
derived from the chapter source, so the lecture, the book and the labs cannot
drift apart — and a correction made in the handout reaches the slides on the
next build.

**Current state:** 24 decks, 1,014 slides, every slide carrying speaker notes.
All 42 handout figures are embedded as vector art. Static checks clean.

## Build

```powershell
python figures\extract_figures.py     # TikZ -> SVG + PNG (only what changed)
python build_slides.py --all          # -> out\*.pptx
python overview.py                    # -> out\VSS_00_Course_Overview.pptx
python qa.py --render                 # check, then look at render\*_sheet.png
```

`--ch 12` does one chapter, `--part 3` a part. `qa.py` without `--render` runs
the static checks only, which is fast enough for every edit; `--deck <text>`
filters by deck name.

The handout must have been built first: the figure numbers, section numbers and
page references all come from `../VSS_Fall2026_Handout.aux`.

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
examples, tables · pitfalls · lab · checkpoints · summary · review questions ·
close.

A divider is a pause, so it has to be worth pausing for: only sections that put
more than one slide on screen get one.

The distinctive slide is the **lab walk-through**, which splits the handout's
listing at its own `# --- n.` step markers, one step per slide, with Morph
tweening between them. This course's labs are shell, YAML, Dockerfiles,
Terraform and one C program rather than Python, so the step markers may open
with `#`, `//` or `/*` — `../checklabs.py` enforces the same set the builder
recognises.

**This course has no Four Lenses.** `lens_slides()` in `dsdeck.py` and
`fourlenses` in `texparse.py` are dead code, kept so that the four courses'
copies of the pipeline stay diffable. The overview deck's corresponding slide
is *One idea, six times over* — the six resources this course virtualizes.

## Design

Palette and part colours are lifted from `../vsshandout.sty`, so the decks and
the book match: Part I primary, II secondary, III green, IV purple, V teal,
VI gold. `ground()` darkens a hue until it carries white type, which is what
makes the gold of Part VI usable full-bleed.

Segoe UI for prose, Cambria for inline maths, Consolas for code, nothing below
10.5 pt.

## This pipeline is a copy

Copied from `../../Artificial Intelligence Fall 2026/slides/`, which was copied
from Machine Learning, which was copied from Data Science. Keeping the four
diffable is deliberate. What was changed here:

| File | Change |
| --- | --- |
| `texparse.py` | `MASTER`/`AUX` point at this handout. **`PART_TITLES` and the new `PART_RANGES` are now derived from the master's own `\part` and `\input` lines** instead of being hard-coded — the siblings still list them, and would silently print another course's part names if copied carelessly, which is exactly what happened here before it was caught. |
| `texparse.py` | `strip_comments` now takes the newline with the comment, as TeX does. Without it a comment-only line inside a `tikzpicture` option list becomes a blank line, which ends the paragraph — and the failure appears only when the figure is *extracted*, never when the handout builds. Ported from Data Science; **worth porting back to AI and ML, which still have the bug.** |
| `dsdeck.py` | Course strings. Lab briefing now **paginates** its preamble instead of shrinking it: a long preamble used to push the fact cards off the bottom of the slide, and clamping the size instead drove the text below the 10.5 pt floor. |
| `build_slides.py` | `PARTS` for this course's six parts; deck names `VSS_ChNN_*`. |
| `qa.py` | `VSS_Ch` prefix (9 characters, not 8); `--deck` filter ported from Data Science. |
| `overview.py` | Ported from Data Science, which is the only sibling that has one. Four Lenses slide replaced; part ranges now read from `PART_RANGES`. |
| `figures/` | Preamble generated from `vsshandout.sty`. |

## Things worth knowing before you edit

- **A `\\` in a TikZ node needs `align` or `text width`** — on the node, or on
  a style it names, transitively. Without one it is the ordinary LaTeX line
  break and illegal in a node.
- **A `\\` *inside* a nested group of an aligned node is the same fault**, and
  far worse to find: `{Docker\\\emph{the image,\\not the VM}}` looks right, and
  the error appears at the end of a figure three pages later. `../checkfigs.py`
  catches both, and has a self-test.
- **Do not name a TikZ style `cap`, `step`, `pos`, `in` or `out`.** They shadow
  built-in keys, and the error names the key rather than your style.
  `../checkfigs.py` has the list.
- **The Bash tool's heredocs silently collapse `\\` to `\`.** Every `.tex` edit
  in this course was made with an editor or a Python file, never a shell
  heredoc. The same warning is in the Advanced Research Methodology README and
  it is not advice to ignore.
- **Measure, do not guess.** `measure.py` reads the actual Windows font files;
  a line is exactly `1.2 × size × line_spacing` tall, and Pillow's advance
  widths match PowerPoint's to within 0.3%.
- **Morph pairs shapes by name when the name starts with `!!`.** Every shape
  meant to travel between slides is named that way, and `qa.py` checks that
  consecutive Morph slides share at least one.
- **Rendering each slide and looking at it finds what no checker can.** The
  part dividers carried *Artificial Intelligence*'s part titles through a
  clean build and a clean static-check run. Only the contact sheet showed it.
