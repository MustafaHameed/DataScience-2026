# CCNA 200-301 v2.0 — Lecture Decks

PowerPoint decks generated from the book in `../parts/`. One deck per chapter.
Nothing here is authored twice: every slide is derived from the chapter source,
so the lecture, the notes and the labs cannot drift apart.

**Current state:** complete. All six parts, **34 decks, 1439 slides**, 522
carrying speaker notes. Ten of the book's figures are progressive builds; the
rest render as stills. Static checks clean; every slide rendered and reviewed.

## Build

```powershell
python figures\extract_figures.py            # TikZ -> PNG (all figures)
python build_slides.py --part 1               # -> out\*.pptx
python qa.py --render                                  # check, then look
```

`--all` does all 34. `--ch 8` does one. `qa.py` without `--render` runs the
static checks only, which is fast enough to run on every edit.
`extract_figures.py --check-crop` additionally verifies that nothing falls
outside a figure's pinned bounding box, at the cost of one extra compile per
build level.

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

Title · **where this chapter sits** · outcomes · before-you-start · key terms ·
on-the-exam · concept and definition slides · figures · configure · verify ·
pitfalls · **Break & Fix** · lab briefing · checkpoints · summary · next, with
a **section divider** wherever the book starts a new `\section`.

A divider is a pause, so it has to be worth pausing for: only top-level
sections get one, only when the section puts more than one slide on screen, and
never in front of the end-of-chapter furniture, which inherits whatever heading
it happens to sit under. That comes to 116 dividers across the 34 decks.

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

The artwork lives in `../assets`, shared with the book's title page, and is
generated from the originals in `../assets/src` by `../tools/prepare_logos.py`.
Neither original could be used as delivered: the department shield had no
transparency at all, so on a coloured slide it was a white box, and the
university crest is a single navy that vanishes against every one of the five
domain colours.

So the script strips the outer background by flooding inward from the border --
keying out white everywhere would hollow out the shield, whose interior is
white by design -- and produces a reversed white crest for dark grounds. That
knockout keys on luminance rather than painting every pixel white, or the
calligraphy inside the dome fills in and disappears.

Remove the files and the decks still build: the marks simply do not appear, and
the typographic lockup carries the attribution on its own.

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
- **`str.capitalize()` lower-cases everything after the first letter.** The
  book's configuration titles are written as sentence tails —
  `{a router that leases its own WAN address}` — so a slide has to capitalise
  the first letter. `.capitalize()` did that and turned WAN into "wan", on 31
  of the book's 58 configuration slides.
- **Size a chip by its characters, not by an average.** `fit_pt`'s flat 0.50 em
  is fair for a sentence and far too narrow for an all-capital term: the
  "DHCPDISCOVER" key-term chip wrapped onto a second line inside its own box.
  `text_w()` measures per character, and `_est_h(..., tight=True)` uses it
  wherever a decoration is sized to text that is already laid out.
- **Hang the indent under a bullet.** The marker is an inline run, not a
  PowerPoint list marker, so a wrapped item put its second line hard against
  the margin and read as a new paragraph. `marL` plus a negative `indent`.
- **Content should hang from the title rule, not float in the body.** Optical
  centring looked better on one slide and worse on forty: the eye has to
  re-find the text on every click, and the "(continued)" half of a split list
  began two inches below the first half. Short content now gets a 0.42 in
  nudge and no more.
- **Say it once.** The command was the slide title *and* the terminal panel's
  prompt bar; the bar now carries only the device. A stepped figure narrates
  itself on every beat, so the book's caption printed underneath repeated the
  last step — it moved to the speaker notes.
- **A build must not answer its own question.** Chapter 25's five defences each
  arrive as the answer to "and what stops this?", so the question leads by one
  beat: level *n* poses attack *n*, level *n+1* fills the slot.
- **Some figures must not be revealed by accumulation.** The reachability
  ladder's *shape* is the lesson, and adding one rung at a time spends three
  steps showing a ladder that appears to have three rungs. `\stepdim{n}{...}`
  is the complement of `\stepvis`: a dim/solid pair draws every rung faint from
  the start and lights each as it is reached.
- **A pinned bounding box crops silently.** `\useasboundingbox` is what keeps
  Morph fading instead of sliding, but anything outside it is cut with no
  warning — DORA's first caption lost "Who can hear it?" mid-word.
  `--check-crop` compiles each level again with the pin released and compares
  page sizes.
- **Looking at the figures found four collisions the book's page-by-page
  review had missed.** Figure 30.2's brace labels printed through the syslog
  description, 30.1's legend ran through both SNMP arrows, and 21.1's two DNS
  labels sat on top of each other, and 33.1's loop label printed through the
  ELIMINATE box. All four are obvious at 300 dpi on their own and were not on
  a page at 130. Both copies of each are fixed.
- **Align a build's captions with its steps.** Every caption in the
  troubleshooting cycle named the box to the *left* of the one that had just
  lit up, because they were written as though step 1 were an introduction. The
  figure compiled, the check passed, and only reading slide 10 showed it.
- **A conditional node cannot be referred to by name from a lower level.**
  Arrows between stepped nodes are drawn by name — which works, because the
  arrow into station *n* is drawn at level *n*, when both ends are solid. The
  faint placeholders underneath must use coordinates, and must be retired by
  the level that draws the real arrow over them, or the two render as a double.
- **The last deck has nowhere to point.** Chapter 34's closing slide carried an
  empty NEXT band and simply stopped. It now ends the course — with counts the
  gates can verify, not sentiment.
- **Two items from the approved plan were built but never wired in.** The
  roadmap slide did not exist, and `section_slide` sat in `ccnadeck.py`
  unreferenced for six parts of building — the decks were reviewed slide by
  slide and nothing that is *absent* shows up in that review. Both are in now.
- **A divider named after the wrong heading is worse than none.** Using the
  nearest heading of any level put one in front of chapter 20's two-slide
  `\subsection{DORA}`; using the nearest `\section` and requiring more than one
  slide gives the chapter's real structure. Blocks now carry both levels.
- **Normalise the ground, not the palette.** Four domain colours sit near
  luminance 80; the gold of domain 5 is at 143, so its title slide washed out
  white type and swallowed the shield's gold ribbon. `_ground()` scales any hue
  down to a common luminance for full-bleed use and never lightens one, so the
  four that were already right are untouched and still match the book.
- **A split numbered list has to keep counting.** The second half of chapter
  22's nine-task lab restarted at "1. Extension" — and the book's prose refers
  to those tasks by number, so it was wrong, not merely untidy.
