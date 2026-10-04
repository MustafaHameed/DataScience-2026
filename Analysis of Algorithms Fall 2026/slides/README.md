# Lecture decks

24 decks, 1,140 slides, generated from the handout. Nothing here is authored
by hand: change `../parts/*.tex` and rebuild.

```powershell
python texparse.py --all              # every chapter parses; Q/A counts match
python figures\extract_figures.py     # 39 figures -> SVG + PNG
python build_slides.py --all          # 23 chapter decks
python overview.py                    # the course overview deck
python qa.py --render                 # static checks, then LOOK at the sheets
```

---

## The pipeline

| Stage | File | Does |
|---|---|---|
| parse | `texparse.py` | reads `../parts/*.tex` into typed elements |
| figures | `figures/extract_figures.py` | compiles each `\dsfig` body standalone |
| design | `dsdeck.py` | the house design system: palette, styles, layouts |
| build | `build_slides.py` | one deck per chapter |
| overview | `overview.py` | the course-level deck |
| polish | `pptxfx.py` | OOXML injection: Morph, animation, SVG, alt-text |
| check | `qa.py` | overflow detection, picture and notes counts, rendering |

Supporting: `measure.py` (text metrics), `mathrender.py` (LuaLaTeX for display
maths), `icons.py` (Font Awesome as recolourable vectors).

---

## What differs from the sibling courses' copies

This pipeline was copied from *Virtual Systems and Services* and changed in
seven places. If you are porting it on to a seventh course, these are the
changes that are about **this** course rather than about the pipeline.

**1. Three result environments.** `theorem`, `lemma` and `corollary` were added
to `BOX_OPT` in `texparse.py` and to `STYLES` in `dsdeck.py`, with the
handout's own colours and icons. They share one counter, so the parser
reconstructs the number and the card reads **"THEOREM 3.1"** — the prose cites
results by number, and a card headed only "THEOREM" would break the link.

**2. Proofs are folded into speaker notes.** `proof` is in `BOX_PLAIN` but
never becomes slides: `_fold_proofs()` attaches each proof to the nearest
preceding result. Without this, Chapter 3 produced nine slides all titled
"Proving a Bound From the Definition". Chapter 3 went 66 → 57 slides when this
landed.

**3. The course's maths macros.** `_MATH_WRAP` renders `\bigO{...}`,
`\bigOm{...}` and `\bigTh{...}`; `_MATH_OPNAME` gained `lg`, `gcd`, `lcm`.
Without them **every asymptotic bound in the course reached the slides with its
O, Ω or Θ silently removed** — `\bigO{n^{2}}` became "n²". This was the single
most damaging defect found, and only the contact sheet found it.

**4. Superscripts survive in titles.** `Box.title_text` goes through
`plain_scripts()`, which maps sub/superscript runs to Unicode. "A quadratic is
O(n2)" became "A quadratic is O(n²)".

**5. C++.** `CppLexer` was added to the Pygments map (under `c++`, `cpp` and
`cppaoa`) and to `_lang_name`. The listing-language regex became
`language=([A-Za-z+]+)` — without the `+` in the class, `language=C++` parsed
as `C` and the lab cards read "73 lines of C".

**6. Algorithms compile standalone.** `figures/extract_figures.py` generates a
preamble from `aoahandout.sty`; it now also emits `tcolorbox`, `algpseudocode`,
the lifted `\algrenewcommand` block and a standalone `\aoaalg` with a fixed
width (`\linewidth` is meaningless on a standalone canvas).

**7. `.cpp` is a lab extension.** `lab_code_file()` looks for `.cpp` as well as
`.py` and `.R`.

**8. `\ref` is resolved before LuaLaTeX sees it.** `resolve_refs()` now expands
plain `\ref{...}` from the handout's `.aux`, not only `\chref` and `\dsref`.
This course cites results by number in prose — "Theorem~\ref{thm:qsworst}
predicted the split" — and a paragraph heavy enough to be re-rendered by
LuaLaTeX is compiled as a *standalone* document with no `.aux`, so an
unresolved `\ref` printed **"Theorem ??"** on the slide while the handout was
perfectly correct. Sixteen chapters cite a theorem or lemma this way.

Also: four new icons (`award`, `puzzle-piece`, `angle-double-right`,
`list-ol`), and `texparse.py --all` subtracts folded proofs before reporting
dropped content — otherwise 19 of 23 chapters warn, and an audit that always
warns is an audit nobody reads.

---

## Deck grammar

Each chapter deck runs: title → where this chapter sits → outcomes → key terms
→ agenda → *(per section: divider, then content)* → lab briefing → lab code
walkthrough → checkpoints with answers → chapter summary → review questions →
end card.

Content slides follow the handout's furniture:

| Source | Becomes |
|---|---|
| `definitionbox` / `conceptbox` / `alertbox` | a card, revealed point by point |
| `theorem` / `lemma` / `corollary` | a card headed with its number |
| `proof` | **speaker notes on the preceding result's card** |
| `worked` | a sequence of step slides |
| `\dsfig{\aoaalg{...}}` | the algorithm as **vector art**, not a code panel |
| `pitfall` | a struck-through list |
| `lab` | a briefing card, then one slide per `// --- n.` step |
| `checkpoint` | question slide, then answer slide from Appendix C |

Branding is the department's: Dr. Mustafa Hameed, IUB crest and DIT shield,
BSIT · Fall 2026.

---

## The gate that matters

`qa.py` without `--render` runs static checks — text overflow, picture counts,
notes present. It passes on decks that are badly wrong.

`qa.py --render` rasterises every slide and tiles each deck into a contact
sheet in `render/`. **Look at them — all 24.** In this course that gate caught:

- every asymptotic bound missing its `O`, `Ω` or `Θ`
- proofs appearing as body slides
- `\lg` dropped from `⌊lg n⌋`, changing the claim
- "Theorem **??**" wherever a heavy paragraph cited a result
- lab cards reading "lines of C" for C++
- the previous course's subtitle on every title slide
- **three blocks of the previous course's prose in `overview.py`** — a slide
  about hypervisors and GPUs, and a "bring a laptop with virtualization
  enabled" card. These survived the retarget because they never used a word it
  searched for; `scanprose` in the scratchpad was written afterwards to look
  for the *vocabulary* of the sibling courses rather than their names.
- Chapter 10's title overflowing the chapter chip

Not one of those was caught by a clean LaTeX build or by the static checks.
Two of them — the missing `Θ` and the "??" — were inside **rendered images**,
so even a text scan of the `.pptx` could not see them.

---

## If a figure fails to compile

`extract_figures.py` compiles each `\dsfig` body as a standalone document, so a
figure can fail there while the handout builds. Two traps, both now caught by
`../checkfigs.py`:

- **`\\` in a TikZ node with no `align`** — "Not allowed in LR mode". Add
  `align=left` or `align=center`.
- **A style named after a reserved key** — `cap`, `step`, `pos`, `in`, `out`,
  `grid`. Rename it.
