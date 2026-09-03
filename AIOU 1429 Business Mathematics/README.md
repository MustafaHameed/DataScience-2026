# AIOU 1429 / 5405 — Business Mathematics: Solved Question Bank

Full worked solutions to the end-of-unit **Self Assessment Questions** of the
official AIOU course book *Business Mathematics*, Course Code 5405/1429,
Units 1–9, Department of Mathematics, AIOU Islamabad (2023 printing).

**Three editions**, same source and same palette, at three lengths:

| File | Language | Pages | Scope |
| --- | --- | --- | --- |
| `AIOU_1429_Solved_QuestionBank.pdf` | English | 84 | All nine units, every question, full working |
| `AIOU_1429_Urdu_Guess_Paper.pdf` | اردو | 35 | 27 expected questions, concise answers |
| `AIOU_1429_Urdu_Brief.pdf` | اردو | 19 | 16 highest-yield questions, capped at 21 pp |

The Urdu editions share the English one's box vocabulary —
سوال / حل / اہم اصول / عام غلطی / جواب / پڑتال.

**How the brief edition selects.** Weighted by what the paper actually asks
for, not spread evenly: unit 6 (determinants, inverse, Cramer) gets three
questions; units 1, 4, 7 and 9 get two each; the lighter units get one. Every
mistake named on its closing checklist is demonstrated by a question in the
document — that is why the Unit 2 question carries a third part on sampling
*without* replacement.

## Why these questions rather than a scanned past paper

Copies of the Spring 2019, Spring 2020, Spring 2022 and Autumn 2023 papers for
1429 circulate on study sites, but every one is a photograph or a locked PDF
with no recoverable text, so none could be transcribed accurately. The course
book is published by AIOU itself
([online.aiou.edu.pk](https://online.aiou.edu.pk/LIVE_SITE/SoftBooks/8405.pdf))
with a real text layer, and its end-of-unit sets are the pool the semester
assignments and the final paper are drawn from — one step upstream of the paper.

## Contents

| Unit | Topic | Source |
| --- | --- | --- |
| 1 | Probability theory | §1.9, pp. 25–26 |
| 2 | Random variables | §2.7, pp. 42–43 |
| 3 | Equations and inequalities | §3.9, pp. 64–65 |
| 4 | Linear equations | §4.13, pp. 93–94 |
| 5 | Matrices | §5.8, pp. 107–108 |
| 6 | Determinants and inverses | §6.8, pp. 131–132 |
| 7 | Limits and differentiation | §7.14, pp. 167–169 |
| 8 | Partial derivatives | §8.9, pp. 183–184 |
| 9 | Optimization | §9.9, pp. 205–206 |
| A | Formula sheet | — |
| B | Method in the exam hall | — |

Each solution carries a boxed **Answer**, most carry a **Check** line that
substitutes the result back into the original, and each question type carries a
**Where marks are lost** box naming the specific recurring mistake.

## Two defects in the printed book, handled explicitly

- **Unit 6, Q.5** (electronics materials) prints identical requirements for
  resistors and transistors, making the coefficient matrix singular and the
  system inconsistent. The document shows the determinant vanishing, explains
  what that means for a production plan, then solves a repaired version.
- **Unit 8, Q.7** (production function) prints `0.8l³`, which leaves `P`
  unbounded above, so no maximum exists. The document works the second-order
  test on the printed function, states that plainly, then solves the intended
  `0.8l² − 0.02l³` form.

Several sub-parts of Units 7 and 9 are illegible in the distributed PDF
(collapsed radicals and fraction bars in the text layer). They are named where
they occur and skipped rather than guessed at.

## Build

```powershell
.\build.ps1                # English edition, 84 pp (pdfLaTeX)
.\build_urdu.ps1           # Urdu, 27 questions, 35 pp (XeLaTeX)
.\build_urdu.ps1 -Brief    # Urdu brief, 16 questions, 19 pp (XeLaTeX)
.\build.ps1 -Clean         # remove aux files first
```

Requires MiKTeX (or TeX Live) with `latexmk`. On Windows both scripts borrow
the Perl bundled with Git for Windows if no `perl` is on `PATH`, since MiKTeX
ships `latexmk` without an interpreter. Each ends with a build report and exits
non-zero on any LaTeX error.

**The Urdu edition needs XeLaTeX, not pdfLaTeX** — pdfTeX has neither OpenType
shaping (Nastaliq ligatures) nor a bidirectional algorithm. It also needs the
**Urdu Typesetting** font, which ships with Windows; `build_urdu.ps1` checks
for it up front and fails with a clear message rather than emitting hundreds of
`nullfont` lines. On other platforms, substitute a Nastaliq or Urdu Naskh face
(e.g. Noto Nastaliq Urdu) in `bmurdu.sty`. That build report carries an extra
metric, **missing glyphs**, because the signature failure for an Urdu document
is that the font silently fails to load and every Arabic-script character is
dropped from the page.

`build_urdu.ps1 -Brief` additionally enforces the brief edition's 21-page cap
and fails the build if it is exceeded — the point of that edition is that it
fits in one sitting, so the constraint is checked rather than trusted.

Current status — English: **84 pages, 0 errors, 0 undefined refs, 0 overfull
boxes, 0 font warnings.** Urdu: **35 pages, 0 errors, 0 missing glyphs.**
Urdu brief: **19 pages, 0 errors, 0 missing glyphs, 0 overfull boxes.**

## Layout

```text
AIOU_1429_Solved_QuestionBank.tex   English master
bmsolutions.sty                     English style and pedagogy macros
parts/00_frontmatter.tex            title page, provenance, how to use
parts/unit1_probability.tex … unit9_optimization.tex
parts/appendix_formulae.tex         formula sheet + exam method

AIOU_1429_Urdu_Guess_Paper.tex      Urdu master, 27 questions
AIOU_1429_Urdu_Brief.tex            Urdu master, 16 questions  [compact]
bmurdu.sty                          Urdu style (fontspec + polyglossia + bidi)
parts_ur/00_front_ur.tex            title page, how to use, exam map
parts_ur/01_units1to4_ur.tex        units 1-4
parts_ur/02_units5to9_ur.tex        units 5-9
parts_ur/03_formula_ur.tex          formula sheet + eight common mistakes
parts_ur_brief/00_front_brief.tex   title page + one orientation page
parts_ur_brief/01_questions_brief.tex   the sixteen selected questions
parts_ur_brief/02_formula_brief.tex     condensed formula sheet + mistakes

build.ps1  build_urdu.ps1
```

Both Urdu documents share `bmurdu.sty`. The brief edition loads it as
`\usepackage[compact]{bmurdu}`, which tightens margins, leading (1.55 → 1.38)
and box padding. The option defaults to **off**, so the 35-page edition is
unaffected — verified by rebuilding it after the option was added and
confirming identical metrics (a byte-comparison is meaningless here, since
PDFs embed a creation timestamp).

All colours, boxes and macros live in the `.sty` files — `question`,
`solution`, `keyidea`, `pitfall`, `\ans`, `\vcheck` in `bmsolutions.sty`;
`sawal`, `hal`, `nukta`, `ghalti`, `\jawab`, `\parakh` in `bmurdu.sty`. Do not
duplicate a preamble into part files.

### Two constraints worth knowing before editing `bmurdu.sty`

- **Package order.** `polyglossia` pulls in `bidi`, which insists on being
  loaded last. Every other package must be required *above* it; anything
  loaded afterwards that patches sectioning, floats or tabulars silently
  breaks the right-to-left handling.
- **Tables.** `bidi` reverses the column order of every `tabular` to match the
  page direction. That is correct for a table of Urdu labels and wrong for one
  whose columns carry an intended left-to-right sequence, so order-sensitive
  tables are wrapped in the `ltrtab` environment.

## Status

Prepared for private study. Not an AIOU publication and carries no official
endorsement. It is intended for checking your own working — not for copying
into an assignment for submission.
