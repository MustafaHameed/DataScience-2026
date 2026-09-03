# AIOU 1429 / 5405 — Business Mathematics: Solved Question Bank

Full worked solutions to the end-of-unit **Self Assessment Questions** of the
official AIOU course book *Business Mathematics*, Course Code 5405/1429,
Units 1–9, Department of Mathematics, AIOU Islamabad (2023 printing).

**Output:** `AIOU_1429_Solved_QuestionBank.pdf` — 84 pages.

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
|---|---|---|
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
.\build.ps1           # normal build
.\build.ps1 -Clean    # remove aux files first
```

Requires MiKTeX (or TeX Live) with `latexmk`. On Windows, `build.ps1` borrows
the Perl bundled with Git for Windows if no `perl` is on `PATH`, since MiKTeX
ships `latexmk` without an interpreter. The script ends with a build report
(pages, errors, undefined refs, overfull/underfull boxes, font warnings) and
exits non-zero on any LaTeX error.

Current status: **84 pages, 0 errors, 0 undefined refs, 0 overfull boxes,
0 font warnings.**

## Layout

```
AIOU_1429_Solved_QuestionBank.tex   master file, inputs the parts
bmsolutions.sty                     all style and pedagogy macros
parts/00_frontmatter.tex            title page, provenance, how to use
parts/unit1_probability.tex … unit9_optimization.tex
parts/appendix_formulae.tex         formula sheet + exam method
build.ps1
```

All colours, boxes (`question`, `solution`, `keyidea`, `pitfall`) and macros
(`\ans`, `\vcheck`, `\srcnote`, `\saq`, `\dt`, `\mat`) live in
`bmsolutions.sty`. Do not duplicate the preamble into part files.

## Status

Prepared for private study. Not an AIOU publication and carries no official
endorsement. It is intended for checking your own working — not for copying
into an assignment for submission.
