# Analysis of Algorithms — Fall 2026

Complete teaching materials for the BSIT *Analysis of Algorithms* course:
a 242-page handout, 24 lecture decks, a one-page schedule, and 23 runnable
C++ experiments whose results are quoted in the text.

**Dr. Mustafa Hameed · Department of Information Technology · The Islamia
University of Bahawalpur**

---

## What is here

| Artefact | Size | Build with |
|---|---|---|
| `AoA_Fall2026_Handout.pdf` | 242 pp., 23 chapters, 7 appendices, 39 figures | `.\build.ps1` |
| `slides/out/*.pptx` | 24 decks, 1,140 slides | `python slides\build_slides.py --all` |
| `AoA_Fall2026_Lecture_Schedule.pdf` | 1 page, 16 weeks | `pdflatex AoA_Fall2026_Lecture_Schedule.tex` |
| `code/*.cpp` | 23 programs, 4,573 lines | `python code\run_all.py` |

---

## HEC provenance

The course is **compulsory and specified**, so the content is bounded by a
published outline rather than improvised.

| Edition | Title | BSIT status | Semester | Credits | Prerequisite |
|---|---|---|---|---|---|
| 2017 | Design & Analysis of Algorithms | absent (BSCS core only) | IV | 3-0 | Data Structures and Algorithms |
| **2023** | **Analysis of Algorithms** | **Computing Core, compulsory** | **7** | **3 (3-0)** | **Data Structures** |
| 2025 | Design and Analysis of Algorithms | n/a (CS-only booklet) | IV | 3+0 | none published |

**HEC 2023 (p. 65) is the authority for a BSIT course file**, and its outline
is word for word identical to 2017's (p. 89). Eight CLOs are published in both.

### Four defects in the published documents, recorded rather than inherited

All four are stated in the handout's front matter, with the correction used:

1. **The 2023 CLO table's labels are mis-aligned with their own text** —
   `CLO-1` sits against the first half of outcome 1 and `CLO-2` against its
   continuation, leaving 3, 5 and 7 dangling. The correctly-numbered 2017 list
   is used.
2. **The Bloom's Taxonomy column is blank in both editions**, for all eight
   outcomes, although the neighbouring Data Structures table fills it in. The
   handout assigns levels and says it has done so.
3. **"Corman"** is HEC's misspelling of **Cormen**, in both editions.
4. **CLRS 3rd ed. (2009) is cited; the 4th ed. (2022) is current.**
   Appendix E cites the 4th and explains the discrepancy.

### Three topics taught that the outline does not name

Justified in the Course Specification, because an accreditation reviewer will
ask:

- **Ch 10, the comparison lower bound.** The first result in the course about
  a *problem* rather than an algorithm; without it "merge sort is optimal" is
  an assertion.
- **Ch 17, amortised analysis.** Ch 13's hash-table claim and Ch 19's
  union–find cannot be stated honestly without it, and CLO-3 asks for the time
  complexity of *sequences* of operations.
- **Ch 19, minimum spanning trees.** Absent from the outline text and
  **required by CLO-7 by name**.

---

## The three strands

Every chapter does the same three things, and the handout's subtitle names
them.

**Design** — the algorithm, as pseudocode, in the house typography.

**Proof** — theorems in boxes, proofs as prose. Where a proof teaches
something it is given in full; where it does not, the result is stated and
cited. Proofs are deliberately *not* boxed, so that the slide pipeline carries
them to the speaker notes rather than onto the screen.

**Measurement** — one experiment per chapter, compiled and run on the machine
described in Appendix A, with the output pasted into the text. This is the
strand that distinguishes the course from the textbook, and `code/run_all.py`
is what keeps it honest.

### Results the measurement strand produced

A sample, to show what the strand is for:

| Chapter | Measured |
|---|---|
| 2 | Random input costs **exactly half** the reversed input (0.54, 0.56, 0.54, 0.50, 0.49) — the inversion count, confirmed |
| 3 | $T/n^{2.0}$ flat within 8% while $T/n^{1.9}$ rises 46% and $T/n^{2.1}$ falls 27% |
| 4 | A bug with **zero** wrong answers and **zero** invariant violations in 132,311,650 audited iterations — that never terminates |
| 8 | The 3.78× sorted-to-random gap factors as **1.89× (comparisons) × 2.0× (branch misprediction)**, confirmed by a branchless rewrite |
| 9 | Depth exactly $n-1$; at $n=16{,}000$ the program used **99.9%** of its 1 MB stack, and at 32,000 it crashed |
| 10 | Merge sort is within **1.0%** of the information-theoretic floor at $n=10^6$ |
| 13 | Chain lengths match Poisson to within 2% across seven lengths; a power-of-two modulus used **1.6%** of the table |
| 16 | Eight doublings of knapsack's $W$, each adding **one bit** of input, each doubling the time |
| 19 | Union–find cost per find: 1.62 → 1.88 across a **512-fold** range of $V$ |

**Five first attempts failed and the failures are in the handout**, because
they are more instructive than the successes: a benchmark that allocated
inside the timed region; a bit-packed matrix slower than the byte version; a
load-factor experiment that sampled in insertion order; a Rabin–Karp worst case
that produced zero comparisons; and a Dijkstra counter-example on which
Dijkstra was right.

---

## Building

```powershell
.\build.ps1                      # handout; gates on 0 errors/refs/boxes/warnings
python checklog.py               # quiet build report
python checkfigs.py              # TikZ line-break and reserved-name traps
python checklabs.py              # lab listings: ASCII, <= 76 cols, step markers
python code\run_all.py           # compile, run and check all 23 experiments
python code\sync_labs.py         # handout listings really are in the sources
python slides\texparse.py --all  # every chapter parses; Q/A counts match
python slides\figures\extract_figures.py
python slides\build_slides.py --all
python slides\overview.py
python slides\qa.py --render     # static checks, then LOOK at the sheets
pdflatex AoA_Fall2026_Lecture_Schedule.tex
```

### What "clean" means

| Gate | Required |
|---|---|
| `build.ps1` | 0 errors, 0 undefined refs, 0 overfull, 0 underfull, 0 font warnings, 0 PDF-string warnings |
| `checkfigs.py` / `checklabs.py` | clean |
| `run_all.py` | 42/42 exact checks pass |
| `sync_labs.py` | every line of C++ in the handout is in the source (522 lines) |
| `texparse.py --all` | all 23 chapters parse completely; 3 questions = 3 answers each |
| `qa.py` | static checks clean |
| `qa.py --render` | **all 24 contact sheets looked at** |

The last one is not automatable and is not optional. It is the only gate that
caught the lost `O`/`Ω`/`Θ` in every asymptotic bound, the proofs appearing as
body slides, the dropped `\lg`, "Theorem ??" wherever a heavy paragraph cited a
result, the lab cards reading "lines of C", and three blocks of the previous
course's prose surviving in `slides/overview.py`. Two of those were inside
rendered images, where even a text scan of the `.pptx` could not reach them.

---

## Authoring contract

Each chapter file follows the same shape, parsed by position:

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{<part title>}
\begin{outcomes} ... \end{outcomes}      % 5 Bloom-verb outcomes
\begin{prereq}   ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
\section{...}                            % 4-6 sections, mixing:
  definitionbox / conceptbox / alertbox
  theorem / lemma / corollary            % boxed -> card slides
  proof                                  % NOT boxed -> speaker notes
  \dsfig{\aoaalg{...}{...}}{...}{...}    % algorithms as figures
  \begin{worked} ... \end{worked}        % every arithmetic step shown
  longtable with L{} columns
\begin{pitfall}        ... \end{pitfall}
\begin{lab}[title]     ... \end{lab}     % C++, steps marked // --- n.
\begin{checkpoint}     ... \end{checkpoint}   % exactly 3
\begin{chaptersummary} ... \end{chaptersummary}
\begin{reviewq}        ... \end{reviewq}      % 7: 2 MCQ, 3 Short, 2 Applied
```

### Two decisions worth knowing before editing

**Algorithms are figures, not code blocks.** Every algorithm is authored once
as `\dsfig{\aoaalg{Name}{algpseudocode body}}`. The handout gets textbook
typography; the slide pipeline compiles the same source to SVG and gets vector
art with free line-by-line reveal through `\stepvis`. No parser change was
needed — it is a figure.

**Proofs are prose, not boxes.** `theorem`, `lemma` and `corollary` are
`tcolorbox`es and reach the slides as cards. `proof` is plain `amsthm` and is
folded by `texparse.py` into the notes of the result it proves. A full proof on
a slide is a bad slide.

### Table widths

`\tabcolsep` eats about 0.42 cm per internal gap. Totals that fit:

| Columns | At top level | Inside a `conceptbox`/`alertbox` |
|---|---|---|
| 2 | 15.0 cm | 14.5 cm |
| 3 | 14.5 cm | **14.3 cm** |
| 4 | 14.5 cm | 14.3 cm |

The in-box 3-column limit is the one that bites; six tables overflowed by
exactly 1.32 pt before being narrowed.

---

## Boundary with other courses

Appendix F states it in full. In short:

- **Data Structures** owns the structures and their operations; this course
  owns their *analysis*. About a third of the HEC outline overlaps, and those
  chapters (12, 13) re-analyse rather than re-implement.
- **Theory of Automata** owns computability — decidability, the halting
  problem. This course owns complexity: P, NP, NP-completeness.
- **Discrete Structures** supplies the mathematics, collected in Appendix B.

Appendix F also records that HEC's 2017 postgraduate elective *Advanced
Algorithm Analysis* was deleted in 2023 and 2025 with no replacement, leaving
network flow, linear programming and approximation algorithms uncovered in the
BSIT programme.

---

## Layout

```
Analysis of Algorithms Fall 2026/
├── AoA_Fall2026_Handout.tex        master: 6 parts, 23 chapters, 7 appendices
├── aoahandout.sty                  house style + theorem/proof/pseudocode
├── AoA_Fall2026_Lecture_Schedule.tex   one page, generated from Appendix G
├── build.ps1                       handout build with the log gate
├── checklog.py  checkfigs.py  checklabs.py
├── parts/                          00_frontmatter, 6 openers, 23 chapters, A–G
├── code/                           23 experiments + run_all.py + sync_labs.py
├── assets/                         logos for the schedule page
└── slides/                         the deck pipeline — see slides/README.md
```

**`parts/G_schedule.tex` is the source of record for the schedule.** Change it
first; the slide pipeline parses that appendix, not the one-page PDF.

**`parts/C_answers.tex` has a hard format contract**: `\ansch{<N> --- <title>}`
followed immediately by an `ans` environment of **exactly three** items.
`texparse.py --all` checks it.

---

## Not done

- The curriculum map in `../HEC AI-DS Curriculum Map/` has **no entry for this
  course at all** — no status row, no course card, no overlap row. Out of scope
  here and recorded as a separate, smaller job.
- No git commit has been made; nothing here is staged.
