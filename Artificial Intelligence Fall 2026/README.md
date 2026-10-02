# Artificial Intelligence — Fall 2026

Unified lecture handout and lecture slides for the BSIT **Artificial Intelligence**
course, written to the HEC curriculum. The whole course is one document: six parts,
22 chapters, seven appendices, from what makes a program intelligent to a working
rule engine that explains itself. Every chapter has a matching PowerPoint deck
generated from the handout, and a runnable Python lab.

**Current build:** 290 pages · 52 figures, all numbered, captioned and listed ·
22 decks, 1,133 slides · zero LaTeX errors, zero undefined references, zero
overfull or underfull boxes, zero font warnings · slide QA static checks clean ·
all 22 labs run green.

**Status: complete.** All 22 chapters are written in full — prose, figures,
worked examples, pitfalls, checkpoints, summaries, review questions and a runnable
lab apiece — across 290 pages. Every number quoted in the text is produced by that
chapter's lab and checked by an assertion, so the prose and the code cannot drift
apart. All seven appendices are written, including worked answers to every
checkpoint.

**This course is symbolic, on purpose.** It teaches the agent that is *told* the
rules and must reason with them: agents, search, constraints, games, planning,
logic, rule engines, uncertainty and fuzzy logic. Learning is one signpost chapter
(17) — the three paradigms, the perceptron worked by hand, and the exclusive-or
limit — and then the course hands over to *Machine Learning* in Semester 6.
Appendix F states every boundary.

**One running example.** Every example comes from IT project management: task
dependencies and release plans, sprints and staffing, tickets and escalations,
defects and risk. There are no other application domains and no "four lenses"
panels — the front matter lists which recurring example belongs to which chapter.

## What the course is built on

The content follows the HEC material collected in `../HEC AI-DS Curriculum Map/`
(the three curriculum booklets and the department's analysis of them):

- **HEC 2017** publishes the more detailed outline — it names propositional logic,
  first-order logic, constraint satisfaction, the three learning paradigms and
  fuzzy logic explicitly. Every clause is taught, and the Course Specification
  lists each one against the chapter and the named section that covers it.
- **HEC 2023** makes the course a **computing core for every degree**, Semester 3,
  3 (2-3), prerequisite *Object Oriented Programming*. It drops the formal-logic
  emphasis and adds the four case studies (General Problem Solver, Eliza, Student,
  Macsyma), artificial neural networks, natural language processing, and mandates
  Python. All of it is covered.
- **HEC 2025** publishes five learning outcomes and nothing else. Their union with
  the earlier editions gives the seven course CLOs.

Three places where the handout deliberately departs from the outlines' *grouping*
(not their content) are stated in the specification: search is taught before logic,
planning sits with search rather than with knowledge, and learning gets one
signpost chapter rather than a unit.

**Four topics go beyond the letter of both outlines** — planning, semantic
networks, rule-based expert systems and Bayesian networks. The specification lists
each with the clause it rests on, and Appendix F sets out the boundary with the
elective *Knowledge Representation and Reasoning*, whose 2023 outline overlaps
Parts IV and V.

## Structure

**Part I — Foundations of Intelligent Systems** (1 What AI Is · 2 Intelligent
Agents and Task Environments · **3 Representations and Search in Python**)

**Part II — Problem Solving by Search** (**4 Problem Formulation and Uninformed
Search** · **5 Informed Search and Heuristics** · 6 Local Search and Optimisation)

**Part III — Constraints, Games and Plans** (7 Constraint Satisfaction ·
8 Adversarial Search and Game Playing · 9 Planning)

**Part IV — Knowledge, Logic and Reasoning** (10 Knowledge Representation ·
**11 Propositional Logic and Inference** · 12 First-Order Logic and Resolution ·
13 Rule-Based Expert Systems)

**Part V — Reasoning Under Uncertainty** (14 Uncertainty and Decision Making ·
15 Bayesian Networks · 16 Fuzzy Logic)

**Part VI — Learning, Language and Responsible AI** (17 Learning from Examples ·
18 Natural Language Processing · 19 Classic AI Systems and How to Analyse Them ·
20 Recent Trends and Applications · 21 Responsible AI and Its Limits ·
22 The AI Project)

**Appendices** A Python Reference · B Algorithm and Formula Sheet · C Answers to
Checkpoints · D Glossary · E Resources · **F The Boundary with the Neighbouring
Courses** · G Sixteen-Week Schedule

Chapters 3, 4, 5 and 11 are the load-bearing ones: the search toolkit is reused by
constraint satisfaction, adversarial search and planning, and propositional
inference underpins first-order logic and the rule engine. Week 8 carries Chapter 9
alongside the midterm on purpose — STRIPS progression search reuses the frontier of
Chapters 4 and 5, so the lecture is light in a revision week.

Every worked example was computed, not typed. Where a lab's measurement disagreed
with the prose, the prose was corrected: Chapter 3's representation saving is 8×
(22 states against 184 sequences), Chapter 5's A\* expands 33 states against
uniform-cost search's 46, and Chapter 6's four local-search methods are compared at
an equal budget of 4,000 objective evaluations.

## Build

```powershell
.\build.ps1            # build AI_Fall2026_Handout.pdf
.\build.ps1 -Clean     # remove aux files, then build
```

Requires MiKTeX (or TeX Live) with `pgfplots`, `tcolorbox`, `fontawesome5`,
`booktabs`, `longtable`, `ragged2e`, `needspace`, `float`, `placeins` and
`caption`. `build.ps1` borrows Git for Windows' `perl.exe` for `latexmk` when none
is on `PATH`, prints a report (pages, errors, undefined references, bad boxes, font
warnings) and exits non-zero on any error.

Two standalone one-pagers build independently:

```powershell
pdflatex AI_Fall2026_Lecture_Schedule.tex
pdflatex ..\AI_ML_DS_Fall2026_ProgressionAtAGlance.tex
```

The second spans all three courses and lives at the repository root. It shows the
AI → ML → Data Science progression and, in one table, which course **owns** each
shared topic and which merely signposts or recaps it.

### Labs

```powershell
cd code
python sync_labs.py          # extract every chapter's lab into code/
python sync_labs.py --run    # ... and run them all
python sync_labs.py --run 5  # ... run chapter 5 only
```

The handout is the single source of the lab code: `sync_labs.py` copies each
chapter's *Try It Yourself* listing to `code/chNN_<short-name>.py`, checks that it
is pure ASCII (pdfLaTeX listings cannot take UTF-8) and no wider than 78 columns,
and runs it. **Edit the chapter, not the extracted file.** The labs need only
Python 3.10 or later; Chapters 3, 6, 15 and 16 also use NumPy, and everything else
is the standard library. No lab needs a GPU or network access. Chapter 6's lab
takes about 18 seconds (200 runs × 4 methods); the rest finish in under a second.

### Slides

```powershell
cd slides
python texparse.py --all                 # audit: every chapter parses completely
python figures\extract_figures.py        # compile the handout's figures (cached)
python build_slides.py --all             # all 22 decks into slides\out
python build_slides.py --ch 5            # one chapter
python qa.py                             # static checks on every deck
python qa.py --render --ch 5             # also export PNGs and a contact sheet
```

Build the handout first: slide page references come from its `.aux`. Maths is
typeset with LuaLaTeX on demand; figures are compiled from the handout's own TikZ,
so a figure corrected in the handout reaches the slides on the next build.

## Layout

| Path | Contents |
| --- | --- |
| `AI_Fall2026_Handout.tex` | Master file; `\input`s everything |
| `aihandout.sty` | Single source of style; a sibling of `../Machine Learning Fall 2026/mlhandout.sty` |
| `AI_Fall2026_Lecture_Schedule.tex` | Standalone one-page weekly schedule |
| `parts/00_frontmatter.tex` | Title, how to use, **Course Specification (HEC)**, course map, running example, notation, self-check |
| `parts/00_partN_opener.tex` | One reading-path opener per part |
| `parts/01_*.tex` … `parts/22_*.tex` | The 22 chapters, in teaching order |
| `parts/A_*.tex` … `parts/G_*.tex` | The seven appendices |
| `code/` | `sync_labs.py` and the extracted labs |
| `slides/` | The deck pipeline, its assets, figure renders and the decks in `slides/out/` |

## Authoring conventions

Every chapter uses the same furniture, in the same order, as the Machine Learning
and Data Science handouts — and the slide pipeline depends on it:

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{<part title>}
\begin{outcomes} ... \end{outcomes}
\begin{prereq} ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
...content: definitionbox / conceptbox / alertbox / examplebox / \dsfig ...
\begin{worked}[title] ... \end{worked}
\begin{pitfall} ... \end{pitfall}
\begin{lab}[title] ... one lstlisting ... \end{lab}
\begin{checkpoint} ... three items ... \end{checkpoint}
\begin{chaptersummary} ... \end{chaptersummary}
\begin{reviewq} ... \end{reviewq}
```

- Every figure goes through `\dsfig` (or `\dsfigh` to force it in place); all are
  native TikZ/pgfplots. Inside a figure use only `\term` among the custom macros —
  the slide pipeline compiles figures standalone.
- Checkpoint answers go in `parts/C_answers.tex` under `\ansch{N --- Title}`,
  three per chapter; the slides reveal them after each question.
- The schedule table in `parts/G_schedule.tex` must keep its five columns (week,
  topic, chapters, lab, assessment): the slides read it for the week and lab. The
  week must be a bare integer, the chapters column only digits and commas, an empty
  assessment cell literally `---`, and **no chapter may appear in two weeks** —
  `week_of()` returns the first match.
- Paragraph columns use `L{<width>}`, never `p{<width>}`. A six-column table needs
  its widths to sum to about 12.8 cm, not 15.6 — the bold header cells are usually
  what overflows first.
- Displayed maths belongs outside list items (the slide parser flattens items).
- No `\textsc` inside bold text (Latin Modern has no bold small caps), and no TikZ
  style named `pos`, `out` or `in` — they are built-in keys.

### Three traps found while writing this course

- **A `\\` inside a nested `{...}` group in an aligned TikZ node breaks the
  picture**, with a misleading error deep in `\tikz@finish` and a
  `Division by 0` from `graphics`. Keep every `\\` at the top level of the node
  text; wrap each line in its own group instead, or use `text width` and let it
  wrap. This cost an hour in Chapter 2.
- **`above=1pt of {(x,y)}`** positions a node but leaks its closing parenthesis as
  text in `nullfont`. Use `anchor=` with an explicit `at (x,y)`.
- **`\owns`, `\th`, `\sign` and `\nil` are already defined** by LaTeX or amssymb,
  so `\newcommand` for any of them fails silently and the old meaning is used. The
  progression one-pager uses `\tOwn`, `\tHead`, `\tSign`, `\tNil`.

## The pipeline is a copy

`slides/` is a copy of `../Machine Learning Fall 2026/slides/`, not a shared
module, so that each course can be rebuilt without touching the others. File names
are unchanged to keep a diff between the copies meaningful. The adaptations:

| File | What changed |
| --- | --- |
| `texparse.py` | `MASTER`/`AUX` paths, `PART_TITLES`, **`F_schedule.tex` → `G_schedule.tex`** (Appendix F is the boundary note here), and the maths symbol table |
| `dsdeck.py` | `course=`, core-properties title/subject/keywords, the title-slide subtitle |
| `build_slides.py` | `PARTS = [(1,3),(4,6),(7,9),(10,13),(14,16),(17,22)]`, deck filename prefix `AI_` |
| `qa.py` | the `AI_Ch(\d\d)_` regex and chapter filter |
| `figures/extract_figures.py` | `STY` points at `aihandout.sty` |
| `pptxfx.py` | the default **theme name** — without this every AI deck carries a PowerPoint theme called "Machine Learning Fall 2026" |
| `code/sync_labs.py` | verbatim; it contains no course-specific string |

Two improvements here are worth porting back to `mlhandout.sty` and
`dshandout.sty`:

- The `\rowcolors` bleed patch covers **`array` and `aligned`** as well as
  `cases`/`matrix`/`pmatrix`/`bmatrix`. Truth tables and resolution proofs use
  `array`, and without the patch every one following a striped `longtable` came out
  painted.
- `texparse.py`'s `_MATH_SYM` gains the **logic operators** (`\models`, `\vdash`,
  `\top`, `\bot`, `\wedge`, `\vee`, `\lnot`, `\implies`, `\iff`, `\therefore` and
  the rest). An unlisted symbol reaches the slides as *nothing at all* — the same
  fault as the `\varnothing` bug recorded in the Machine Learning README.

`dsdeck.py` still carries the Data Science pipeline's four-lens card code
(`lens_slides`, and the `fourlenses` box in `texparse.py`). No chapter here uses
it, so it never runs; it is kept only to keep this copy diffable against the
others.
