# Machine Learning — Fall 2026

Unified lecture handout and lecture slides for the BSIT **Machine Learning**
course, written to the HEC curriculum. The whole course is one document: six
parts, 22 chapters, six appendices, from what learning means to reinforcement
learning, responsible practice and an end-to-end project. Every chapter has a
matching PowerPoint deck generated from the handout, and a runnable Python lab.

**Current build:** 210 pages · 72 figures, all numbered, captioned and listed ·
22 labs, all run · 22 decks, 954 slides · zero LaTeX errors, zero undefined
references, zero overfull or underfull boxes, zero font warnings · slide QA
static checks clean.

**One running example.** Every example in the handout comes from IT project
management: effort and schedules, tasks and sprints, tickets and releases, risk,
staffing and technical debt. There are no other application domains and no "four
lenses" panels — the front matter lists which recurring example belongs to which
chapter. The set text's classic tables appear with their attributes renamed
(EnjoySport becomes release approvals in Chapter 4; PlayTennis becomes sprint
outcomes in Chapters 9 and 10), so every number can still be checked against
Mitchell.

## What the course is built on

The content follows the HEC material collected in
`../HEC AI-DS Curriculum Map/` (the three curriculum booklets and the
department's analysis of them):

- **HEC 2023** (BS Computing Disciplines) is the only edition that publishes an
  outline for *Machine Learning*: 3 (2-3) credit hours, prerequisite *Artificial
  Intelligence*, five CLOs and the Mitchell/Murphy set texts. **Every topic of
  that outline is taught** — the Course Specification at the front of the
  handout lists each one against the chapter that covers it.
- **HEC 2025** (BS Computer Science) publishes learning outcomes only, a
  different list in each specialization that carries the course. Their union
  adds evaluation metrics, bias–variance reasoning, PCA, modern toolkits,
  end-to-end solutions and ethics (A3: Valuing); all are met, and the
  specification traces each one. The seven course CLOs are the union of both
  editions.
- The curriculum map's **overlap analysis** decides the boundaries: this course
  owns prediction (generalisation, model selection, evaluation, ensembles,
  regularisation); association rules, outlier mining and graph mining are left
  to Data Mining, and deep architectures to Deep Learning.

Three places where the handout deliberately departs from the HEC outline's
*grouping* (not its content) are stated in the specification: k-NN is taught
with the supervised learners, HMMs open Part V rather than sitting under
reinforcement learning, and overfitting and evaluation come before the
algorithms rather than after them.

The **Summer 2026** handout (`../Summer_2026/Machine_Learning_Handout.pdf`) is a
24-page overview for an eight-week session. It omits concept learning, version
spaces, self-organizing maps, semi-supervised EM, HMMs and MDPs, which the HEC
outline requires. Appendix E says so and positions it as a revision summary.

## Structure

**Part I — Foundations** (1 What ML Is · 2 Mathematics · 3 Workflow and Toolkit · 4 Concept Learning and the Version Space)

**Part II — Linear Models and Generalization** (5 Linear Regression · 6 Logistic Regression · **7 Overfitting, Bias–Variance and Regularisation** · **8 Measuring Classifier Performance**)

**Part III — Supervised Learning Algorithms** (9 Decision Trees · 10 Bayesian Learning and Naive Bayes · 11 k-Nearest Neighbours · 12 Support Vector Machines · 13 Artificial Neural Networks · 14 Ensemble Learning)

**Part IV — Learning Without Full Labels** (15 Clustering · 16 PCA and Self-Organizing Maps · 17 Semi-Supervised Learning and EM)

**Part V — Sequences, Decisions and Reinforcement** (18 Hidden Markov Models · 19 Markov Decision Processes · 20 Reinforcement Learning)

**Part VI — Responsible Practice** (21 Responsible and Ethical ML · 22 The ML Project)

**Appendices** A Python and scikit-learn · B Formula Sheet · C Answers to
Checkpoints · D Glossary · E Resources · F Sixteen-Week Schedule

Chapters 7 and 8 are the hinge of the course and are taught before the
algorithms on purpose. The schedule maps the chapters onto sixteen weeks of two
lecture and three lab hours, with the midterm after Week 8 and the project from
Week 11.

Every worked example was computed, not typed: the numbers in the handout were
checked in Python before they were written, and the renamed classic tables of the
set text reproduce Mitchell's published values (information gains 0.246/0.151/
0.048/0.029, the six-hypothesis version space, naive Bayes at 0.795). Figures
that show results (the polynomial fits, error curves, tree size, bandits,
ensembles, the scree plot) are drawn from real computed data.

## Build

```powershell
.\build.ps1            # build ML_Fall2026_Handout.pdf
.\build.ps1 -Clean     # remove aux files, then build
```

Requires MiKTeX (or TeX Live) with `pgfplots`, `tcolorbox`, `fontawesome5`,
`booktabs`, `longtable`, `ragged2e`, `needspace`, `float`, `placeins` and
`caption`. `build.ps1` borrows Git for Windows' `perl.exe` for `latexmk` when
none is on `PATH`, prints a report (pages, errors, undefined references, bad
boxes, font warnings) and exits non-zero on any error.

### Labs

```powershell
cd code
python sync_labs.py          # extract every chapter's lab into code/
python sync_labs.py --run    # ... and run them all
python sync_labs.py --run 9  # ... run chapter 9 only
```

The handout is the single source of the lab code: `sync_labs.py` copies each
chapter's *Try It Yourself* listing to `code/chNN_<short-name>.py`, checks that it
is pure ASCII (pdfLaTeX listings cannot take UTF-8) and no wider than 78
columns, and runs it. **Edit the chapter, not the extracted file.** The labs need
only a standard Anaconda installation (NumPy, pandas, SciPy, scikit-learn); none
needs a GPU or network access. The Chapter 14 lab takes a little over a minute and
Chapters 18 and 20 take 10–20 seconds; the rest finish within a few seconds.

### Slides

```powershell
cd slides
python texparse.py --all                 # audit: every chapter parses completely
python figures\extract_figures.py        # compile the handout's figures (cached)
python build_slides.py --all             # all 22 decks into slides\out
python build_slides.py --ch 7            # one chapter
python build_slides.py --part 3          # one part
python qa.py                             # static checks on every deck
python qa.py --render --ch 7             # also export PNGs and a contact sheet
```

Build the handout first: slide page references come from its `.aux`. Maths is
typeset with LuaLaTeX on demand; figures are compiled from the handout's own
TikZ, so a figure corrected in the handout reaches the slides on the next build.
`qa.py --render` drives the installed PowerPoint through COM.

Each deck has the same shape: title slide, where the chapter sits in the course,
outcomes, vocabulary, a section agenda, one slide per figure and box (lists and
worked-example steps build one click at a time), the lab as a
stepped code walk-through with its runnable file, checkpoint question-then-answer
pairs, the summary, the multiple-choice review, and a closing slide naming the
next chapter and the week's lab. Speaker notes carry the handout's prose.

## Layout

| Path | Contents |
| --- | --- |
| `ML_Fall2026_Handout.tex` | Master file; `\input`s everything |
| `mlhandout.sty` | Single source of style; a sibling of `../Data Science Fall 2026/dshandout.sty` |
| `parts/00_frontmatter.tex` | Title, how to use, **Course Specification (HEC)**, course map, domains, notation, self-check |
| `parts/00_partN_opener.tex` | One reading-path opener per part |
| `parts/01_*.tex` … `parts/22_*.tex` | The 22 chapters, in teaching order |
| `parts/A_*.tex` … `parts/F_*.tex` | Appendices |
| `code/` | `sync_labs.py` and the 22 extracted labs |
| `slides/` | The deck pipeline, its assets, figure renders and the decks in `slides/out/` |

## Authoring conventions

Every chapter uses the same furniture, in the same order, as the Data Science
handout — and the slide pipeline depends on it:

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{<part title>}
\begin{outcomes} ... \end{outcomes}
\begin{prereq} ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
...content: definitionbox / conceptbox / alertbox / \dsfig ...
\begin{worked}[title] ... \end{worked}
\begin{pitfall} ... \end{pitfall}
\begin{lab}[title] ... one lstlisting ... \end{lab}
\begin{checkpoint} ... three items ... \end{checkpoint}
\begin{chaptersummary} ... \end{chaptersummary}
\begin{reviewq} ... \end{reviewq}
```

- Every figure goes through `\dsfig` (or `\dsfigh` to force it in place); all are
  native TikZ/pgfplots. Inside a figure use only `\term` among the custom
  macros — the slide pipeline compiles figures standalone.
- Checkpoint answers go in `parts/C_answers.tex` under `\ansch{N --- Title}`,
  three per chapter; the slides reveal them after each question.
- The schedule table in `parts/F_schedule.tex` must keep its five columns (week,
  topic, chapters, lab, assessment): the slides read it for the week and lab.
- Paragraph columns use `L{<width>}`, never `p{<width>}`; widths sum to about
  15.6 cm for two columns and 15.0 cm for three.
- Displayed maths belongs outside list items (the slide parser flattens items).
- No `\textsc` inside bold text (Latin Modern has no bold small caps), and no
  TikZ style named `pos`, `out` or `in` — they are built-in keys.

Two things are handled by the style file rather than by hand:

- **Section headings reserve space** (`\Needspace` is prepended to `\section`
  and `\subsection`), so no heading is stranded at the foot of a page.
- **Table stripes are switched off inside `cases` and matrix environments.**
  `\rowcolors` is global, so without this every matrix after a striped table
  was painted. *The same fault exists in `dshandout.sty`* and is worth porting
  back.

## The slide pipeline is a copy

`slides/` is a copy of `../Data Science Fall 2026/slides/`, not a shared module,
so that each course can be rebuilt without touching the other. File names are
unchanged to keep a diff between the two copies meaningful. The adaptations:
master file and part titles (`texparse.py`), appendix file names (answers and
schedule), course strings and the 22-chapter count (`dsdeck.py`), the `ML_`
output prefix (`build_slides.py`, `qa.py`), and the style file the figure
preamble is generated from (`figures/extract_figures.py`). Improvements made
here that are worth porting back:

- `texparse.py --all` no longer counts a table or formula drawn *inside* a
  figure as dropped slide content.
- The roadmap slide's chapter label shrinks to fit a long title. Its "Used later
  in" list, when the prerequisite boxes give fewer than five chapters, is topped
  up from later chapters' `\chref`s. Only the final chapter says "the course ends
  here".
- A lab step with a heading is sized to the room left under the heading. Before,
  an 18-line step lost its last line off the bottom of the panel.
- A long lab step is split at a blank line or before a top-level statement,
  never inside a bracketed expression. Before, a hard cut at 20 lines could leave
  one continuation line alone on its own slide.
- Lab slides number their lines as the runnable file does, and a set-up of three
  lines or fewer joins the first step instead of taking a slide of its own.
- Display formulas that are slightly too wide shrink to fit instead of wrapping
  mid-equation (`mathrender.py`).
- `\varnothing` in prose reaches the slides as ∅ instead of vanishing
  (`texparse.py`'s symbol table).

The deck builder still carries the Data Science pipeline's four-lens card code
(`lens_slides` in `dsdeck.py`, the `fourlenses` box in `texparse.py`). No chapter
here uses it, so it never runs; it is kept only to keep this copy diffable
against the Data Science one.
