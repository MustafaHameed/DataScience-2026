# Data Science — Fall 2026

Unified master lecture handout for the BSIT Data Science course. The entire
course is one document: six parts, 32 chapters, seven appendices, built from
basics to advanced in a single deliberate sequence.

This edition replaces the thirteen separate handouts used in Spring 2026
(archived unchanged in `../Data Science Spring 2026/`). Appendix F maps each
old handout to the chapters that now cover it.

**Current build:** 262 pages · 103 figures, all numbered, captioned and listed ·
zero LaTeX errors, zero undefined references, zero undefined font shapes, zero
overfull boxes, zero underfull boxes.

The log still carries `Infinite glue shrinkage found in box being split` once
per longtable that spans a page break. That is upstream behaviour in current
longtable, not a fault in this document — a bare `report` with one long table
reproduces it. TeX drops the offending shrink component and the rules and
repeated headers come out correct, so `build.ps1` filters the line out of its
error count. Do not chase it.

Figure conventions worth knowing before you edit one:

- Multi-item comparison figures use a **2×2 grid** (or 3+2 for five items), not a
  single row. A four-across row forces a 3.9 cm column and a `\tiny` font, which
  then overflows the text block and gets shrunk again by `\dsfitw`.
- Wide explanatory notes inside a figure are `\scriptsize` and anchored
  `north west`, so they hang *below* their coordinate and cannot grow upwards
  into the diagram.
- Front-matter figures are numbered `F.n`; part-opener maps are `Pn.1`. Chapter
  figures use the normal `chapter.n`. `\locator` restores chapter numbering, so
  every chapter file is unaffected.

## Build

```powershell
.\build.ps1            # build DataScience_Fall2026_Handout.pdf
.\build.ps1 -Clean     # remove aux files, then build
```

Requires MiKTeX (or TeX Live) with `pgfplots`, `tcolorbox`, `fontawesome5`,
`booktabs`, `longtable`, `ragged2e`, `needspace`, `float` and `caption`.
`latexmk` handles the multiple passes needed to resolve the table of contents,
list of figures and cross-references.

Equivalent manual build:

```powershell
latexmk -pdf -interaction=nonstopmode DataScience_Fall2026_Handout.tex
```

`latexmk` is a Perl script and MiKTeX ships no interpreter, so `build.ps1`
borrows the `perl.exe` bundled with Git for Windows (or Strawberry Perl) when
none is on `PATH`. Without that, `latexmk` fails with *could not find the script
engine 'perl'*.

`build.ps1` prints a report afterwards: page count, errors, undefined
references, overfull hboxes and vboxes, underfull boxes, font warnings and
PDF-string warnings. Every one of them should read `0`; it exits non-zero if any
error survives, so it is safe to use as a gate. It also refuses to report on a
log older than half an hour, so a build that never ran cannot masquerade as a
clean one.

## Layout

| Path | Contents |
| --- | --- |
| `DataScience_Fall2026_Handout.tex` | Master file. `\input`s everything; produces one PDF. |
| `dshandout.sty` | **Single source of style.** Colours, boxes, figure wrapper, all pedagogy macros. Change appearance here and nowhere else. |
| `parts/00_frontmatter.tex` | Title page, how-to-use, course map, the four domains, notation, prerequisite self-check, table of contents, list of figures. |
| `parts/00_partN_opener.tex` | One concept-map opener per part. |
| `parts/01_*.tex` … `parts/32_*.tex` | The 32 chapters, numbered in teaching order. |
| `parts/A_*.tex` … `parts/G_*.tex` | Appendices A–G. |
| `code/` | Runnable Python and R extracted from the chapter labs. |

## Structure

**Part I — Foundations of Data** (1 Landscape · 2 Lifecycle · 3 Data · 4 Mathematics · 5 Descriptive Statistics · 6 Probability)

**Part II — From Data to Insight** (7 Wrangling · 8 EDA · 9 Inference · 10 Causality)

**Part III — Machine Learning Core** (11 Fundamentals · 12 Regression · 13 Classification · 14 Evaluation · 15 Features & Ensembles · 16 Unsupervised)

**Part IV — Deep Learning and Modern AI** (17 Neural Networks · 18 Architectures · 19 Time Series & Streaming · 20 NLP · 21 Generative AI · 22 Agentic AI)

**Part V — Data Science in Systems Context** (23 Data Engineering · 24 Cloud & Edge · **25 Data Science for IoT** · **26 Data Science for Cybersecurity** · 27 MLOps)

**Part VI — Professional Practice** (28 Ethics · 29 Research Methods · 30 Project Management · 31 Communication · 32 Capstone)

IoT and cybersecurity sit in Part V because both depend on machinery taught
earlier — streaming and time series (Ch 19), edge computing (Ch 24), anomaly
detection (Ch 16) and evaluation under class imbalance (Ch 14). They are not
absent before then: every chapter from Chapter 1 onwards carries a **Four
Lenses** panel applying its technique to learning analytics, project
management, IoT and cybersecurity.

## Authoring conventions

Every chapter uses the same furniture, in this order:

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{<part title>}
\begin{outcomes} ... \end{outcomes}
\begin{prereq}   ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
...content: definitionbox / conceptbox / \dsfig ...
\begin{worked}[title]   ... \end{worked}
\begin{pitfall}         ... \end{pitfall}
\begin{fourlenses}{topic}
  \lensLA{...} \lensPM{...} \lensIOT{...} \lensSEC{...}
\end{fourlenses}
\begin{lab}[title]      ... \end{lab}
\begin{checkpoint}      ... \end{checkpoint}
\begin{chaptersummary}  ... \end{chaptersummary}
\begin{reviewq}         ... \end{reviewq}
```

Rules:

- **Every figure goes through `\dsfig{<tikz>}{<caption>}{<label>}`** so it is
  numbered, captioned and collected in the List of Figures. No bare
  `tikzpicture` in a `center` block.
- All figures are drawn natively in TikZ/pgfplots. No external image files —
  the document is fully self-contained.
- Cross-reference chapters with `\chref{shortname}` and figures with
  `\dsref{label}`.
- Checkpoint answers go in `parts/D_answers.tex`, keyed by chapter.
- New terms go in `parts/E_glossary.tex`.

### Tables

- **Paragraph columns use `L{<width>}`, never `p{<width>}`.** `L` is `p` set
  ragged right, with hyphens made expensive. The 183 paragraph columns in this
  handout are mostly under 4 cm; justifying them produced rivers, 111 underfull
  lines and column heads reading `Legal opera-tions`. All 56 longtables follow
  the same skeleton:

  ```latex
  \rowcolors{2}{white}{lightbg}
  \begin{longtable}{@{}L{3.4cm}L{4.4cm}L{6.4cm}@{}}
  \toprule
  \rowcolor{primary!15}
  \textbf{Head} & \textbf{Head} & \textbf{Head} \\
  \midrule
  \endhead
  ...rows...
  \bottomrule
  \end{longtable}
  ```

- Column widths must sum to about **15.0 cm for three columns**, 15.6 cm for
  two, 15.8 cm for five — the text block is 16.2 cm and `\tabcolsep` eats
  roughly 0.42 cm between each pair. Check the widest single word in each head
  against its column: `\textbf{Hyperparameter}` needs 3.15 cm and silently
  overflowed a 3.0 cm column for several editions.

### Typography rules living in `dshandout.sty`

- **Part and chapter kickers are medium-weight small caps, not bold.** Latin
  Modern Roman has no `T1/lmr/bx/sc`, so `\bfseries\textsc` had been silently
  losing the small caps for the whole run. `\mdseries` plus microtype tracking
  gives the letterspaced kicker the design was always asking for. Do not put
  `\bfseries` back in a `\titleformat` label that also uses `\textsc`.
- **`\_` carries an `\allowbreak`.** `snake_case` identifiers appear inline in
  the prose about eighty times and TeX has no break point inside one. A line may
  now break after an underscore, with no hyphen inserted. Maths is exempt.
- **Widows and orphans are forbidden** (`\widowpenalty`/`\clubpenalty` at
  10000); `\brokenpenalty` is deliberately left finite at 4000, because
  forbidding a break after a hyphenated line as well leaves TeX no legal break
  at all on some figure-heavy pages.
- **A section that opens straight into a tall figure needs
  `\Needspace*{0.4\textheight}` before it.** Otherwise the heading strands
  itself at the foot of a page, the float has nowhere to go, and it prints over
  the footer rule. Section 19.4 is the worked instance.
- **Maths in a heading needs `\texorpdfstring`.** PDF bookmarks are plain
  strings and cannot hold a math shift; the navigation pane goes down to
  subsection level, so subsection titles are in scope too.
