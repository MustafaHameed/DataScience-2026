# Data Science — Fall 2026

Unified master lecture handout for the BSIT Data Science course. The entire
course is one document: six parts, 32 chapters, seven appendices, built from
basics to advanced in a single deliberate sequence.

This edition replaces the thirteen separate handouts used in Spring 2026
(archived unchanged in `../Data Science Spring 2026/`). Appendix F maps each
old handout to the chapters that now cover it.

**Current build:** 259 pages · 103 figures, all numbered, captioned and listed ·
zero LaTeX errors, zero undefined references, zero overfull boxes.

## Build

```powershell
.\build.ps1            # build DataScience_Fall2026_Handout.pdf
.\build.ps1 -Clean     # remove aux files, then build
```

Requires MiKTeX (or TeX Live) with `pgfplots`, `tcolorbox`, `fontawesome5`,
`booktabs`, `longtable`, `float` and `caption`. `latexmk` handles the multiple
passes needed to resolve the table of contents, list of figures and
cross-references.

Equivalent manual build:

```powershell
latexmk -pdf -interaction=nonstopmode DataScience_Fall2026_Handout.tex
```

`build.ps1` prints a report afterwards: page count, error count, undefined
references and overfull hboxes.

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
