# HEC Curriculum Map — AI, Data Science and Networking

How one group of courses — artificial intelligence, machine learning, data
mining, data science, deep learning, NLP, and the infrastructure and networking
courses they run on — is treated across three editions of the HEC computing
curriculum. Ends with a pathway, five semester schemes for BSIT and BSCS (IT
specialization), a CCNA v2.0 alignment, and the published contents of every
course.

**Current build:** 52 pages · zero LaTeX errors, zero undefined references, zero
overfull boxes, zero underfull boxes, zero font warnings.

The document opens with a **Summary** — the findings and a five-item "what to do
first". Read that before anything else; the rest is evidence for it.

## Sources

The three booklets sit in the repository root and are read, not shipped:

| Edition | File | Covers |
| --- | --- | --- |
| 2017 | `../BS & MS (CS,SE & IT) Curriculum 2017.pdf` | BSCS, BSSE, BSIT and the MS programmes, 171 pp. |
| 2023 | `../BS Curriculm Computing Disciplines-2023.pdf` | Eleven BS degrees including BSAI, BSDS and BSIT, 100 pp. |
| 2025 | `../HEC CURRICULUM OF COMPUTER SCIENCE-2025.pdf` | BS Computer Science only, with fourteen specializations, 86 pp. |

Every claim about the curriculum is cited to a page range in the closing
**Sources** section. Page references are to the PDF page, not the printed page
number. Part 7 additionally cites Cisco's own published CCNA blueprint.

## What the document contains

The order is evidence (1–3), analysis (4–7), design (8), practice (9),
reference (10).

| Part | Contents |
| --- | --- |
| — | **Summary** — what the three editions did, the four findings that change what a department should do, the CCNA position, what to do first, and a table routing each kind of reader to the right part |
| 1 | What each of the three booklets is, and where the elective credit actually sits in BSIT and BSCS |
| 2 | The status matrix — 43 courses × 3 editions, showing compulsory / elective / postgraduate-only / absent, in three blocks: the AI/DS cluster, the infrastructure it runs on, and the networking depth |
| 3 | Course cards: for each major course, what each edition says about credits, prerequisites and content, and what changed |
| 4 | What stays the same, and what changes — the invariants, then 2017→2023 and 2023→2025, plus a naming-drift table |
| 5 | The overlap map: twelve pairs of courses that duplicate each other, with the recommended split |
| 6 | The pathway — one dependency graph reconciling all three editions, with the infrastructure lane in teal |
| 7 | **Networking depth and the CCNA** — what networking courses exist in each edition, the CCNA 200-301 **v2.0** blueprint, a domain-by-domain mapping onto HEC courses, and why the gap is practice rather than knowledge |
| 8 | **Five semester schemes** (see below) — four data tracks and one CCNA track, together so the choice can be made side by side |
| 9 | Risks, the certification requirement, and running the three editions side by side |
| 10 | **The published course contents** — every topic outline, prerequisite, learning outcome and reading list, verbatim, for 43 courses across all three editions. This is the part to open when writing an actual course file. |

Parts 1–9 are the analysis; Part 10 is the raw material. The gaps in Part 10 are
recorded rather than passed over: 2017 specifies all 75 of its bachelor courses,
2023 specifies only 46 (those common to every computing degree plus the AI, Data
Science and Cyber Security cluster tables — the Information Technology cluster is
left entirely unspecified), and 2025 publishes learning outcomes only, with no
prerequisite and no outline for any of its 215 specialization courses.

Part 7 is the only part that draws on a non-HEC source: Cisco's own published
exam-topics document for **CCNA 200-301 v2.0**. v2.0 replaces v1.1 on 3 February
2027, so any student entering from Fall 2026 sits v2.0 — the whole part is
aligned to it, with a table in §7.5 for anyone carrying a v1.1 course file
forward.

The planning fact that matters is not the domain weights: of the 29 published
topics, **8 begin *Troubleshoot* or *Diagnose* and 8 begin *Configure*, so 55%
of the exam is hands-on**, against 21% *Describe*. Under v1.1 the word
"troubleshoot" appeared in no topic statement at all.

The closing note in **Sources** says plainly that the mapping is a planning
judgement rather than a Cisco-certified equivalence. Cisco revises blueprints
without renaming the exam — re-check `cisco.com/go/certroadmap` before committing
a scheme of studies.

## The five schemes

- **Scheme A** — BSIT under HEC 2023, all seven domain electives on the cluster.
- **Scheme B** — BSIT under HEC 2023: four data courses, Cloud Computing as the
  platform, two IT electives. **This is the recommended default.**
- **Scheme C** — BSCS with an IT specialization under HEC 2025, blending the IT
  and Data Science clusters and exploiting the fact that Computer Networks and
  Cloud Computing are compulsory major courses that cost no elective credit.
- **Scheme D** — the three-course minimum viable strand.
- **Scheme E** — a CCNA **v2.0** track under HEC 2025: six networking
  electives, two AI electives because domain 5.0 now examines agentic AI and
  prompt selection, with the compulsory professional-certification credit spent
  on CCNA itself.

Every scheme spends only elective credit and leaves the compulsory parts of the
degree untouched, so none of them requires HEC approval of a new programme.

## Build

```powershell
.\build.ps1            # build HEC_AI_DS_Curriculum_Map.pdf
.\build.ps1 -Clean     # remove aux files, then build
```

Requires MiKTeX (or TeX Live) with `tikz`, `tcolorbox`, `fontawesome5`,
`booktabs`, `longtable`, `ragged2e`, `needspace`, `float` and `caption`.

`latexmk` is a Perl script and MiKTeX ships no interpreter, so `build.ps1`
borrows the `perl.exe` bundled with Git for Windows (or Strawberry Perl) when
none is on `PATH`. The script prints a report afterwards and exits non-zero if
any error survives, so it is safe to use as a gate.

## Layout

| Path | Contents |
| --- | --- |
| `HEC_AI_DS_Curriculum_Map.tex` | The whole document. Single file. |
| `hecmap.sty` | Style. A deliberate sibling of `../Data Science Fall 2026/dshandout.sty` — same palette, same typographic decisions, much smaller. |

## Authoring conventions

Carried over from the Fall 2026 handout, and for the same reasons:

- **Paragraph columns use `L{<width>}`, never `p{<width>}`.** `L` is `p` set
  ragged right with hyphens made expensive.
- Column widths must sum to about **15.3 cm for four columns**, 15.8 for three,
  16.2 for two, 14.9 for five — the text block is 16.6 cm and `\tabcolsep` eats
  roughly 0.42 cm between each pair.
- **Kickers are medium-weight small caps.** Latin Modern Roman has no
  `T1/lmr/bx/sc`, so `\bfseries\textsc` silently loses the small caps.
- **A section opening straight into a tall figure needs `\Needspace`** before it,
  or the heading strands itself at the foot of the page.

Each course comparison uses the `coursecard` environment, which always carries
the three editions in the same order and the same colours — red for 2017, blue
for 2023, teal for 2025:

```latex
\begin{coursecard}{Course name}{one-line description}
  \edrow{2017}{status}{detail}
  \edrow{2023}{status}{detail}
  \edrow{2025}{status}{detail}
  \verdict{what actually changed}
\end{coursecard}
```
