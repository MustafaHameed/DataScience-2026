# CCNA 200-301 v2.0 — Course Volume

One book containing the syllabus, the notes and the labs for a course taught to
Cisco CCNA 200-301 **version 2.0** — written from zero, with no prior networking
assumed, teaching protocol theory and device configuration together.

**Current build:** 91 pages · zero LaTeX errors, zero undefined references, zero
overfull boxes, zero underfull boxes, zero font warnings.

> **Status: in progress.** Front matter, Part I (chapters 1–7) and chapter 8 are
> written. Chapters 9–34 and Appendices A–G are stubs that compile but have no
> content yet. See *Progress* below.

## Why v2.0 and not v1.1

v1.1 retires 2 February 2027; v2.0 goes live the next day, keeping the same exam
number. Anyone entering a course from Fall 2026 sits v2.0. The two blueprints are
different enough that a v1.1 book teaches the wrong exam — QoS, NTP, TFTP, REST
APIs, JSON and Terraform are gone; OSPFv3, SFTP/SCP, DNS record diagnosis, storm
control, RA guard, PoE and a whole AI domain have arrived.

**The planning fact that drives the whole design:** of the 29 published topics,
8 begin *Troubleshoot* or *Diagnose* and 8 begin *Configure* — so **55% of the
exam is hands-on**, against 21% *Describe*. Under v1.1 the word "troubleshoot"
appeared in no topic statement at all. That is why every chapter has `config`,
`verify` and `breakfix` blocks rather than prose with a lab bolted on the end.

## Source of record

`200-301_CCNA_v2.0_Exam_Topics_PDF.pdf` in this folder — Cisco's own published
document, marked Cisco Public. Topic numbers, weights and wording come from it
and nothing else. `200-301-CCNA-v1.1.pdf` is kept for the historical comparison
in the front matter only.

Cisco revises blueprints without renaming exams. Re-check
`cisco.com/go/certroadmap` before teaching from this.

## Layout

| Path | Contents |
| --- | --- |
| `CCNA_v2_Handout.tex` | Master file. `\input`s everything; produces one PDF. |
| `ccna.sty` | **Single source of style.** Palette, boxes, IOS listings, topology macros, blueprint tags. |
| `parts/00_frontmatter.tex` | Title page, exam facts, how to use the volume, syllabus with CLOs and assessment. |
| `parts/00_partN_opener.tex` | One concept-map opener per part. |
| `parts/01_*.tex` … `parts/34_*.tex` | The 34 chapters, in teaching order. |
| `parts/A_*.tex` … `parts/G_*.tex` | Appendices A–G. |
| `tools/check_coverage.py` | Proves every blueprint topic is covered. |
| `build.ps1` | Build + quality report; exits non-zero on any error. |

## Build

```powershell
.\build.ps1            # build CCNA_v2_Handout.pdf
.\build.ps1 -Clean     # remove aux files, then build
python tools\check_coverage.py   # prove blueprint coverage
```

Requires MiKTeX (or TeX Live) with `tikz`, `tcolorbox`, `listings`,
`fontawesome5`, `booktabs`, `longtable`, `ragged2e`, `needspace`, `float` and
`caption`. `latexmk` is a Perl script and MiKTeX ships no interpreter, so
`build.ps1` borrows the `perl.exe` bundled with Git for Windows.

Both gates must pass: the build report must read `0` across the board, and
`check_coverage.py` must exit 0.

## Chapter furniture

```latex
\chapter{...}\label{ch:short}
\bp{2.1,2.2}                       % blueprint topics -- NO SPACES in the list
\begin{outcomes} ... \end{outcomes}
\begin{prereq} ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
...definitionbox / conceptbox / alertbox / \ccnafig...
\begin{config}{what this builds} ... \end{config}
\begin{verify}{SW1}{show vlan brief} ... \end{verify}
\begin{pitfall} ... \end{pitfall}
\begin{breakfix}[the symptom] ... \end{breakfix}
\begin{lab}{title}{Packet Tracer} ... \end{lab}
\begin{examnote}[On the exam] ... \end{examnote}
\begin{checkpoint} ... \end{checkpoint}
\begin{chaptersummary} ... \end{chaptersummary}
\begin{reviewq} ... \end{reviewq}
```

## Authoring conventions

Carried over from the two sibling projects in this repo
(`../Data Science Fall 2026/dshandout.sty`, `../HEC AI-DS Curriculum Map/hecmap.sty`)
for the reasons those files document:

- **Paragraph columns use `L{<width>}`, never `p{<width>}`** — ragged right with
  expensive hyphens.
- **Column widths.** The text block here is **16.2 cm** (2.4 cm margins), *not*
  the 16.6 cm of the curriculum map. `\tabcolsep` eats ~0.42 cm per gap, so the
  budget is **15.78 cm for two columns, 15.36 for three, 14.93 for four, 14.51
  for five.**
- **Kickers are medium-weight small caps** — Latin Modern has no `T1/lmr/bx/sc`.
- `\widowpenalty`/`\clubpenalty` 10000, `\brokenpenalty` 4000.
- `\_` carries an `\allowbreak`; all figures are TikZ-native, no external images.

Two traps specific to this volume:

- **`verify` takes two arguments, device then command** — `\begin{verify}{SW1}{show vlan brief}`.
  The `#` of the prompt is generated internally, because a raw `#` in a tcolorbox
  title is an "Illegal parameter number" error.
- **`^` inside `\cmd{}` is a superscript.** Write `\cmd{| exclude \textasciicircum!}`.

## Accuracy

Configurations are written to be syntactically correct for IOS and IOS XE and to
follow the command order a device requires. **They have not been verified on live
hardware.** Run each lab once on your own platform before issuing it — that is
worth more than a validation claim this book cannot honestly make. The front
matter says so to the reader too.

## Progress

| Part | Chapters | State |
| --- | --- | --- |
| Front matter | — | **Written** — title, exam facts, how-to-use, syllabus, CLOs, assessment |
| I Foundations | 1–7 | **Written** |
| II Switching and Network Access | 8 | **Written** |
| II Switching and Network Access | 9–14 | Stubs |
| III IP Routing | 15–19 | Stubs |
| IV Network Services and Security | 20–26 | Stubs |
| V Wireless, Virtualisation and Operations | 27–32 | Stubs |
| VI Sitting the Exam | 33–34 | Stubs |
| Appendices A–G | — | Stubs |

Stubs carry their chapter heading, label and `\bp{}` tag, so the master compiles
and `check_coverage.py` already reports all 29 topics accounted for. Filling a
stub does not change the structure.
