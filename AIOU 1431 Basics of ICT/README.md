# AIOU 1431 / 5403 — Basics of ICT: اہم ترین سوالات (Urdu brief)

Sixteen highest-yield questions from the official AIOU course book, answered
concisely in Urdu.

**Output:** `AIOU_1431_Urdu_Brief.pdf` — 20 pages (capped at 21).

## Source

Every question is reproduced from the end-of-unit **Self-Assessment Questions**
of:

> *Basics of Information and Communication Technology*, Course Code 1431/5403,
> Units 1–9. Department of Computer Science, Allama Iqbal Open University,
> Islamabad. 1st edition, 2013. 310 pp.

Answers are written against that book's own terminology and section numbering —
each question carries its source reference (e.g. `2.9.5, Q.1`) so it can be
checked against the book.

Note that AIOU's own `online.aiou.edu.pk/LIVE_SITE/SoftBooks/5403.pdf` serves
only a **2-page cover**, not the book. The full text came from a circulating
copy of the same official edition.

## Selection

Weighted by exam load rather than spread evenly across units:

| Unit | Topic | Questions |
| --- | --- | --- |
| 1 | Introduction to ICT | 2 |
| 2 | Overview and organization of computer | 3 |
| 3 | Input devices | 2 |
| 4 | Output devices | 2 |
| 5 | Computer software | 2 |
| 6 | Operating system | 1 |
| 7 | Data communication and networking | 2 |
| 8 | Multimedia | 1 |
| 9 | Computer languages | 1 |

Technical terms are kept in English — that is what the paper expects — with an
Urdu explanation alongside, plus a four-part **اصطلاحات کی فہرست** (term
glossary) at the back. Most marks in this subject are lost in the gap between
the English term and the Urdu explanation, so the glossary is not an appendix
but a study aid in its own right.

## Build

```powershell
.\build_urdu.ps1           # normal build
.\build_urdu.ps1 -Clean    # remove aux files first
```

**XeLaTeX, not pdfLaTeX** — pdfTeX has neither OpenType shaping (Nastaliq
ligatures) nor a bidirectional algorithm. Needs the **Urdu Typesetting** font,
which ships with Windows; the script checks for it up front and fails with a
clear message rather than emitting hundreds of `nullfont` lines. On other
platforms substitute a Nastaliq or Urdu Naskh face in `aiouurdu.sty`.

The build report carries two guards specific to these documents:

- **Missing glyphs** — the signature Urdu failure is the font not loading and
  every Arabic-script character being silently dropped. Counted as a
  first-class metric, and a non-zero count fails the build.
- **Page cap** — 21 pages. The point of a brief edition is that it fits in one
  sitting, so the constraint is enforced, not trusted.

Current status: **20 pages, 0 errors, 0 missing glyphs, 0 overfull boxes.**

## Layout

```text
AIOU_1431_Urdu_Brief.tex     master
aiouurdu.sty                 shared Urdu style (fontspec + polyglossia + bidi)
parts_ur/00_front.tex        title page + orientation page
parts_ur/01_questions.tex    the sixteen questions
parts_ur/02_glossary.tex     term glossary + eight common mistakes
build_urdu.ps1
```

`aiouurdu.sty` is the same machinery as `bmurdu.sty` in the
*AIOU 1429 Business Mathematics* folder, renamed because it is
course-agnostic — the box names (`sawal`, `hal`, `nukta`, `ghalti`) and macros
(`\jawab`, `\parakh`) are Urdu words, not subject-specific. Each course folder
keeps its own copy so it stays self-contained, matching the convention of the
other course folders in this repository.

Two constraints worth knowing before editing it:

- **Package order.** `polyglossia` pulls in `bidi`, which insists on being
  loaded last. Everything else must be required above it.
- **Tables.** `bidi` reverses tabular column order to match page direction.
  That is right for Urdu label tables and wrong for order-carrying ones; those
  are wrapped in the `ltrtab` environment.

## Status

Prepared for private study. Not an AIOU publication and carries no official
endorsement. Intended for checking your own work — not for copying into an
assignment for submission.
