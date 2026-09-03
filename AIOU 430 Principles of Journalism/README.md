# AIOU 430 — اصولِ صحافت (Principles of Journalism): Urdu brief

Fourteen most-expected questions with concise Urdu answers.

**Output:** `AIOU_430_Urdu_Brief.pdf` — 19 pages (cap 21).

## Source — past papers, not the book

The official 430 course book could not be obtained (the circulating copy
returns HTTP 500; AIOU's own SoftBooks endpoint 404s on every filename tried).
So every question is transcribed from two genuine AIOU past papers, read from
scans:

- BA, اصولِ صحافت (430), **بہار 2013** — 8 questions
- BA, اصولِ صحافت (430), **خزاں 2013** — 7 questions

Fifteen questions merged to fourteen distinct topics. Each question in the PDF
carries its provenance, e.g. `(خزاں 2013، سوال 2)`.

Stated openly on the title page: the guide is built on past papers, and answers
follow established journalism principles and documented Pakistani press history
rather than the book's own wording.

## Paper pattern (read off the papers)

100 marks · 3 hours · seven or eight questions set · answer any five · 20 marks
each — about 35 minutes per answer.

## Topics in both papers (highest yield)

Evolution of journalism (subcontinent and Pakistan) · sub-editor's qualities and
duties · code of ethics and press laws · editorial — definition, types, and its
difference from the feature.

Includes a dated reference table (Bengal Gazette 1780, جامِ جہاں نما 1822,
تہذیب الاخلاق 1870, الہلال 1912 …), since history answers are won on names and
dates, and a خبر / اداریہ / کالم / فیچر distinction table.

## Build

```powershell
.\build_urdu.ps1           # normal build
.\build_urdu.ps1 -Clean    # remove aux files first
```

XeLaTeX only (Urdu needs OpenType shaping and bidi). Requires the **Urdu
Typesetting** font; the script checks for it and fails clearly if absent. The
build report counts **missing glyphs** (the signature Urdu failure) and enforces
a **21-page cap**.

Current status: **19 pages, 0 errors, 0 missing glyphs, 0 overfull boxes.**

## Status

Private study only. Not an AIOU publication. Not for copying into an assignment.
