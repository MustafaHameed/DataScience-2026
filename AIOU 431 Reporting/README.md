# AIOU 431 — خبر نگاری (Reporting): Urdu brief

Fourteen most-expected questions with concise Urdu answers.

**Output:** `AIOU_431_Urdu_Brief.pdf` — 19 pages (cap 21).

## Source — past papers, not the book

The official 431 course book could not be obtained (the circulating copy
returns HTTP 500; AIOU's own SoftBooks endpoint 404s). So every question is
transcribed from two genuine AIOU past papers, read from scans:

- BA, خبر نگاری (431), **بہار 2013** — 8 questions
- BA, خبر نگاری (431), **خزاں 2013** — 8 questions

Sixteen questions merged to fourteen distinct topics. Each question in the PDF
carries its provenance, e.g. `(بہار 2013، سوال 4)`.

This is stated openly on the title page: the guide is built on past papers, and
answers follow established reporting principles rather than the book's own
wording.

## Paper pattern (read off the papers)

100 marks · 3 hours · eight questions set · answer any five · 20 marks each
— about 35 minutes per answer.

## Topics in both papers (highest yield)

Reporter's duties and the reporter/correspondent/representative distinction ·
meaning and role of news · مکتوب نگاری · news agencies · analytical and
investigative reporting.

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
