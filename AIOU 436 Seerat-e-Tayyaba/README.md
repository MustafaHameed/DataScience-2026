# AIOU 436 — سیرتِ طیبہ (Seerat-e-Tayyaba): Urdu brief

Seventeen most-expected questions with concise Urdu answers, covering
the Makki period, the Madani period and its battles, the Prophet's
multifaceted role (the "بحیثیت ..." question type), and a short-answer
glossary.

**Output:** `AIOU_436_Urdu_Brief.pdf` — 17 pages (cap 21).

## Source — past papers, not the book

The prescribed book (787 pages) is a scanned, image-only PDF with no
text layer — unusable as a source. Every question is transcribed from
four genuine AIOU past papers, read from scans:

- BA, سیرتِ طیبہ (436), **بہار 2013** — 8 questions
- BA, سیرتِ طیبہ (436), **خزاں 2012** — 8 questions
- BA, سیرتِ طیبہ (436), **خزاں 2017** — 8 questions
- BA/ADP, سیرتِ طیبہ (436), **بہار 2022** — 8 questions

Thirty-two questions merge to seventeen distinct topics.

## Accuracy note

Dates, numbers and names were checked against standard seerah sources
before being typed in. Where a paper's short-answer item touched a
detail this session could not verify with confidence (one obscure
name among the ten short-answer items in the خزاں 2017 paper), that
item was left out of the glossary rather than guessed — stated openly
on the title page.

## Paper pattern (read off the papers)

100 marks · 3 hours · eight questions set · answer any five · 20 marks
each. One change across eras: **pass marks 40 → 50** for the بہار
2022 paper, flagged in its own warning box.

## Topics in three or four papers (highest yield)

Sources of seerah (مآخذ و مصادر — in **all four** papers) · the
Prophet's ﷺ multifaceted role — teacher, administrator, lawgiver, head
of family (the recurring "بحیثیت ..." essay format) · his economic
life and teachings · the conquest of Makkah · Badr and the Constitution
of Medina.

## Build

```powershell
.\build_urdu.ps1           # normal build
.\build_urdu.ps1 -Clean    # remove aux files first
```

XeLaTeX only (Urdu needs OpenType shaping and bidi). Requires the
**Urdu Typesetting** font for the body text and **Traditional Arabic**
for Quranic and hadith quotations; the script checks for both and
fails clearly if either is absent. The build report counts **missing
glyphs** (the signature Urdu failure) and enforces a **21-page cap**.

Current status: **17 pages, 0 errors, 0 missing glyphs, 0 overfull
boxes.**

## Status

Private study only. Not an AIOU publication. Not for copying into an
assignment.
