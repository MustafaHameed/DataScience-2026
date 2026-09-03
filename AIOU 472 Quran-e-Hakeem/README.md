# AIOU 472 — مطالعہ قرآنِ حکیم (Study of Quran-e-Hakeem): Urdu brief

Eighteen most-expected questions with concise Urdu translation and
tafsir, covering Surah Al-Baqarah, Surah Aal-e-Imran, and cross-surah
themes (riba, oaths, Bani Israel).

**Output:** `AIOU_472_Urdu_Brief.pdf` — 15 pages (cap 21).

## Source — past papers, not the book

The prescribed book (318 pages) is a scanned, image-only PDF with no
text layer, and the copy obtained doesn't even show clear chapter
markers — unusable as a source. Every question is transcribed from
four genuine AIOU past papers, read from scans:

- BA, مطالعہ قرآن حکیم (472), **بہار 2013** — 8 questions
- BA, مطالعہ قرآن حکیم (472), **خزاں 2013** — 8 questions
- ADP, مطالعہ قرآن حکیم (472), **خزاں 2021** — 8 questions
- BS, مطالعہ قرآن حکیم (472), **بہار 2022** — 8 questions

Thirty-two questions merge to eighteen distinct topics. Each question
carries its provenance, e.g. `(خزاں 2021، سوال 5)`.

## Accuracy note

Every verse is cited by surah and ayah number and was checked
individually against a standard mushaf before being typed in.
Translations are literal, exam-purpose renderings, not a claim to any
published translator's exact wording. Quranic Arabic is set in a
**Naskh** face (Traditional Arabic), not the Nastaliq used for the
Urdu body text — Nastaliq's steep slope collides fully-vowelled
harakat into the letters below them.

## Paper pattern (read off the papers)

100 marks · 3 hours · eight questions set · answer any five · 20 marks
each. One change across eras: **pass marks 40 → 50** for the
Associate-Degree/BS-era papers, flagged in its own warning box.

## Topics in three or four papers (highest yield)

Al-Kitab's characteristics and tahrif (concealing/distorting the
scripture) · riba and oath rulings · muhkamat/mutashabihat, unity of
the ummah, and condemnation of bukhl · hypocrites' traits · talaq/khula
and hajj rulings · word and verb-form identification (a fixed-format
question in two of the four papers).

## Build

```powershell
.\build_urdu.ps1           # normal build
.\build_urdu.ps1 -Clean    # remove aux files first
```

XeLaTeX only (Urdu needs OpenType shaping and bidi). Requires the
**Urdu Typesetting** font for the body text and **Traditional Arabic**
for Quranic quotations; the script checks for both and fails clearly
if either is absent. The build report counts **missing glyphs** (the
signature Urdu failure) and enforces a **21-page cap**.

Current status: **15 pages, 0 errors, 0 missing glyphs, 0 overfull
boxes.**

## Status

Private study only. Not an AIOU publication. Not for copying into an
assignment.
