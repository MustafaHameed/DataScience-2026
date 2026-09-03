# AIOU 437 — اسلامیات، اختیاری (Islamiat, Elective): Urdu brief

Sixteen most-expected questions with concise Urdu answers, covering
Quranic verses, hadith, fiqh of salah, and Islamic history.

**Output:** `AIOU_437_Urdu_Brief.pdf` — 21 pages (cap 21).

## "(E)" means Elective, not English

The course listing shows "437 Islamiat (E)". Both past-paper scans read
for this guide head themselves **اسلامیات (اختیاری)** — Islamiat
*Elective* — and are set entirely in Urdu, including the Quranic and
hadith citations. This guide is therefore in Urdu, matching the other
five in this set.

## Source — four past papers, spanning a level change

The official course book could not be obtained (Google Drive returns
HTTP 500 on every copy found). Every question is transcribed from four
genuine AIOU past papers, read from scans:

- BA, اسلامیات (اختیاری) 437, **بہار 2012** — 8 questions
- BA, اسلامیات (اختیاری) 437, **بہار 2013** — 8 questions
- ADP, اسلامیات (اختیاری) 437, **خزاں 2021** — 8 questions
- ADP, اسلامیات (اختیاری) 437, **بہار 2022** — 8 questions

Four papers rather than the two used for 430 and 431: the course was
re-levelled from BA to Associate Degree between 2013 and 2021, and the
two eras differ in ways that matter for exam prep — pass marks moved
40→50, and whether the first question is compulsory changed. A
two-paper sample from one era would have misrepresented the current
pattern. Thirty-two questions merge to sixteen distinct topics.

## Accuracy note

Every Quranic verse is cited by surah and ayah number; every hadith is
attributed to the collection that actually grades it (Bukhari, Muslim,
Tirmidhi, or the Muwatta for the one hadith qudsi used). Where the
schools of fiqh differ — the timing of zuhr, the shafaq that starts
isha, istisqa prayer — the Hanafi position is given first (the
Pakistani and AIOU default) and the difference is stated openly rather
than flattened away.

Quranic and hadith Arabic is set in a **Naskh** face (Traditional
Arabic), not the Nastaliq used for the Urdu body text — Nastaliq's
steep slope collides fully-vowelled Arabic harakat into the letters
below them, checked on a rendered page before choosing.

## Paper pattern (read off the papers)

100 marks · 3 hours · eight questions set · answer any five · 20 marks
each. Two changes between eras: **pass marks 40→50**, and the first
question is compulsory in the 2021 paper but not in 2022 — read the
"نوٹ" line on the actual paper, don't assume.

## Topics in all four papers (highest yield)

Aal-e-Imran verse translation/tafsir (every paper's Q1 or Q2) · hadith
translation/tafsir · a topical note on Aal-e-Imran (Ghazwa Uhud, tauheed,
tafarruqa, deen-e-kainat) · Umayyad and Abbasid history. A ten-term
short-definition question was compulsory in two of the four papers.

## Build

```powershell
.\build_urdu.ps1           # normal build
.\build_urdu.ps1 -Clean    # remove aux files first
```

XeLaTeX only (Urdu needs OpenType shaping and bidi). Requires the
**Urdu Typesetting** font for the body text and **Traditional Arabic**
for Quranic/hadith quotations; the script checks for both and fails
clearly if either is absent. The build report counts **missing
glyphs** (the signature Urdu failure) and enforces a **21-page cap**.

Current status: **21 pages, 0 errors, 0 missing glyphs, 0 overfull
boxes.**

## Status

Private study only. Not an AIOU publication. Not for copying into an
assignment.
