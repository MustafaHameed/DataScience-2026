# Advanced Research Methodology — Fall 2026

Master lecture handout for the MS IT and PhD IT course in Advanced Research
Methodology, Department of Information Technology, The Islamia University of
Bahawalpur.

The course is not optional for doctoral scholars: **PhD Degree Regulations-2024
§10(iv)(a)** requires every PhD student to take a 3-credit-hour course in
*"Advanced Statistical Analysis / Hypothetical Testing / Research
Methodology"*. This volume is that course.

**One PDF, 279 pages: 6 parts, 33 chapters, 8 appendices, 82 references.**

---

## Build

```powershell
.\build.ps1            # normal build
.\build.ps1 -Clean     # remove aux files first
```

Requires MiKTeX, `latexmk`, and **`biber`** (`biber --version` to check;
`mpm --install=biber` if missing). The script borrows Perl from Git for
Windows, since MiKTeX ships `latexmk` without an interpreter.

The build report prints pages, errors, undefined refs, overfull/underfull
boxes and font warnings. **The standard for this volume is zero of all of
them** — it has been held at every commit.

### If the build says `I can't write on file ... .pdf`

That is a file lock, not a LaTeX error: a PDF viewer has the output open.
Close it. To verify the source without closing your viewer:

```powershell
latexmk -pdf -interaction=nonstopmode -file-line-error `
        -jobname=ARM_check AdvancedResearchMethodology_Handout.tex
```

### Reading the log

Do **not** grep for `^!` alone — `-file-line-error` reformats errors as
`file:line: message`. Use:

```bash
grep -nE "^!|^.*:[0-9]+: |Overfull|Underfull|Citation .* undefined" \
     AdvancedResearchMethodology_Handout.log | grep -v "Infinite glue shrinkage"
```

`Infinite glue shrinkage found in box being split` is upstream `longtable`
behaviour on any table spanning a page break. The rules and repeated headers
still come out correct. `build.ps1` filters it; do not chase it.

---

## Layout

```
AdvancedResearchMethodology_Handout.tex   spine: \part + \input only
armhandout.sty                            all style and pedagogy macros
refs.bib                                  87 entries (82 cited), DOI on all 57 that have one
build.ps1                                 latexmk wrapper + build report
parts/                                    frontmatter, 6 openers, 33 chapters, 8 appendices
code/                                     runnable R and Python for Appendix B
slides/                                   PowerPoint decks, generated from parts/*.tex
assets/                                   iub-logo.png, dit-logo.png
```

The lecture decks in [slides/](slides/) are **generated from this source**, not
written alongside it: `slides/build_slides.py` parses `parts/*.tex` and emits
one deck per chapter, so a correction here reaches the slides on the next run.
See [slides/README.md](slides/README.md). Because they parse the chapter
furniture, **changing that structure means updating `slides/texparse.py` too**.

`armhandout.sty` is derived from `../Data Science Fall 2026/dshandout.sty`.
The typographic engineering is carried over unchanged, with its comments; what
differs is the teal palette, the running heads, biblatex, the four-traditions
block, and four discipline-specific boxes.

---

## Chapter furniture — fixed order

Every chapter carries the same apparatus in the same order. A chapter missing
any of it is unfinished.

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{Part~N --- <part title>}
\begin{outcomes}   ... \end{outcomes}     % Bloom verbs, what they can DO
\begin{prereq}     ... \end{prereq}
\begin{keyterms}   ... \end{keyterms}
  ... body: definitionbox / conceptbox / alertbox / regbox / paperbox ...
\begin{worked}[title]     ... \end{worked}
\begin{pitfall}           ... \end{pitfall}
\begin{traditions}{topic} \tradEMP{} \tradDSR{} \tradQUAL{} \tradFORM{}
\end{traditions}
\begin{templatebox}[Cn --- title] ... \end{templatebox}   % where the chapter has one
\begin{lab}[title]        ... \end{lab}
\begin{checkpoint}        ... \end{checkpoint}   % sketches go in Appendix E
\begin{chaptersummary}    ... \end{chaptersummary}
\begin{reviewq}           ... \end{reviewq}      % open, no sketches
```

### Boxes specific to this volume

| Box | Use |
|---|---|
| `regbox[clause]` | A binding IUB/HEC rule, **quoted**, with its exact clause. Never paraphrase a regulation into a `conceptbox` — students must see at a glance what is advice and what can fail them. |
| `paperbox[handle]` | Anchor reading. Body order: full citation, *Why it matters*, *What to extract*. |
| `templatebox[Cn — title]` | A fill-in worksheet. Use `\filllines{n}` for ruled lines. Indexed in Appendix C. |
| `vignette[title]` | A worked end-to-end study. **Everything inside is invented and says so.** |

Macros: `\reg{doc}{clause}` for an inline rule reference, `\oa` for the
open-access badge, `\term{}`, `\chref{}`, `\figref{}`, `\armfig{}{}{}`.

---

## Two content rules

**Worked studies are constructed.** Every number in a `vignette` or a `worked`
box is invented, and the box says so. No result is ever attributed to a real
study without a citation.

**`\oa` means publisher open access.** An author copy on a personal page is
*not* open access and is labelled "Author copy" in Appendix F. A book that
teaches provenance has to state its own precisely. The `note = {openaccess}`
field in `refs.bib` is the source of truth.

---

## Authoring rules that will otherwise bite

Carried from `../Data Science Fall 2026/README.md` and `../CCNA/README.md`,
plus two learned building this volume:

- Use `L{<width>}` not `p{<width>}` for paragraph columns. Budget **≈15.0 cm
  for 3 columns, 15.6 cm for 2** at a 16.2 cm text block — `\tabcolsep` eats
  ~0.42 cm per gap. A 3-column table at 15.4 cm overflows by about 1.2 pt.
- Kickers are `\mdseries\textsc`, never `\bfseries\textsc` — Latin Modern has
  no `T1/lmr/bx/sc`.
- Raw `_ # $ % & ~` in a **tcolorbox title** errors at the `\end{}`, not the
  title.
- Maths in a heading needs `\texorpdfstring`.
- Figures are TikZ/pgfplots inline. The only external images are the two logos,
  guarded by `\IfFileExists` so the volume builds without `assets/`.
- **Do not name a TikZ style `cap`** — it collides with the built-in line-cap
  key and fails with a `pgfkeys` error pointing at the `\begin{tikzpicture}`
  line, not the offending node.
- **Do not edit `.tex` with `sed`.** Backslash and `&` handling mangles LaTeX
  silently. Use the editor tools or a Python script with raw strings.

---

## Verification

1. `.\build.ps1 -Clean` → 0 errors, 0 undefined refs, 0 overfull, 0 underfull,
   0 font warnings, ~279 pages.
2. `grep -c "entry{" *.bbl` → 82. No `WARN` in `.blg`.
3. **Regulation audit.** Every `regbox` clause number checked against the
   source PDFs. Appendix A's table is the checklist. A wrong threshold in a
   compliance appendix is worse than no appendix.
4. **Link check** on `refs.bib`: resolve every DOI; `\oa` only where the
   publisher version is genuinely free. Downgrade the rest.
   Done 2 Sep 2026. Method, because it is easy to get wrong: do **not** judge a
   DOI by fetching its landing page — Elsevier, Wiley, SAGE, ACM and IEEE all
   return `403` to a script, which looks identical to a dead link. Query the
   metadata instead (`https://api.openalex.org/works/doi:<doi>`), compare the
   returned title with the entry's, and read `open_access.oa_status`:
   `gold`/`diamond`/`hybrid`/`bronze` mean the published version is free on the
   publisher's site and earn `\oa`; `green` is a repository copy only, so
   "Author copy"; `closed` is "Library". Crossref's API was rate-limiting and
   then unreachable during this pass — OpenAlex mirrors the same registry
   metadata and answered reliably.
5. **Furniture check**: every chapter has all elements above.
6. **Logo guard**: rename `assets/` away; the build must still succeed.

---

## Source documents

The compliance content is transcribed from, and must be checked against:

- *Doctor of Philosophy (PhD) Degree Regulations-2024*, IUB
- *MPhil/MS/LLM/MSc (Hons)/MBA Degree Regulations-2024*, IUB
- *IUB Theses/Synopses Guidelines-2023*
- HEC *Anti-Plagiarism Policy*; HJRS journal categories

Both regulation sets were approved by the Syndicate on 13 June 2024. **When a
figure in Appendix A disagrees with the current document from the Directorate
of Advanced Studies and Research, the Directorate's document wins** — and the
appendix should be corrected.

---

## Related material in this repository

| Path | Relation |
|---|---|
| `../Data Science Spring 2026/Research_Methods_in_IT_Handout.tex` | The introductory MSIT predecessor. This volume is its graduate successor, not its replacement. |
| `../Data Science Spring 2026/MSIT_Synopsis_Guide_Demo_IUB_2023.tex` | The institutional synopsis format. Chapter 30 and template C14 derive from it. |
| `../Data Science Fall 2026/parts/29_research.tex` | The BSIT-level precursor chapter. |
| `../Data Science Fall 2026/dshandout.sty` | Parent of `armhandout.sty`. |
