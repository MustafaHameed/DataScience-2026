# E-Commerce Application Development — Fall 2026

Teaching materials for the BSIT *E-Commerce Application Development* course
(**INFT-4041**, 5th semester): a handout, lecture decks, a running shop the
whole course is built on, and the measurements that keep its claims honest.

**Dr. Mustafa Hameed · Department of Information Technology · The Islamia
University of Bahawalpur**

> **Status: in progress.** Chapters 1–8 of 23 are written — all of Parts I and II —
> together with the front matter, all six part openers, and Appendices A, B,
> C and G. See
> [Not done](#not-done) for exactly what is missing.

---

## What is here

| Artefact | Size now | Build with |
|---|---|---|
| `probe_master.pdf` | 90 pp., 8 chapters, 4 appendices, 12 figures | `pdflatex probe_master.tex` |
| `slides/out/*.pptx` | 9 decks, 331 slides | `python slides\build_slides.py --all` |
| `code/` | a running OpenCart shop, a catalogue seeder, 4 measurement harnesses | `docker compose up -d --build` |

`ECAD_Fall2026_Handout.tex` is the real master and **does not build yet** — it
`\input`s all 23 chapters and seven appendices. `probe_master.tex` builds the
chapters that exist and is deleted when `parts/` is complete.

---

## HEC provenance

The course is **specified once and then withdrawn**, so the outline has to be
traced carefully rather than assumed.

| Edition | Title | Status | Contents published? |
|---|---|---|---|
| **2017** | **E-Commerce** | **Elective in BS *Software Engineering*** (p. 38); absent from BSIT. 3(3,0); prerequisite *Web Engineering* | **Yes — p. 94, the only outline in any edition** |
| 2023 | — | **Absent entirely**; one incidental domain mention (p. 23) | No |
| 2025 | Digital Marketing & E-Commerce | Interdisciplinary/Allied list, item 11 (p. 16) | No — title and credit hours only |

**HEC 2017 p. 94 is the sole published specification.** The department teaches
the course at 5th semester per the registrar's CBT sheet; the previous
`Summer_2026` handout's title page says 3rd semester, and that discrepancy is
recorded in the Course Specification rather than silently corrected.

### Five defects in the published document, recorded rather than inherited

All five are stated in the handout's front matter with the correction used:

1. **Only two CLOs**, both too broad to assess. The siblings publish eight.
   Six are added by the department, each traced back to one of HEC's two and
   each marked as departmental.
2. **The Bloom's Taxonomy column is blank** for both outcomes. Levels are
   assigned here and said to be.
3. **The "Course Content" paragraph is the table of contents of its own
   reference book** — Michael Peacock, *PHP 5 E-commerce Development* (Packt,
   2010), reference 2 on the same page. HEC did not write an outline; it
   pasted a contents page.
4. **Two clauses are duplicated** by that paste: "Checkout" and "Managing
   Products and Categories" each appear twice. 24 clauses, 22 distinct.
5. **The prescribed technology is end-of-life.** PHP 5 lost security support
   on 31 December 2018. Laudon & Traver is cited at the 13th ed. (2017),
   Rayport at 2007.

A sixth mismatch, recorded in Appendix F: HEC's prerequisite is *Web
Engineering*; the department's course of that content is **CS-207 / INFT-4005
Web Technologies** (4th semester). The name differs; the dependency holds.

### Three topics taught that the outline does not name

Justified in the Course Specification, because a reviewer will ask:

- **Ch 6, the data model.** The outline names the admin screens but never the
  tables. *Database Systems* is a prerequisite that should be paid off.
- **Ch 16, security as its own chapter.** HEC buries it inside "Deploying,
  Security, and Maintenance". CLO-7 asks for a security posture to be
  *analysed*; one third of one clause will not do it.
- **Ch 19, module development.** The outline is configuration end to end. The
  department's title says *Application Development*, and HEC's own teaching
  methodology requires a project.

---

## The platform decision

**OpenCart 4.1.0.4 (PHP 8.2 + MariaDB 11), built from a pinned release and run
through Docker Compose.**

HEC's clause list *is* a storefront admin menu — "Managing Products and
Categories", "Product Variations", "The Shopping Basket", "Shipping and Tax",
"Discounts, Vouchers, and Referrals", "Administration: Dashboard… Orders,
Customers, Refunds" — so tracing the outline to the chapters is nearly
mechanical, and it keeps HEC's PHP/MySQL lineage without teaching PHP 5.

Three things found while building it, all of which became teaching material:

- **There is no OpenCart image to pull.** Bitnami withdrew theirs and there has
  never been an official one, so every `docker pull opencart` instruction on
  the web is now dead. `code/shop/Dockerfile` builds it from `php:8.2-apache`
  plus the release, pinned by version **and SHA-256** — a version alone is not
  reproducible, because a release asset can be replaced.
- **OpenCart's own `Dockerfile` does not build.** It contains a shell syntax
  error (`docker-php-ext-install zip && && docker-php-ext-enable zip`).
- **Its own `docker-compose.yml` pins `mysql:5.7`**, which left support in
  October 2023, and starts Redis, Memcached and PostgreSQL that a stock
  install never uses.

Chapter 5 and Chapter 20 use all three: an artefact being official is evidence,
not authority.

### Running it

```powershell
cd code
copy .env.example .env
docker compose up -d --build     # first run: ~5.5 min, ~1.5 GB
python waitfor.py                # three URLs must answer 200
```

| | URL | Credentials |
|---|---|---|
| Storefront | <http://localhost:8090/> | — |
| Admin | <http://localhost:8090/admin/> | `admin` / `admin123` |
| Database console | <http://localhost:8091/> | `shop` / `shoppw` / `opencart` |

Ports are 8090/8091 rather than the usual 8080/8081 because both of those were
already taken on the development machine — which is why they are variables in
`.env` and not literals in the Compose file.

`docker compose down` keeps your shop; `down -v` destroys it. Chapter 5's lab
makes students do both, deliberately, in a week when it costs them nothing.

---

## The measured strand

Twelve of the twenty-three chapters end with a number that was **produced, not
asserted**; the other eleven argue in prose, because dressing up a legal
obligation with a stopwatch would be theatre. Measurements live in `code/` and
are quoted inside `\begin{measured}` boxes.

Chapter 5's, the only one written so far, on the machine in Appendix A:

| Start | To first HTTP 200 |
|---|---|
| From scratch (no images at all) | **338.4 s** (214.6 s of it the build) |
| Cold (images built, empty volumes) | **34.1 s** |
| Warm (the shop already exists) | **16.4 s** |

A ratio of **20.6×** between the cost paid once and the cost paid daily, which
is why quoting a single "startup time" would describe an experience nobody has.
Footprint: 1,508 MB of images, 221 MB of volumes.

Chapter 6's, which is the most interesting:

| Products | Page | Queries | DB time | DB as % of page |
|---|---|---|---|---|
| 100 | 112 / 111 ms | **144** | 60 / 35 ms | 54% / 32% |
| 1,000 | 203 / 101 ms | **144** | 60 / 60 ms | 30% / 60% |
| 10,000 | 171 / 302 ms | **144** | 114 / 128 ms | 67% / 43% |
| 50,000 | 541 / 476 ms | **144** | 402 / 477 ms | 74% / 100% |

Two independent runs. **One category page runs 144 queries**, and that is the
same at a hundred products as at fifty thousand — the N+1 is per *displayed*
product. Adding the three indexes `oc_product` does not ship helps at 1,000 and
10,000 products and **makes the database slower at 50,000**, in both runs: an
index that removes a filesort replaces sequential reads with scattered ones,
and past some size that loses.

**Three** earlier versions of that measurement were wrong, and all three are
recorded in the handout. The first timed `docker compose exec` rather than the
query. The second timed a hand-written query that only *resembled* OpenCart's.
The third — the worst — seeded 50,000 products with `store_id = 1` when the
default store is `0`, so every one of them was invisible to the storefront and
the measurement was timing a page showing the nineteen demo products. Nothing
failed; the page returned 200 and meant something else.

Chapter 8's finding is the counter-intuitive one:

| Page | Requests | CSS+JS | Images | Total |
|---|---|---|---|---|
| Home | 25 | 570 KB | 293 KB | **895 KB** |
| Category | 20 | 570 KB | 122 KB | **738 KB** |
| Product | 18 | 597 KB | 61 KB | **713 KB** |

**The framework is the page.** CSS and JavaScript are ~570 KB on *every* page
regardless of content — 84% of the product page. Images, which everyone
optimises first, are 9–33%. The largest single file is `bootstrap.css` at
265 KB, shipped **unminified with no minified copy anywhere in the tree**.

**An honesty note that is in the handout too.** An earlier run of the same
measurement on the same machine gave 17.6 s cold and 7.7 s warm — roughly half
— because less was running. Students' numbers will differ by more than they
expect; the *ratio* should not.

---

## Building

```powershell
pdflatex probe_master.tex             # until parts/ is complete
.\build.ps1                           # the real handout, once it is
python checkfigs.py                   # TikZ line-break and reserved-name traps
python checklabs.py                   # listings: ASCII, <= 76 cols, step markers
python slides\texparse.py --all       # every chapter parses; Q/A counts match
python slides\scanprose.py            # no sibling course's prose survives
python slides\figures\extract_figures.py
python slides\build_slides.py --all
python slides\overview.py
python slides\qa.py --render          # static checks, then LOOK at the sheets
```

### What "clean" means

| Gate | Required | Now |
|---|---|---|
| handout build | 0 errors, 0 overfull, 0 underfull, 0 font warnings | ✅ 90 pp. |
| `checkfigs.py` | clean | ✅ |
| `checklabs.py` | clean | ✅ |
| `texparse.py --all` | every chapter parses; 3 questions = 3 answers | ✅ 8/8 |
| `scanprose.py` | **0 suspicious literals** | ✅ |
| `qa.py` | static checks clean | ✅ |
| `qa.py --render` | **every contact sheet looked at** | ✅ 9/9 |

Undefined references are expected until all 23 chapters exist: the written
chapters forward-reference ones that do not yet.

The last gate is not automatable and is not optional. In this course so far it
caught **the Algorithms course's tagline, "Design, Proof and Measurement", on
every title slide** — a clean build, clean static checks and a passing parse
all said nothing.

---

## Authoring contract

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{<part title>}
\begin{outcomes} ... \end{outcomes}      % 5 Bloom-verb outcomes
\begin{prereq}   ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
\section{...}                            % 4-6 sections, mixing:
  definitionbox / conceptbox / alertbox / examplebox
  \begin{measured} ... \end{measured}    % 12 chapters only
  \dsfig{<TikZ wireframe>}{...}{...}
  \begin{worked} ... \end{worked}        % every arithmetic step shown
  longtable with L{} columns
\begin{pitfall}        ... \end{pitfall}
\begin{lab}[title]     ... \end{lab}     % steps marked # --- n.
\begin{checkpoint}     ... \end{checkpoint}   % exactly 3
\begin{chaptersummary} ... \end{chaptersummary}
\begin{reviewq}        ... \end{reviewq}      % 7: 2 MCQ, 3 Short, 2 Applied
```

### Three style additions no sibling has

- **`\adminpath{Catalog, Products, Add New}`** — a click path, rendered with
  chevrons. A GUI course states one on nearly every page, and the slide
  pipeline cannot style what it cannot recognise. Comma list; `\uiel` trims the
  items, so a space after the comma is harmless.
- **`\uiel{Save}`** — one control, in sans against the body serif, so a reader
  skimming for "where do I click" finds every control without reading a word.
- **`\begin{measured}`** — graphite, because every other hue in the palette is
  taken. The optional argument *extends* the title rather than replacing it.

### Table widths

`\tabcolsep` eats about 0.42 cm per internal gap. 2-col 15.0 cm at top level,
3-col 14.5 cm, and **14.3 cm inside a `conceptbox`/`alertbox`** — that last one
is the one that bites.

---

## Layout

```
E-Commerce Application Development Fall 2026/
├── ECAD_Fall2026_Handout.tex    master: 6 parts, 23 chapters, 7 appendices
├── probe_master.tex             scratch master; delete when parts/ is done
├── ecadhandout.sty              house style + \uiel, \adminpath, measured
├── build.ps1  checklog.py  checkfigs.py  checklabs.py
├── parts/                       00_frontmatter, 6 openers, chapters, A–G
├── code/                        docker-compose.yml, shop/, waitfor.py, mNN_*.py
├── assets/                      logos
└── slides/                      the deck pipeline — see slides/README.md
```

**`parts/G_schedule.tex` is the source of record for the schedule.** The slide
pipeline parses that appendix, not the one-page PDF.

**`parts/C_answers.tex` has a hard format contract**: `\ansch{<N> --- <title>}`
followed immediately by an `ans` environment of **exactly three** items.
`texparse.py --all` checks it.

---

## Not done

- **Chapters 9–23.** Fifteen of twenty-three; Parts I and II are complete.
- **Appendices D, E and F.** A, B, C and G exist; C covers chapters 1–8 only.
- **Eight of the twelve measurement harnesses.** `m05_coldstart.py`,
  `m06_catalogue.py`, `m07_variations.py` and `m08_pageweight.py` are written.
  `code/run_all.py` and `code/sync_labs.py` do not exist yet.
- **The one-page lecture schedule** (`ECAD_Fall2026_Lecture_Schedule.tex`),
  generated from Appendix G.
- **`slides/README.md`** and **`code/README.md`**.
- **The ~15 screenshots.** Playwright is not installed; the TikZ wireframes
  carry the course meanwhile and no chapter's argument depends on a screenshot.
- The curriculum map in `../HEC AI-DS Curriculum Map/` has no entry for this
  course. Out of scope here, by instruction.
- No git commit has been made; nothing here is staged.
