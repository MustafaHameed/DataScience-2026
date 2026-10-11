# E-Commerce Application Development — Fall 2026

Teaching materials for the BSIT *E-Commerce Application Development* course
(**INFT-4041**, 5th semester): a handout, lecture decks, a running shop the
whole course is built on, and the measurements that keep its claims honest.

**Dr. Mustafa Hameed · Department of Information Technology · The Islamia
University of Bahawalpur**

> **Status: complete.** All 23 chapters, all seven appendices, 24 decks, the
> one-page schedule, 15 measurement harnesses and a per-chapter PDF for every
> chapter and appendix. See [Not done](#not-done) for the two things that
> remain.

---

## What is here

| Artefact | Size | Build with |
|---|---|---|
| `ECAD_Fall2026_Handout.pdf` | **219 pp.**, 23 chapters, 7 appendices, 20 figures | `.\build.ps1` |
| `ECAD_Fall2026_Lecture_Schedule.pdf` | one page, 16 weeks | `pdflatex ECAD_Fall2026_Lecture_Schedule.tex` |
| `handouts/*.pdf` | **30 class extracts** — one per chapter and appendix | `python makechapters.py` |
| `slides/out/*.pptx` | **24 decks, 1,000 slides** | `python slides\build_slides.py --all` |
| `code/` | a running shop, a mock payment provider, 15 harnesses, a PHP extension | `docker compose up -d --build` |

### Per-chapter PDFs, for handing out in class

The master is the right object for a student to keep and the wrong one to hand
out in a ninety-minute class. `makechapters.py` builds the same `parts/` files
one chapter at a time, each with a cover sheet naming the week it is taught,
that week's laboratory and that week's assessment — all three read out of
`parts/G_schedule.tex`, which stays the single source of record.

```powershell
python makechapters.py            # all 23 chapters and 7 appendices
python makechapters.py 9 10       # just those
python makechapters.py A B        # appendices
```

Page numbers restart at 1 per extract; chapter, figure and table numbers are
the full volume's, so a reference followed in an extract lands where it does in
the book. Cross-chapter references resolve through the master's `.aux`,
filtered to drop the labels the extract defines itself — which is why each one
is built twice.

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

### The rest, one line each

| Ch | Finding |
|---|---|
| 9 | Adding three products left the cookie at **49 B** and the session record at **18 B**, and added three rows to `oc_cart`. **The basket is not in the session.** Changing the session engine moves the 18 bytes and nothing else |
| 9 | An **empty** basket is read six times per page and a full one once, because the cache test cannot tell "empty" from "not loaded" |
| 10 | **A stock install cannot complete a guest checkout.** `config_checkout_payment_address` is read by ten controllers and written by none, so no shipping address is ever stored. Every page still returned 200 |
| 10 | A checkout is **607 queries**; `confirm` performs **25 of its 55 writes** |
| 12 | Moving the coupon line after the tax line changed the customer's bill from **770.16 to 784.20** — one integer in an unlabelled admin field |
| 14 | Gateway latency is **bimodal**: median 321 ms, p95 3,774 ms. A 1-second timeout discards 11.7% of payments and a 3-second timeout 10.7% — the middle buys nothing |
| 15 | bcrypt cost 8→14 ran 13.7 … 874 ms, an **exact doubling**. A wrong password on a registered address took **88.6 ms** and on an unknown one **23.0 ms** — a 65 ms gap, which *is* one bcrypt verify. **The leak is caused by the defence** |
| 16 | The shop passed **5 of 15** checks as installed and **14 of 15** after one configuration file. The PHP error log was readable over HTTP. SQL injection and XSS were already handled |
| 16 | Three hardening controls were applied correctly, accepted by Apache without a warning, and **did nothing** |
| 17 | The same six orders were reported as **6, 5 and 2** by the table, the widgets and the chart. All three correct; none labelled with its filter |
| 18 | The admin API ships enabled with an **empty** address allow-list, which permits nothing — so the admin panel cannot change an order |
| 18 | Status 11 *Refunded* returns stock; status 12 *Reversed* does not. A return marked "Refunded" writes two rows and **does nothing else** |
| 19 | A module instance costs 4 ms and **exactly 4 queries** — of which **one is the module's** |
| 20 | Throughput **peaked at 5 concurrent clients and fell at 20** while the median grew 5.6×. opcache doubled throughput at the saturation point |
| 20 | A 1.0 s backup took **6.4 s to restore** |
| 21 | Turning on readable URLs without the server rewrite points **every link in the shop at a 404** — and the cause is Chapter 16's `AllowOverride None` |

---

## Building

```powershell
.\build.ps1                           # the handout
pdflatex ECAD_Fall2026_Lecture_Schedule.tex
python makechapters.py                # the 30 class extracts
python checkfigs.py                   # TikZ line-break and reserved-name traps
python checklabs.py                   # listings: ASCII, <= 76 cols, step markers
python code\sync_labs.py              # every command the labs print exists
python code\run_all.py --quick        # every quoted result still holds
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
| handout build | 0 errors, 0 undefined refs, 0 overfull, 0 underfull, 0 font warnings | ✅ **219 pp.** |
| schedule build | one page, 0 warnings | ✅ |
| `makechapters.py` | 30 extracts, 0 errors each | ✅ |
| `checkfigs.py` | clean | ✅ |
| `checklabs.py` | clean | ✅ |
| `code/sync_labs.py` | every command, service, path and PHP line in the labs exists | ✅ 128 checked |
| `code/run_all.py --quick` | every quoted result still reproduces | ✅ **24/24** |
| `texparse.py --all` | every chapter parses; 3 questions = 3 answers | ✅ **23/23** |
| `scanprose.py` | **0 suspicious literals** | ✅ |
| `qa.py` | static checks clean | ✅ |
| `qa.py --render` | **every contact sheet looked at** | ✅ 24/24 |

The last gate is not automatable and is not optional. In this course it caught
**the Algorithms course's tagline, "Design, Proof and Measurement", on every
title slide**; Chapter 6's title overflowing Chapter 5's chip by 156%; a
15-row table running off the bottom of a slide; and a 30-line PHP listing doing
the same. A clean build, clean static checks and a passing parse said nothing
about any of them.

`run_all.py` checks the *shape* of each result, not the number — the numbers
are facts about one laptop. If a check fails, the handout and the software have
diverged, and the first job is to find out which one moved.

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
├── ECAD_Fall2026_Lecture_Schedule.tex   the one-page schedule
├── ecadhandout.sty              house style + \uiel, \adminpath, measured
├── ecadchapter.sty              cover sheet for a single-chapter extract
├── makechapters.py              builds handouts/, one PDF per chapter
├── build.ps1  checklog.py  checkfigs.py  checklabs.py
├── parts/                       00_frontmatter, 6 openers, chapters, A–G
├── handouts/                    30 class extracts, one per chapter/appendix
├── code/                        the stack, 15 harnesses, modules/ — see its README
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

- **The ~15 screenshots.** Playwright is not installed (`pip install
  playwright && playwright install chromium`). The twenty TikZ wireframes
  carry the course meanwhile, and **no chapter's argument depends on a
  screenshot** — they were always the degradable part of the plan.
- The curriculum map in `../HEC AI-DS Curriculum Map/` has no entry for this
  course. Out of scope here, by instruction.
- No git commit has been made; nothing here is staged.

## Four faults in the shipped software that the course turns into teaching

Every one of them is a feature that is present, configured, enabled and inert,
and every one of them leaves every page returning 200.

1. **`config_checkout_payment_address` is never written by the installer**, so
   a stock install cannot complete a guest checkout (Ch 10).
2. **`.htaccess.txt` is never renamed, and `AllowOverride None` means it would
   not be read if it were**, so the PHP error log is served over HTTP (Ch 16)
   — and, two chapters later, readable URLs point at 404s (Ch 21).
3. **`oc_api_ip` ships empty**, and an empty allow-list permits nothing, so the
   admin panel cannot change an order (Ch 18).
4. **The seeder's own `store_id` bug**, ours rather than the platform's, made
   50,000 products invisible to every customer while the admin panel showed
   them. It is recorded in Ch 6 rather than quietly fixed, because it is the
   best example in the course of a system that answered correctly and meant
   something else.
