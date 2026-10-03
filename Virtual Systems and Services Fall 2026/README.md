# Virtual Systems and Services — Fall 2026

Complete teaching materials for the BSIT elective **Virtual Systems and
Services**, Department of Information Technology, The Islamia University of
Bahawalpur: a 239-page handout, 24 lecture decks, and a runnable lab for every
chapter.

## Where this course comes from

HEC published a full outline for this course exactly once: **HEC 2017**
(*Curriculum of BS & MS CS, SE & IT*, pp. 71–130) — 3 credit hours,
prerequisite *Programming Fundamentals*, a hypervisor course covering VMware,
Xen and Hyper-V, para-virtualized components and processor support for
virtualization. HEC 2023 keeps the title as an unspecified BSIT domain
elective (credit count only, no published contents). **HEC 2025 drops the
course entirely** — its ground is split between *Cloud Computing* (a new
major-core course, Semester VI) and *Microservices Architecture and Docker
Containers* (Network Infrastructure & Cloud Computing specialization). See
[`../HEC AI-DS Curriculum Map/`](../HEC%20AI-DS%20Curriculum%20Map/) for the
full edition-by-edition trace.

This handout modernises the 2017 outline rather than reviving it verbatim: the
hypervisor foundations (Parts I–II) are taught close to the original
specification, then carried forward (Parts III–VI) into containers, Kubernetes,
cloud service models and confidential computing — the ground HEC 2025 hands to
*Cloud Computing* and *Microservices Architecture and Docker Containers*.

The curriculum map's own verdict is that where both courses exist, this one
should be retired. **Appendix F answers that directly** rather than ignoring
it: it states where the line falls against all four neighbouring courses, and
the specific case in which this one still earns three credit hours.

The 2017 set texts (Stanney's *Handbook of Virtual Environments*, Burdea's
*Virtual Reality Technology*) are virtual-reality references, not
virtualization references, and are not used — Appendix E says why in full.

## Build

```powershell
.\build.ps1              # the handout: latexmk, then nine checks on the log
python checkfigs.py      # TikZ faults that only surface as cascades
python checklabs.py      # lab listings: ASCII, <= 76 columns, step markers
python checklog.py       # the build report again, without the latexmk noise
```

The gate is that every count is zero:

```
Pages                 239
Errors                0
Undefined refs        0
Overfull hboxes       0
Overfull vboxes       0
Underfull boxes       0
Font warnings         0
PDF-string warns      0
```

Requires MiKTeX with `latexmk` and a Perl interpreter (`build.ps1` finds Git's
if none is on the path), plus `fontawesome5`, `tcolorbox`, `pgfplots`,
`listings`, `microtype` and `needspace`.

## Layout

```
VSS_Fall2026_Handout.tex        master: 6 parts, 23 chapters, 7 appendices
vsshandout.sty                  single source of style and pedagogy
build.ps1                       build + log gate
checkfigs.py                    TikZ line-break and reserved-name checks
checklabs.py                    lab listing checks
checklog.py                     quiet build report
parts/                          front matter, part openers, chapters, appendices
code/                           every lab, extracted from the handout
slides/                         the deck pipeline — see slides/README.md
VSS_Fall2026_Lecture_Schedule.tex   the one-page schedule
```

`parts/G_schedule.tex` is the **source of record** for the schedule. The
one-page PDF is generated from it and the slide pipeline parses it; change the
appendix first.

## Structure

**Part I — Foundations of Virtualization** (1 What Virtualization Is ·
2 Virtualization Architectures · 3 The Virtual Machine Monitor)

**Part II — Hypervisors and Compute Virtualization** (4 Hardware Support ·
5 Major Hypervisor Platforms · 6 CPU and Memory Virtualization ·
7 Paravirtualization and Device Virtualization)

**Part III — Containers and Lightweight Virtualization** (8 From Virtual
Machines to Containers · 9 Docker and the Container Ecosystem · 10 Kubernetes ·
11 Lightweight and Secure Virtualization)

**Part IV — Storage, Network and Cloud Virtualization** (12 Storage ·
13 Network Virtualization and SDN · 14 Cloud Service and Deployment Models ·
15 Desktop and Application Virtualization)

**Part V — Management, Performance and Security** (16 Management Platforms ·
17 High Availability, Live Migration and DR · 18 Performance Monitoring and
Capacity Planning · 19 Security in Virtualized Environments)

**Part VI — Emerging Hardware and Practice** (20 GPU and Accelerator
Virtualization · 21 Confidential Computing · 22 Case Studies · 23 The Semester
Project)

Appendices: A Laboratory Toolkit · B Command Reference · C Answers to the
Checkpoints · D Glossary · E Resources · F The Boundary with the Neighbouring
Courses · G The Sixteen-Week Schedule.

## Authoring conventions

Every chapter has the same furniture, in the same order. **The slide pipeline
parses by position, so a chapter missing any of it is unfinished.**

```latex
\chapter{...}\label{ch:shortname}
\locator{<part 1-6>}{<part title>}
\begin{outcomes}  ... \end{outcomes}     % 5 Bloom-verb outcomes
\begin{prereq}    ... \end{prereq}       % \chref{} back-references
\begin{keyterms}  ... \end{keyterms}     % $\bullet$-separated
\section{...}                            % 4-6 sections, mixing:
  definitionbox / conceptbox / alertbox / worked / longtable / \dsfig
\begin{pitfall}         ... \end{pitfall}
\begin{lab}[title]      ... \end{lab}    % one listing, steps marked # --- n.
\begin{checkpoint}      ... \end{checkpoint}   % exactly 3 questions
\begin{chaptersummary}  ... \end{chaptersummary}
\begin{reviewq}         ... \end{reviewq}      % 7: 2 MCQ, 3 Short, 2 Applied
```

- **Figures** go through `\dsfig{<tikz>}{<caption>}{<label>}`, never a bare
  `tikzpicture`. All native TikZ/pgfplots — no external images.
- **Cross-references** with `\chref{shortname}` and `\dsref{label}`.
- **Checkpoint answers** go in `parts/C_answers.tex`, keyed
  `\ansch{<N> --- <title>}` with exactly three items. The slide builder
  regexes this; renaming the macro breaks every deck.
- **New terms** go in `parts/D_glossary.tex`.
- **Tables** use `L{<width>}` never `p{<width>}`; widths sum to about 15.6 cm
  for two columns, 15.0 for three, 14.7 for four, 14.4 for five.
- `vsshandout.sty` is the single source of style. Change appearance there and
  nowhere else.

### Three traps, each of which cost real time here

1. **`\\` in a TikZ node with no `align` or `text width`** — on the node or on
   a style it names. `checkfigs.py` catches it.
2. **`\\` inside a nested group of an aligned node** —
   `{Docker\\\emph{a\\b}}`. The node looks correct and the error surfaces at
   the end of the figure, as a cascade of dozens of unrelated messages.
   `checkfigs.py` catches this too, and has a self-test.
3. **A TikZ style named `cap`, `step`, `pos`, `in` or `out`** shadows a
   built-in key, and the error names the key rather than your style.

And one about tooling rather than LaTeX: **do not edit `.tex` through a shell
heredoc.** Git Bash collapses `\\` to `\` silently. Use an editor or a Python
file.

## Lab code

`code/` holds every lab, **extracted from the handout** by
`code/sync_labs.py` — the handout is the source, and editing a file in `code/`
is pointless because the next sync overwrites it.

```powershell
python code\sync_labs.py            # write the files
python code\sync_labs.py --check    # fail if anything has drifted
```

All 20 shell labs pass `bash -n`; all listings are ASCII-only and within 76
columns. They have **not** been executed end to end on the machine that
produced this handout, which has no hypervisor, container engine or C compiler.
`code/README.md` states this and lists what each lab needs. Three labs need
hardware not everyone has and ship captured output instead.

## Slides

24 decks, 1,014 slides, generated from `parts/`. See
[`slides/README.md`](slides/README.md) for the build, the deck grammar and the
differences from the sibling courses' copies of the same pipeline.

## Dated material

Three chapters date faster than the rest and every dated claim in them was
checked at the time of writing (October 2026), not recalled:

- **Chapter 14** — cloud pricing and reservation discounts, from the provider's
  own figures. Re-check before quoting.
- **Chapter 19** — the 2025 ESXi escape chain and CVE-2026-47876, a VM escape in
  a paravirtual network adapter; and the 9.8-scored vCenter flaws, which are
  the chapter's argument that blast radius beats severity score.
- **Chapter 21** — the state of SEV-SNP, TDX and ARM CCA availability, which
  changed several times during 2026.

Appendix E's final section says where to check each of them.
