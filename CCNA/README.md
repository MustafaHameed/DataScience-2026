# CCNA 200-301 v2.0 — Course Volume

One book containing the syllabus, the notes and the labs for a course taught to
Cisco CCNA 200-301 **version 2.0** — written from zero, with no prior networking
assumed, teaching protocol theory and device configuration together.

**Current build:** 291 pages · zero LaTeX errors, zero undefined references, zero
overfull boxes, zero underfull boxes, zero font warnings, no orphaned pages.
32 figures, all TikZ.

> **Status: complete.** All 34 chapters and all seven appendices are written.
> Every blueprint topic is covered and every quality gate reads zero.

## Why v2.0 and not v1.1

v1.1 retires 2 February 2027; v2.0 goes live the next day, keeping the same exam
number. Anyone entering a course from Fall 2026 sits v2.0. The two blueprints are
different enough that a v1.1 book teaches the wrong exam — QoS, NTP, TFTP, REST
APIs, JSON and Terraform are gone; OSPFv3, SFTP/SCP, DNS record diagnosis, storm
control, RA guard, PoE and a whole AI domain have arrived.

**The planning fact that drives the whole design:** of the 29 published topics,
8 begin *Troubleshoot* or *Diagnose* and 8 begin *Configure* — so **55% of the
exam is hands-on**, against 21% *Describe*. Under v1.1 the word "troubleshoot"
appeared in no topic statement at all.

That is why **every chapter carries a `breakfix`** — a broken device to diagnose —
and why 32 of the 34 carry a lab, rather than prose with a lab bolted on the end.
`config` and `verify` blocks appear wherever there is something to configure:
chapters covering *describe*-only topics (wireless principles, VPNs, AI) have
none, because the blueprint asks for none and inventing some would teach the
wrong exam.

## Source of record

`200-301_CCNA_v2.0_Exam_Topics_PDF.pdf` in this folder — Cisco's own published
document, marked Cisco Public. Topic numbers, weights and wording come from it
and nothing else. `200-301-CCNA-v1.1.pdf` is kept for the historical comparison
in the front matter only.

Cisco revises blueprints without renaming exams. Re-check
`cisco.com/go/certroadmap` before teaching from this.

## Layout

| Path | Contents |
| --- | --- |
| `CCNA_v2_Handout.tex` | Master file. `\input`s everything; produces one PDF. |
| `ccna.sty` | **Single source of style.** Palette, boxes, IOS listings, topology macros, blueprint tags. |
| `parts/00_frontmatter.tex` | Title page, exam facts, how to use the volume, syllabus with CLOs and assessment. |
| `parts/00_partN_opener.tex` | One concept-map opener per part. |
| `parts/01_*.tex` … `parts/34_*.tex` | The 34 chapters, in teaching order. |
| `parts/A_*.tex` … `parts/G_*.tex` | Appendices A–G. |
| `tools/checkall.py` | **Runs all seven gates and reports one verdict.** |
| `tools/logcheck.py` | Reads the build log; the seven typographic gates. |
| `tools/lintsource.py` | Source traps LaTeX reports badly or in the wrong place. |
| `tools/check_coverage.py` | Every blueprint topic is covered by a chapter. |
| `tools/check_furniture.py` | Every chapter carries its required pedagogy blocks. |
| `tools/check_answers.py` | Appendix E answers every checkpoint question. |
| `tools/check_sectionrefs.py` | Literal "Section N.M" references point where they claim. |
| `tools/check_pages.py` | No page is left almost empty by a bad break. |
| `tools/prepare_logos.py` | Turns the supplied institutional artwork into usable assets. |
| `tools/gen_appendices.py` | Generates Appendix D and the lab index. |
| `tools/gen_drills.py` | Generates Appendix C's drills with computed answers. |
| `build.ps1` | Build + quality report; exits non-zero on any error. |

## Build

```powershell
.\build.ps1            # build CCNA_v2_Handout.pdf
.\build.ps1 -Clean     # remove aux files, then build
python tools\checkall.py   # all seven gates; exits non-zero on any failure
```

Each gate exists because the error it catches actually happened here and
nothing else caught it — a raw `_` reported at the wrong line, a chapter with no
Break & Fix, three literal section references left pointing at the wrong
sections after content moved, and five chapters that ended with a single review
question stranded on an otherwise blank page.

**Do not check the log with `grep '^!'`.** latexmk runs with
`-file-line-error`, which reformats every error as `file:line: message` with no
leading `!`, so that pattern silently reports a clean build over real errors. It
did exactly that once here. `build.ps1` and `logcheck.py` both use the correct
pattern; use one of them.

Requires MiKTeX (or TeX Live) with `tikz`, `tcolorbox`, `listings`,
`fontawesome5`, `booktabs`, `longtable`, `ragged2e`, `needspace`, `float` and
`caption`. `latexmk` is a Perl script and MiKTeX ships no interpreter, so
`build.ps1` borrows the `perl.exe` bundled with Git for Windows.

`checkall.py` must exit 0 before this ships. There is no gate that is
optional and none that is advisory.

## Chapter furniture

```latex
\chapter{...}\label{ch:short}
\bp{2.1,2.2}                       % blueprint topics -- NO SPACES in the list
\begin{outcomes} ... \end{outcomes}
\begin{prereq} ... \end{prereq}
\begin{keyterms} ... \end{keyterms}
...definitionbox / conceptbox / alertbox / \ccnafig...
\begin{config}{what this builds} ... \end{config}
\begin{verify}{SW1}{show vlan brief} ... \end{verify}
\begin{pitfall} ... \end{pitfall}
\begin{breakfix}[the symptom] ... \end{breakfix}
\begin{lab}{title}{Packet Tracer} ... \end{lab}
\begin{examnote}[On the exam] ... \end{examnote}
\begin{checkpoint} ... \end{checkpoint}
\begin{chaptersummary} ... \end{chaptersummary}
\begin{reviewq} ... \end{reviewq}
```

## Authoring conventions

Carried over from the two sibling projects in this repo
(`../Data Science Fall 2026/dshandout.sty`, `../HEC AI-DS Curriculum Map/hecmap.sty`)
for the reasons those files document:

- **Paragraph columns use `L{<width>}`, never `p{<width>}`** — ragged right with
  expensive hyphens.
- **Column widths.** The text block here is **16.2 cm** (2.4 cm margins), *not*
  the 16.6 cm of the curriculum map. `\tabcolsep` eats ~0.42 cm per gap, so the
  budget is **15.78 cm for two columns, 15.36 for three, 14.93 for four, 14.51
  for five.**
- **Kickers are medium-weight small caps** — Latin Modern has no `T1/lmr/bx/sc`.
- `\widowpenalty`/`\clubpenalty` 10000, `\brokenpenalty` 4000.
- `\_` carries an `\allowbreak`; all figures are TikZ-native, no external images.

Two traps specific to this volume:

- **`verify` takes two arguments, device then command** — `\begin{verify}{SW1}{show vlan brief}`.
  The `#` of the prompt is generated internally, because a raw `#` in a tcolorbox
  title is an "Illegal parameter number" error.
- **`^` inside `\cmd{}` is a superscript.** Write `\cmd{| exclude \textasciicircum!}`.
- **A raw `_` in any box title is a subscript**, and the error surfaces at the
  `\end{}` of the environment rather than at the title. Write `ERR\_DISABLE`.
  The same applies to `#`, `$`, `%`, `&` and `~`.
- **Passing a tinted colour into a TikZ style argument breaks xcolor.** A style
  written `f/.style={fill=#1!10, text=#1!40!black}` invoked as `f=dom4!60`
  expands to `dom4!60!40!black`, and xcolor reads `40` as a colour name. Pass
  the base colour and override the tint on the node: `[f=dom4, fill=dom4!28]`.
- **Part openers set `\structuralfigures{PN.}`; `\chapter` takes it back**
  automatically via an `AddToHook` in `ccna.sty`. Chapter files should not call
  `\normalfigures` themselves.
- **Two tall `\ccnafig` floats in one chapter can produce `Overfull \vbox … while
  \output is active`** — a message with no file and no line number, so it cannot
  be located by reading the log. Bisect by commenting chapters, then figures, out
  of the build. The fix is `\ccnafigh` (forced placement) for the tall ones;
  adjusting `\topfraction`/`\bottomfraction` does **not** help, which is itself
  the clue that it is the float *queue* and not float *sizing*.

`tools/lintsource.py` mechanically catches the raw-`_`-in-title and
tinted-colour-into-a-style traps. Run it before building; both were made twice
before it existed.

### Drawing figures

All figures are TikZ, drawn inline, numbered through `\ccnafig` (floating) or
`\ccnafigh` (forced in place). Three rules learned the hard way:

- **Give every prose node an explicit `text width`.** Without one a node is as
  wide as its text, which both collides with its neighbours and drags the
  picture's bounding box out sideways, shifting the whole figure off centre.
- **Never name a style `cap`** — it collides with PGF's built-in line-cap key
  and fails with "the key '/tikz/cap' requires a value".
- **A label with a white background, placed *on* a short arrow, erases it.**
  Put flow labels above their arrows, not on them.

Device colours follow the domain palette: `dom2` teal for switches, `dom3`
purple for routers, `dom1` blue for servers, `dom5` gold for wireless and
telephony, `dom4` red for security boundaries, `neutral` grey for hosts.

Listing environments, all of which take their title as an argument:
`config` (IOS you type), `verify{device}{command}` (IOS output), `hostcmd{OS}{command}`
(a command run on a Windows/macOS/Linux host), `playbook{filename}` (YAML).

## Accuracy

Configurations are written to be syntactically correct for IOS and IOS XE and to
follow the command order a device requires. **They have not been verified on live
hardware.** Run each lab once on your own platform before issuing it — that is
worth more than a validation claim this book cannot honestly make. The front
matter says so to the reader too.

## Progress

| Part | Chapters | State |
| --- | --- | --- |
| Front matter | — | **Written** — title, exam facts, how-to-use, syllabus, CLOs, assessment |
| I Foundations | 1–7 | **Written** |
| II Switching and Network Access | 8–14 | **Written** |
| III IP Routing | 15–19 | **Written** |
| IV Network Services and Security | 20–26 | **Written** |
| V Wireless, Virtualisation and Operations | 27–32 | **Written** |
| VI Sitting the Exam | 33–34 | **Written** |
| Appendices A–G | — | **Written** — D and the lab index generated from source |

Stubs carry their chapter heading, label and `\bp{}` tag, so the master compiles
and `check_coverage.py` already reports all 29 topics accounted for. Filling a
stub does not change the structure.
