"""Build one PDF per chapter, for handing out in class.

The master handout is one volume. That is the right object for a student to
keep and the wrong one to hand out in a 90-minute class, so the same files in
parts/ are also built one chapter at a time, each with its own cover sheet
naming the week it is taught, that week's laboratory and that week's
assessment -- all three read out of parts/G_schedule.tex, which stays the
single source of record.

    python makechapters.py              every chapter and appendix that exists
    python makechapters.py 5 6 7        those chapters
    python makechapters.py A B          those appendices
    python makechapters.py --list       what is available, and from where
    python makechapters.py --keep       leave the generated drivers in place

Two things are worth knowing about how it works.

Cross-references. A chapter cites its neighbours constantly -- "as Chapter 16
shows" -- and a chapter compiled on its own cannot resolve any of them, so
every one would print as "??". Instead the master's .aux is reused as an
external reference source (xr-hyper), filtered to drop the labels this chapter
defines itself; without that filtering every one of its own labels would be
multiply defined. That is why each extract is built twice: the first pass is
what tells us which labels are local.

Page numbers restart at 1 per extract and are local to it. Chapter, figure and
table numbers are the full volume's, because a reference that lands somewhere
different in the extract than in the book is worse than no reference.
"""
import argparse
import os
import re
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
PARTS = os.path.join(ROOT, "parts")
OUT = os.path.join(ROOT, "handouts")
SCHEDULE = os.path.join(PARTS, "G_schedule.tex")

# The master whose .aux supplies cross-chapter numbers. The real master is
# preferred; probe_master.tex is the scratch one that builds only the chapters
# written so far, and is what exists while the book is being written.
MASTERS = ("ECAD_Fall2026_Handout", "probe_master")

PART_OF = {1: 1, 2: 1, 3: 1, 4: 1,
           5: 2, 6: 2, 7: 2, 8: 2,
           9: 3, 10: 3, 11: 3, 12: 3,
           13: 4, 14: 4, 15: 4, 16: 4,
           17: 5, 18: 5, 19: 5, 20: 5,
           21: 6, 22: 6, 23: 6}


# --------------------------------------------------------------------- TeX
def braced(text: str, start: int) -> tuple[str, int]:
    """Contents of the {...} group beginning at text[start], and the index
    just past it. Counts braces, because chapter titles contain them."""
    assert text[start] == "{", text[start:start + 20]
    depth, i = 0, start
    while i < len(text):
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
            if depth == 0:
                return text[start + 1:i], i + 1
        i += 1
    raise ValueError("unbalanced braces")


def tidy(cell: str) -> str:
    """A schedule cell as one line of LaTeX. The table wraps its cells across
    source lines for readability; the cover wants them in one piece."""
    return " ".join(cell.split()).strip()


# ---------------------------------------------------------------- discovery
def discover() -> dict:
    """{key: {...}} for every chapter and appendix file in parts/."""
    found = {}
    for name in sorted(os.listdir(PARTS)):
        m = re.fullmatch(r"(\d\d|[A-G])_([a-z]+)\.tex", name)
        if not m:
            continue
        tag, slug = m.groups()
        src = os.path.join(PARTS, name)
        with open(src, encoding="utf-8") as fh:
            text = fh.read()

        c = text.find("\\chapter{")
        if c < 0:
            continue
        title, after = braced(text, c + len("\\chapter"))

        lab = re.search(r"\\label\{(?:ch|app):([a-z]+)\}", text[after:after + 200])

        partline = ""
        loc = text.find("\\locator{")
        if loc >= 0:
            _, i = braced(text, loc + len("\\locator"))
            partline, _ = braced(text, i)

        if tag.isdigit():
            num = int(tag)
            found[str(num)] = dict(kind="Chapter", num=num, shown=str(num),
                                   slug=slug, file=name, title=title,
                                   partline=partline, sort=(0, num),
                                   label=lab.group(1) if lab else slug)
        else:
            found[tag] = dict(kind="Appendix", num=ord(tag) - ord("A") + 1,
                              shown=tag, slug=slug, file=name, title=title,
                              partline=partline, sort=(1, tag),
                              label=lab.group(1) if lab else slug)
    return found


def schedule() -> dict:
    """{chapter number: (week, lab, assessment)} from the first longtable of
    parts/G_schedule.tex -- the same table the slide pipeline parses."""
    if not os.path.exists(SCHEDULE):
        return {}
    with open(SCHEDULE, encoding="utf-8") as fh:
        text = fh.read()
    body = text.split("\\begin{longtable}", 1)[-1].split("\\end{longtable}", 1)[0]
    body = body.split("\\endhead", 1)[-1]

    out = {}
    for row in body.split("\\\\"):
        row = row.replace("\\bottomrule", "").strip()
        if not row or row.startswith("%"):
            continue
        cells = [tidy(c) for c in row.split("&")]
        if len(cells) != 5 or not cells[0].isdigit():
            continue
        week, _topic, chapters, lab, assess = cells
        for item in chapters.split(","):
            item = item.strip()
            if "--" in item:
                lo, hi = (int(x) for x in item.split("--"))
                rng = range(lo, hi + 1)
            elif item.isdigit():
                rng = [int(item)]
            else:
                continue
            for n in rng:
                out[n] = (f"Week {week}", lab, assess)
    return out


def master_aux() -> str | None:
    """Whichever master .aux carries the most labels. The real master wins
    once it builds; until then the scratch one does."""
    best, count = None, 0
    for stem in MASTERS:
        path = os.path.join(ROOT, stem + ".aux")
        if not os.path.exists(path):
            continue
        with open(path, encoding="utf-8", errors="replace") as fh:
            n = sum(1 for line in fh if line.startswith("\\newlabel"))
        if n > count:
            best, count = path, n
    return best


# ------------------------------------------------------------------ writing
def driver(info: dict, facts: tuple, stem: str) -> str:
    """The generated standalone master for one chapter."""
    kind, shown, title = info["kind"], info["shown"], info["title"]
    week, lab, assess = facts
    foot = f"{kind} {shown} extract"
    # \setcounter{chapter} is one less than the number we want, because
    # \chapter increments before it prints. For an appendix \appendix has
    # already reset the counter and switched \thechapter to letters.
    if kind == "Chapter":
        preamble = f"\\setcounter{{chapter}}{{{info['num'] - 1}}}"
    else:
        preamble = f"\\appendix\n\\setcounter{{chapter}}{{{info['num'] - 1}}}"

    safe = title.replace("&", "\\&")
    return f"""% Generated by makechapters.py -- do not edit; edit parts/{info['file']}.
% {kind} {shown} of the E-Commerce Application Development handout, built on
% its own for distribution in class.
\\documentclass[11pt,a4paper,oneside]{{report}}
% xr-hyper before hyperref, which ecadhandout loads.
\\usepackage{{xr-hyper}}
\\usepackage{{ecadhandout}}
\\usepackage{{ecadchapter}}
% Cross-chapter numbers come from the master, minus the labels this file
% defines itself -- see makechapters.py for why the list is filtered.
\\externaldocument{{_xref_{stem}}}
\\hypersetup{{
  pdftitle={{E-Commerce Application Development --- {kind} {shown}: {safe}}},
  pdfsubject={{Class extract: {kind} {shown} of the Fall 2026 unified handout
               for INFT-4041, BSIT 5th semester}}}}
\\extractfooter{{{foot}}}
\\begin{{document}}
\\pagenumbering{{arabic}}
\\chapterextract{{{kind}}}{{{shown}}}{{{safe}}}{{{info['partline']}}}%
  {{{week}}}{{{lab}}}{{{assess}}}
{preamble}
\\input{{parts/{info['file'][:-4]}}}
\\end{{document}}
"""


def xref(stem: str, local: set[str]) -> None:
    """Write _xref_<stem>.aux: the master's labels, minus this chapter's own.

    Without the subtraction every label the chapter defines is also defined
    externally, and LaTeX reports each one as multiply defined -- which is a
    warning, and this course's build gate counts warnings."""
    path = os.path.join(ROOT, f"_xref_{stem}.aux")
    src, kept = master_aux(), []
    if src:
        with open(src, encoding="utf-8", errors="replace") as fh:
            for line in fh:
                if not line.startswith("\\newlabel{"):
                    continue
                name, _ = braced(line, line.index("{"))
                if name not in local:
                    kept.append(line)
    with open(path, "w", encoding="utf-8") as fh:
        fh.writelines(kept)


def labels_of(stem: str) -> set[str]:
    """Every label defined by the extract itself, read back from its .aux."""
    path = os.path.join(ROOT, stem + ".aux")
    out = set()
    if not os.path.exists(path):
        return out
    with open(path, encoding="utf-8", errors="replace") as fh:
        for line in fh:
            if line.startswith("\\newlabel{"):
                name, _ = braced(line, line.index("{"))
                out.add(name)
    return out


# ----------------------------------------------------------------- building
NOISE = re.compile(r"Infinite glue shrinkage")


def report(stem: str) -> dict:
    """The same metrics build.ps1 reports, for one extract."""
    path = os.path.join(ROOT, stem + ".log")
    if not os.path.exists(path):
        return dict(pages=0, errors=["no log"], undef=0, over=0, under=0,
                    fonts=0, pdfstr=0)
    with open(path, encoding="utf-8", errors="replace") as fh:
        log = fh.read()
    lines = log.splitlines()
    errors = [ln for ln in lines
              if (ln.startswith("!") or re.match(r"^.*:\d+: ", ln))
              and not NOISE.search(ln)]
    pages = re.findall(r"Output written on .*\((\d+) pages", log)
    return dict(
        pages=int(pages[-1]) if pages else 0,
        errors=errors,
        undef=len(re.findall(r"Reference `[^']*' on page [^ ]* undefined", log)),
        over=len(re.findall(r"Overfull \\[hv]box", log)),
        under=len([ln for ln in lines if ln.startswith("Underfull")]),
        fonts=len(re.findall(r"LaTeX Font Warning", log)),
        pdfstr=len(re.findall(r"Token not allowed", log)))


def latexmk(stem: str) -> int:
    return subprocess.run(
        ["latexmk", "-pdf", "-interaction=nonstopmode", "-file-line-error",
         "-silent", stem],
        cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode


def sweep(stem: str) -> None:
    for ext in (".tex", ".aux", ".log", ".out", ".toc", ".lof", ".lot",
                ".fls", ".fdb_latexmk", ".synctex.gz"):
        for path in (os.path.join(ROOT, stem + ext),
                     os.path.join(ROOT, f"_xref_{stem}" + ext)):
            if os.path.exists(path):
                os.remove(path)


def build(key: str, info: dict, facts: tuple, keep: bool) -> dict:
    stem = f"_ch{info['shown'].zfill(2) if info['kind'] == 'Chapter' else info['shown']}"
    tex = os.path.join(ROOT, stem + ".tex")
    with open(tex, "w", encoding="utf-8") as fh:
        fh.write(driver(info, facts, stem))

    # Pass 1 resolves nothing external, but its .aux names every local label.
    xref(stem, set())
    latexmk(stem)
    xref(stem, labels_of(stem))
    latexmk(stem)

    metrics = report(stem)
    pdf = os.path.join(ROOT, stem + ".pdf")
    name = (f"ECAD_Ch{info['shown'].zfill(2)}_{info['slug']}.pdf"
            if info["kind"] == "Chapter"
            else f"ECAD_App{info['shown']}_{info['slug']}.pdf")
    dest = os.path.join(OUT, name)
    if os.path.exists(pdf):
        os.makedirs(OUT, exist_ok=True)
        shutil.move(pdf, dest)
        metrics["pdf"] = name
    else:
        metrics["pdf"] = None
    if not keep:
        sweep(stem)
    return metrics


# --------------------------------------------------------------------- main
def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("which", nargs="*", help="chapter numbers or appendix letters")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--keep", action="store_true",
                    help="leave the generated drivers and logs in place")
    args = ap.parse_args()

    found, weeks = discover(), schedule()
    aux = master_aux()

    if args.list:
        print(f"cross-references from: {os.path.basename(aux) if aux else '(none)'}")
        for key in sorted(found, key=lambda k: found[k]["sort"]):
            i = found[key]
            w = weeks.get(i["num"], ("", "", ""))[0] if i["kind"] == "Chapter" else ""
            print(f"  {i['kind'][:3]} {i['shown']:>2}  {w:<8} {i['title']}")
        return 0

    keys = args.which or sorted(found, key=lambda k: found[k]["sort"])
    missing = [k for k in keys if k not in found]
    if missing:
        print(f"not written yet: {', '.join(missing)}", file=sys.stderr)
        return 1

    print(f"cross-references from {os.path.basename(aux) if aux else '(none)'}\n")
    head = f"{'':>4} {'pp':>3} {'err':>3} {'??':>3} {'ovr':>3} {'unf':>3} {'fnt':>3}  file"
    print(head)
    print("-" * len(head))
    bad = 0
    for key in keys:
        info = found[key]
        facts = weeks.get(info["num"], ("", "", "")) \
            if info["kind"] == "Chapter" else ("", "", "")
        m = build(key, info, facts, args.keep)
        flag = "" if not m["errors"] and m["pdf"] else "  <-- FAILED"
        if flag:
            bad += 1
        print(f"{info['shown']:>4} {m['pages']:>3} {len(m['errors']):>3} "
              f"{m['undef']:>3} {m['over']:>3} {m['under']:>3} {m['fonts']:>3}  "
              f"{m['pdf'] or '(none)'}{flag}")
        for line in m["errors"][:3]:
            print(f"       {line}")

    print(f"\n{len(keys) - bad}/{len(keys)} built into handouts/")
    if any(weeks.get(found[k]['num'], ('', '', ''))[0] == ''
           and found[k]['kind'] == 'Chapter' for k in keys):
        print("note: a chapter with no week is one G_schedule.tex does not place")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())
