#!/usr/bin/env python3
"""Render the handout's TikZ figures as slide-ready vector and bitmap files.

The handout already draws 101 figures in TikZ and pgfplots. Redrawing them for
the slides would guarantee that they drift, so they are compiled from the
handout's own source instead::

    \\dsfig{..}{..}{label}  ->  build/<label>.tex (standalone)  ->  PDF
                             ->  out/<label>.svg   (what PowerPoint shows)
                             ->  out/<label>.png   (300 dpi fallback)

The standalone preamble is *generated from mlhandout.sty* on every run -- the
colour definitions and the TikZ/pgfplots libraries are lifted out of it, not
copied -- so a palette change in the handout reaches the slides next build.
The .sty itself cannot be loaded: it sets up geometry, fancyhdr and titlesec
for a report class that standalone does not provide.

Compiles are cached by content hash in out/manifest.json; only figures whose
source (or the preamble) changed are rebuilt.

Progressive builds
------------------
A figure with an annotated copy at src/<label>.tex is compiled once per build
level. Wrap elements in \\stepvis{n}{..} (appears from level n), \\steponly{n}
{..} (level n alone -- for swapping a caption) or \\stepdim{n}{..} (shown only
*before* level n, for a faint placeholder). The copy records the hash of the
handout figure it was made from; if the handout figure changes afterwards the
build warns, because the stepped version is now out of date.

    python extract_figures.py              # everything that changed
    python extract_figures.py --force      # recompile everything
    python extract_figures.py --label rag  # one figure
    python extract_figures.py --check-crop # verify pinned bounding boxes
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor

HERE = os.path.dirname(os.path.abspath(__file__))
SLIDES = os.path.dirname(HERE)
ROOT = os.path.dirname(SLIDES)
STY = os.path.join(ROOT, "mlhandout.sty")
SRC = os.path.join(HERE, "src")
BUILD = os.path.join(HERE, "build")
OUT = os.path.join(HERE, "out")
MANIFEST = os.path.join(OUT, "manifest.json")
DPI = 300

sys.path.insert(0, SLIDES)
import texparse as T          # noqa: E402


# ---------------------------------------------------------------------------
# Preamble, generated from mlhandout.sty
# ---------------------------------------------------------------------------

def _balanced_after(s: str, i: int) -> str:
    g, _ = T.find_group(s, s.index("{", i))
    return g


def preamble() -> str:
    sty = open(STY, encoding="utf-8").read()
    sty = T.strip_comments(sty)
    libs = [_balanced_after(sty, m.start())
            for m in re.finditer(r"\\usetikzlibrary\{", sty)]
    plots = [_balanced_after(sty, m.start())
             for m in re.finditer(r"\\usepgfplotslibrary\{", sty)]
    compat = re.search(r"\\pgfplotsset\{compat=[^}]*\}", sty)
    colours = "\n".join(re.findall(
        r"\\definecolor\{[^}]*\}\{[^}]*\}\{[^}]*\}", sty))
    lines = [
        r"\documentclass[border=4pt]{standalone}",
        r"\usepackage{lmodern}",
        r"\usepackage[T1]{fontenc}",
        r"\usepackage[utf8]{inputenc}",
        r"\usepackage{textcomp}",
        r"\usepackage{amsmath,amssymb}",
        r"\usepackage{xcolor}",
        r"\usepackage{tikz}",
        r"\usepackage{pgfplots}",
        r"\usepackage{fontawesome5}",
        r"\usepackage{ragged2e}",
    ]
    lines += [r"\usetikzlibrary{%s}" % re.sub(r"\s+", "", l) for l in libs]
    lines += [r"\usepgfplotslibrary{%s}" % re.sub(r"\s+", "", p)
              for p in plots]
    lines.append(compat.group(0) if compat else r"\pgfplotsset{compat=1.18}")
    lines += [
        "% --- colours, lifted from mlhandout.sty",
        colours,
        "% --- helpers the figure bodies use",
        r"\newcommand{\term}[1]{\textbf{\textcolor{primary}{#1}}}",
        r"\newcommand{\dsfitw}[1]{#1}",
        r"% --- progressive builds",
        r"\newcount\buildlevel \buildlevel=99",
        r"\newcommand{\stepvis}[2]{\ifnum#1>\buildlevel\relax\else#2\fi}",
        r"\newcommand{\steponly}[2]{\ifnum#1=\buildlevel #2\fi}",
        r"\newcommand{\stepdim}[2]{\ifnum#1>\buildlevel #2\fi}",
        r"\begin{document}",
    ]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Compiling
# ---------------------------------------------------------------------------

def _sha(*parts: str) -> str:
    h = hashlib.sha1()
    for p in parts:
        h.update(p.encode("utf-8"))
    return h.hexdigest()[:16]


def load_manifest() -> dict:
    if os.path.exists(MANIFEST):
        try:
            return json.load(open(MANIFEST, encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_manifest(m: dict) -> None:
    os.makedirs(OUT, exist_ok=True)
    with open(MANIFEST, "w", encoding="utf-8") as fh:
        json.dump(m, fh, indent=1, sort_keys=True)


def compile_one(name: str, body: str, pre: str, level: int | None = None
                ) -> tuple[bool, str, tuple[float, float]]:
    """Compile *body* to out/<name>.{svg,png}. Returns (ok, why, size_pt)."""
    os.makedirs(BUILD, exist_ok=True)
    os.makedirs(OUT, exist_ok=True)
    tex = os.path.join(BUILD, name + ".tex")
    setlvl = "" if level is None else "\\buildlevel=%d\n" % level
    with open(tex, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(pre + "\n" + setlvl + body + "\n\\end{document}\n")
    r = subprocess.run(
        ["pdflatex", "-interaction=nonstopmode", "-halt-on-error",
         "-output-directory", BUILD, tex],
        capture_output=True, text=True, cwd=BUILD,
        encoding="utf-8", errors="replace")
    pdf = os.path.join(BUILD, name + ".pdf")
    if r.returncode != 0 or not os.path.exists(pdf):
        why = ""
        log = os.path.join(BUILD, name + ".log")
        if os.path.exists(log):
            txt = open(log, encoding="utf-8", errors="replace").read()
            m = re.search(r"^! (.+)$", txt, re.M)
            why = m.group(1) if m else "compile failed"
        return False, why, (0, 0)
    subprocess.run(["pdftocairo", "-svg", pdf,
                    os.path.join(OUT, name + ".svg")],
                   capture_output=True)
    subprocess.run(["pdftocairo", "-png", "-r", str(DPI), "-transp",
                    "-singlefile", pdf, os.path.join(OUT, name)],
                   capture_output=True)
    size = page_size(pdf)
    ok = os.path.exists(os.path.join(OUT, name + ".svg")) and \
        os.path.exists(os.path.join(OUT, name + ".png"))
    return ok, "" if ok else "conversion failed", size


def page_size(pdf: str) -> tuple[float, float]:
    try:
        import fitz
        with fitz.open(pdf) as d:
            r = d[0].rect
            return (round(r.width, 2), round(r.height, 2))
    except Exception:
        out = subprocess.run(["pdfinfo", pdf], capture_output=True,
                             text=True).stdout
        m = re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", out)
        return (float(m.group(1)), float(m.group(2))) if m else (0, 0)


def steps_for(body: str) -> int:
    levels = [int(x) for x in
              re.findall(r"\\step(?:vis|only|dim)\{(\d+)\}", body)]
    return max(levels) if levels else 1


def annotated(label: str) -> tuple[str | None, str | None]:
    """(body, recorded source hash) of figures/src/<label>.tex, if present."""
    p = os.path.join(SRC, label + ".tex")
    if not os.path.exists(p):
        return None, None
    body = open(p, encoding="utf-8").read()
    m = re.search(r"source-hash:\s*([0-9a-f]+)", body)
    return body, (m.group(1) if m else None)


def figure_source(fig) -> str:
    """The body as it should compile: cross references resolved."""
    return T.resolve_refs(fig.body)


def jobs(labels: set[str] | None = None) -> list[dict]:
    """Everything that has to exist in out/, one dict per compile."""
    out = []
    for stem, fig in T.all_figures():
        if labels and fig.label not in labels:
            continue
        src = figure_source(fig)
        ann, recorded = annotated(fig.label)
        base = {"label": fig.label, "stem": stem,
                "src_hash": _sha(fig.body)}
        if ann is not None:
            base["stale"] = recorded is not None and \
                recorded != base["src_hash"]
            n = steps_for(ann)
            for lv in range(1, n + 1):
                out.append(dict(base, name=f"{fig.label}-{lv}", body=ann,
                                level=lv, steps=n))
        out.append(dict(base, name=fig.label, body=src, level=None, steps=1))
    return out


def crop_check(job: dict, pre: str) -> str:
    body = job["body"]
    if "useasboundingbox" not in body:
        return ""
    loose = body.replace("\\useasboundingbox", "\\path")
    ok1, _, a = compile_one("chk-" + job["name"] + "-pin", body, pre,
                            job["level"])
    ok2, _, b = compile_one("chk-" + job["name"] + "-loose", loose, pre,
                            job["level"])
    for f in os.listdir(OUT):
        if f.startswith("chk-"):
            os.remove(os.path.join(OUT, f))
    if not (ok1 and ok2):
        return ""
    dw, dh = b[0] - a[0], b[1] - a[1]
    if dw > 2.0 or dh > 2.0:
        return f"content extends {dw:.0f}x{dh:.0f}pt beyond the pinned box"
    return ""


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--label", nargs="*")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--clean", action="store_true")
    ap.add_argument("--check-crop", action="store_true")
    ap.add_argument("-j", type=int, default=6)
    a = ap.parse_args()

    if a.clean and os.path.isdir(BUILD):
        shutil.rmtree(BUILD)
    os.makedirs(SRC, exist_ok=True)
    pre = preamble()
    pre_hash = _sha(pre)
    man = load_manifest()
    todo = jobs(set(a.label) if a.label else None)

    if a.list:
        for j in todo:
            print(f"  {j['stem']:24} {j['name']}")
        print(f"\n  {len(todo)} outputs")
        return 0

    def need(j):
        h = _sha(pre_hash, j["body"], str(j["level"]))
        e = man.get(j["name"])
        have = os.path.exists(os.path.join(OUT, j["name"] + ".svg")) and \
            os.path.exists(os.path.join(OUT, j["name"] + ".png"))
        return a.force or not have or not e or e.get("hash") != h, h

    work = []
    for j in todo:
        n, h = need(j)
        if n:
            work.append((j, h))

    print(f"  {len(todo)} figure outputs, {len(work)} to compile")
    failed, stale, cropped = [], set(), []

    def run(item):
        j, h = item
        ok, why, size = compile_one(j["name"], j["body"], pre, j["level"])
        return j, h, ok, why, size

    with ThreadPoolExecutor(max_workers=a.j) as ex:
        for j, h, ok, why, size in ex.map(run, work):
            if ok:
                man[j["name"]] = {"hash": h, "w": size[0], "h": size[1],
                                  "label": j["label"], "stem": j["stem"],
                                  "level": j["level"], "steps": j["steps"]}
                print(f"    ok   {j['name']:32} {size[0]:>6.0f} x "
                      f"{size[1]:<5.0f} pt")
            else:
                failed.append(f"{j['name']}: {why}")
                print(f"    FAIL {j['name']:32} {why}")
    for j in todo:
        if j.get("stale"):
            stale.add(j["label"])
    # sizes for outputs that were already current
    save_manifest(man)

    if a.check_crop:
        for j in todo:
            if j["level"] is not None:
                why = crop_check(j, pre)
                if why:
                    cropped.append(f"{j['name']}: {why}")

    print()
    if stale:
        print("  STALE annotated copies (handout figure changed since the "
              "stepped version was made):")
        for s in sorted(stale):
            print("    figures/src/" + s + ".tex")
    if cropped:
        print("  CROPPED:")
        for c in cropped:
            print("    " + c)
    if failed:
        print("  FAILED:")
        for f in failed:
            print("    " + f)
        return 1
    print("  all figures current")
    return 1 if (cropped or stale) else 0


if __name__ == "__main__":
    sys.exit(main())
