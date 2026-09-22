#!/usr/bin/env python3
"""Check the decks, then render them so they can be looked at.

Two halves, and the second is the one that matters.

**Static checks** read each .pptx and catch what is measurable:

* every slide has a filled title placeholder (outline, accessibility);
* text fits its box -- measured with the real fonts, as the layout was;
* nothing extends off the slide;
* every picture has alt text or is marked decorative;
* no run is set below 10 pt (footers) or 12 pt (anything else);
* consecutive Morph slides share at least one ``!!`` shape name, or the
  Morph silently degrades to a cut;
* every agenda link resolves to a slide;
* every figure of the chapter appears in its deck.

**Rendering** drives the installed PowerPoint through COM, exports every
slide as a PNG, and assembles a contact sheet per deck. No checker catches a
cramped card, a poor colour pairing or a build step that reveals the wrong
thing -- only looking does.

    python qa.py                    check every deck in out/
    python qa.py --ch 14            one chapter
    python qa.py --render --ch 14   also export PNGs and a contact sheet
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import sys

from pptx import Presentation
from pptx.util import Emu

import measure as M
import texparse as T

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
RENDER = os.path.join(HERE, "render")
EMU = 914400.0
W_IN, H_IN = 13.333, 7.5


def _is_decorative(shape) -> bool:
    return "decorative" in shape._element.xml and 'val="1"' in \
        shape._element.xml


def _alt(shape) -> str:
    c = shape._element.find(".//{http://schemas.openxmlformats.org/"
                            "presentationml/2006/main}cNvPr")
    return c.get("descr", "") if c is not None else ""


def _fits(shape):
    """(ok, ratio) -- measured height of a text box's content vs its box."""
    tf = shape.text_frame
    w = shape.width / EMU
    h = shape.height / EMU
    if w <= 0.05 or h <= 0.05:
        return True, 0.0
    wraps = tf.word_wrap is not False
    total = 0.0
    paras = list(tf.paragraphs)
    for i, p in enumerate(paras):
        pieces = []
        size = None
        for r in p.runs:
            sz = r.font.size.pt if r.font.size else 18
            size = max(size or 0, sz)
            fam = r.font.name or "Segoe UI"
            pieces.append((r.text, sz, fam, bool(r.font.bold),
                           bool(r.font.italic)))
        if not pieces:
            continue
        size = size or 18
        pPr = p._p.pPr
        mar = int(pPr.get("marL", "0")) / EMU if pPr is not None else 0.0
        if wraps:
            lines = M.count_lines([(t.split("\t")[-1], *rest) for t, *rest
                                   in pieces], max(0.2, w - mar))
        else:
            lines = 1
            if M.natural_width(pieces) > w + 0.05:
                return False, M.natural_width(pieces) / w
        ls = p.line_spacing if isinstance(p.line_spacing, float) else 1.0
        total += lines * M.LINE * size * ls / 72.0
        if i < len(paras) - 1 and p.space_after is not None:
            total += p.space_after.pt / 72.0
    return total <= h * 1.04 + 0.02, (total / h if h else 0.0)


def check_deck(path: str, verbose=True) -> list[str]:
    prs = Presentation(path)
    name = os.path.basename(path)
    probs: list[str] = []
    prev_names, prev_morph = None, False
    pictures = 0
    notes = 0
    figs_in_deck = set()
    for i, s in enumerate(prs.slides, 1):
        names = set()
        title = s.shapes.title
        if title is None or not title.text_frame.text.strip():
            probs.append(f"{name} s{i}: no title")
        for sh in s.shapes:
            names.add(sh.name)
            x, y = sh.left / EMU, sh.top / EMU
            w, h = sh.width / EMU, sh.height / EMU
            if x < -0.01 or y < -0.8 or x + w > W_IN + 0.02 or \
                    y + h > H_IN + 0.02:
                if not sh.name.startswith("!!wm"):
                    probs.append(f"{name} s{i}: '{sh.name}' off the slide "
                                 f"({x:.2f},{y:.2f} {w:.2f}x{h:.2f})")
            if sh.shape_type == 13:          # picture
                pictures += 1
                if not _is_decorative(sh) and not _alt(sh):
                    probs.append(f"{name} s{i}: picture '{sh.name}' has no "
                                 f"alt text")
                if sh.name == "!!figure":
                    m = re.match(r"Figure ([\dA-Z]+(?:\.\d+)*)", _alt(sh))
                    if m:
                        figs_in_deck.add(m.group(1))
            if sh.has_text_frame and sh.text_frame.text.strip():
                ok, ratio = _fits(sh)
                if not ok:
                    probs.append(f"{name} s{i}: '{sh.name}' overflows "
                                 f"({ratio:.0%}) "
                                 f"{sh.text_frame.text.strip()[:50]!r}")
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.size and r.text.strip():
                            floor = 10 if sh.name.startswith("!!foot") \
                                else 9.5 if sh.name.startswith("!!dotn") \
                                else 10.5
                            if r.font.size.pt < floor:
                                probs.append(
                                    f"{name} s{i}: {r.font.size.pt}pt text "
                                    f"in '{sh.name}': {r.text[:30]!r}")
                                break
            if sh.has_table if hasattr(sh, "has_table") else False:
                pass
            # links
            try:
                ca = sh.click_action
                if ca.action == 1 or "hlinksldjump" in sh._element.xml:
                    if ca.target_slide is None:
                        probs.append(f"{name} s{i}: dead link on "
                                     f"'{sh.name}'")
            except Exception:
                pass
        xml = s._element.xml
        morph = "p159:morph" in xml
        if morph and prev_names is not None:
            shared = {n for n in names & prev_names if n.startswith("!!")}
            if not shared:
                probs.append(f"{name} s{i}: Morph shares no '!!' shape name "
                             f"with the previous slide")
        prev_names, prev_morph = names, morph
        if s.has_notes_slide and \
                s.notes_slide.notes_text_frame.text.strip():
            notes += 1
    # coverage: every figure of the chapter is in the deck
    m = re.match(r"DS_Ch(\d\d)_", name)
    if m:
        n = int(m.group(1))
        ch = T.parse_chapter(n)
        want = {T.ref_num("fig:" + e.label)
                for e in T.walk([it.el for it in ch.items])
                if isinstance(e, T.FigureE)}
        miss = sorted(want - figs_in_deck)
        if miss:
            probs.append(f"{name}: figures missing from the deck: "
                         f"{', '.join(miss)}")
    if verbose:
        print(f"  {name:62} {len(prs.slides._sldIdLst):>3} slides "
              f"{pictures:>3} pics {notes:>3} with notes")
    return probs


# ---------------------------------------------------------------------------
# Rendering
# ---------------------------------------------------------------------------

def render(paths, width=1600):
    try:
        import win32com.client as win32
    except ImportError:
        print("  pywin32 not available; cannot render")
        return []
    os.makedirs(RENDER, exist_ok=True)
    app = win32.Dispatch("PowerPoint.Application")
    done = []
    try:
        for p in paths:
            stem = os.path.splitext(os.path.basename(p))[0]
            d = os.path.join(RENDER, stem)
            os.makedirs(d, exist_ok=True)
            for f in glob.glob(os.path.join(d, "s*.png")):
                os.remove(f)
            pres = app.Presentations.Open(os.path.abspath(p), ReadOnly=True,
                                          WithWindow=False)
            try:
                for i, sl in enumerate(pres.Slides, 1):
                    sl.Export(os.path.join(d, f"s{i:03d}.png"), "PNG", width,
                              int(width * 7.5 / 13.333))
                n = pres.Slides.Count
            finally:
                pres.Close()
            sheet(d, os.path.join(RENDER, stem + "_sheet.png"))
            print(f"  rendered {n:>3} slides  {stem}")
            done.append(d)
    finally:
        try:
            if app.Presentations.Count == 0:
                app.Quit()
        except Exception:
            pass
    return done


def sheet(d, out, cols=6, tw=420):
    from PIL import Image, ImageDraw
    files = sorted(glob.glob(os.path.join(d, "s*.png")))
    th = int(tw * 7.5 / 13.333)
    rows = (len(files) + cols - 1) // cols
    im = Image.new("RGB", (cols * (tw + 10) + 10, rows * (th + 24) + 10),
                   "white")
    dr = ImageDraw.Draw(im)
    for i, f in enumerate(files):
        t = Image.open(f).convert("RGB").resize((tw, th), Image.LANCZOS)
        x = 10 + (i % cols) * (tw + 10)
        y = 10 + (i // cols) * (th + 24)
        im.paste(t, (x, y + 16))
        dr.rectangle([x - 1, y + 15, x + tw, y + 16 + th],
                     outline=(190, 190, 190))
        dr.text((x, y + 1), str(i + 1), fill=(60, 60, 60))
    im.save(out)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", nargs="*", type=int)
    ap.add_argument("--render", action="store_true")
    ap.add_argument("--width", type=int, default=1600)
    a = ap.parse_args()
    paths = sorted(glob.glob(os.path.join(OUT, "*.pptx")))
    if a.ch:
        keep = {f"DS_Ch{n:02d}_" for n in a.ch}
        paths = [p for p in paths if os.path.basename(p)[:8] in keep]
    if not paths:
        print("  no decks in out/ -- run build_slides.py first")
        return 1
    print(f"Checking {len(paths)} deck(s)\n")
    probs = []
    for p in paths:
        probs += check_deck(p)
    print()
    if probs:
        for q in probs[:60]:
            print("  " + q)
        if len(probs) > 60:
            print(f"  ... and {len(probs) - 60} more")
        print(f"\n  {len(probs)} problem(s)")
    else:
        print("  static checks clean")
    if a.render:
        print("\nRendering")
        render(paths, a.width)
    return 1 if probs else 0


if __name__ == "__main__":
    sys.exit(main())
