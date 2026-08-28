#!/usr/bin/env python3
"""Turn a chapter of the CCNA volume into structured slide content.

The book is the single source of truth: every slide is derived from
``parts/NN_*.tex`` so the lecture, the notes and the labs cannot drift apart.
This module does the reading. ``build_slides.py`` does the drawing.

Two things it produces:

* **Blocks** -- one per pedagogy environment (outcomes, definitionbox,
  config, breakfix, ...), in document order, each with a title and body.
* **Runs** -- inline markup resolved into styled fragments, because
  ``\\textbf{Never}`` has to arrive at PowerPoint as a bold run, not as the
  five characters ``\\textbf``.

Run it directly to inspect what a chapter parses to::

    python texparse.py 8          # chapter 8, summary
    python texparse.py 8 --full   # every block, every run
"""
from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass, field

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                 # ...\CCNA
PARTS = os.path.join(ROOT, "parts")
MASTER = os.path.join(ROOT, "CCNA_v2_Handout.tex")


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Run:
    """A fragment of text with one style. The unit PowerPoint actually wants."""
    text: str
    bold: bool = False
    italic: bool = False
    mono: bool = False
    term: bool = False          # \term{} -- a defined term, accent coloured
    chip: str | None = None     # a blueprint topic tag such as "2.1"

    def __repr__(self) -> str:          # keeps --full output readable
        f = "".join(c for c, on in
                    (("b", self.bold), ("i", self.italic),
                     ("m", self.mono), ("t", self.term)) if on)
        return f"{self.text!r}{'/' + f if f else ''}"


@dataclass
class Block:
    """One pedagogy environment, or a run of prose between them."""
    kind: str                                   # 'definitionbox', 'config', ...
    title: str = ""                             # optional argument, if any
    args: list[str] = field(default_factory=list)   # mandatory args
    paras: list[list[Run]] = field(default_factory=list)   # prose before the list
    items: list[list[Run]] = field(default_factory=list)   # list items
    tail: list[list[Run]] = field(default_factory=list)
    # Prose *after* the list. Kept separate because merging it into `paras`
    # printed a box's closing sentence as if it introduced the box -- "Ties
    # are broken by..." belongs under the port roles, not above them.
    numbered: bool = False                      # items came from tightnum
    verbatim: str = ""                          # code, kept exactly as written
    section: str = ""                           # section heading in force
    children: list["Block"] = field(default_factory=list)
    # Nested blocks, in order. A breakfix is a narrative that embeds the
    # `verify` output it is diagnosing and the `config` that fixes it; the
    # guided-diagnosis slides need those staged separately from the prose.


@dataclass
class Chapter:
    number: int
    label: str
    title: str
    topics: list[str]
    blocks: list[Block]
    stem: str


# ---------------------------------------------------------------------------
# Chapter ordering -- needed so \chref{ospfv2} can become "Chapter 17"
# ---------------------------------------------------------------------------

def chapter_index() -> tuple[dict[str, int], list[tuple[int, str]]]:
    """Return ({label: number}, [(number, stem)]) in master-file order."""
    text = open(MASTER, encoding="utf-8").read()
    labels: dict[str, int] = {}
    order: list[tuple[int, str]] = []
    n = 0
    for stem in re.findall(r"\\input\{parts/([^}]+)\}", text):
        path = os.path.join(PARTS, stem + ".tex")
        if not os.path.exists(path):
            continue
        body = open(path, encoding="utf-8").read()
        m = re.search(r"\\chapter\{.*?\}\\label\{ch:([^}]+)\}", body)
        if not m:
            continue
        n += 1
        labels[m.group(1)] = n
        order.append((n, stem))
    return labels, order


_LABELS: dict[str, int] = {}


# ---------------------------------------------------------------------------
# Inline markup -> runs
# ---------------------------------------------------------------------------

# Order matters: multi-character sequences before their prefixes.
_LITERALS = [
    (r"\%", "%"), (r"\_", "_"), (r"\&", "&"), (r"\#", "#"),
    (r"\$", "$"), (r"\{", "{"), (r"\}", "}"),
    (r"$\rightarrow$", "\u2192"), (r"$\Rightarrow$", "\u21d2"),
    (r"$\Longrightarrow$", "\u27f9"), (r"$\leftarrow$", "\u2190"),
    (r"$\bullet$", "\u2022"), (r"\ldots", "\u2026"), (r"\dots", "\u2026"),
    (r"---", "\u2014"), (r"--", "\u2013"),
    (r"``", "\u201c"), (r"''", "\u201d"), (r"`", "\u2018"),
    # A lone apostrophe is a right single quote in LaTeX output ("the
    # interface's subnet"); a straight one on a slide reads as a typo
    # beside the em dashes and curly doubles. Mono runs get it back.
    (r"'", "\u2019"),
    (r"\,", "\u2009"), (r"\;", " "), (r"\ ", " "), (r"~", "\u00a0"),
    (r"\%", "%"),
]

# Commands whose braces are dropped but whose content is kept unstyled.
_TRANSPARENT = ("textcolor", "textnormal", "mdseries", "normalfont",
                "footnotesize", "scriptsize", "tiny", "small", "large")

_DROP = ("faIcon", "label", "index", "vspace", "hspace", "medskip",
         "smallskip", "bigskip", "needspace", "Needspace", "centering",
         "RaggedRight", "rowcolor", "toprule", "midrule", "bottomrule",
         "endhead", "arraybackslash", "allowbreak", "noindent",
         # \begin{tight} must lose its argument too, or the environment
         # name arrives in the slide as the literal word "tight".
         "begin", "end")


def _find_group(s: str, i: int) -> tuple[str, int]:
    """Read a balanced {...} starting at s[i] == '{'. Returns (inner, next_i)."""
    assert s[i] == "{"
    depth, j = 0, i
    while j < len(s):
        if s[j] == "{" and (j == 0 or s[j - 1] != "\\"):
            depth += 1
        elif s[j] == "}" and s[j - 1] != "\\":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return s[i + 1:], len(s)          # unbalanced; salvage what we can


_STRAIGHTEN = str.maketrans({"\u2019": "'", "\u2018": "'",
                             "\u201c": '"', "\u201d": '"'})


def to_runs(tex: str, bold=False, italic=False, mono=False,
            term=False) -> list[Run]:
    """Convert a fragment of LaTeX into styled runs."""
    runs: list[Run] = []
    buf: list[str] = []

    def flush():
        if buf:
            txt = "".join(buf)
            if mono:
                # Code is quoted literally: a curly quote in an IOS
                # command would be wrong, and would not paste.
                txt = txt.translate(_STRAIGHTEN)
            if txt:
                runs.append(Run(txt, bold, italic, mono, term))
            buf.clear()

    i = 0
    while i < len(tex):
        # literal sequences
        for src, dst in _LITERALS:
            if tex.startswith(src, i):
                buf.append(dst)
                i += len(src)
                break
        else:
            if tex[i] == "\\":
                m = re.match(r"\\([a-zA-Z]+)\*?", tex[i:])
                if not m:
                    i += 1
                    continue
                name = m.group(1)
                j = i + m.end()
                # skip an optional [...] argument
                if j < len(tex) and tex[j] == "[":
                    k = tex.find("]", j)
                    if k != -1:
                        j = k + 1
                if name in _DROP:
                    if j < len(tex) and tex[j] == "{":
                        _, j = _find_group(tex, j)
                    i = j
                    continue
                if j < len(tex) and tex[j] == "{":
                    inner, j2 = _find_group(tex, j)
                    flush()
                    if name in ("textbf", "bfseries"):
                        runs += to_runs(inner, True, italic, mono, term)
                    elif name in ("emph", "textit", "itshape"):
                        runs += to_runs(inner, bold, True, mono, term)
                    elif name in ("cmd", "texttt", "prompt", "ttfamily"):
                        runs += to_runs(inner, bold, italic, True, term)
                    elif name == "term":
                        runs += to_runs(inner, True, italic, mono, True)
                    elif name == "chref":
                        num = _LABELS.get(inner)
                        runs.append(Run(f"Chapter {num}" if num
                                        else "another chapter", bold, italic))
                    elif name == "tp":
                        runs.append(Run(inner, chip=inner))
                    elif name == "bp":
                        for t in inner.split(","):
                            if t.strip():
                                runs.append(Run(t.strip(), chip=t.strip()))
                    elif name == "ccnaref":
                        runs.append(Run("the figure", bold, italic))
                    elif name in _TRANSPARENT:
                        # \textcolor{c}{t}: first group is the colour, drop it
                        if name == "textcolor" and j2 < len(tex) and tex[j2] == "{":
                            inner2, j2 = _find_group(tex, j2)
                            runs += to_runs(inner2, bold, italic, mono, term)
                        else:
                            runs += to_runs(inner, bold, italic, mono, term)
                    else:
                        runs += to_runs(inner, bold, italic, mono, term)
                    i = j2
                    continue
                i = j
                continue
            if tex[i] == "$":
                # inline maths: keep the text, strip the delimiters
                k = tex.find("$", i + 1)
                if k == -1:
                    i += 1
                    continue
                inner = tex[i + 1:k]
                inner = (inner.replace(r"\times", "\u00d7")
                              .replace(r"\rightarrow", "\u2192")
                              .replace(r"\ge", "\u2265").replace(r"\le", "\u2264"))
                inner = re.sub(r"\^\{?(\d+)\}?", lambda m: _sup(m.group(1)), inner)
                inner = re.sub(r"[\\{}]", "", inner)
                buf.append(inner.strip())
                i = k + 1
                continue
            buf.append(tex[i])
            i += 1
    flush()
    return _tidy(runs)


_SUPS = str.maketrans("0123456789", "\u2070\u00b9\u00b2\u00b3\u2074"
                                    "\u2075\u2076\u2077\u2078\u2079")


def _sup(digits: str) -> str:
    return digits.translate(_SUPS)


def _tidy(runs: list[Run]) -> list[Run]:
    """Collapse whitespace and merge adjacent runs sharing a style."""
    out: list[Run] = []
    for r in runs:
        r.text = re.sub(r"[ \t\n]+", " ", r.text)
        if not r.text:
            continue
        if (out and r.chip is None and out[-1].chip is None
                and (out[-1].bold, out[-1].italic, out[-1].mono, out[-1].term)
                == (r.bold, r.italic, r.mono, r.term)):
            out[-1].text += r.text
        else:
            out.append(r)
    if out:
        out[0].text = out[0].text.lstrip()
        out[-1].text = out[-1].text.rstrip()
    return [r for r in out if r.text or r.chip]


def plain(runs: list[Run]) -> str:
    """Runs back to bare text -- for speaker notes and QA measurement."""
    return "".join(r.text for r in runs).strip()


# ---------------------------------------------------------------------------
# Block extraction
# ---------------------------------------------------------------------------

# Environments carrying an optional [title].
_OPT_TITLE = {"examnote", "definitionbox", "conceptbox", "alertbox",
              "worked", "pitfall", "breakfix"}
# Environments carrying mandatory {args}.
_ARGS = {"config": 1, "verify": 2, "hostcmd": 2, "playbook": 1, "lab": 2}
# Environments whose body is code and must not be touched.
_VERBATIM = {"config", "verify", "hostcmd", "playbook"}
# Plain environments.
_PLAIN = {"outcomes", "prereq", "keyterms", "checkpoint", "chaptersummary",
          "reviewq"}

_ALL_ENVS = _OPT_TITLE | set(_ARGS) | _PLAIN

_ENV_RE = re.compile(r"\\begin\{(" + "|".join(sorted(_ALL_ENVS)) + r")\}")


_NESTED_RE = re.compile(
    r"\\begin\{(config|verify|hostcmd|playbook)\}(.*?)\\end\{\1\}", re.S)


def _extract_children(body: str) -> tuple[list[Block], str]:
    """Lift nested code environments out of a narrative body.

    Returns (children, body_without_them). A breakfix embeds the `verify`
    it is diagnosing; those must be staged as their own slides rather than
    silently dropped by the prose cleaner.
    """
    children: list[Block] = []

    def take(m: re.Match) -> str:
        kind, rest = m.group(1), m.group(2)
        i, args = 0, []
        for _ in range(_ARGS.get(kind, 0)):
            while i < len(rest) and rest[i] in " \n":
                i += 1
            if i < len(rest) and rest[i] == "{":
                a, i = _find_group(rest, i)
                args.append(plain(to_runs(a)))
        children.append(Block(kind, args=args, verbatim=rest[i:].strip("\n")))
        return "\n\n"

    return children, _NESTED_RE.sub(take, body)


def _split_items(body: str):
    """Return (items, numbered, lead_paras, tail_paras)."""
    """Pull a tight/tightnum/tasks list out of a body.

    Returns (items, numbered, leading_paragraphs).
    """
    pat = re.compile(r"\\begin\{(tight|tightnum|tasks)\}(.*?)\\end\{\1\}", re.S)
    matches = list(pat.finditer(body))
    if not matches:
        return [], False, _paras(body), []

    # A box can hold more than one list -- chapter 4's default gateway has
    # two. Taking only the first left the second flattened into prose.
    items: list[list[Run]] = []
    numbered = False
    lead_src: list[str] = []
    pos = 0
    for m in matches:
        lead_src.append(body[pos:m.start()])
        numbered = numbered or m.group(1) == "tightnum"
        for it in re.split(r"\\item\s", m.group(2))[1:]:
            it = re.sub(r"^\[[^\]]*\]\s*", "", it.strip())   # \item[(a)] label
            runs = to_runs(it)
            if runs:
                items.append(runs)
        pos = m.end()

    lead: list[list[Run]] = []
    for chunk in lead_src:
        lead += _paras(chunk)
    return items, numbered, lead, _paras(body[pos:])


def _paras(body: str) -> list[list[Run]]:
    """Split prose into paragraphs, dropping tables and pictures."""
    body = re.sub(r"\\begin\{longtable\}.*?\\end\{longtable\}", " ", body, flags=re.S)
    body = re.sub(r"\\begin\{tikzpicture\}.*?\\end\{tikzpicture\}", " ", body, flags=re.S)
    body = re.sub(r"\\begin\{(center|quote|tabular|topology)\}", " ", body)
    body = re.sub(r"\\end\{(center|quote|tabular|topology)\}", " ", body)
    body = re.sub(r"\\ccnafigh?\{.*", " ", body, flags=re.S)
    out = []
    for chunk in re.split(r"\n\s*\n", body):
        if not chunk.strip():
            continue
        runs = to_runs(chunk)
        if runs and plain(runs):
            out.append(runs)
    return out


_FIG_RE = re.compile(r"\\ccnafigh?\{")
_SEC_RE = re.compile(r"\\(?:sub)?section\{(.+?)\}")


def _gap_blocks(gap: str, section: str) -> tuple[list[Block], str]:
    """Prose between environments, with any figure surfaced as its own block.

    Figures matter to the slides -- each becomes a slide, in the place the
    book puts it -- so they cannot just be stripped out with the tables.

    Section headings are tracked *positionally*. Reading them all up front
    and applying the last one gave every figure the heading of the section
    that follows it, so Figure 12.1 -- which illustrates why a loop is fatal
    -- arrived on a slide titled "How the tree is chosen".

    Returns (blocks, section in force at the end of the gap).
    """
    out: list[Block] = []
    pos = 0

    def flush(upto: int) -> None:
        """Prose from pos to upto, updating the heading as we pass one."""
        nonlocal section
        chunk = gap[pos:upto]
        last = 0
        for sm in _SEC_RE.finditer(chunk):
            for p in _paras(chunk[last:sm.start()]):
                out.append(Block("prose", paras=[p], section=section))
            section = plain(to_runs(sm.group(1)))
            last = sm.end()
        for p in _paras(chunk[last:]):
            out.append(Block("prose", paras=[p], section=section))

    while True:
        m = _FIG_RE.search(gap, pos)
        if not m:
            break
        flush(m.start())
        i = m.end() - 1
        _, j = _find_group(gap, i)               # the picture itself
        while j < len(gap) and gap[j] in " \n%":
            j += 1
        caption, j = _find_group(gap, j) if j < len(gap) and gap[j] == "{" \
            else ("", j)
        while j < len(gap) and gap[j] in " \n%":
            j += 1
        label, j = _find_group(gap, j) if j < len(gap) and gap[j] == "{" \
            else ("", j)
        out.append(Block("figure", title=plain(to_runs(caption)),
                         args=[label.strip()], section=section))
        pos = j
    flush(len(gap))
    return out, section


def parse_chapter(number: int) -> Chapter:
    """Parse chapter *number* (1-34) into a Chapter."""
    global _LABELS
    if not _LABELS:
        _LABELS, _ = chapter_index()
    labels, order = chapter_index()
    stem = dict(order)[number]
    path = os.path.join(PARTS, stem + ".tex")
    src = open(path, encoding="utf-8").read()

    m = re.search(r"\\chapter\{(.+?)\}\\label\{ch:([^}]+)\}", src)
    title, label = (m.group(1), m.group(2)) if m else (stem, stem)
    title = plain(to_runs(title))
    bp = re.search(r"\\bp\{([^}]*)\}", src)
    topics = [t.strip() for t in bp.group(1).split(",")] if bp else []

    blocks: list[Block] = []
    section = ""
    pos = 0
    while True:
        m = _ENV_RE.search(src, pos)
        if not m:
            break
        kind = m.group(1)

        # prose since the previous environment -> notes material
        gap_blocks, section = _gap_blocks(src[pos:m.start()], section)
        blocks.extend(gap_blocks)

        # the environment's own arguments
        i = m.end()
        opt_title, args = "", []
        if kind in _OPT_TITLE and i < len(src) and src[i] == "[":
            depth, j = 0, i
            while j < len(src):
                if src[j] == "[":
                    depth += 1
                elif src[j] == "]":
                    depth -= 1
                    if depth == 0:
                        break
                j += 1
            opt_title = plain(to_runs(src[i + 1:j]))
            i = j + 1
        for _ in range(_ARGS.get(kind, 0)):
            while i < len(src) and src[i] in " \n":
                i += 1
            if i < len(src) and src[i] == "{":
                a, i = _find_group(src, i)
                args.append(plain(to_runs(a)))

        end = src.find("\\end{" + kind + "}", i)
        if end == -1:
            pos = i
            continue
        body = src[i:end]

        blk = Block(kind, opt_title, args, section=section)
        if kind in _VERBATIM:
            blk.verbatim = body.strip("\n")
        else:
            blk.children, body = _extract_children(body)
            blk.items, blk.numbered, blk.paras, blk.tail = _split_items(body)
        blocks.append(blk)
        pos = end + len("\\end{" + kind + "}")

    # trailing prose
    tail_blocks, section = _gap_blocks(src[pos:], section)
    blocks.extend(tail_blocks)

    return Chapter(number, label, title, topics, blocks, stem)


# ---------------------------------------------------------------------------
# Checkpoint answers, from Appendix E
# ---------------------------------------------------------------------------

def checkpoint_answers(number: int) -> list[list[Run]]:
    """The answers to chapter *number*'s checkpoint, from parts/E_answers.tex."""
    src = open(os.path.join(PARTS, "E_answers.tex"), encoding="utf-8").read()
    m = re.search(r"\\subsection\*\{Chapter " + str(number) +
                  r" --- .+?\}\s*\\begin\{tightnum\}(.*?)\\end\{tightnum\}",
                  src, re.S)
    if not m:
        return []
    return [to_runs(re.sub(r"^\[[^\]]*\]\s*", "", it.strip()))
            for it in re.split(r"\\item\s", m.group(1))[1:]]


# ---------------------------------------------------------------------------
# Inspection
# ---------------------------------------------------------------------------

def audit() -> int:
    """Parse every chapter and check nothing was silently dropped.

    Each environment the book contains must surface as a block or a child.
    A mismatch means the parser lost content, which would show up as a
    missing slide rather than as an error.
    """
    labels, order = chapter_index()
    bad = 0
    print(f"{'ch':>3}  {'title':38} blocks  envs  answers  status")
    print("-" * 78)
    for n, stem in order:
        if n > 34:
            break
        src = open(os.path.join(PARTS, stem + ".tex"), encoding="utf-8").read()
        raw = {e: len(re.findall(r"\\begin\{" + e + r"\}", src))
               for e in _ALL_ENVS}
        ch = parse_chapter(n)
        got: dict[str, int] = {}
        for b in ch.blocks:
            if b.kind in _ALL_ENVS:
                got[b.kind] = got.get(b.kind, 0) + 1
            for c in b.children:
                got[c.kind] = got.get(c.kind, 0) + 1
        miss = {e: (raw[e], got.get(e, 0))
                for e in _ALL_ENVS if raw[e] != got.get(e, 0)}
        cp = [b for b in ch.blocks if b.kind == "checkpoint"]
        nq = len(cp[0].items) if cp else 0
        na = len(checkpoint_answers(n))
        ok = not miss and nq == na and nq > 0
        if not ok:
            bad += 1
        note = "ok" if ok else ("; ".join(f"{k} {r}!={g}"
                                         for k, (r, g) in miss.items())
                                or f"Q{nq}!=A{na}")
        print(f"{n:>3}  {ch.title[:38]:38} {len(ch.blocks):>6} "
              f"{sum(raw.values()):>5} {na:>8}  {note}")
    print()
    print("all chapters parse completely" if not bad
          else f"{bad} chapter(s) with dropped content")
    return 1 if bad else 0


def main() -> int:
    if "--all" in sys.argv:
        return audit()
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    full = "--full" in sys.argv
    ch = parse_chapter(n)
    print(f"Chapter {ch.number}: {ch.title}")
    print(f"  label ch:{ch.label}   topics {', '.join(ch.topics) or '(none)'}")
    print(f"  {len(ch.blocks)} blocks\n")

    counts: dict[str, int] = {}
    for b in ch.blocks:
        counts[b.kind] = counts.get(b.kind, 0) + 1
        for c in b.children:
            counts[f"  +{c.kind}"] = counts.get(f"  +{c.kind}", 0) + 1
    for k in sorted(counts, key=lambda k: -counts[k]):
        print(f"  {k:16} {counts[k]}")

    ans = checkpoint_answers(n)
    cp = [b for b in ch.blocks if b.kind == "checkpoint"]
    nq = len(cp[0].items) if cp else 0
    print(f"\n  checkpoint: {nq} questions, {len(ans)} answers "
          f"{'OK' if nq == len(ans) else '<-- MISMATCH'}")

    if full:
        print("\n" + "=" * 72)
        for b in ch.blocks:
            head = f"[{b.kind}]" + (f" {b.title}" if b.title else "")
            if b.args:
                head += " " + " / ".join(b.args)
            print("\n" + head)
            if b.verbatim:
                for line in b.verbatim.split("\n")[:6]:
                    print("    | " + line)
                extra = len(b.verbatim.split("\n")) - 6
                if extra > 0:
                    print(f"    | ... {extra} more lines")
            for p in b.paras:
                print("    " + plain(p)[:150])
            for it in b.items:
                print("    - " + plain(it)[:140])
    return 0


if __name__ == "__main__":
    sys.exit(main())
