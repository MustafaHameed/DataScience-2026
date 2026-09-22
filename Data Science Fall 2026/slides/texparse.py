#!/usr/bin/env python3
"""Read the Data Science handout into structured slide content.

The handout is the single source of truth. Every slide is derived from
``parts/NN_*.tex`` so the lecture, the book and the labs cannot drift apart;
a correction made in the handout reaches the slides on the next build.

The CCNA parser this descends from gave each box three flat lists (lead
paragraphs, items, tail). That is not enough here: a worked example in this
handout interleaves prose, a list, display maths, a table and sometimes a code
listing, and the order *is* the argument. So a box body is parsed into a
sequence of typed elements instead:

    Para    a paragraph of styled runs (plus its source, for maths rendering)
    ListE   a tight/tightnum list; each item is a Para
    MathE   display maths, \\[ ... \\]
    TableE  a longtable or tabular: header, rows, column widths, alignment
    CodeE   an lstlisting, verbatim, with its language
    FigureE a \\dsfig: TikZ body, caption, label
    Box     a pedagogy environment holding its own element sequence

Inspect what a chapter parses to::

    python texparse.py 14           # summary
    python texparse.py 14 --full    # every element
    python texparse.py --all        # audit: nothing silently dropped
"""
from __future__ import annotations

import os
import re
import sys
from dataclasses import dataclass, field
from functools import lru_cache

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)                      # ...\Data Science Fall 2026
PARTS = os.path.join(ROOT, "parts")
MASTER = os.path.join(ROOT, "DataScience_Fall2026_Handout.tex")
AUX = os.path.join(ROOT, "DataScience_Fall2026_Handout.aux")
CODE = os.path.join(ROOT, "code")

PART_TITLES = {
    1: "Foundations of Data",
    2: "From Data to Insight",
    3: "Machine Learning Core",
    4: "Deep Learning and Modern AI",
    5: "Data Science in Systems Context",
    6: "Professional Practice",
}


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class Run:
    """A fragment of text with one style -- the unit PowerPoint wants."""
    text: str
    bold: bool = False
    italic: bool = False
    mono: bool = False
    term: bool = False          # \term{}: a defined term, primary coloured
    math: bool = False          # came from $...$; set in the maths face
    script: str | None = None   # 'sub' | 'sup' for maths indices
    color: str | None = None    # \textcolor{name}{...}

    def style(self):
        return (self.bold, self.italic, self.mono, self.term, self.math,
                self.script, self.color)

    def __repr__(self) -> str:
        f = "".join(c for c, on in (("b", self.bold), ("i", self.italic),
                                    ("m", self.mono), ("t", self.term),
                                    ("$", self.math)) if on)
        if self.script:
            f += "_" if self.script == "sub" else "^"
        return f"{self.text!r}{'/' + f if f else ''}"


@dataclass
class Para:
    runs: list[Run]
    tex: str = ""               # source, refs resolved -- for maths rendering
    heavy: bool = False         # carries maths that Unicode cannot set well
    kind = "para"

    @property
    def text(self) -> str:
        return plain(self.runs)


@dataclass
class ListE:
    items: list[Para]
    numbered: bool = False
    kind = "list"


@dataclass
class MathE:
    tex: str
    kind = "math"


@dataclass
class TableE:
    header: list[list[Run]]
    rows: list[list[list[Run]]]
    widths: list[float]         # relative column widths (from L{..}/p{..})
    align: list[str]            # 'l' | 'r' | 'c' per column
    band: str = "lightbg"       # \rowcolors{2}{white}{<band>}
    marks: list[str] = field(default_factory=list)
    kind = "table"


@dataclass
class CodeE:
    code: str
    lang: str = ""              # 'python' | 'r' | 'bash' | '' (plain text)
    kind = "code"


@dataclass
class FigureE:
    label: str
    caption: list[Run]
    caption_tex: str
    body: str                   # TikZ source, verbatim from the handout
    here: bool = False          # \dsfigh
    kind = "figure"


@dataclass
class Box:
    env: str                    # definitionbox, worked, lab, fourlenses, ...
    title: list[Run] = field(default_factory=list)
    title_tex: str = ""
    body: list = field(default_factory=list)       # elements, in order
    children: list["Box"] = field(default_factory=list)   # lenses
    arg: str = ""               # lens id (LA/PM/IOT/SEC); fourlenses topic
    kind = "box"

    @property
    def title_text(self) -> str:
        return plain(self.title)


@dataclass
class Item:
    """One top-level element of a chapter, with where it sits."""
    el: object                  # Para | ListE | MathE | TableE | CodeE | FigureE | Box
    section: str = ""           # nearest heading, any level
    top: str = ""               # nearest \section only
    secnum: str = ""            # e.g. "14.2"
    page: int = 0               # handout page of that section


@dataclass
class Section:
    num: str
    title: str
    page: int
    level: int                  # 1 section, 2 subsection


@dataclass
class Chapter:
    number: int
    label: str
    title: str
    stem: str
    part: int
    page: int
    items: list[Item]
    sections: list[Section]

    def boxes(self, env: str) -> list[Box]:
        return [i.el for i in self.items
                if isinstance(i.el, Box) and i.el.env == env]

    def first(self, env: str) -> Box | None:
        b = self.boxes(env)
        return b[0] if b else None

    @property
    def last_page(self) -> int:
        """The page before the next chapter (or appendix) starts."""
        pages = sorted(v[1] for k, v in aux().labels.items()
                       if (k.startswith("ch:") or k.startswith("app:"))
                       and v[1] > self.page)
        own = max([s.page for s in self.sections] + [self.page])
        return max(own, pages[0] - 1) if pages else own


# ---------------------------------------------------------------------------
# Brace and argument helpers
# ---------------------------------------------------------------------------

def find_group(s: str, i: int, open_="{", close="}") -> tuple[str, int]:
    """Read a balanced group starting at s[i] == open_. Returns (inner, next)."""
    assert s[i] == open_, (s[i:i + 30], open_)
    depth, j = 0, i
    while j < len(s):
        c = s[j]
        if c == "\\":
            j += 2
            continue
        if c == open_:
            depth += 1
        elif c == close:
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    return s[i + 1:], len(s)


def skip_ws(s: str, i: int) -> int:
    while i < len(s) and s[i] in " \t\n":
        i += 1
    return i


def read_args(s: str, i: int, n: int, opt: bool = False
              ) -> tuple[str | None, list[str], int]:
    """Read an optional [..] (if opt) then n mandatory {..} starting at i."""
    o = None
    j = skip_ws(s, i) if opt else i
    if opt and j < len(s) and s[j] == "[":
        o, j = find_group(s, j, "[", "]")
        i = j
    args = []
    for _ in range(n):
        j = skip_ws(s, i)
        if j < len(s) and s[j] == "{":
            a, i = find_group(s, j)
            args.append(a)
        else:
            args.append("")
    return o, args, i


def strip_comments(tex: str) -> str:
    """Remove LaTeX comments. Code listings must be protected first."""
    return re.sub(r"(?<!\\)%.*", "", tex)


# ---------------------------------------------------------------------------
# Cross references, from the .aux of the last handout build
# ---------------------------------------------------------------------------

@dataclass
class Aux:
    labels: dict                # key -> (num, page, title)
    toc: list                   # (level, num, title_tex, page)


def _aux_groups(s: str, i: int, n: int) -> list[str]:
    out = []
    for _ in range(n):
        i = skip_ws(s, i)
        if i >= len(s) or s[i] != "{":
            break
        g, i = find_group(s, i)
        out.append(g)
    return out


@lru_cache(maxsize=1)
def aux() -> Aux:
    labels, toc = {}, []
    if not os.path.exists(AUX):
        print("WARNING: no .aux -- build the handout first; numbers will be "
              "missing", file=sys.stderr)
        return Aux(labels, toc)
    src = open(AUX, encoding="utf-8", errors="replace").read()
    for m in re.finditer(r"\\newlabel\{([^}]+)\}", src):
        g = _aux_groups(src, m.end(), 1)
        if not g:
            continue
        inner = _aux_groups(g[0], 0, 3)
        if len(inner) >= 2:
            try:
                page = int(re.sub(r"\D", "", inner[1]) or 0)
            except ValueError:
                page = 0
            labels[m.group(1)] = (inner[0], page,
                                  inner[2] if len(inner) > 2 else "")
    lv = {"chapter": 0, "section": 1, "subsection": 2}
    for m in re.finditer(r"\\contentsline \{(chapter|section|subsection)\}",
                         src):
        g = _aux_groups(src, m.end(), 2)
        if len(g) < 2:
            continue
        nm = re.match(r"\\numberline \{([^}]*)\}(.*)", g[0], re.S)
        if not nm:
            continue
        try:
            page = int(g[1])
        except ValueError:
            page = 0
        toc.append((lv[m.group(1)], nm.group(1), nm.group(2).strip(), page))
    return Aux(labels, toc)


def ref_num(key: str) -> str:
    v = aux().labels.get(key)
    return v[0] if v else ""


def ref_page(key: str) -> int:
    v = aux().labels.get(key)
    return v[1] if v else 0


# ---------------------------------------------------------------------------
# Maths -> Unicode runs
# ---------------------------------------------------------------------------

_MATH_SYM = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "lambda": "λ",
    "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "rho": "ρ", "sigma": "σ",
    "tau": "τ", "phi": "φ", "varphi": "φ", "chi": "χ", "psi": "ψ",
    "omega": "ω", "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ",
    "Sigma": "Σ", "Phi": "Φ", "Omega": "Ω", "Pi": "Π",
    "times": "×", "cdot": "·", "pm": "±", "mp": "∓", "div": "÷",
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "neq": "≠", "ne": "≠",
    "approx": "≈", "sim": "∼", "simeq": "≃", "propto": "∝", "equiv": "≡",
    "gg": "≫", "ll": "≪", "in": "∈", "notin": "∉", "subset": "⊂",
    "supset": "⊃", "subseteq": "⊆", "cap": "∩", "cup": "∪",
    "rightarrow": "→", "to": "→", "Rightarrow": "⇒", "leftarrow": "←",
    "Leftarrow": "⇐", "longrightarrow": "⟶", "Longrightarrow": "⟹",
    "leftrightarrow": "↔", "uparrow": "↑", "downarrow": "↓",
    "infty": "∞", "nabla": "∇", "partial": "∂", "sum": "Σ", "prod": "∏",
    "int": "∫", "ldots": "…", "cdots": "⋯", "dots": "…", "circ": "∘",
    "blacksquare": "■", "square": "□", "otimes": "⊗", "oplus": "⊕",
    "forall": "∀", "exists": "∃", "neg": "¬", "land": "∧", "lor": "∨",
    "emptyset": "∅", "star": "⋆", "ast": "∗", "bullet": "•", "prime": "′",
    "lvert": "|", "rvert": "|", "vert": "|", "|": "‖", "angle": "∠",
    "degree": "°", "checkmark": "✓", "%": "%", "$": "$", "#": "#",
    "&": "&", "_": "_", "{": "{", "}": "}", "lbrace": "{", "rbrace": "}",
    "langle": "⟨", "rangle": "⟩", "lfloor": "⌊", "rfloor": "⌋",
    "lceil": "⌈", "rceil": "⌉",
}
_MATH_OPNAME = {"log", "ln", "exp", "max", "min", "sin", "cos", "tan", "arg",
                "argmax", "argmin", "det", "lim", "sup", "inf", "Pr", "var",
                "sgn", "tanh"}
_MATH_SPACE = {",": "\u2009", ";": " ", ":": " ", "!": "", " ": " ",
               "quad": "  ", "qquad": "    "}
_MATH_DROP = {"left", "right", "big", "Big", "bigg", "Bigg", "displaystyle",
              "textstyle", "limits", "nolimits", "mathstrut", "strut", "!",
              "phantom", "hphantom", "vphantom", "allowbreak", "nonumber"}
_ACCENT = {"hat": "\u0302", "widehat": "\u0302", "bar": "\u0304",
           "overline": "\u0304", "tilde": "\u0303", "widetilde": "\u0303",
           "vec": "\u20d7", "dot": "\u0307", "ddot": "\u0308"}
_UPRIGHT = {"text", "mathrm", "textrm", "operatorname", "textnormal",
            "mbox", "textup", "mathsf", "textsf"}


def math_runs(m: str, bold: bool = False, script: str | None = None,
              upright: bool = False) -> list[Run]:
    """Convert inline LaTeX maths to styled runs.

    Single Latin letters are italic (they are variables), digits and
    operators upright, \\text{} upright in the text face. Indices become runs
    with a baseline offset, which PowerPoint renders as real sub- and
    superscripts -- Unicode has subscript forms for only half the alphabet,
    and "P_attack" on a slide reads as a typo.
    """
    out: list[Run] = []

    def emit(t, it=False, b=bold, sc=script, is_math=True):
        if t:
            out.append(Run(t, bold=b, italic=it and not upright,
                           math=is_math, script=sc))

    i = 0
    n = len(m)
    while i < n:
        c = m[i]
        if c == "\\":
            mm = re.match(r"\\([A-Za-z]+|.)", m[i:])
            name = mm.group(1)
            j = i + mm.end()
            if name in _MATH_SPACE:
                emit(_MATH_SPACE[name])
                i = j
                continue
            if name in _MATH_DROP:
                i = j
                if name in ("phantom", "hphantom", "vphantom"):
                    k = skip_ws(m, j)
                    if k < n and m[k] == "{":
                        _, i = find_group(m, k)
                continue
            if name in _UPRIGHT:
                k = skip_ws(m, j)
                if k < n and m[k] == "{":
                    inner, j = find_group(m, k)
                    for r in text_runs(inner, bold=bold):
                        r.script = script
                        r.math = False
                        out.append(r)
                i = j
                continue
            if name in ("mathbf", "boldsymbol", "bm", "textbf"):
                k = skip_ws(m, j)
                if k < n and m[k] == "{":
                    inner, j = find_group(m, k)
                    out += math_runs(inner, True, script, upright)
                i = j
                continue
            if name in ("mathit", "textit", "emph"):
                k = skip_ws(m, j)
                if k < n and m[k] == "{":
                    inner, j = find_group(m, k)
                    out += math_runs(inner, bold, script, False)
                i = j
                continue
            if name in ("frac", "dfrac", "tfrac"):
                _, (a, b), j = read_args(m, j, 2)
                ra = math_runs(a, bold, script, upright)
                rb = math_runs(b, bold, script, upright)
                wrap_a = len(plain(ra)) > 1 and re.search(r"[+\-−·× ]",
                                                           plain(ra))
                wrap_b = len(plain(rb)) > 1 and re.search(r"[+\-−·× ]",
                                                           plain(rb))
                if wrap_a:
                    emit("(")
                out += ra
                if wrap_a:
                    emit(")")
                emit("/")
                if wrap_b:
                    emit("(")
                out += rb
                if wrap_b:
                    emit(")")
                i = j
                continue
            if name == "sqrt":
                _, (a,), j = read_args(m, j, 1)
                ra = math_runs(a, bold, script, upright)
                emit("√")
                if len(plain(ra)) > 1:
                    emit("(")
                    out += ra
                    emit(")")
                else:
                    out += ra
                i = j
                continue
            if name in _ACCENT:
                _, (a,), j = read_args(m, j, 1)
                ra = math_runs(a, bold, script, upright)
                if ra:
                    ra[-1].text += _ACCENT[name]
                out += ra
                i = j
                continue
            if name in ("underbrace", "overbrace"):
                _, (a,), j = read_args(m, j, 1)
                out += math_runs(a, bold, script, upright)
                # drop the brace label (^{..} or _{..})
                k = skip_ws(m, j)
                if k < n and m[k] in "^_":
                    k = skip_ws(m, k + 1)
                    if k < n and m[k] == "{":
                        _, k = find_group(m, k)
                    else:
                        k += 1
                    j = k
                i = j
                continue
            if name in _MATH_OPNAME:
                emit(name)
                i = j
                continue
            if name in _MATH_SYM:
                sym = _MATH_SYM[name]
                if name == "mid":
                    sym = "∣"
                emit(sym, it=False)
                i = j
                continue
            if name == "mid":
                emit(" ∣ ")
                i = j
                continue
            if name == "not":
                emit("\u0338")
                i = j
                continue
            # unknown command: keep an argument if it has one
            k = skip_ws(m, j)
            if k < n and m[k] == "{":
                inner, j = find_group(m, k)
                out += math_runs(inner, bold, script, upright)
            i = j
            continue
        if c in "^_":
            sc = "sup" if c == "^" else "sub"
            k = skip_ws(m, i + 1)
            if k < n and m[k] == "{":
                inner, j = find_group(m, k)
            elif k < n and m[k] == "\\":
                mm = re.match(r"\\([A-Za-z]+|.)", m[k:])
                inner, j = m[k:k + mm.end()], k + mm.end()
            else:
                inner, j = m[k:k + 1], k + 1
            if sc == "sup" and inner.strip() in ("\\circ", "\\circ "):
                emit("°")
            elif sc == "sup" and inner.strip() in ("\\prime", "'"):
                emit("′")
            else:
                out += math_runs(inner, bold, sc, upright)
            i = j
            continue
        if c == "{":
            inner, j = find_group(m, i)
            out += math_runs(inner, bold, script, upright)
            i = j
            continue
        if c == "}":
            i += 1
            continue
        if c == "~":
            emit("\u00a0")
            i += 1
            continue
        if c.isalpha():
            # a letter run: single letters are variables
            k = i
            while k < n and m[k].isalpha():
                k += 1
            word = m[i:k]
            for ch in word:
                emit(ch, it=True)
            i = k
            continue
        if c == "-":
            emit("−")
        elif c == "*":
            emit("∗")
        elif c == "'":
            emit("′")
        elif c in " \n\t":
            emit(" ")
        else:
            emit(c)
        i += 1
    return _tidy(out, strip=False)


_HEAVY = re.compile(r"\\(dfrac|underbrace|overbrace|int|iint|prod|binom|"
                    r"begin|lim|overset|underset|stackrel|cases|matrix)\b|"
                    r"\\sum_\{|\\sum_[a-z]|\\sqrt\{[^}]*\\(frac|sum)")


def math_is_heavy(m: str) -> bool:
    """True when Unicode runs would mangle the maths -- render it instead."""
    if _HEAVY.search(m):
        return True
    # a fraction inside a fraction
    for fm in re.finditer(r"\\t?frac", m):
        _, (a, b), _ = read_args(m, fm.end(), 2)
        if "\\frac" in a or "\\frac" in b:
            return True
    return False


# ---------------------------------------------------------------------------
# Inline markup -> runs
# ---------------------------------------------------------------------------

_LITERALS = [
    ("\\%", "%"), ("\\_", "_"), ("\\&", "&"), ("\\#", "#"), ("\\$", "$"),
    ("\\{", "{"), ("\\}", "}"), ("\\textbackslash", "\\"),
    ("\\ldots", "…"), ("\\dots", "…"), ("\\textellipsis", "…"),
    ("\\textendash", "–"), ("\\textemdash", "—"), ("\\textdegree", "°"),
    ("\\textquotesingle", "'"), ("\\copyright", "©"), ("\\pounds", "£"),
    ("\\euro", "€"), ("\\checkmark", "✓"), ("\\LaTeX", "LaTeX"),
    ("\\TeX", "TeX"),
    ("---", "—"), ("--", "–"), ("``", "“"), ("''", "”"), ("`", "‘"),
    ("'", "’"),
    ("\\,", "\u2009"), ("\\;", " "), ("\\:", " "), ("\\ ", " "),
    ("\\@", ""), ("\\-", ""), ("\\/", ""), ("~", "\u00a0"),
]

_TRANSPARENT = {"textnormal", "mdseries", "normalfont", "upshape",
                "footnotesize", "scriptsize", "tiny", "small", "large",
                "Large", "LARGE", "huge", "Huge", "normalsize", "mbox",
                "makebox", "hbox", "textup", "textrm", "textsf", "RaggedRight",
                "raggedright", "centering", "nolinkurl", "textsc", "sffamily",
                "rmfamily", "text"}
_DROP_WITH_ARG = {"label", "index", "vspace", "hspace", "needspace",
                  "phantom", "hphantom", "vphantom",
                  "Needspace", "rowcolor", "cellcolor", "faIcon", "rowcolors",
                  "setlength", "addcontentsline", "markboth", "thispagestyle",
                  "pagestyle", "setcounter", "arrayrulecolor", "structuralfigures"}
_DROP = {"medskip", "smallskip", "bigskip", "centering", "noindent", "par",
         "toprule", "midrule", "bottomrule", "endhead", "endfirsthead",
         "endfoot", "endlastfoot", "hline", "arraybackslash", "allowbreak",
         "newpage", "clearpage", "item", "quad", "qquad", "hfill", "vfill",
         "normalfigures", "relax", "linebreak", "nolinebreak", "protect",
         "leavevmode", "unskip", "ignorespaces", "raggedleft", "break",
         "cmidrule", "addlinespace", "tabularnewline", "footnotemark"}

_STRAIGHTEN = str.maketrans({"’": "'", "‘": "'", "“": '"', "”": '"',
                             "\u00a0": " ", "—": "--", "–": "-"})


def _chref_text(key: str) -> str:
    n = ref_num("ch:" + key)
    if n:
        return f"Chapter {n}"
    n = ref_num("app:" + key)
    return f"Appendix {n}" if n else "another chapter"


def _dsref_text(key: str) -> str:
    n = ref_num("fig:" + key)
    return f"Figure {n}" if n else "the figure"


def text_runs(tex: str, bold=False, italic=False, mono=False, term=False,
              color=None) -> list[Run]:
    """Convert a fragment of LaTeX prose into styled runs."""
    runs: list[Run] = []
    buf: list[str] = []

    def flush():
        if buf:
            txt = "".join(buf)
            if mono:
                txt = txt.translate(_STRAIGHTEN)
            runs.append(Run(txt, bold, italic, mono, term, color=color))
            buf.clear()

    i, n = 0, len(tex)
    while i < n:
        # literal sequences (never inside \texttt for the quote ligatures)
        hit = False
        for src, dst in _LITERALS:
            if tex.startswith(src, i):
                if mono and src in ("---", "--", "``", "''", "`", "'"):
                    buf.append(src)
                else:
                    buf.append(dst)
                i += len(src)
                hit = True
                break
        if hit:
            continue
        c = tex[i]
        if c == "\\":
            if tex.startswith("\\\\", i):
                # a hard line break inside a run of prose
                buf.append(" ")
                i += 2
                if i < n and tex[i] == "[":
                    _, i = find_group(tex, i, "[", "]")
                continue
            mm = re.match(r"\\([A-Za-z]+)\*?", tex[i:])
            if not mm:
                i += 2
                continue
            name = mm.group(1)
            j = i + mm.end()
            if name in _DROP:
                i = j
                continue
            if name in _DROP_WITH_ARG:
                k = skip_ws(tex, j)
                if k < n and tex[k] == "[":
                    _, k = find_group(tex, k, "[", "]")
                    k = skip_ws(tex, k)
                cnt = 3 if name == "rowcolors" else \
                    2 if name == "setlength" else 1
                for _ in range(cnt):
                    k = skip_ws(tex, k)
                    if k < n and tex[k] == "{":
                        _, k = find_group(tex, k)
                i = k
                continue
            # declarations that style the rest of their group: {\bfseries x}
            if name in ("bfseries", "itshape", "ttfamily", "em"):
                flush()
                rest = tex[j:]
                sub = text_runs(rest, bold or name == "bfseries",
                                italic or name in ("itshape", "em"),
                                mono or name == "ttfamily", term, color)
                runs += sub
                i = n
                continue
            if name == "color":
                k = skip_ws(tex, j)
                col = None
                if k < n and tex[k] == "{":
                    col, k = find_group(tex, k)
                flush()
                runs += text_runs(tex[k:], bold, italic, mono, term,
                                  _colname(col))
                i = n
                continue
            k = skip_ws(tex, j) if name not in ("\\",) else j
            if k < n and tex[k] == "[" and name not in ("item",):
                _, k = find_group(tex, k, "[", "]")
            if k < n and tex[k] == "{":
                inner, k2 = find_group(tex, k)
                flush()
                if name == "textbf":
                    runs += text_runs(inner, True, italic, mono, term, color)
                elif name in ("emph", "textit", "textsl"):
                    runs += text_runs(inner, bold, not italic, mono, term,
                                      color)
                elif name in ("texttt", "code", "path", "url", "lstinline"):
                    runs += text_runs(inner, bold, italic, True, term, color)
                elif name == "term":
                    runs += text_runs(inner, True, italic, mono, True, color)
                elif name == "chref":
                    runs.append(Run(_chref_text(inner), bold, italic,
                                    color=color))
                elif name == "dsref":
                    runs.append(Run(_dsref_text(inner), bold, italic,
                                    color=color))
                elif name == "ref":
                    runs.append(Run(ref_num(inner) or "?", bold, italic))
                elif name == "textcolor":
                    k3 = skip_ws(tex, k2)
                    if k3 < n and tex[k3] == "{":
                        inner2, k2 = find_group(tex, k3)
                        runs += text_runs(inner2, bold, italic, mono, term,
                                          _colname(inner))
                    else:
                        runs += text_runs(inner, bold, italic, mono, term,
                                          color)
                elif name == "footnote":
                    runs.append(Run(" (" + plain(text_runs(inner)) + ")",
                                    bold, italic, mono, term, color=color))
                elif name == "href":
                    k3 = skip_ws(tex, k2)
                    if k3 < n and tex[k3] == "{":
                        inner2, k2 = find_group(tex, k3)
                        runs += text_runs(inner2, bold, italic, mono, term,
                                          color)
                elif name == "texorpdfstring":
                    k3 = skip_ws(tex, k2)
                    if k3 < n and tex[k3] == "{":
                        _, k2 = find_group(tex, k3)
                    runs += text_runs(inner, bold, italic, mono, term, color)
                elif name == "textsuperscript":
                    for r in text_runs(inner, bold, italic, mono, term, color):
                        r.script = "sup"
                        runs.append(r)
                elif name == "textsubscript":
                    for r in text_runs(inner, bold, italic, mono, term, color):
                        r.script = "sub"
                        runs.append(r)
                else:
                    runs += text_runs(inner, bold, italic, mono, term, color)
                i = k2
                continue
            # a command with no argument that we do not know: drop it
            i = j
            continue
        if c == "$":
            flush()
            k = i + 1
            while k < n and not (tex[k] == "$" and tex[k - 1] != "\\"):
                k += 1
            for r in math_runs(tex[i + 1:k], bold=bold):
                if color and not r.color:
                    r.color = color
                runs.append(r)
            i = k + 1
            continue
        if c == "{":
            inner, k = find_group(tex, i)
            flush()
            runs += text_runs(inner, bold, italic, mono, term, color)
            i = k
            continue
        if c == "}":
            i += 1
            continue
        if c == "&":
            buf.append(" ")
            i += 1
            continue
        buf.append(c)
        i += 1
    flush()
    return _tidy(runs)


def _colname(c: str | None) -> str | None:
    """Reduce a xcolor expression (accent!75!black) to its base name."""
    if not c:
        return None
    return c.split("!")[0].strip()


def _tidy(runs: list[Run], strip: bool = True) -> list[Run]:
    """Collapse whitespace and merge adjacent runs that share a style."""
    out: list[Run] = []
    for r in runs:
        if not r.mono:
            r.text = re.sub(r"[ \t\n]+", " ", r.text)
        else:
            r.text = re.sub(r"[\t\n]+", " ", r.text)
        if not r.text:
            continue
        if out and out[-1].style() == r.style():
            out[-1].text += r.text
        else:
            out.append(r)
    # a space between two runs should not be doubled
    for a, b in zip(out, out[1:]):
        if a.text.endswith(" ") and b.text.startswith(" "):
            b.text = b.text.lstrip(" ")
    out = [r for r in out if r.text]
    if strip and out:
        out[0].text = out[0].text.lstrip()
        out[-1].text = out[-1].text.rstrip()
        out = [r for r in out if r.text]
    return out


def plain(runs) -> str:
    """Runs back to bare text -- for notes, titles and measuring."""
    if isinstance(runs, str):
        return runs
    return "".join(r.text for r in runs).strip()


def resolve_refs(tex: str) -> str:
    """Replace \\chref/\\dsref with their text, for LaTeX re-rendering."""
    tex = re.sub(r"\\chref\{([^}]*)\}", lambda m: _chref_text(m.group(1)),
                 tex)
    tex = re.sub(r"\\dsref\{([^}]*)\}", lambda m: _dsref_text(m.group(1)),
                 tex)
    return tex


# ---------------------------------------------------------------------------
# Body parsing: a box body (or a whole chapter) -> element sequence
# ---------------------------------------------------------------------------

BOX_OPT = {"definitionbox", "conceptbox", "alertbox", "examplebox", "worked",
           "pitfall", "lab"}
BOX_PLAIN = {"outcomes", "prereq", "keyterms", "checkpoint", "chaptersummary",
             "reviewq"}
BOX_ENVS = BOX_OPT | BOX_PLAIN | {"fourlenses"}
LIST_ENVS = {"tight", "tightnum", "itemize", "enumerate", "ans"}
LENSES = ("LA", "PM", "IOT", "SEC")

_TOKEN = re.compile(
    r"\\begin\{(?P<env>[A-Za-z*]+)\}"
    r"|\\(?P<fig>dsfigh?)\{"
    r"|(?P<dmath>(?<!\\)\\\[)"
    r"|\\(?P<lens>lens(?:LA|PM|IOT|SEC))\{"
    r"|\\(?P<lst>LSTPH)\{(?P<lstn>\d+)\}"
    r"|\\(?P<sec>(?:sub)*section)\*?\{"
    r"|\\(?P<rowc>rowcolors)\{")


class _Ctx:
    """Per-chapter parse state: protected listings and heading tracking."""

    def __init__(self, listings):
        self.listings = listings
        self.band = "lightbg"


def _protect_listings(src: str) -> tuple[str, list[tuple[str, str]]]:
    listings = []

    def take(m):
        opts = m.group(1) or ""
        listings.append((opts, m.group(2)))
        return "\\LSTPH{%d}" % (len(listings) - 1)

    src = re.sub(r"\\begin\{lstlisting\}(\[[^\]]*\])?\n?(.*?)\\end\{lstlisting\}",
                 take, src, flags=re.S)
    return src, listings


def _lang(opts: str) -> str:
    m = re.search(r"language=([A-Za-z]+)", opts)
    if not m:
        return ""
    l = m.group(1).lower()
    return {"pythonds": "python"}.get(l, l)


def parse_body(tex: str, ctx: _Ctx, headings=None) -> list:
    """Parse LaTeX into a flat list of elements (and headings, if tracked).

    *headings*, when given, is a callable receiving (level, title_tex) for
    each \\section / \\subsection met, so the caller can track position.
    Returned list may then contain ('heading', level, title) tuples.
    """
    out: list = []
    pos = 0
    n = len(tex)
    while True:
        m = _TOKEN.search(tex, pos)
        if not m:
            out += _prose(tex[pos:], ctx)
            break
        out += _prose(tex[pos:m.start()], ctx)
        if m.group("env"):
            env = m.group("env")
            end_tag = "\\end{%s}" % env
            start = m.end()
            end = _find_env_end(tex, start, env)
            inner = tex[start:end]
            pos = end + len(end_tag)
            if env in BOX_ENVS:
                out.append(_box(env, inner, ctx))
            elif env in LIST_ENVS:
                out.append(_list(inner, env in ("tightnum", "enumerate", "ans"),
                                 ctx))
            elif env in ("longtable", "tabular", "tabularx"):
                out.append(_table(env, inner, ctx))
            elif env in ("center", "quote", "flushleft", "minipage",
                         "tcolorbox", "multicols"):
                if env in ("tcolorbox", "minipage", "multicols"):
                    # skip the option/width argument
                    k = skip_ws(inner, 0)
                    if k < len(inner) and inner[k] == "[":
                        _, k = find_group(inner, k, "[", "]")
                    if env in ("minipage", "multicols"):
                        k = skip_ws(inner, k)
                        if k < len(inner) and inner[k] == "{":
                            _, k = find_group(inner, k)
                    inner = inner[k:]
                out += parse_body(inner, ctx, headings)
            elif env in ("equation", "equation*", "align", "align*",
                         "gather", "gather*"):
                out.append(MathE(inner.strip()))
            elif env == "tikzpicture":
                pass            # bare TikZ (only \locator uses one)
            else:
                out += parse_body(inner, ctx, headings)
        elif m.group("fig"):
            i = m.end() - 1
            body, i = find_group(tex, i)
            _, (cap, lab), i = read_args(tex, i, 2)
            out.append(FigureE(lab.strip(), text_runs(cap),
                               resolve_refs(cap), body.strip(),
                               here=m.group("fig") == "dsfigh"))
            pos = i
        elif m.group("dmath"):
            k = tex.find("\\]", m.end())
            k = n if k == -1 else k
            out.append(MathE(resolve_refs(tex[m.end():k].strip())))
            pos = k + 2
        elif m.group("lens"):
            i = m.end() - 1
            inner, i = find_group(tex, i)
            b = Box("lens", arg=m.group("lens")[4:])
            b.body = parse_body(inner, ctx)
            out.append(b)
            pos = i
        elif m.group("lst"):
            opts, code = ctx.listings[int(m.group("lstn"))]
            out.append(CodeE(code.strip("\n"), _lang(opts)))
            pos = m.end()
        elif m.group("sec"):
            i = m.end() - 1
            title, i = find_group(tex, i)
            level = m.group("sec").count("sub") + 1
            if headings is not None:
                out.append(("heading", level, title))
            pos = i
        elif m.group("rowc"):
            _, (a, b, c), i = read_args(tex, m.end() - 1, 3)
            ctx.band = _colname(c) or "lightbg"
            pos = i
    return out


def _find_env_end(tex: str, start: int, env: str) -> int:
    """Index of the \\end{env} matching a \\begin{env}, honouring nesting."""
    depth = 1
    pat = re.compile(r"\\(begin|end)\{" + re.escape(env) + r"\}")
    for m in pat.finditer(tex, start):
        depth += 1 if m.group(1) == "begin" else -1
        if depth == 0:
            return m.start()
    return len(tex)


_BREAKS = re.compile(r"\n\s*\n|\\(?:smallskip|medskip|bigskip|par)\b"
                     r"|(?<!\\)\\\\(?:\[[^\]]*\])?|\\newline\b")


def _prose(tex: str, ctx: _Ctx) -> list[Para]:
    """Split prose into paragraphs; line breaks start a new paragraph."""
    out = []
    for chunk in _BREAKS.split(tex):
        if not chunk or not chunk.strip():
            continue
        p = make_para(chunk)
        if p is not None:
            out.append(p)
    return out


def make_para(chunk: str) -> Para | None:
    runs = text_runs(chunk)
    if not runs or not plain(runs):
        return None
    heavy = any(math_is_heavy(mm.group(1))
                for mm in re.finditer(r"(?<!\\)\$(.+?)(?<!\\)\$", chunk,
                                      re.S))
    return Para(runs, resolve_refs(chunk.strip()), heavy)


def _list(inner: str, numbered: bool, ctx: _Ctx) -> ListE:
    items = []
    # split on \item at depth 0 of nested lists
    parts = _split_items(inner)
    for it in parts:
        it = re.sub(r"^\s*\[[^\]]*\]\s*", "", it)
        # nested list inside an item: flatten its items after this one
        nested = []
        for nm in re.finditer(r"\\begin\{(tight|tightnum|itemize|enumerate)\}"
                              r"(.*?)\\end\{\1\}", it, re.S):
            nested.append(nm)
        head = it
        if nested:
            head = it[:nested[0].start()]
        p = make_para(head)
        if p is not None:
            items.append(p)
        for nm in nested:
            for sub in _split_items(nm.group(2)):
                q = make_para(sub)
                if q is not None:
                    q.runs.insert(0, Run("– "))
                    items.append(q)
    return ListE(items, numbered)


def _split_items(inner: str) -> list[str]:
    out, depth, last = [], 0, None
    for m in re.finditer(r"\\begin\{(tight|tightnum|itemize|enumerate)\}|"
                         r"\\end\{(tight|tightnum|itemize|enumerate)\}|"
                         r"\\item\b", inner):
        tok = m.group(0)
        if tok.startswith("\\begin"):
            depth += 1
        elif tok.startswith("\\end"):
            depth -= 1
        elif depth == 0:
            if last is not None:
                out.append(inner[last:m.start()])
            last = m.end()
    if last is not None:
        out.append(inner[last:])
    return [o for o in out if o.strip()]


def _box(env: str, inner: str, ctx: _Ctx) -> Box:
    title_tex = ""
    i = 0
    if env in BOX_OPT or env == "fourlenses":
        k = skip_ws(inner, 0)
        if env == "fourlenses" and k < len(inner) and inner[k] == "{":
            title_tex, i = find_group(inner, k)
        elif k < len(inner) and inner[k] == "[":
            title_tex, i = find_group(inner, k, "[", "]")
    b = Box(env, text_runs(title_tex), resolve_refs(title_tex))
    body = parse_body(inner[i:], ctx)
    if env == "fourlenses":
        b.arg = plain(b.title)
        b.children = [e for e in body if isinstance(e, Box) and e.env == "lens"]
        b.body = [e for e in body if not (isinstance(e, Box)
                                          and e.env == "lens")]
    elif env == "reviewq":
        b.body = [_list(inner, True, ctx)]
    else:
        b.body = body
    return b


# -- tables -----------------------------------------------------------------

def _colspec(spec: str) -> tuple[list[float], list[str]]:
    widths, align = [], []
    i = 0
    while i < len(spec):
        c = spec[i]
        if c == "@" or c == ">" or c == "<":
            i += 1
            if i < len(spec) and spec[i] == "{":
                _, i = find_group(spec, i)
            continue
        if c in "LpXmb":
            i += 1
            w = 3.0
            if i < len(spec) and spec[i] == "{":
                g, i = find_group(spec, i)
                mm = re.match(r"\s*([\d.]+)\s*(cm|mm|in|pt)?", g)
                if mm:
                    w = float(mm.group(1))
                    unit = mm.group(2) or "cm"
                    w *= {"cm": 1, "mm": 0.1, "in": 2.54, "pt": 0.0353}[unit]
                elif "linewidth" in g or "textwidth" in g:
                    w = 5.0
            widths.append(w)
            align.append("l")
            continue
        if c in "lrc":
            widths.append(0.0)          # natural width -- decided later
            align.append(c)
        i += 1
    return widths, align


def _split_top(s: str, sep: str) -> list[str]:
    """Split on *sep* outside braces (and not escaped)."""
    out, depth, cur, i = [], 0, [], 0
    while i < len(s):
        c = s[i]
        if c == "\\":
            if s.startswith(sep, i) and sep == "\\\\" and depth == 0:
                out.append("".join(cur))
                cur = []
                i += 2
                if i < len(s) and s[i] == "[":
                    _, i = find_group(s, i, "[", "]")
                continue
            cur.append(s[i:i + 2])
            i += 2
            continue
        if c == "{":
            depth += 1
        elif c == "}":
            depth -= 1
        if depth == 0 and sep != "\\\\" and s.startswith(sep, i):
            out.append("".join(cur))
            cur = []
            i += len(sep)
            continue
        cur.append(c)
        i += 1
    out.append("".join(cur))
    return out


def _table(env: str, inner: str, ctx: _Ctx) -> TableE:
    k = skip_ws(inner, 0)
    if env == "tabularx" and k < len(inner) and inner[k] == "{":
        _, k = find_group(inner, k)
        k = skip_ws(inner, k)
    spec = ""
    if k < len(inner) and inner[k] == "{":
        spec, k = find_group(inner, k)
    widths, align = _colspec(spec)
    body = inner[k:]
    # the header is everything before the first \midrule
    head_src, _, rows_src = body.partition("\\midrule")
    if not rows_src:
        head_src, rows_src = "", body
    rows_src = rows_src.replace("\\endhead", "").replace("\\endfirsthead", "")

    def cells(row: str) -> list[list[Run]]:
        row = re.sub(r"\\(rowcolor|cellcolor)(\[[^\]]*\])?\{[^}]*\}", "", row)
        row = re.sub(r"\\(toprule|bottomrule|hline|midrule)", "", row)
        return [text_runs(c) for c in _split_top(row, "&")]

    header = []
    for r in _split_top(head_src, "\\\\"):
        if r.strip() and re.sub(r"\\(toprule|rowcolor\{[^}]*\})", "",
                                r).strip():
            header = cells(r)
    rows = []
    marks = []
    for r in _split_top(rows_src, "\\\\"):
        rr = re.sub(r"\\(toprule|bottomrule|midrule|hline|endhead)", "", r)
        if not rr.strip():
            continue
        row = cells(rr)
        if any(plain(c) for c in row):
            rows.append(row)
    ncol = max([len(widths)] + [len(r) for r in rows] + [len(header)])
    widths += [0.0] * (ncol - len(widths))
    align += ["l"] * (ncol - len(align))
    for r in rows:
        r += [[] for _ in range(ncol - len(r))]
    if header:
        header += [[] for _ in range(ncol - len(header))]
    return TableE(header, rows, widths, align, ctx.band, marks)


# ---------------------------------------------------------------------------
# Chapters
# ---------------------------------------------------------------------------

@lru_cache(maxsize=1)
def chapter_order() -> list[tuple[int, str]]:
    """[(number, stem)] for the 32 chapters, in master-file order."""
    master = open(MASTER, encoding="utf-8").read()
    out, n = [], 0
    for stem in re.findall(r"^\\input\{parts/([^}]+)\}", master, re.M):
        p = os.path.join(PARTS, stem + ".tex")
        if not os.path.exists(p) or not re.match(r"\d\d_", stem) \
                or stem.startswith("00_"):
            continue
        n += 1
        out.append((n, stem))
    return out


def _norm(t: str) -> str:
    t = plain(text_runs(t)) if "\\" in t or "$" in t else t
    return re.sub(r"[^a-z0-9]+", "", t.lower())


@lru_cache(maxsize=64)
def parse_chapter(number: int) -> Chapter:
    stem = dict(chapter_order())[number]
    raw = open(os.path.join(PARTS, stem + ".tex"), encoding="utf-8").read()
    src, listings = _protect_listings(raw)
    src = strip_comments(src)
    ctx = _Ctx(listings)

    m = re.search(r"\\chapter\{", src)
    title_tex, j = find_group(src, m.end() - 1)
    lm = re.match(r"\s*\\label\{ch:([^}]+)\}", src[j:])
    label = lm.group(1) if lm else stem
    title = plain(text_runs(title_tex))
    loc = re.search(r"\\locator\{(\d)\}", src)
    part = int(loc.group(1)) if loc else 0
    body = src[j + (lm.end() if lm else 0):]
    body = re.sub(r"\\locator\{\d\}\{[^}]*\}", "", body)

    # sections of this chapter, from the toc lines in the .aux
    toc = aux().toc
    chap_num = str(number)
    secs: list[Section] = []
    for lvl, num, t, page in toc:
        if lvl in (1, 2) and num.split(".")[0] == chap_num:
            secs.append(Section(num, plain(text_runs(t)), page, lvl))
    ch_page = ref_page("ch:" + label)

    items: list[Item] = []
    section = top = secnum = ""
    page = ch_page
    sec_iter = {(_norm(s.title), s.level): s for s in secs}
    for el in parse_body(body, ctx, headings=True):
        if isinstance(el, tuple) and el[0] == "heading":
            _, level, t = el
            ttl = plain(text_runs(t))
            s = sec_iter.get((_norm(t), level))
            section = ttl
            if level == 1:
                top = ttl
            if s:
                secnum, page = s.num, s.page
            continue
        items.append(Item(el, section, top, secnum, page))
    return Chapter(number, label, title, stem, part, ch_page, items, secs)


@lru_cache(maxsize=1)
def all_stems() -> list[str]:
    """Every file the master \\inputs, in order."""
    master = open(MASTER, encoding="utf-8").read()
    return re.findall(r"^\\input\{parts/([^}]+)\}", master, re.M)


def parse_file(stem: str) -> list:
    """Any handout file -> flat element list (headings dropped)."""
    raw = open(os.path.join(PARTS, stem + ".tex"), encoding="utf-8").read()
    src, listings = _protect_listings(raw)
    src = strip_comments(src)
    return [e for e in parse_body(src, _Ctx(listings), headings=True)
            if not isinstance(e, tuple)]


def all_figures() -> list[tuple[str, FigureE]]:
    """(stem, figure) for every \\dsfig in the handout, in book order."""
    out = []
    for stem in all_stems():
        if not os.path.exists(os.path.join(PARTS, stem + ".tex")):
            continue
        for e in walk(parse_file(stem)):
            if isinstance(e, FigureE):
                out.append((stem, e))
    return out


# ---------------------------------------------------------------------------
# Appendices: checkpoint answers and the teaching schedule
# ---------------------------------------------------------------------------

@lru_cache(maxsize=1)
def _answers() -> dict[int, list[Para]]:
    raw = open(os.path.join(PARTS, "D_answers.tex"), encoding="utf-8").read()
    src, listings = _protect_listings(raw)
    src = strip_comments(src)
    ctx = _Ctx(listings)
    out = {}
    for m in re.finditer(r"\\ansch\{(\d+)[^}]*\}\s*\\begin\{ans\}(.*?)"
                         r"\\end\{ans\}", src, re.S):
        out[int(m.group(1))] = _list(m.group(2), True, ctx).items
    return out


def checkpoint_answers(number: int) -> list[Para]:
    return _answers().get(number, [])


@dataclass
class Week:
    week: int
    topic: str
    chapters: list[int]
    lab: str
    assessment: str


@lru_cache(maxsize=1)
def schedule() -> list[Week]:
    raw = open(os.path.join(PARTS, "G_schedule.tex"), encoding="utf-8").read()
    src = strip_comments(raw)
    m = re.search(r"\\begin\{longtable\}(.*?)\\end\{longtable\}", src, re.S)
    t = _table("longtable", m.group(1), _Ctx([]))
    out = []
    for r in t.rows:
        wk = plain(r[0])
        chs: list[int] = []
        for part in re.split(r",\s*", plain(r[2])):
            mm = re.match(r"(\d+)\s*[–-]+\s*(\d+)", part)
            if mm:
                chs += list(range(int(mm.group(1)), int(mm.group(2)) + 1))
            elif part.strip().isdigit():
                chs.append(int(part))
        a = plain(r[4])
        out.append(Week(int(wk) if wk.isdigit() else 0, plain(r[1]), chs,
                        plain(r[3]), "" if a in ("—", "-", "---") else a))
    return out


def week_of(number: int) -> Week | None:
    for w in schedule():
        if number in w.chapters:
            return w
    return None


def lab_code_file(number: int) -> str:
    """The runnable extract in code/ for this chapter, if one exists."""
    if not os.path.isdir(CODE):
        return ""
    for f in sorted(os.listdir(CODE)):
        if f.startswith(f"ch{number:02d}_") and f.endswith((".py", ".R")):
            return f
    return ""


# ---------------------------------------------------------------------------
# Inspection and audit
# ---------------------------------------------------------------------------

def walk(elems):
    """Every element, depth first, including box bodies and lenses."""
    for e in elems:
        yield e
        if isinstance(e, Box):
            yield from walk(e.body)
            yield from walk(e.children)


def audit() -> int:
    """Parse every chapter and check that nothing was silently dropped."""
    bad = 0
    print(f"{'ch':>3}  {'title':40} items figs tabs code math boxes  Q/A  status")
    print("-" * 94)
    for n, stem in chapter_order():
        raw = open(os.path.join(PARTS, stem + ".tex"), encoding="utf-8").read()
        src = strip_comments(_protect_listings(raw)[0])
        want = {
            "figure": len(re.findall(r"\\dsfigh?\{", src)),
            "table": len(re.findall(r"\\begin\{(longtable|tabular)\}", src)),
            "code": len(re.findall(r"\\begin\{lstlisting\}", raw)),
            "math": len(re.findall(r"(?<!\\)\\\[", src)),
            "box": sum(len(re.findall(r"\\begin\{%s\}" % e, src))
                       for e in BOX_ENVS),
            "lens": len(re.findall(r"\\lens(LA|PM|IOT|SEC)\{", src)),
        }
        ch = parse_chapter(n)
        got = {k: 0 for k in want}
        for e in walk([i.el for i in ch.items]):
            if isinstance(e, Box):
                got["lens" if e.env == "lens" else "box"] += 1
            elif e.kind in got:
                got[e.kind] += 1
        miss = {k: (want[k], got[k]) for k in want if want[k] != got[k]}
        cp = ch.first("checkpoint")
        nq = sum(len(e.items) for e in (cp.body if cp else [])
                 if isinstance(e, ListE))
        na = len(checkpoint_answers(n))
        unnamed = [i for i in ch.items if not i.secnum]
        ok = not miss and nq == na and nq > 0 and ch.part and ch.page
        if not ok:
            bad += 1
        note = "ok" if ok else ("; ".join(f"{k} {w}!={g}"
                                         for k, (w, g) in miss.items())
                                or f"Q{nq}!=A{na}" if nq != na else
                                "no part/page")
        print(f"{n:>3}  {ch.title[:40]:40} {len(ch.items):>5} {got['figure']:>4}"
              f" {got['table']:>4} {got['code']:>4} {got['math']:>4}"
              f" {got['box']:>5} {nq:>2}/{na:<2} {note}")
    wk = schedule()
    covered = sorted({c for w in wk for c in w.chapters})
    print()
    print(f"schedule: {len(wk)} weeks, chapters covered "
          f"{'all 32' if covered == list(range(1, 33)) else covered}")
    print("all chapters parse completely" if not bad
          else f"{bad} chapter(s) with dropped content")
    return 1 if bad else 0


def _dump(e, ind="  "):
    if isinstance(e, Para):
        flag = " [HEAVY]" if e.heavy else ""
        print(f"{ind}P{flag}: {e.text[:150]}")
    elif isinstance(e, ListE):
        print(f"{ind}LIST{' (numbered)' if e.numbered else ''}:")
        for it in e.items:
            print(f"{ind}  - {'[H] ' if it.heavy else ''}{it.text[:140]}")
    elif isinstance(e, MathE):
        print(f"{ind}MATH: {e.tex[:120]}")
    elif isinstance(e, TableE):
        print(f"{ind}TABLE {len(e.rows)}x{len(e.widths)} w={e.widths} "
              f"a={e.align} band={e.band}")
        print(f"{ind}  H: {[plain(c) for c in e.header]}")
        for r in e.rows[:3]:
            print(f"{ind}  R: {[plain(c)[:30] for c in r]}")
    elif isinstance(e, CodeE):
        lines = e.code.split("\n")
        print(f"{ind}CODE [{e.lang or 'text'}] {len(lines)} lines: "
              f"{lines[0][:80]}")
    elif isinstance(e, FigureE):
        print(f"{ind}FIGURE {e.label} ({ref_num('fig:' + e.label)}): "
              f"{plain(e.caption)[:100]}")
    elif isinstance(e, Box):
        print(f"{ind}[{e.env}] {e.title_text}{(' {' + e.arg + '}') if e.arg else ''}")
        for x in e.body:
            _dump(x, ind + "    ")
        for c in e.children:
            _dump(c, ind + "    ")


def main() -> int:
    if "--all" in sys.argv:
        return audit()
    n = int(sys.argv[1]) if len(sys.argv) > 1 else 1
    ch = parse_chapter(n)
    w = week_of(n)
    print(f"Chapter {ch.number}: {ch.title}  (ch:{ch.label}, part {ch.part}, "
          f"p. {ch.page}-{ch.last_page})")
    if w:
        print(f"  week {w.week}: {w.topic} | lab: {w.lab} | "
              f"assess: {w.assessment or '-'}")
    print(f"  code file: {lab_code_file(n) or '-'}")
    print(f"  sections: " + "; ".join(f"{s.num} {s.title} (p{s.page})"
                                      for s in ch.sections))
    if "--full" in sys.argv:
        cur = None
        for it in ch.items:
            if it.section != cur:
                cur = it.section
                print(f"\n== {it.secnum} {it.section}  [top: {it.top}] "
                      f"p.{it.page}")
            _dump(it.el)
        ans = checkpoint_answers(n)
        print(f"\n== answers ({len(ans)})")
        for a in ans:
            print("  - " + a.text[:140])
    else:
        from collections import Counter
        c = Counter()
        for e in walk([i.el for i in ch.items]):
            c[e.env if isinstance(e, Box) else e.kind] += 1
        for k, v in c.most_common():
            print(f"  {k:16} {v}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
