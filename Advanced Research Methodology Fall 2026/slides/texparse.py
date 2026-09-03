# -*- coding: utf-8 -*-
"""Parse the handout's LaTeX into slide-ready content.

The decks are generated from parts/*.tex rather than written separately, so a
correction to the handout reaches the slides on the next build and the two can
never drift. Everything here is read-only with respect to the handout.

The chapter furniture is fixed and documented in ../README.md, which is what
makes this parse tractable: every chapter has \\locator, outcomes, prereq,
keyterms, ..., traditions, checkpoint, chaptersummary, reviewq in that order.
"""
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PARTS = os.path.join(ROOT, "parts")


# Chapter numbers come from the compiled .aux so the slides always agree with
# the printed volume; falling back to file order if the handout is unbuilt.
def chapter_numbers():
    aux = os.path.join(ROOT, "AdvancedResearchMethodology_Handout.aux")
    m = {}
    if os.path.exists(aux):
        txt = open(aux, encoding="utf-8", errors="replace").read()
        for k, n in re.findall(r"newlabel\{ch:([a-z]+)\}\{\{(\d+)\}", txt):
            m[k] = int(n)
    return m


PART_TITLES = {
    1: "Foundations of Inquiry in Computing",
    2: "Finding and Synthesising Knowledge",
    3: "Research Designs for Computing",
    4: "Measurement, Data and Analysis",
    5: "Rigour, Reproducibility and Ethics",
    6: "Producing and Defending Research",
}

# --------------------------------------------------------------- text
ACCENTS = {
    r'\{\\"u\}': "u", r'\{\\"o\}': "o", r'\{\\"a\}': "a", r'\{\\"i\}': "i",
    r"\{\\'e\}": "e", r"\{\\'a\}": "a", r"\{\\'o\}": "o", r"\{\\'i\}": "i",
    r"\{\\`e\}": "e", r"\{\\`a\}": "a", r"\{\\`o\}": "o",
    r"\{\\^e\}": "e", r"\{\\^o\}": "o", r"\{\\~n\}": "n",
    r"\{\\c\{c\}\}": "c", r"\{\\o\}": "o", r"\{\\aa\}": "a",
    r"\{\\ss\}": "ss",
}
# unicode replacements applied after the ascii fallbacks above
ACCENT_UNI = {
    '{\\"u}': "\u00fc", '{\\"o}': "\u00f6", '{\\"a}': "\u00e4",
    "{\\'e}": "\u00e9", "{\\'a}": "\u00e1", "{\\'o}": "\u00f3",
    "{\\`e}": "\u00e8", "{\\`a}": "\u00e0", "{\\`o}": "\u00f2",
    "{\\^e}": "\u00ea", "{\\~n}": "\u00f1", "{\\c{c}}": "\u00e7",
    "{\\ss}": "\u00df", "{\\o}": "\u00f8",
}

SYMBOLS = [
    (r"\$\\bullet\$", "\u2022"), (r"\\ldots", "\u2026"), (r"\\dots", "\u2026"),
    (r"\\%", "%"), (r"\\&", "&"), (r"\\_", "_"), (r"\\#", "#"),
    (r"\\\$", "$"), (r"\\S", "\u00a7"), (r"\\,", " "), (r"\\ ", " "),
    (r"\\@", ""), (r"\\/", ""), (r"\\medskip", ""), (r"\\smallskip", ""),
    (r"\\bigskip", ""), (r"\\par\b", " "), (r"\\noindent", ""),
    (r"\\centering", ""), (r"\\newpage", ""), (r"\\clearpage", ""),
    (r"\\toprule", ""), (r"\\midrule", ""), (r"\\bottomrule", ""),
    (r"\\hline", ""), (r"\\endfirsthead", ""), (r"\\endhead", ""),
    (r"\\endfoot", ""),
]

# macros whose single argument survives, macro name discarded
KEEP_ARG = ["term", "emph", "textbf", "textit", "texttt", "textsc", "mbox",
            "text", "textsuperscript", "underline", "textup", "textmd",
            "acrshort", "acrlong"]
# macros discarded entirely, argument and all
DROP_ALL = ["label", "index", "vspace", "hspace", "rule", "caption",
            "addcontentsline", "phantomsection", "Needspace", "armfitw",
            "setlength", "footnotesize", "small", "normalsize", "footnote"]


def _strip_comments(s):
    out = []
    for line in s.split("\n"):
        i, esc = 0, False
        cut = None
        while i < len(line):
            c = line[i]
            if c == "\\" and not esc:
                esc = True
            elif c == "%" and not esc:
                cut = i
                break
            else:
                esc = False
            i += 1
        out.append(line if cut is None else line[:cut])
    return "\n".join(out)


def _balanced(s, start):
    """Given s[start] == '{', return index just past the matching '}'."""
    depth, i = 0, start
    while i < len(s):
        if s[i] == "{" and (i == 0 or s[i - 1] != "\\"):
            depth += 1
        elif s[i] == "}" and s[i - 1] != "\\":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return len(s)


def _apply_macro(s, name, keep):
    out, i, tok = [], 0, "\\" + name
    while True:
        j = s.find(tok, i)
        if j == -1:
            out.append(s[i:])
            break
        # must not be a longer macro name
        after = j + len(tok)
        if after < len(s) and (s[after].isalpha()):
            out.append(s[i:after])
            i = after
            continue
        out.append(s[i:j])
        # optional [..]
        k = after
        while k < len(s) and s[k] == " ":
            k += 1
        if k < len(s) and s[k] == "[":
            e = s.find("]", k)
            k = e + 1 if e != -1 else k
        if k < len(s) and s[k] == "{":
            e = _balanced(s, k)
            if keep:
                out.append(s[k + 1:e - 1])
            i = e
        else:
            i = k
    return "".join(out)


def tex_to_text(s, chnums=None, keep_newlines=False):
    """LaTeX fragment -> readable plain text."""
    if not s:
        return ""
    s = _strip_comments(s)

    # cross references become readable chapter pointers
    if chnums:
        def _ch(m):
            n = chnums.get(m.group(1))
            return ("Ch. %d" % n) if n else "the relevant chapter"
        s = re.sub(r"\\chref\{([^}]*)\}", _ch, s)
    s = re.sub(r"\\chref\{[^}]*\}", "the relevant chapter", s)
    s = re.sub(r"\\(figref|ref|autoref)\{[^}]*\}", "", s)
    s = re.sub(r"\\cite[a-z]*\s*(\[[^\]]*\])?\{[^}]*\}", "", s)
    s = re.sub(r"\\reg\{([^}]*)\}\{([^}]*)\}", r"\1 \2", s)

    for pat, rep in ACCENT_UNI.items():
        s = s.replace(pat, rep)
    for name in DROP_ALL:
        s = _apply_macro(s, name, keep=False)
    for name in KEEP_ARG:
        s = _apply_macro(s, name, keep=True)

    # maths: keep it simple and readable
    s = re.sub(r"\$([^$]*)\$", lambda m: _math(m.group(1)), s)

    for pat, rep in SYMBOLS:
        s = re.sub(pat, rep, s)

    s = s.replace("``", "\u201c").replace("''", "\u201d")
    s = s.replace("---", "\u2014").replace("--", "\u2013")
    s = s.replace("~", " ")
    # any leftover macro without args
    s = re.sub(r"\\[a-zA-Z]+\*?", "", s)
    s = s.replace("{", "").replace("}", "")

    if keep_newlines:
        s = re.sub(r"[ \t]+", " ", s)
        s = re.sub(r"\n{3,}", "\n\n", s)
        return s.strip()
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _math(m):
    m = m.strip()
    rep = {r"\alpha": "\u03b1", r"\beta": "\u03b2", r"\delta": "\u03b4",
           r"\omega": "\u03c9", r"\kappa": "\u03ba", r"\chi": "\u03c7",
           r"\mu": "\u03bc", r"\sigma": "\u03c3", r"\times": "\u00d7",
           r"\leq": "\u2264", r"\geq": "\u2265", r"\approx": "\u2248",
           r"\pm": "\u00b1", r"\bullet": "\u2022", r"\to": "\u2192",
           r"\rightarrow": "\u2192", r"\ldots": "\u2026"}
    for k, v in rep.items():
        m = m.replace(k, v)
    m = re.sub(r"\\text\{([^}]*)\}", r"\1", m)
    m = re.sub(r"[_^]\{([^}]*)\}", r"\1", m)
    m = re.sub(r"\\[a-zA-Z]+", "", m)
    return m.replace("{", "").replace("}", "").strip()


# ------------------------------------------------------------ structure
def find_envs(text, name):
    """All \\begin{name}[opt]...\\end{name}, nesting-aware. -> [(opt, body)]"""
    res, i = [], 0
    b, e = "\\begin{%s}" % name, "\\end{%s}" % name
    while True:
        j = text.find(b, i)
        if j == -1:
            return res
        k = j + len(b)
        opt = ""
        if k < len(text) and text[k] == "[":
            depth, p = 1, k + 1
            while p < len(text) and depth:
                if text[p] == "[":
                    depth += 1
                elif text[p] == "]":
                    depth -= 1
                p += 1
            opt = text[k + 1:p - 1]
            k = p
        elif k < len(text) and text[k] == "{":
            p = _balanced(text, k)
            opt = text[k + 1:p - 1]
            k = p
        # find matching end, honouring nesting of the same environment
        depth, p = 1, k
        while p < len(text) and depth:
            nb, ne = text.find(b, p), text.find(e, p)
            if ne == -1:
                break
            if nb != -1 and nb < ne:
                depth += 1
                p = nb + len(b)
            else:
                depth -= 1
                p = ne + len(e)
        res.append((opt, text[k:p - len(e)] if depth == 0 else text[k:]))
        i = p


def items_of(body):
    """\\item entries of a list environment."""
    inner = body
    for env in ("tight", "tightnum", "itemize", "enumerate"):
        found = find_envs(body, env)
        if found:
            inner = found[0][1]
            break
    parts = re.split(r"\\item\b", inner)[1:]
    return [p.strip() for p in parts if p.strip()]


def first_para(body):
    """The opening prose of a box.

    Boxes often open with a short framing line ("This is worth being precise
    about.") before the substance. Cutting at the first break would put that
    lead-in on a slide and drop the point, so keep extending until there is
    enough text to be worth showing.
    """
    body = body.strip()
    body = re.sub(r"^\s*\\medskip\s*", "", body)
    stops = ("\\medskip", "\\begin{tight", "\\begin{itemize",
             "\\begin{enumerate", "\\begin{center", "\n\n")
    best = body
    for limit in (40, 140, 260):
        cut = len(body)
        for stop in stops:
            k = body.find(stop)
            if k > limit:
                cut = min(cut, k)
        best = body[:cut]
        if len(re.sub(r"\\[a-zA-Z]+|[{}]", "", best).strip()) >= 120:
            break
    return best


def _prose_only(body):
    """Section text with every environment and table removed."""
    out, i = [], 0
    while True:
        j = body.find("\\begin{", i)
        if j == -1:
            out.append(body[i:])
            break
        out.append(body[i:j])
        name = body[j + 7:body.find("}", j + 7)]
        e = body.find("\\end{%s}" % name, j)
        i = (e + len("\\end{%s}" % name)) if e != -1 else len(body)
    return "".join(out)


def condense(text, max_sentences=2, max_chars=240):
    """Prose -> a slide-sized statement.

    Whole sentences only. Cutting mid-sentence leaves a slide ending in
    something like "(Ch." which reads as a bug to the room, so drop trailing
    sentences until it fits and only chop a lone over-long sentence.
    """
    text = text.strip()
    if not text:
        return ""
    sents = [s for s in re.split(r"(?<=[.!?])\s+", text) if s.strip()]
    picked = sents[:max_sentences]
    while len(picked) > 1 and len(" ".join(picked)) > max_chars:
        picked.pop()
    out = " ".join(picked).strip()
    if len(out) > max_chars:
        cut = out[:max_chars]
        sp = cut.rfind(" ")
        out = (cut[:sp] if sp > 60 else cut).rstrip(" ,;:(") + "\u2026"
    return out


def parse_chapter(path, chnums):
    """One parts/NN_*.tex -> a dict of slide-ready content."""
    raw = open(path, encoding="utf-8", errors="replace").read()

    def t(s, **kw):
        return tex_to_text(s, chnums, **kw)

    m = re.search(r"\\chapter\{(.+?)\}\\label\{ch:([a-z]+)\}", raw, re.S)
    if not m:
        return None
    title, key = t(m.group(1)), m.group(2)

    lm = re.search(r"\\locator\{(\d+)\}\{(.+?)\}", raw)
    part_no = int(lm.group(1)) if lm else 0

    ch = {
        "key": key, "number": chnums.get(key, 0), "title": title,
        "part": part_no, "part_title": PART_TITLES.get(part_no, ""),
        "outcomes": [], "prereq": "", "keyterms": [], "sections": [],
        "worked": None, "pitfall": "", "pitfall_title": "",
        "pitfall_items": [], "traditions": None,
        "summary": [], "reviewq": [], "checkpoint": [], "lab": None,
        "regboxes": [], "papers": [],
    }

    env = find_envs(raw, "outcomes")
    if env:
        ch["outcomes"] = [t(i) for i in items_of(env[0][1])]
    env = find_envs(raw, "prereq")
    if env:
        ch["prereq"] = condense(t(env[0][1]), 3, 300)
    env = find_envs(raw, "keyterms")
    if env:
        terms = t(env[0][1])
        ch["keyterms"] = [x.strip() for x in terms.split("\u2022")
                          if x.strip()]

    # section bodies, each with the teaching boxes it contains
    secs = list(re.finditer(r"\\section\{(.+?)\}", raw, re.S))
    for i, sm in enumerate(secs):
        start = sm.end()
        end = secs[i + 1].start() if i + 1 < len(secs) else len(raw)
        body = raw[start:end]
        points = []
        for envname, kind in (("definitionbox", "Definition"),
                              ("conceptbox", "Concept"),
                              ("alertbox", "Watch out"),
                              ("examplebox", "Example"),
                              ("regbox", "Regulation")):
            for opt, eb in find_envs(body, envname):
                head = t(opt) if opt else kind
                txt = condense(t(first_para(eb)), 2, 230)
                if txt:
                    points.append({"kind": kind, "head": head, "text": txt})
        # A section built entirely from a table or figure has no boxes to
        # quote. Fall back to its prose so it still reaches a slide rather
        # than appearing as an empty heading.
        if not points:
            txt = condense(t(_prose_only(body)), 3, 300)
            if txt:
                points.append({"kind": "", "head": "", "text": txt})
        ch["sections"].append({"title": t(sm.group(1)), "points": points})

    env = find_envs(raw, "worked")
    if env:
        opt, body = env[0]
        ch["worked"] = {"title": t(opt) or "Worked example",
                        "text": condense(t(first_para(body)), 3, 300),
                        "items": [t(x) for x in items_of(body)][:6]}
    env = find_envs(raw, "pitfall")
    if env:
        opt, body = env[0]
        its = [t(i) for i in items_of(body)]
        ch["pitfall_title"] = t(opt) or "Common mistakes"
        ch["pitfall_items"] = its[:6]
        ch["pitfall"] = (its[0] if its
                         else condense(t(first_para(body)), 3, 300))

    env = find_envs(raw, "traditions")
    if env:
        body = env[0][1]
        quad = {}
        # \tradEMP{...} carries its text as a braced argument
        for macro, lab in (("tradEMP", "Empirical"),
                           ("tradDSR", "Design science"),
                           ("tradQUAL", "Qualitative"),
                           ("tradFORM", "Formal")):
            mm = re.search(r"\\%s\s*\{" % macro, body)
            if mm:
                e = _balanced(body, mm.end() - 1)
                quad[lab] = condense(t(body[mm.end():e - 1]), 2, 250)
        if quad:
            ch["traditions"] = {"topic": t(env[0][0]), "quadrants": quad}

    env = find_envs(raw, "chaptersummary")
    if env:
        its = [t(i) for i in items_of(env[0][1])]
        ch["summary"] = its or [condense(t(env[0][1]), 4, 420)]
    env = find_envs(raw, "reviewq")
    if env:
        ch["reviewq"] = [t(i) for i in items_of(env[0][1])][:8]
    env = find_envs(raw, "checkpoint")
    if env:
        ch["checkpoint"] = [t(i) for i in items_of(env[0][1])][:6]
    env = find_envs(raw, "lab")
    if env:
        ch["lab"] = {"title": t(env[0][0]) or "Lab",
                     "items": [t(i) for i in items_of(env[0][1])][:6],
                     "text": condense(t(first_para(env[0][1])), 2, 220)}
    for opt, body in find_envs(raw, "regbox"):
        ch["regboxes"].append({"clause": t(opt),
                               "text": condense(t(first_para(body)), 3, 300)})
    for opt, body in find_envs(raw, "paperbox"):
        ch["papers"].append({"handle": t(opt),
                             "text": condense(t(first_para(body)), 2, 240)})
    return ch


def chapter_files():
    return sorted(f for f in os.listdir(PARTS)
                  if re.match(r"^(0[1-9]|[12]\d|3[0-3])_.*\.tex$", f))


if __name__ == "__main__":
    import json
    import sys
    nums = chapter_numbers()
    fn = sys.argv[1] if len(sys.argv) > 1 else "10_experiments.tex"
    ch = parse_chapter(os.path.join(PARTS, fn), nums)
    print(json.dumps(ch, indent=2, ensure_ascii=False)[:4000])
