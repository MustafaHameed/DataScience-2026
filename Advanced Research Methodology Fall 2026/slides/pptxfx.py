# -*- coding: utf-8 -*-
"""Slide transitions and entrance animations for python-pptx.

python-pptx has no animation API, so the <p:timing> and <p:transition> trees
are built here and grafted onto the slide XML. The shapes are addressed by
shape id, which python-pptx exposes as shape.shape_id.

The timing tree PowerPoint expects is deeper than it looks. Every slide needs
one tmRoot -> one mainSeq, and inside mainSeq one <p:par> per click. Effects
sharing a click sit as siblings inside that click's childTnLst, the first
carrying nodeType="clickEffect" and the rest "withEffect" or "afterEffect".
Get that nesting wrong and PowerPoint silently drops every animation on the
slide rather than reporting an error, so keep the shape of build_timing().
"""
from lxml import etree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}


def _q(tag):
    pre, local = tag.split(":")
    return "{%s}%s" % (NS[pre], local)


def _el(tag, **attrs):
    e = etree.SubElement if False else etree.Element
    node = e(_q(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


def _sub(parent, tag, **attrs):
    node = etree.SubElement(parent, _q(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


# Preset ids and filters matching PowerPoint's own entrance effects, so the
# animation pane shows a familiar name rather than "Custom".
EFFECTS = {
    "fade":      dict(presetID=10, subtype=0,  filter="fade"),
    "wipeUp":    dict(presetID=22, subtype=4,  filter="wipe(up)"),
    "wipeRight": dict(presetID=22, subtype=8,  filter="wipe(right)"),
    "flyBottom": dict(presetID=2,  subtype=4,  filter="slide(fromBottom)"),
    "flyLeft":   dict(presetID=2,  subtype=8,  filter="slide(fromLeft)"),
    "zoom":      dict(presetID=23, subtype=16, filter="zoom(in)"),
    "split":     dict(presetID=13, subtype=26, filter="barn(inVertical)"),
}

TRIGGER_NODE = {"click": "clickEffect", "with": "withEffect",
                "after": "afterEffect"}


class _Ids(object):
    """PowerPoint requires globally increasing cTn ids within the slide."""

    def __init__(self, start=1):
        self.n = start

    def next(self):
        self.n += 1
        return self.n


def _effect(parent, ids, spid, effect, trigger, dur, delay, para):
    """One entrance effect on one shape (or one paragraph of it)."""
    spec = EFFECTS.get(effect, EFFECTS["fade"])
    par = _sub(parent, "p:par")
    ctn = _sub(par, "p:cTn", id=ids.next(), presetID=spec["presetID"],
               presetClass="entr", presetSubtype=spec["subtype"],
               fill="hold", grpId=0, nodeType=TRIGGER_NODE[trigger])
    st = _sub(ctn, "p:stCondLst")
    _sub(st, "p:cond", delay=delay)
    ch = _sub(ctn, "p:childTnLst")

    def target(el):
        tgt = _sub(el, "p:tgtEl")
        sp = _sub(tgt, "p:spTgt", spid=spid)
        if para is not None:
            tx = _sub(sp, "p:txEl")
            _sub(tx, "p:pRg", st=para, end=para)

    # make it visible at the start of its own effect
    s = _sub(ch, "p:set")
    cb = _sub(s, "p:cBhvr")
    c1 = _sub(cb, "p:cTn", id=ids.next(), dur=1, fill="hold")
    s1 = _sub(c1, "p:stCondLst")
    _sub(s1, "p:cond", delay=0)
    target(cb)
    al = _sub(cb, "p:attrNameLst")
    an = _sub(al, "p:attrName")
    an.text = "style.visibility"
    to = _sub(s, "p:to")
    _sub(to, "p:strVal", val="visible")

    # the visible transition itself
    ae = _sub(ch, "p:animEffect", transition="in", filter=spec["filter"])
    cb2 = _sub(ae, "p:cBhvr")
    _sub(cb2, "p:cTn", id=ids.next(), dur=dur)
    target(cb2)


def build_timing(steps):
    """steps: list of clicks; each click is a list of effect dicts.

    An effect dict is {spid, effect, trigger, dur, delay, para}. The first
    effect of a click should use trigger "click"; later ones "with"/"after".
    """
    ids = _Ids(1)
    timing = _el("p:timing")
    tnLst = _sub(timing, "p:tnLst")
    par0 = _sub(tnLst, "p:par")
    ctn0 = _sub(par0, "p:cTn", id=ids.next(), dur="indefinite",
                restart="never", nodeType="tmRoot")
    ch0 = _sub(ctn0, "p:childTnLst")
    seq = _sub(ch0, "p:seq", concurrent=1, nextAc="seek")
    ctn1 = _sub(seq, "p:cTn", id=ids.next(), dur="indefinite",
                nodeType="mainSeq")
    ch1 = _sub(ctn1, "p:childTnLst")

    for click in steps:
        if not click:
            continue
        outer = _sub(ch1, "p:par")
        octn = _sub(outer, "p:cTn", id=ids.next(), fill="hold")
        ost = _sub(octn, "p:stCondLst")
        _sub(ost, "p:cond", delay="indefinite")
        och = _sub(octn, "p:childTnLst")
        inner = _sub(och, "p:par")
        ictn = _sub(inner, "p:cTn", id=ids.next(), fill="hold")
        ist = _sub(ictn, "p:stCondLst")
        _sub(ist, "p:cond", delay=0)
        ich = _sub(ictn, "p:childTnLst")
        for e in click:
            _effect(ich, ids, e["spid"], e.get("effect", "fade"),
                    e.get("trigger", "click"), e.get("dur", 500),
                    e.get("delay", 0), e.get("para"))

    prev = _sub(seq, "p:prevCondLst")
    c = _sub(prev, "p:cond", evt="onPrev", delay=0)
    t = _sub(c, "p:tgtEl")
    _sub(t, "p:sldTgt")
    nxt = _sub(seq, "p:nextCondLst")
    c2 = _sub(nxt, "p:cond", evt="onNext", delay=0)
    t2 = _sub(c2, "p:tgtEl")
    _sub(t2, "p:sldTgt")
    return timing


def animate(slide, steps):
    """Attach an animation sequence to a slide."""
    steps = [s for s in steps if s]
    if not steps:
        return
    sld = slide._element
    for old in sld.findall(_q("p:timing")):
        sld.remove(old)
    sld.append(build_timing(steps))


def transition(slide, kind="fade", speed="med"):
    """Slide transition. Kept to effects every PowerPoint build supports."""
    sld = slide._element
    for old in sld.findall(_q("p:transition")):
        sld.remove(old)
    tr = _el("p:transition", spd=speed)
    if kind == "fade":
        _sub(tr, "p:fade")
    elif kind == "push":
        _sub(tr, "p:push", dir="u")
    elif kind == "wipe":
        _sub(tr, "p:wipe", dir="d")
    elif kind == "cut":
        _sub(tr, "p:cut")
    else:
        _sub(tr, "p:fade")
    # p:transition must precede p:timing in the slide element
    timing = sld.find(_q("p:timing"))
    if timing is not None:
        timing.addprevious(tr)
    else:
        sld.append(tr)


def click(*effects):
    """Convenience: one click revealing the given effects together."""
    out = []
    for i, e in enumerate(effects):
        e = dict(e)
        e["trigger"] = "click" if i == 0 else "with"
        out.append(e)
    return out


def fx(shape, effect="fade", dur=500, delay=0, para=None, trigger="click"):
    return dict(spid=shape.shape_id, effect=effect, dur=dur, delay=delay,
                para=para, trigger=trigger)


def cascade(shape, n, effect="wipeUp", dur=400, gap=0):
    """Reveal a text shape one paragraph per click."""
    return [[dict(spid=shape.shape_id, effect=effect, trigger="click",
                  dur=dur, delay=gap, para=i)] for i in range(n)]
