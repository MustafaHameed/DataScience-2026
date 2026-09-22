#!/usr/bin/env python3
"""PowerPoint features python-pptx does not expose, grafted on as XML.

* **Entrance animations** -- the <p:timing> tree: one click per build step,
  effects in a step as siblings (first ``clickEffect``, rest ``withEffect``),
  optionally one paragraph of a text box at a time. Get the nesting wrong and
  PowerPoint silently drops every animation on the slide, so keep the shape of
  ``_timing()``. (Descends from the ARM decks' pptxfx.)
* **Transitions** -- fade, push, wipe, and **Morph** (with a fade fallback for
  PowerPoint before 2016). Morph pairs shapes across slides by name when the
  name starts with ``!!``; the deck names every shape that is meant to travel
  that way.
* **SVG pictures** -- a PNG blip with an ``asvg:svgBlip`` extension, which is
  exactly how PowerPoint itself stores an inserted SVG: 2016+ draws the vector,
  anything older draws the PNG.
* **Accessibility** -- alt text (``descr``) and the *Mark as decorative* flag,
  so a screen reader announces the figure and skips the ornament.
* **Sections**, **slide hyperlinks**, a live **slide-number field**, and the
  theme's **fonts and colours**.
"""
from __future__ import annotations

import copy
import re
import uuid

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
SVG = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
ADEC = "http://schemas.microsoft.com/office/drawing/2017/decorative"
NS = {"p": P, "a": A, "r": R, "p14": P14, "mc": MC}


def q(tag: str) -> str:
    pre, local = tag.split(":")
    return "{%s}%s" % ({"p": P, "a": A, "r": R, "p14": P14, "mc": MC,
                        "asvg": SVG, "adec": ADEC}[pre], local)


def _sub(parent, tag, **attrs):
    node = etree.SubElement(parent, q(tag))
    for k, v in attrs.items():
        node.set(k, str(v))
    return node


# ---------------------------------------------------------------------------
# Entrance animations
# ---------------------------------------------------------------------------

EFFECTS = {
    "appear":    dict(presetID=1,  subtype=0,  filter=None),
    "fade":      dict(presetID=10, subtype=0,  filter="fade"),
    "wipeUp":    dict(presetID=22, subtype=4,  filter="wipe(up)"),
    "wipeRight": dict(presetID=22, subtype=8,  filter="wipe(right)"),
    "wipeDown":  dict(presetID=22, subtype=1,  filter="wipe(down)"),
    "flyBottom": dict(presetID=2,  subtype=4,  filter="slide(fromBottom)"),
    "flyLeft":   dict(presetID=2,  subtype=8,  filter="slide(fromLeft)"),
    "zoom":      dict(presetID=23, subtype=16, filter="zoom(in)"),
    "split":     dict(presetID=13, subtype=26, filter="barn(inVertical)"),
}
_NODE = {"click": "clickEffect", "with": "withEffect", "after": "afterEffect"}


class _Ids:
    def __init__(self):
        self.n = 1

    def next(self):
        self.n += 1
        return self.n


def _effect(parent, ids, e):
    spec = EFFECTS.get(e.get("effect", "fade"), EFFECTS["fade"])
    par = _sub(parent, "p:par")
    ctn = _sub(par, "p:cTn", id=ids.next(), presetID=spec["presetID"],
               presetClass="entr", presetSubtype=spec["subtype"],
               fill="hold", grpId=0, nodeType=_NODE[e.get("trigger",
                                                         "click")])
    st = _sub(ctn, "p:stCondLst")
    _sub(st, "p:cond", delay=e.get("delay", 0))
    ch = _sub(ctn, "p:childTnLst")

    def target(el):
        tgt = _sub(el, "p:tgtEl")
        sp = _sub(tgt, "p:spTgt", spid=e["spid"])
        if e.get("para") is not None:
            tx = _sub(sp, "p:txEl")
            _sub(tx, "p:pRg", st=e["para"], end=e.get("para_end", e["para"]))

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
    if spec["filter"]:
        ae = _sub(ch, "p:animEffect", transition="in", filter=spec["filter"])
        cb2 = _sub(ae, "p:cBhvr")
        _sub(cb2, "p:cTn", id=ids.next(), dur=e.get("dur", 400))
        target(cb2)


def _timing(steps):
    ids = _Ids()
    timing = etree.Element(q("p:timing"))
    tn = _sub(timing, "p:tnLst")
    par0 = _sub(tn, "p:par")
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
        for i, e in enumerate(click):
            e = dict(e)
            e.setdefault("trigger", "click" if i == 0 else "with")
            _effect(ich, ids, e)
    prev = _sub(seq, "p:prevCondLst")
    c = _sub(prev, "p:cond", evt="onPrev", delay=0)
    _sub(_sub(c, "p:tgtEl"), "p:sldTgt")
    nxt = _sub(seq, "p:nextCondLst")
    c2 = _sub(nxt, "p:cond", evt="onNext", delay=0)
    _sub(_sub(c2, "p:tgtEl"), "p:sldTgt")
    # build list: text shapes animated by paragraph must declare it
    bld = _sub(timing, "p:bldLst")
    seen = set()
    for click in steps:
        for e in click or []:
            if e.get("para") is not None and e["spid"] not in seen:
                seen.add(e["spid"])
                _sub(bld, "p:bldP", spid=e["spid"], grpId=0, build="p")
    if not len(bld):
        timing.remove(bld)
    return timing


def fx(shape, effect="fade", dur=400, delay=0, para=None, trigger=None):
    d = dict(spid=shape.shape_id, effect=effect, dur=dur, delay=delay,
             para=para)
    if trigger:
        d["trigger"] = trigger
    return d


def animate(slide, steps):
    """Attach click steps: a list of lists of fx() dicts."""
    steps = [s for s in steps if s]
    sld = slide._element
    for old in sld.findall(q("p:timing")):
        sld.remove(old)
    if steps:
        _place(sld, _timing(steps), "timing")


# ---------------------------------------------------------------------------
# Transitions
# ---------------------------------------------------------------------------

def _place(sld, node, kind):
    """Insert transition/timing in schema order: cSld, clrMapOvr,
    transition, timing, extLst."""
    after = [sld.find(q("p:cSld")), sld.find(q("p:clrMapOvr"))]
    if kind == "timing":
        tr = sld.find(q("mc:AlternateContent"))
        if tr is None:
            tr = sld.find(q("p:transition"))
        after.append(tr)
    anchor = [a for a in after if a is not None][-1]
    anchor.addnext(node)


def transition(slide, kind="fade", ms=450):
    sld = slide._element
    for old in sld.findall(q("p:transition")):
        sld.remove(old)
    for old in sld.findall(q("mc:AlternateContent")):
        sld.remove(old)
    if kind == "none":
        return
    if kind == "morph":
        xml = (f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:p="{P}" '
               f'xmlns:p14="{P14}"><mc:Choice xmlns:p159="{P159}" '
               f'Requires="p159"><p:transition spd="slow" p14:dur="{ms}">'
               f'<p159:morph option="byObject"/></p:transition></mc:Choice>'
               f'<mc:Fallback><p:transition spd="med"><p:fade/>'
               f'</p:transition></mc:Fallback></mc:AlternateContent>')
    else:
        inner = {"fade": "<p:fade/>", "push": '<p:push dir="u"/>',
                 "pushLeft": '<p:push dir="l"/>', "wipe": '<p:wipe dir="r"/>',
                 "cut": "<p:cut/>"}.get(kind, "<p:fade/>")
        xml = (f'<mc:AlternateContent xmlns:mc="{MC}" xmlns:p="{P}" '
               f'xmlns:p14="{P14}"><mc:Choice Requires="p14">'
               f'<p:transition spd="med" p14:dur="{ms}">{inner}'
               f'</p:transition></mc:Choice><mc:Fallback>'
               f'<p:transition spd="med">{inner}</p:transition>'
               f'</mc:Fallback></mc:AlternateContent>')
    node = etree.fromstring(xml)
    timing = sld.find(q("p:timing"))
    if timing is not None:
        timing.addprevious(node)
    else:
        _place(sld, node, "transition")


# ---------------------------------------------------------------------------
# Pictures: SVG, alt text, decorative
# ---------------------------------------------------------------------------

def add_svg(slide, pic, svg_path: str) -> None:
    """Give a picture a vector twin; PowerPoint 2016+ draws the SVG."""
    pkg = slide.part.package
    partname = pkg.next_partname("/ppt/media/image%d.svg")
    with open(svg_path, "rb") as fh:
        part = Part(partname, "image/svg+xml", pkg, fh.read())
    rid = slide.part.relate_to(part, RT.IMAGE)
    blip = pic._element.find(".//" + q("a:blip"))
    lst = blip.find(q("a:extLst"))
    if lst is None:
        lst = _sub(blip, "a:extLst")
    ext = _sub(lst, "a:ext", uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}")
    sb = etree.SubElement(ext, q("asvg:svgBlip"), nsmap={"asvg": SVG})
    sb.set("{%s}embed" % R, rid)


def _cnvpr(shape):
    el = shape._element
    for path in (".//p:cNvPr",):
        found = el.find(path, NS)
        if found is not None:
            return found
    return None


def alt_text(shape, text: str) -> None:
    c = _cnvpr(shape)
    if c is not None and text:
        c.set("descr", re.sub(r"\s+", " ", text).strip()[:1500])


def decorative(shape) -> None:
    """Mark a shape decorative (Office 2019+): screen readers skip it."""
    c = _cnvpr(shape)
    if c is None:
        return
    lst = c.find(q("a:extLst"))
    if lst is None:
        lst = _sub(c, "a:extLst")
    ext = _sub(lst, "a:ext", uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}")
    d = etree.SubElement(ext, q("adec:decorative"), nsmap={"adec": ADEC})
    d.set("val", "1")


# ---------------------------------------------------------------------------
# Shape styling helpers
# ---------------------------------------------------------------------------

def shadow(shape, blur_pt=14, dist_pt=3, alpha=16, color="1C252B"):
    """A soft drop shadow, below and slightly right."""
    spPr = shape._element.find(q("p:spPr"))
    if spPr is None:
        return
    for old in spPr.findall(q("a:effectLst")):
        spPr.remove(old)
    eff = etree.SubElement(spPr, q("a:effectLst"))
    sh = _sub(eff, "a:outerShdw", blurRad=int(blur_pt * 12700),
              dist=int(dist_pt * 12700), dir=5400000, algn="t",
              rotWithShape=0)
    clr = _sub(sh, "a:srgbClr", val=color)
    _sub(clr, "a:alpha", val=alpha * 1000)
    # effectLst must follow a:ln in spPr
    ln = spPr.find(q("a:ln"))
    if ln is not None:
        ln.addnext(eff)


def letter_spacing(run, hundredths_pt: int) -> None:
    run._r.get_or_add_rPr().set("spc", str(hundredths_pt))


def baseline(run, script: str | None) -> None:
    if script:
        run._r.get_or_add_rPr().set("baseline",
                                    "-25000" if script == "sub" else "30000")


def fill_alpha(shape, alpha_pct: int) -> None:
    """Transparency on a solid fill (100 = opaque)."""
    sf = shape._element.find(".//" + q("a:solidFill"))
    if sf is None:
        return
    clr = sf[0]
    for old in clr.findall(q("a:alpha")):
        clr.remove(old)
    _sub(clr, "a:alpha", val=alpha_pct * 1000)


def no_autofit(text_frame) -> None:
    bp = text_frame._txBody.find(q("a:bodyPr"))
    for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
        for old in bp.findall(q(tag)):
            bp.remove(old)
    _sub(bp, "a:noAutofit")


# ---------------------------------------------------------------------------
# Fields, links, sections
# ---------------------------------------------------------------------------

def slide_number_field(paragraph, size_pt=10, color="7B8A97",
                       font="Segoe UI", placeholder="1"):
    """Append a live slide-number field: it renumbers when slides move."""
    p = paragraph._p
    fld = _sub(p, "a:fld", id="{%s}" % str(uuid.uuid4()).upper(),
               type="slidenum")
    rpr = _sub(fld, "a:rPr", lang="en-GB", sz=int(size_pt * 100))
    sf = _sub(rpr, "a:solidFill")
    _sub(sf, "a:srgbClr", val=color)
    _sub(rpr, "a:latin", typeface=font)
    t = _sub(fld, "a:t")
    t.text = placeholder
    # endParaRPr must stay last
    end = p.find(q("a:endParaRPr"))
    if end is not None:
        p.remove(end)
        p.append(end)


def to_back(shape) -> None:
    """Send a shape to the bottom of the z-order (behind the title
    placeholder, which every slide carries from its layout)."""
    el = shape._element
    tree = el.getparent()
    tree.remove(el)
    # spTree starts with nvGrpSpPr and grpSpPr; shapes follow
    tree.insert(2, el)


def link_to_slide(shape, target_slide) -> None:
    shape.click_action.target_slide = target_slide


def sections(prs, groups) -> None:
    """groups: [(name, [slide, ...])] -> PowerPoint sections."""
    pres = prs.part._element
    ext_lst = pres.find(q("p:extLst"))
    if ext_lst is None:
        ext_lst = _sub(pres, "p:extLst")
    uri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
    for e in ext_lst.findall(q("p:ext")):
        if e.get("uri") == uri:
            ext_lst.remove(e)
    ext = _sub(ext_lst, "p:ext", uri=uri)
    lst = etree.SubElement(ext, q("p14:sectionLst"), nsmap={"p14": P14})
    for name, slides in groups:
        if not slides:
            continue
        sec = _sub(lst, "p14:section", name=name[:60],
                   id="{%s}" % str(uuid.uuid4()).upper())
        ids = _sub(sec, "p14:sldIdLst")
        for s in slides:
            _sub(ids, "p14:sldId", id=s.slide_id)
    ext_lst.getparent().remove(ext_lst)
    pres.append(ext_lst)


# ---------------------------------------------------------------------------
# Theme and master
# ---------------------------------------------------------------------------

def set_theme(prs, colours: dict, major="Segoe UI Semibold",
              minor="Segoe UI", name="Data Science Fall 2026") -> None:
    """Write the course palette and faces into the theme, so a slide a
    colleague adds by hand already matches."""
    master = prs.slide_master
    theme_part = master.part.part_related_by(RT.THEME)
    root = etree.fromstring(theme_part.blob)
    root.set("name", name)
    scheme = root.find(".//" + q("a:clrScheme"))
    scheme.set("name", name)
    for key, hexv in colours.items():
        slot = scheme.find(q("a:" + key))
        if slot is None:
            continue
        for ch in list(slot):
            slot.remove(ch)
        _sub(slot, "a:srgbClr", val=hexv)
    fs = root.find(".//" + q("a:fontScheme"))
    fs.set("name", name)
    for tag, face in (("a:majorFont", major), ("a:minorFont", minor)):
        lat = fs.find(q(tag)).find(q("a:latin"))
        lat.set("typeface", face)
        lat.attrib.pop("panose", None)
    theme_part._blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)


def set_title_style(prs, color="1A5276", size_pt=30) -> None:
    """Left-aligned, theme heading face, course colour -- master-wide."""
    master = prs.slide_master._element
    ts = master.find(".//" + q("p:titleStyle"))
    lvl = ts.find(q("a:lvl1pPr"))
    lvl.set("algn", "l")
    d = lvl.find(q("a:defRPr"))
    d.set("sz", str(size_pt * 100))
    d.set("b", "0")
    for old in d.findall(q("a:solidFill")):
        d.remove(old)
    sf = etree.Element(q("a:solidFill"))
    _sub(sf, "a:srgbClr", val=color)
    d.insert(0, sf)
