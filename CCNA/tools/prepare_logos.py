#!/usr/bin/env python3
"""Turn the supplied institutional logos into usable assets.

The two files as delivered cannot be placed on anything but a white page:

* ``DIT Logo New.png`` has no transparency at all -- a white rectangle behind
  the shield -- so on a coloured slide it renders as a white box.
* ``IubLogo.png`` is transparent above the arch but white below it, and the
  mark itself is a single dark navy, which is invisible on a dark ground.

So this produces, from the originals in ``assets/src``:

    iub-logo.png         navy mark, background removed, trimmed
    iub-logo-white.png   the same mark knocked out to white, for dark grounds
    dit-logo.png         outer background removed, trimmed; the shield's own
                         white interior is part of the design and is kept

The outer background is removed by flood-filling inward from the border rather
than by keying out every white pixel, which would hollow out the DIT shield.

    python tools/prepare_logos.py
"""
from __future__ import annotations

import os
import sys
from collections import deque

try:
    from PIL import Image
except ImportError:                                    # pragma: no cover
    sys.exit("Pillow is required: pip install pillow")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(ROOT, "assets")
SRC = os.path.join(ASSETS, "src")

# How close to the background colour a pixel must be to be flooded away.
TOL = 34
# Alpha below this counts as already transparent.
CLEAR = 24


def _near(a, b, tol=TOL) -> bool:
    return all(abs(x - y) <= tol for x, y in zip(a[:3], b[:3]))


def strip_background(im: Image.Image) -> Image.Image:
    """Flood the outer background to transparent, from every border pixel.

    Keying out white everywhere would take the DIT shield's interior with it;
    the shield is white by design. Only background connected to the edge goes.
    """
    im = im.convert("RGBA")
    w, h = im.size
    px = im.load()
    # The background colour is whatever the corners agree on.
    corners = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
    opaque = [c for c in corners if c[3] > CLEAR]
    bg = opaque[0] if opaque else (255, 255, 255, 255)

    seen = bytearray(w * h)
    q: deque[tuple[int, int]] = deque()
    for x in range(w):
        for y in (0, h - 1):
            q.append((x, y))
    for y in range(h):
        for x in (0, w - 1):
            q.append((x, y))

    while q:
        x, y = q.popleft()
        if x < 0 or y < 0 or x >= w or y >= h or seen[y * w + x]:
            continue
        p = px[x, y]
        if p[3] > CLEAR and not _near(p, bg):
            continue
        seen[y * w + x] = 1
        px[x, y] = (p[0], p[1], p[2], 0)
        q.extend(((x + 1, y), (x - 1, y), (x, y + 1), (x, y - 1)))
    return im


def trim(im: Image.Image, pad: int = 2) -> Image.Image:
    """Crop to the visible mark, so two logos scale to the same optical size."""
    box = im.getchannel("A").point(lambda v: 255 if v > CLEAR else 0).getbbox()
    if not box:
        return im
    l, t, r, b = box
    w, h = im.size
    return im.crop((max(0, l - pad), max(0, t - pad),
                    min(w, r + pad), min(h, b + pad)))


def knockout(im: Image.Image, rgb=(255, 255, 255)) -> Image.Image:
    """Reverse a one-colour mark out of its ground, keeping its counters.

    Painting every opaque pixel white loses the shapes that were *left white*
    in the original -- on the IUB crest that is the calligraphy inside the
    dome, which filled in and vanished. So the ink is keyed by luminance
    instead: dark pixels become the new colour at full strength, light ones
    become transparent, and the dark ground shows through them exactly as the
    white page used to.

    Only sound for a single-colour mark. The DIT shield is navy, gold and red,
    and would lose its ribbon entirely.
    """
    out = im.convert("RGBA")
    px = out.load()
    w, h = out.size

    def lum(p):
        return (p[0] * 299 + p[1] * 587 + p[2] * 114) // 1000

    # The ink is navy, not black, so a plain (255 - luminance) leaves the whole
    # mark at about three-quarters alpha and the background shows through it.
    # Stretch the range so the darkest ink present reaches full opacity.
    floor = min((lum(px[x, y]) for y in range(h) for x in range(w)
                 if px[x, y][3] > 128), default=0)
    span = max(1, 255 - floor)

    for y in range(h):
        for x in range(w):
            p = px[x, y]
            if not p[3]:
                continue
            t = min(255, (255 - lum(p)) * 255 // span)
            px[x, y] = (*rgb, p[3] * t // 255)
    return out


JOBS = [
    ("IubLogo.png", "iub-logo.png", None),
    ("IubLogo.png", "iub-logo-white.png", (255, 255, 255)),
    ("DIT Logo New.png", "dit-logo.png", None),
]


def main() -> int:
    if not os.path.isdir(SRC):
        sys.exit(f"no source logos in {SRC}")
    for name, out, ko in JOBS:
        p = os.path.join(SRC, name)
        if not os.path.exists(p):
            print(f"  missing {name} -- skipped")
            continue
        im = trim(strip_background(Image.open(p)))
        if ko:
            im = knockout(im, ko)
        dest = os.path.join(ASSETS, out)
        im.save(dest)
        print(f"  {out:22} {im.size[0]}x{im.size[1]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
