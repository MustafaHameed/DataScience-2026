"""Chapter 8 -- what a page weighs, and which part of it is the weight.

A page is not one request. The browser fetches the HTML, reads it, and then
fetches every stylesheet, script and image it names -- and the customer waits
for all of it. "The page takes 100 ms" is a statement about the first request
only, and it is the least interesting number on this list.

What this measures, for any page:

  * the HTML, and how much of it is the page's own content;
  * every asset the HTML names, its size and how long it took;
  * the totals, split by kind, so you can see where the weight actually is.

    python m08_pageweight.py                     the category page
    python m08_pageweight.py --url <url>         any page of your shop
    python m08_pageweight.py --compare           home, category, product
"""
import argparse
import os
import re
import sys
import time
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from waitfor import wait_for_shop        # noqa: E402

SHOP = "http://localhost:8090/"
PAGES = {
    "home": SHOP,
    "category": SHOP + "index.php?route=product/category&path=20",
    "product": SHOP + "index.php?route=product/product&product_id=42",
}

# src/href attributes the browser would follow for a *render-blocking or
# displayed* asset. Deliberately not everything: a preload hint or an
# anchor is not weight the customer waits for.
ASSET = re.compile(
    r'<img[^>]+src="([^"]+)"'
    r'|<link[^>]+rel="stylesheet"[^>]+href="([^"]+)"'
    r'|<link[^>]+href="([^"]+)"[^>]+rel="stylesheet"'
    r'|<script[^>]+src="([^"]+)"',
    re.I)


def fetch(url: str) -> tuple[int, float, str]:
    """(bytes, ms, content-type) for one request."""
    t0 = time.perf_counter()
    try:
        with urllib.request.urlopen(url, timeout=60) as r:
            body = r.read()
            return len(body), (time.perf_counter() - t0) * 1000, \
                r.headers.get("Content-Type", "")
    except Exception:
        return 0, (time.perf_counter() - t0) * 1000, "error"


def kind(url: str, ctype: str) -> str:
    u = url.lower().split("?")[0]
    if u.endswith((".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".ico")):
        return "image"
    if u.endswith(".css") or "css" in ctype:
        return "css"
    if u.endswith(".js") or "javascript" in ctype:
        return "js"
    if u.endswith((".woff", ".woff2", ".ttf", ".eot")):
        return "font"
    return "other"


def weigh(url: str, verbose: bool = True) -> dict:
    html_bytes, html_ms, _ = fetch(url)
    with urllib.request.urlopen(url, timeout=60) as r:
        html = r.read().decode("utf-8", "replace")

    # One asset may be named many times -- the same thumbnail on twenty
    # product cards is ONE request, because the browser caches it within the
    # page. Counting it twenty times would overstate the weight fivefold and
    # is the most common error in a hand-rolled page-weight script.
    seen, assets = set(), []
    for m in ASSET.finditer(html):
        raw = next(g for g in m.groups() if g)
        full = urllib.parse.urljoin(url, raw)
        if full in seen:
            continue
        seen.add(full)
        assets.append(full)

    totals = {"html": [1, html_bytes, html_ms]}
    for a in assets:
        n, ms, ct = fetch(a)
        k = kind(a, ct)
        row = totals.setdefault(k, [0, 0, 0.0])
        row[0] += 1
        row[1] += n
        row[2] += ms

    if verbose:
        print(f"\n{url}")
        print(f"  {'kind':<8} {'requests':>9} {'KB':>10} {'ms':>9}")
        for k in ("html", "css", "js", "image", "font", "other"):
            if k in totals:
                c, b, ms = totals[k]
                print(f"  {k:<8} {c:>9} {b / 1024:>10.1f} {ms:>9.0f}")
        c = sum(v[0] for v in totals.values())
        b = sum(v[1] for v in totals.values())
        ms = sum(v[2] for v in totals.values())
        print(f"  {'TOTAL':<8} {c:>9} {b / 1024:>10.1f} {ms:>9.0f}")
        if "image" in totals:
            print(f"  images are {totals['image'][1] / b * 100:.0f}% of the "
                  f"bytes and {totals['image'][0] / c * 100:.0f}% of the "
                  f"requests")
    return totals


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--url")
    ap.add_argument("--compare", action="store_true")
    a = ap.parse_args()
    wait_for_shop(quiet=True)
    if a.compare:
        for name, u in PAGES.items():
            print(f"\n=== {name} ===", end="")
            weigh(u)
    else:
        weigh(a.url or PAGES["category"])
