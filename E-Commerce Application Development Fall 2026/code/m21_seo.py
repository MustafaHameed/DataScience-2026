"""Chapter 21 -- what a search engine would find, measured rather than assumed.

Search engine optimisation is the part of this course with the worst
signal-to-noise ratio in the literature, because almost none of it is
checkable. These things are:

  * whether your URLs are readable, or carry a query string;
  * whether every page has a title and a description, and whether they are
    DIFFERENT from each other;
  * whether a crawler is told which URL is canonical;
  * whether there is a sitemap at all;
  * how long the server takes to send the first byte.

Nothing here is about ranking. It is about whether the pages can be read,
distinguished and found, which is the part you control.

    python m21_seo.py              all of it
    python m21_seo.py --urls       SEO URLs off, then on
    python m21_seo.py --meta       titles and descriptions across the catalogue
    python m21_seo.py --crawl      robots.txt, sitemap, canonical, first byte
    python m21_seo.py --restore    put config_seo_url back as it was
"""
import argparse
import os
import re
import sys
import time
import urllib.error
import urllib.request
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402

SHOP = "http://localhost:8090/"
TITLE = re.compile(r"<title>(.*?)</title>", re.S | re.I)
DESC = re.compile(r'<meta\s+name="description"\s+content="(.*?)"', re.S | re.I)
CANON = re.compile(r'<link[^>]+rel="canonical"[^>]+href="([^"]+)"', re.I)
PRODUCT_LINK = re.compile(r'href="([^"]*product_id=\d+[^"]*)"')


def sql(stmt: str) -> str:
    return S.sql(stmt)


def one(stmt: str, default: str = "0") -> str:
    out = sql(stmt).strip().splitlines()
    if len(out) < 2:
        return default
    v = out[1].split("\t")[0]
    return default if v in ("NULL", "") else v


def get(url: str) -> tuple[int, str, float]:
    """(status, body, milliseconds-to-first-byte)."""
    t0 = time.perf_counter()
    try:
        with urllib.request.urlopen(url, timeout=60) as r:
            first = (time.perf_counter() - t0) * 1000
            return r.status, r.read().decode("utf-8", "replace"), first
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace"), 0.0
    except OSError:
        return 0, "", 0.0


def seo_url(on: bool) -> None:
    sql(f"UPDATE oc_setting SET value={1 if on else 0} "
        f"WHERE `key`='config_seo_url';")


def rewrite(on: bool) -> None:
    """Apply (or remove) the server half of readable URLs.

    A server config, not a .htaccess: Chapter 16 measured AllowOverride None,
    which means .htaccess is never read on this stack.
    """
    import subprocess
    here = os.path.dirname(os.path.abspath(__file__))
    if on:
        dest = "shop:/etc/apache2/conf-available/zz-seourl-ecad.conf"
        subprocess.run(["docker", "compose", "cp", "shop/seourl.conf", dest],
                       cwd=here, capture_output=True)
        cmd = ("a2enmod rewrite >/dev/null 2>&1; "
               "a2enconf zz-seourl-ecad >/dev/null 2>&1; "
               "apachectl -k graceful")
    else:
        cmd = "a2disconf zz-seourl-ecad >/dev/null 2>&1; apachectl -k graceful"
    subprocess.run(["docker", "compose", "exec", "-T", "shop", "sh", "-lc",
                    cmd], cwd=here, capture_output=True)
    wait_for_shop(quiet=True)


def fix() -> None:
    """Measure the broken state, apply both halves, measure again."""
    print("\n=== Readable URLs: the setting alone, then both halves ===")
    category = SHOP + "index.php?route=product/category&path=20"

    def probe(label: str) -> None:
        _c, body, _ms = get(category)
        links = [u for u in re.findall(
            r'href="(http://localhost:8090/[^"]*)"', body)
            if "/image/" not in u and u.rstrip("/") != SHOP.rstrip("/")]
        link = links[0].replace("&amp;", "&") if links else "(none)"
        code, _b, _ms = get(link)
        print(f"  {label:<34}{str('index.php' not in link):>10}"
              f"{code:>8}   {link[:46]}")

    print(f"  {'':<34}{'readable':>10}{'HTTP':>8}")
    seo_url(False)
    rewrite(False)
    probe("setting off, rewrite off")
    seo_url(True)
    probe("setting ON, rewrite off")
    rewrite(True)
    probe("setting ON, rewrite ON")
    print("\n  Neither half is any use alone. The setting without the")
    print("  rewrite is worse than nothing: it points every link in the")
    print("  shop at a page that does not exist.")


def product_ids(limit: int = 19) -> list[int]:
    out = sql(f"SELECT product_id FROM oc_product WHERE status=1 "
              f"ORDER BY product_id LIMIT {limit};")
    return [int(ln.split("\t")[0]) for ln in out.strip().splitlines()[1:]]


# ------------------------------------------------------------------ 1. URLs
def urls() -> None:
    print("\n=== 1. What the URLs look like ===")
    rows = one("SELECT COUNT(*) FROM oc_seo_url;")
    setting = one("SELECT value FROM oc_setting "
                  "WHERE `key`='config_seo_url';")
    print(f"  rows in oc_seo_url: {rows}")
    print(f"  config_seo_url:     {setting}")

    category = SHOP + "index.php?route=product/category&path=20"
    for state in (0, 1):
        seo_url(bool(state))
        _c, body, _ms = get(category)
        # The links the category page itself emits, not the assets.
        links = [u for u in re.findall(r'href="(http://localhost:8090/[^"]*)"',
                                       body)
                 if "/image/" not in u and u.rstrip("/") != SHOP.rstrip("/")]
        sample = links[0].replace("&amp;", "&") if links else "(none)"
        readable = "index.php?route=" not in sample
        print(f"\n  config_seo_url={state}")
        print(f"    a link the page emits: {sample[:66]}")
        print(f"    readable:              {readable}")
        if readable and links:
            code, _b, _ms = get(sample)
            print(f"    and fetching it gives: HTTP {code}"
                  f"{'   <-- every link on the site' if code == 404 else ''}")

    # The application half, reached directly, bypassing the rewrite.
    code, _b, _ms = get(SHOP + "index.php?_route_=htc-touch-hd")
    print(f"\n  the same page via index.php?_route_=...   HTTP {code}")

    print(f"\n  The shop ships {rows} prepared keywords and the setting that")
    print("  uses them turned OFF. Turning it on rewrites every link into a")
    print("  readable one -- and every one of them 404s, because the rewrite")
    print("  from /en-gb/catalog/desktops to index.php?_route_=... lives in")
    print("  .htaccess, which Chapter 16 measured Apache as never reading.")
    print("  The application half works perfectly; the server half was never")
    print("  switched on. ONE misconfiguration, two chapters apart.")


# ------------------------------------------------------------------ 2. meta
def meta() -> None:
    print("\n=== 2. Titles and descriptions across the catalogue ===")
    ids = product_ids()
    titles, descs, missing_desc = [], [], 0
    for pid in ids:
        _c, body, _ms = get(
            SHOP + f"index.php?route=product/product&product_id={pid}")
        t = TITLE.search(body)
        d = DESC.search(body)
        titles.append((t.group(1).strip() if t else "")[:90])
        if d and d.group(1).strip():
            descs.append(d.group(1).strip()[:120])
        else:
            missing_desc += 1

    dupt = [v for v, n in Counter(titles).items() if n > 1]
    dupd = [v for v, n in Counter(descs).items() if n > 1]
    print(f"  {'product pages checked':<34}{len(ids):>6}")
    print(f"  {'distinct titles':<34}{len(set(titles)):>6}")
    print(f"  {'duplicate titles':<34}{len(dupt):>6}")
    print(f"  {'pages with no description':<34}{missing_desc:>6}")
    print(f"  {'distinct descriptions':<34}{len(set(descs)):>6}")
    print(f"  {'duplicate descriptions':<34}{len(dupd):>6}")
    if titles:
        print(f"\n  a title: {titles[0]}")
    for group, label in ((dupt, "titles"), (dupd, "descriptions")):
        for v in group[:2]:
            print(f"  duplicate {label}: {v[:70]}")
    print("\n  A duplicate title is two pages a crawler cannot tell apart,")
    print("  and a missing description is one it will write for you.")


# ----------------------------------------------------------------- 3. crawl
def crawl() -> None:
    print("\n=== 3. What a crawler is told ===")
    code, body, _ms = get(SHOP + "robots.txt")
    rules = [ln for ln in body.splitlines() if ln.strip().lower()
             .startswith("disallow")]
    print(f"  robots.txt            HTTP {code}, {len(rules)} Disallow rules")
    has_sitemap = "sitemap" in body.lower()
    print(f"  names a sitemap?      {has_sitemap}")

    for path in ("sitemap.xml",
                 "index.php?route=extension/opencart/feed/google_sitemap"):
        code, body, _ms = get(SHOP + path)
        print(f"  {path[:46]:<46} HTTP {code}")

    print(f"\n  {'page':<12}{'canonical?':>12}{'first byte':>13}")
    for label, url in (("home", SHOP),
                       ("category", SHOP + "index.php?route=product/category"
                                           "&path=20"),
                       ("product", SHOP + "index.php?route=product/product"
                                          "&product_id=43")):
        samples = []
        canon = None
        for _ in range(5):
            _c, body, ms = get(url)
            samples.append(ms)
            canon = canon or CANON.search(body)
        samples.sort()
        print(f"  {label:<12}{str(bool(canon)):>12}"
              f"{samples[len(samples) // 2]:>11.0f} ms")

    print("\n  No sitemap of any kind ships with this platform, and")
    print("  robots.txt does not name one. A crawler finds pages only by")
    print("  following links, so anything not linked from the home page")
    print("  within a few hops is, in practice, not there.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--urls", action="store_true")
    ap.add_argument("--meta", action="store_true")
    ap.add_argument("--crawl", action="store_true")
    ap.add_argument("--fix", action="store_true")
    ap.add_argument("--restore", action="store_true")
    a = ap.parse_args()
    picked = any((a.urls, a.meta, a.crawl, a.fix))

    wait_for_shop(quiet=True)
    if a.restore:
        seo_url(False)
        rewrite(False)
        print("config_seo_url 0 and the rewrite off, as shipped")
        return 0
    if a.fix:
        fix()
        return 0
    if a.urls or not picked:
        urls()
    if a.meta or not picked:
        meta()
    if a.crawl or not picked:
        crawl()
    seo_url(False)
    return 0


if __name__ == "__main__":
    sys.exit(main())
