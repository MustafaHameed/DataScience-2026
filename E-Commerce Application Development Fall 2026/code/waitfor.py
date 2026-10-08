"""Poll the storefront until it answers, or give up.

Every measurement in this course starts from a shop that is actually up, and
"docker compose up -d" returns long before that is true -- it returns when the
containers have been *started*, which on a first run is several minutes before
OpenCart has finished installing itself.  Measuring against a shop that is
still installing is the single easiest way to produce a number that means
nothing, so every harness calls wait_for_shop() first.

Run on its own it is also the answer to "is my shop up?":

    python waitfor.py
"""
import sys
import time
import urllib.error
import urllib.request

SHOP = "http://localhost:8090/"
ADMIN = "http://localhost:8090/admin/"
ADMINER = "http://localhost:8091/"


def probe(url, timeout=5):
    """Return (status, bytes, seconds) for one GET, or (None, 0, seconds)."""
    t0 = time.perf_counter()
    try:
        with urllib.request.urlopen(url, timeout=timeout) as r:
            body = r.read()
            return r.status, len(body), time.perf_counter() - t0
    except urllib.error.HTTPError as e:
        # A 4xx/5xx is still an answer: the server is up and talking.
        return e.code, 0, time.perf_counter() - t0
    except Exception:
        return None, 0, time.perf_counter() - t0


def wait_for_shop(url=SHOP, limit=600, quiet=False):
    """Block until `url` returns 200. Returns seconds waited, or raises."""
    t0 = time.perf_counter()
    last = None
    while time.perf_counter() - t0 < limit:
        status, size, _ = probe(url)
        if status == 200:
            return time.perf_counter() - t0
        if not quiet and status != last:
            waited = time.perf_counter() - t0
            print(f"  [{waited:6.1f}s] {url} -> {status}")
            last = status
        time.sleep(2)
    raise TimeoutError(f"{url} did not return 200 within {limit}s")


if __name__ == "__main__":
    ok = True
    for name, url in [("storefront", SHOP), ("admin", ADMIN),
                      ("adminer", ADMINER)]:
        status, size, secs = probe(url)
        mark = "ok  " if status == 200 else "DOWN"
        print(f"{mark} {name:<12} {url:<34} {status} "
              f"{size:>7} bytes  {secs * 1000:6.0f} ms")
        ok = ok and status == 200
    sys.exit(0 if ok else 1)
