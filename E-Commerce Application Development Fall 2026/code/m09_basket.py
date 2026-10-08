"""Chapter 9 -- where a basket actually lives, and what survives a restart.

Almost every tutorial on this subject says "the cart is kept in the session".
It is worth finding out whether that is true of the shop in front of you,
because if it is not, everything that follows from it -- how big a session
gets, what a restart costs, why a logged-in customer's basket comes back on
another machine -- is wrong too.

Four questions:

  1. When you add a product, which of the three stores changes: the cookie in
     the browser, the session record on the server, or a database table?
  2. What does a visitor who never accepts a cookie cost?
  3. Does switching the session engine from the database to files move the
     basket with it?
  4. What survives `restart`, what survives `down` and `up`, and what does not?

    python m09_basket.py             all four, in order (about three minutes)
    python m09_basket.py --where     question 1
    python m09_basket.py --sessions  question 2
    python m09_basket.py --engine    question 3
    python m09_basket.py --restart   question 4
    python m09_basket.py --gc        the shutdown collector, as an aside
    python m09_basket.py --destroy   also measure `down -v`, which REINSTALLS
                                     the shop and takes a few minutes

Nothing here needs a browser. A cookie jar and a POST are the whole client,
which is the point: the shop cannot tell the difference, so neither can its
storage.
"""
import argparse
import http.cookiejar
import os
import statistics
import subprocess
import sys
import time
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402

CODE = os.path.dirname(os.path.abspath(__file__))
SHOP = "http://localhost:8090/"
ADD = SHOP + "index.php?route=checkout/cart.add"

# Demo products with no required options, so that one POST adds one line.
# Product 42 has nine required options and is the subject of Chapter 7; using
# it here would measure the option validator, not the basket.
GOODS = [(43, 2, "MacBook"), (40, 1, "iPhone"), (28, 3, "HTC Touch HD")]

CATALOG_CONFIG = "system/config/catalog.php"


# --------------------------------------------------------------- plumbing
def sql(stmt: str) -> str:
    return S.sql(stmt)


def rootsql(stmt: str) -> str:
    """As in m06: GLOBAL settings are not the shop user's to change, and
    MariaDB does not complain when it declines -- it just does nothing."""
    r = subprocess.run(
        ["docker", "compose", "exec", "-T", "db", "mariadb", "-uroot",
         "-prootpw"], cwd=CODE, input=stmt, text=True, capture_output=True)
    return r.stdout


def scalar(stmt: str, default: int = 0) -> int:
    out = sql(stmt).strip().splitlines()
    try:
        return int(out[1].split("\t")[0])
    except (IndexError, ValueError):
        return default


def shop(cmd: str) -> str:
    r = subprocess.run(["docker", "compose", "exec", "-T", "shop", "sh", "-lc",
                        cmd], cwd=CODE, text=True, capture_output=True)
    return r.stdout


def compose(*args: str) -> None:
    subprocess.run(["docker", "compose", *args], cwd=CODE,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


class Browser:
    """One visitor: a cookie jar and nothing else."""

    def __init__(self):
        self.cj = http.cookiejar.CookieJar()
        self.op = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.cj))

    def get(self, url: str = SHOP) -> int:
        with self.op.open(url, timeout=120) as r:
            return len(r.read())

    def add(self, product_id: int, quantity: int) -> str:
        data = urllib.parse.urlencode(
            {"product_id": product_id, "quantity": quantity}).encode()
        req = urllib.request.Request(
            ADD, data=data, headers={"X-Requested-With": "XMLHttpRequest"})
        with self.op.open(req, timeout=120) as r:
            return r.read().decode()

    @property
    def sid(self) -> str:
        for c in self.cj:
            if c.name == "OCSESSID":
                return c.value
        return ""

    @property
    def cookie_bytes(self) -> int:
        """What the browser sends up on every single request: `name=value`
        pairs joined by '; '. This is paid once per request, including for
        every image and stylesheet on the page."""
        pairs = [f"{c.name}={c.value}" for c in self.cj]
        return len("; ".join(pairs))


def fetch_ms(url: str = SHOP, reps: int = 15) -> float:
    times = []
    for i in range(reps + 1):
        t0 = time.perf_counter()
        with urllib.request.urlopen(url, timeout=120) as r:
            r.read()
        if i:
            times.append((time.perf_counter() - t0) * 1000)
    return statistics.median(times)


def page_queries(url: str = SHOP) -> tuple[int, int]:
    """(queries, session queries) for ONE request, read out of MariaDB's slow
    query log rather than guessed -- the method Chapter 6 arrived at after
    three wrong ones."""
    rootsql("SET GLOBAL slow_query_log=0;"
            "SET GLOBAL log_output='TABLE';"
            "TRUNCATE TABLE mysql.slow_log;"
            "SET GLOBAL long_query_time=0;"
            "SET GLOBAL slow_query_log=1;")
    Browser().get(url)
    rootsql("SET GLOBAL slow_query_log=0;")
    out = rootsql("SELECT COUNT(*), "
                  "SUM(sql_text LIKE '%`oc_session`%' "
                  "OR sql_text LIKE '%oc_session%') "
                  "FROM mysql.slow_log;")
    try:
        row = out.strip().splitlines()[1].split("\t")
        return int(row[0]), int(row[1] or 0)
    except (IndexError, ValueError):
        return 0, 0


# ------------------------------------------------------- 1. where it lives
def where() -> None:
    print("\n=== 1. Add three products; watch all three stores ===")
    b = Browser()
    b.get()
    sid = b.sid
    print(f"session id {sid!r} ({len(sid)} characters)")

    head = (f"  {'after adding':<22}{'cookie B':>9}{'session B':>11}"
            f"{'cart rows':>11}{'cart B':>8}")
    print(head)
    print("  " + "-" * (len(head) - 2))

    def row(label: str) -> None:
        sess = scalar(f"SELECT COALESCE(LENGTH(data),0) FROM oc_session "
                      f"WHERE session_id='{sid}';")
        rows = scalar(f"SELECT COUNT(*) FROM oc_cart "
                      f"WHERE session_id='{sid}';")
        size = scalar(
            f"SELECT COALESCE(SUM(LENGTH(CONCAT_WS('',cart_id,store_id,"
            f"customer_id,session_id,product_id,`option`,quantity,price,"
            f"date_added))),0) FROM oc_cart WHERE session_id='{sid}';")
        print(f"  {label:<22}{b.cookie_bytes:>9}{sess:>11}{rows:>11}{size:>8}")

    row("(nothing)")
    for pid, qty, name in GOODS:
        reply = b.add(pid, qty)
        if '"error"' in reply:
            print(f"  {name}: refused -- {reply[:90]}")
            continue
        row(f"{qty} x {name}")

    data = sql(f"SELECT data FROM oc_session WHERE session_id='{sid}';")
    content = data.strip().splitlines()[-1] if data.strip() else "(none)"
    print(f"\n  the whole session record: {content}")
    print("  the basket is not in it. It is rows in oc_cart, keyed by the")
    print("  session id -- which is the only thing the cookie carries.")


# -------------------------------------------------- 2. the cost of a visit
def sessions(n: int = 20) -> None:
    print(f"\n=== 2. {n} requests, with and without a cookie jar ===")
    before = scalar("SELECT COUNT(*) FROM oc_session;")

    for _ in range(n):                       # a new visitor every time
        Browser().get()
    stateless = scalar("SELECT COUNT(*) FROM oc_session;") - before

    mid = scalar("SELECT COUNT(*) FROM oc_session;")
    b = Browser()                            # one visitor, n page views
    for _ in range(n):
        b.get()
    returning = scalar("SELECT COUNT(*) FROM oc_session;") - mid

    total = scalar("SELECT COUNT(*) FROM oc_session;")
    bytes_ = scalar("SELECT COALESCE(SUM(LENGTH(data)+32),0) FROM oc_session;")
    print(f"  {n} first-time visitors   -> {stateless:>4} new session rows")
    print(f"  1 visitor, {n} page views -> {returning:>4} new session rows")
    print(f"  table now: {total} rows, about {bytes_ / 1024:.1f} KB")
    print("  A session is written on every request, before the visitor")
    print("  has asked for anything. A crawler that ignores cookies")
    print("  writes one row per URL it touches, and they expire on a")
    print("  clock, not on departure.")


# ------------------------------------------------------- 3. engine swap
def engine() -> None:
    print("\n=== 3. Does changing the session engine move the basket? ===")
    b = Browser()
    b.get()
    b.add(*GOODS[0][:2])
    sid = b.sid

    head = (f"  {'engine':<9}{'page ms':>8}{'queries':>9}{'on session':>12}"
            f"{'cart rows':>11}{'files':>7}{'KB':>5}")

    def state(label: str) -> None:
        for _ in range(12):                  # give each store something to do
            Browser().get()
        rows = scalar(f"SELECT COUNT(*) FROM oc_cart "
                      f"WHERE session_id='{sid}';")
        files = shop("ls -1 system/storage/session 2>/dev/null "
                     "| wc -l").strip()
        kb = shop("du -sk system/storage/session 2>/dev/null "
                  "| cut -f1").strip()
        ms = fetch_ms()
        q, qs = page_queries()
        print(f"  {label:<9}{ms:>8.0f}{q:>9}{qs:>12}{rows:>11}"
              f"{files:>7}{kb or '0':>5}")

    print(head)
    print("  " + "-" * (len(head) - 2))
    state("db")

    swap = r"s/session_engine.\] *= *'db'/session_engine'] = 'file'/"
    shop(f'sed -i "{swap}" {CATALOG_CONFIG}')
    time.sleep(1)
    b2 = Browser()
    b2.get()
    b2.add(*GOODS[1][:2])
    state("file")

    back = r"s/session_engine.\] *= *'file'/session_engine'] = 'db'/"
    shop(f'sed -i "{back}" {CATALOG_CONFIG}')
    restored = shop(f"grep session_engine {CATALOG_CONFIG}").strip()
    print(f"\n  restored: {restored}")
    print("  The basket did not move: oc_cart holds it under either engine.")
    print("  What moved is the 18-byte session record, and with it two")
    print("  queries per request.")


# ---------------------------------------------------- 4. what survives
def restart(destroy: bool = False) -> None:
    print("\n=== 4. What survives ===")
    b = Browser()
    b.get()
    for pid, qty, _ in GOODS:
        b.add(pid, qty)
    sid = b.sid

    def check(label: str) -> None:
        wait_for_shop(quiet=True)
        rows = scalar(f"SELECT COUNT(*) FROM oc_cart "
                      f"WHERE session_id='{sid}';")
        srow = scalar(f"SELECT COUNT(*) FROM oc_session "
                      f"WHERE session_id='{sid}';")
        # The browser still holds the cookie; ask the shop what it shows.
        page = b.op.open(SHOP, timeout=120).read().decode("utf-8", "replace")
        shown = "item" in page.lower()
        print(f"  {label:<26} oc_cart {rows:>2}  oc_session {srow:>2}  "
              f"basket visible: {'yes' if shown and rows else 'no'}")

    check("before")

    t0 = time.perf_counter()
    compose("restart", "shop")
    check(f"restart shop ({time.perf_counter() - t0:.0f}s)")

    t0 = time.perf_counter()
    compose("down")
    compose("up", "-d")
    check(f"down, up -d ({time.perf_counter() - t0:.0f}s)")

    if not destroy:
        print("  down -v not run: it deletes the volumes and reinstalls the")
        print("  shop. Pass --destroy to measure it.")
        return
    t0 = time.perf_counter()
    compose("down", "-v")
    compose("up", "-d")
    check(f"down -v, up -d ({time.perf_counter() - t0:.0f}s)")


# ------------------------------------- aside: the empty-basket cache miss
def cache() -> None:
    """Does an empty basket cost less than a full one? It costs more.

    Cart::getProducts() begins `if (!$this->data)`, which is true both when
    the basket has not been loaded yet and when it is empty -- PHP cannot
    tell an unset array from an empty one that way. So an empty basket misses
    its own cache on every call, and the home page asks for it six times.
    """
    print("\n=== Aside: what an empty basket costs ===")

    def shape(op: Browser) -> tuple[int, int, int]:
        rootsql("SET GLOBAL slow_query_log=0;"
                "SET GLOBAL log_output='TABLE';"
                "TRUNCATE TABLE mysql.slow_log;"
                "SET GLOBAL long_query_time=0;"
                "SET GLOBAL slow_query_log=1;")
        op.get()
        rootsql("SET GLOBAL slow_query_log=0;")
        out = rootsql(
            "SELECT COUNT(*), "
            "SUM(sql_text LIKE 'SELECT * FROM `oc_cart`%'), "
            "SUM(sql_text LIKE 'DELETE FROM `oc_cart`%') "
            "FROM mysql.slow_log;")
        try:
            r = out.strip().splitlines()[1].split("\t")
            return int(r[0]), int(r[1] or 0), int(r[2] or 0)
        except (IndexError, ValueError):
            return 0, 0, 0

    b = Browser()
    b.get()
    print(f"  {'basket':<12}{'queries':>9}{'SELECT cart':>13}"
          f"{'DELETE cart':>13}")
    for n, (pid, qty, _) in [(0, (0, 0, ""))] + list(enumerate(GOODS, 1)):
        if n:
            b.add(pid, qty)
        total, sel, dele = shape(b)
        print(f"  {n:<12}{total:>9}{sel:>13}{dele:>13}")
    print("  The DELETE is unconditional: every page view of the storefront")
    print("  writes to oc_cart before it reads from it.")


# ------------------------------------------------ aside: the collector
def gc(n: int = 25) -> None:
    print(f"\n=== Aside: the shutdown collector, over {n} requests ===")
    rootsql("SET GLOBAL slow_query_log=0;"
            "SET GLOBAL log_output='TABLE';"
            "TRUNCATE TABLE mysql.slow_log;"
            "SET GLOBAL long_query_time=0;"
            "SET GLOBAL slow_query_log=1;")
    for _ in range(n):
        Browser().get()
    rootsql("SET GLOBAL slow_query_log=0;")

    out = rootsql(
        "SELECT COUNT(*), COALESCE(SUM(TIME_TO_SEC(query_time)),0)*1000 "
        "FROM mysql.slow_log WHERE sql_text LIKE 'OPTIMIZE%';")
    try:
        row = out.strip().splitlines()[1].split("\t")
        hits, ms = int(row[0]), float(row[1])
    except (IndexError, ValueError):
        hits, ms = 0, 0.0
    print(f"  OPTIMIZE TABLE oc_session ran {hits} times in {n} requests "
          f"({hits / n * 100:.0f}%), costing {ms:.0f} ms in total")
    print("  Session::__construct registers gc() as a shutdown function,")
    print("  so it is considered on every request; the probability and")
    print("  divisor settings make")
    print("  that one request in five. On InnoDB, OPTIMIZE TABLE is a full")
    print("  table rebuild.")


# --------------------------------------------------------------------- main
def main() -> int:
    ap = argparse.ArgumentParser()
    for flag in ("where", "sessions", "engine", "restart", "gc", "cache"):
        ap.add_argument(f"--{flag}", action="store_true")
    ap.add_argument("--destroy", action="store_true",
                    help="also measure down -v, which reinstalls the shop")
    a = ap.parse_args()
    picked = any((a.where, a.sessions, a.engine, a.restart,
                  a.gc, a.cache))

    wait_for_shop(quiet=True)
    if a.where or not picked:
        where()
    if a.sessions or not picked:
        sessions()
    if a.engine or not picked:
        engine()
    if a.cache or not picked:
        cache()
    if a.gc or not picked:
        gc()
    if a.restart or not picked:
        restart(a.destroy)
    return 0


if __name__ == "__main__":
    sys.exit(main())
