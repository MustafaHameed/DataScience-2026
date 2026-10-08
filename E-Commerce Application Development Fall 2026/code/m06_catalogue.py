"""Chapter 6 -- what a catalogue costs as it grows, and what an index buys.

Three questions, in order:

  1. How does the category page's response time grow with the catalogue?
  2. What is the database actually doing -- which EXPLAIN answers?
  3. What do the missing indexes buy, measured rather than asserted?

The interesting part is question 3, and the interesting answer is that it
depends entirely on which column the page sorts by. OpenCart's default sort
is `p.sort_order`, which has no index. Sorting by name goes through
`ORDER BY LCASE(pd.name)`, and the LCASE() means the index on `name` cannot
be used however carefully it was built -- a function applied to a column
hides that column from the optimiser.

    python m06_catalogue.py              the full run (several minutes)
    python m06_catalogue.py --quick      100 and 1,000 only

WARNING: this replaces the catalogue. Run seed/catalogue.py --restore
afterwards if you want your own products back.
"""
import argparse
import os
import statistics
import subprocess
import sys
import time
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402

CATEGORY_URL = "http://localhost:8090/index.php?route=product/category&path=20"
SIZES = [100, 1000, 10000, 50000]

# The indexes OpenCart does not ship. oc_product has a PRIMARY KEY and
# nothing else: no index on the column it sorts by, none on the columns
# every storefront query filters by.
INDEXES = [
    ("idx_sort_order", "oc_product", "(sort_order)"),
    ("idx_status_date", "oc_product", "(status, date_available)"),
    ("idx_price", "oc_product", "(price)"),
]


def sql(stmt: str) -> str:
    return S.sql(stmt)


def rootsql(stmt: str) -> str:
    """Run SQL as root.

    Turning the query log on and off is a GLOBAL setting, and the `shop`
    user does not hold the privilege. It does not say so: MariaDB accepts the
    statement, silently declines to act on it, and the measurement then
    reports zero queries per page -- which looks like a result rather than
    like a permissions problem. Instrumentation runs as root for that reason
    and for no other; everything the course actually teaches uses `shop`.
    """
    r = subprocess.run(
        ["docker", "compose", "exec", "-T", "db",
         "mariadb", "-uroot", "-prootpw"],
        cwd=os.path.dirname(os.path.abspath(__file__)),
        input=stmt, text=True, capture_output=True)
    return r.stdout


def fetch_ms(url: str, reps: int = 7) -> float:
    """Median wall-clock milliseconds for a page.

    The median, not the mean: one scheduler hiccup drags a mean and nothing
    drags a median. The first request is discarded because it warms caches
    that every later request then benefits from, and reporting it would
    measure the warm-up rather than the page.
    """
    times = []
    for i in range(reps + 1):
        t0 = time.perf_counter()
        with urllib.request.urlopen(url, timeout=120) as r:
            r.read()
        dt = (time.perf_counter() - t0) * 1000
        if i:
            times.append(dt)
    return statistics.median(times)


def page_sql(url: str) -> tuple[int, float]:
    """(queries, total DB milliseconds) for ONE page request.

    Measured by the server, not by this script. MariaDB's slow-query log with
    long_query_time=0 records every statement it executes together with the
    time it took, so requesting the page with the log on and reading the table
    afterwards gives the real queries the real page ran.

    This replaced a hand-written approximation of OpenCart's listing query,
    which was a mistake worth recording: that query was *similar* to the
    page's and not identical, so timing it and comparing against the page
    compared two different things. At 50,000 products it reported 670 ms of
    "page SQL" inside a 307 ms page -- impossible, and the impossibility was
    the only reason the method was questioned. Never hand-copy the query you
    mean to measure; ask the server what it ran.
    """
    rootsql("SET GLOBAL slow_query_log=0;"
            "SET GLOBAL log_output='TABLE';"
            "TRUNCATE TABLE mysql.slow_log;"
            "SET GLOBAL long_query_time=0;"
            "SET GLOBAL slow_query_log=1;")
    with urllib.request.urlopen(url, timeout=180) as r:
        r.read()
    rootsql("SET GLOBAL slow_query_log=0;")
    out = rootsql("SELECT COUNT(*), "
                  "COALESCE(SUM(TIME_TO_SEC(query_time)),0)*1000 "
                  "FROM mysql.slow_log;")
    rows = [ln.split("\t") for ln in out.strip().splitlines()[1:]]
    if rows and len(rows[0]) >= 2:
        try:
            return int(rows[0][0]), float(rows[0][1])
        except ValueError:
            pass
    return 0, 0.0


def explain(sort_by: str) -> str:
    """EXPLAIN for the listing query, reduced to what matters."""
    order = ("LCASE(pd.name)" if sort_by == "name"
             else "p.sort_order")
    stmt = f"""EXPLAIN SELECT DISTINCT p.product_id, pd.name, p.price
      FROM oc_product_to_category p2c
      INNER JOIN oc_product_to_store p2s ON (p2s.product_id=p2c.product_id)
      LEFT JOIN oc_product p ON (p.product_id=p2s.product_id
            AND p.status=1 AND p.date_available <= NOW())
      LEFT JOIN oc_product_description pd ON (p.product_id=pd.product_id)
      WHERE p2c.category_id=20 AND pd.language_id=1
      ORDER BY {order} LIMIT 20;"""
    out = sql(stmt)
    rows = [ln.split("\t") for ln in out.strip().splitlines()[1:]]
    bits = []
    for r in rows:
        if len(r) >= 10:
            bits.append(f"{r[2]}:{r[3]}" + (f" [{r[-1]}]" if r[-1].strip()
                                            and r[-1] != "NULL" else ""))
    return "; ".join(bits)


def drop_indexes() -> None:
    for name, table, _ in INDEXES:
        sql(f"ALTER TABLE {table} DROP INDEX {name};")


def add_indexes() -> None:
    for name, table, cols in INDEXES:
        sql(f"ALTER TABLE {table} ADD INDEX {name} {cols};")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    a = ap.parse_args()
    sizes = SIZES[:2] if a.quick else SIZES

    wait_for_shop(quiet=True)
    drop_indexes()          # start from the shipped schema whatever happened

    # Warm up before the first measurement, not inside it. PHP's opcache
    # compiles every file on first use, and that cost lands entirely on
    # whichever catalogue size happens to be measured first -- which in an
    # earlier run of this script made 100 products look 45% SLOWER than
    # 1,000. The warm-up is not part of the thing being measured, so it
    # happens here where it cannot be mistaken for a result.
    S.seed(sizes[0], keep=False)
    fetch_ms(CATEGORY_URL, reps=3)

    print("\n=== 1. the category page as the catalogue grows ===")
    print(f"{'products':>10}  {'page ms':>8}  {'queries':>8}  {'DB ms':>8}"
          f"  {'DB/page':>8}")
    base, basesql = {}, {}
    for n in sizes:
        S.seed(n, keep=False)
        ms = fetch_ms(CATEGORY_URL)
        nq, db = page_sql(CATEGORY_URL)
        base[n], basesql[n] = ms, db
        print(f"{n:>10,}  {ms:>8.0f}  {nq:>8d}  {db:>8.0f}"
              f"  {db / ms * 100:>7.0f}%")

    print("\n=== 2. what the database is doing, at the largest size ===")
    print(f"  default sort (p.sort_order): {explain('sort_order')}")
    print(f"  sort by name  (LCASE(name)): {explain('name')}")

    print("\n=== 3. the same page with the missing indexes added ===")
    add_indexes()
    print(f"{'products':>10}  {'page was':>9}  {'page now':>9}  {'':>7}"
          f"  {'DB was':>7}  {'DB now':>7}  {'':>7}")
    for n in sizes:
        S.seed(n, keep=False)
        ms = fetch_ms(CATEGORY_URL)
        _, db = page_sql(CATEGORY_URL)
        dp = (ms - base[n]) / base[n] * 100
        dq = ((db - basesql[n]) / basesql[n] * 100) if basesql[n] else 0.0
        print(f"{n:>10,}  {base[n]:>9.0f}  {ms:>9.0f}  {dp:>+6.1f}%"
              f"  {basesql[n]:>7.0f}  {db:>7.0f}  {dq:>+6.1f}%")
    print(f"\n  with indexes, default sort: {explain('sort_order')}")
    print(f"  with indexes, sort by name: {explain('name')}")

    drop_indexes()
    return 0


if __name__ == "__main__":
    sys.exit(main())
