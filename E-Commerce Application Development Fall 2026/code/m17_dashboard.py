"""Chapter 17 -- what the dashboard counts, which is not one thing.

The admin dashboard shows four headline numbers and a chart. They are computed
by different queries with different filters, they disagree with each other, and
nothing on the screen says so. This builds a small, known set of orders and
then asks each of the dashboard's own questions of it, so that the
disagreement is a number rather than a claim.

    python m17_dashboard.py            build the orders and show the counts
    python m17_dashboard.py --counts   just the counts, change nothing
    python m17_dashboard.py --clean    remove the orders it made

The filters below are not invented. They are copied from the queries the
platform runs:

  order widget   admin/model/sale/order.php          WHERE order_status_id > 0
  sales widget   .../report/sale.php getTotalSales   WHERE order_status_id > 0
  the chart      .../dashboard/chart.php             WHERE order_status_id IN
                                                     (config_complete_status)
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402
import m10_checkout as C                 # noqa: E402


def sql(stmt: str) -> str:
    return S.sql(stmt)


def rows(stmt: str) -> list[list[str]]:
    out = sql(stmt).strip().splitlines()
    return [ln.split("\t") for ln in out[1:]] if len(out) > 1 else []


def complete_statuses() -> list[int]:
    """Which statuses the chart counts -- a SETTING, not a constant."""
    r = rows("SELECT value FROM oc_setting "
             "WHERE `key`='config_complete_status';")
    try:
        return [int(x) for x in json.loads(r[0][0])]
    except (IndexError, ValueError):
        return [5, 3]


def build() -> None:
    """Five paid orders, two of them completed, and one abandonment."""
    print("\n  building a known set of orders...")
    C.clean()
    for _ in range(5):
        C.one_checkout(int.from_bytes(os.urandom(2), "big"))
    C.abandon()
    sql("UPDATE oc_order SET order_status_id=5 WHERE order_id IN "
        "(SELECT * FROM (SELECT order_id FROM oc_order "
        " WHERE order_status_id=1 ORDER BY order_id LIMIT 2) t);")


def counts() -> None:
    comp = complete_statuses()
    in_list = ",".join(str(c) for c in comp)

    print("\n=== What is actually in the table ===")
    print(f"  {'status':>8}{'orders':>9}{'value':>14}")
    for r in rows("SELECT order_status_id, COUNT(*), COALESCE(SUM(total),0) "
                  "FROM oc_order GROUP BY order_status_id "
                  "ORDER BY order_status_id;"):
        name = rows(f"SELECT name FROM oc_order_status "
                    f"WHERE order_status_id={r[0]} AND language_id=1;")
        label = name[0][0] if name else "(not a status)"
        print(f"  {r[0]:>8}{r[1]:>9}{float(r[2]):>14,.2f}   {label}")

    print("\n=== What each part of the dashboard reports ===")
    asks = [
        ("Orders widget", "order_status_id > 0"),
        ("Sales widget", "order_status_id > 0"),
        (f"The chart (status {in_list})", f"order_status_id IN ({in_list})"),
        ("The table itself", "1=1"),
    ]
    print(f"  {'':<30}{'orders':>9}{'value':>14}")
    for label, where in asks:
        r = rows(f"SELECT COUNT(*), COALESCE(SUM(total),0) "
                 f"FROM oc_order WHERE {where};")
        n, v = (r[0][0], float(r[0][1])) if r else ("0", 0.0)
        print(f"  {label:<30}{n:>9}{v:>14,.2f}")

    print("\n  Three different answers to \"how many orders\", all correct,")
    print("  all on one screen, none of them labelled with its filter.")
    print("  The gap between the widgets and the table is the abandonment")
    print("  rate; the gap between the widgets and the chart is the work")
    print("  not yet done.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--counts", action="store_true")
    ap.add_argument("--clean", action="store_true")
    a = ap.parse_args()

    wait_for_shop(quiet=True)
    if a.clean:
        C.clean()
        return 0
    if not a.counts:
        build()
    counts()
    return 0


if __name__ == "__main__":
    sys.exit(main())
