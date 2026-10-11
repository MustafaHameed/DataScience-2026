"""Chapter 10 -- what a checkout costs, step by step.

A checkout is not one page. It is a short sequence of steps, each of which
validates something and saves it, and only the last of which creates an order.
Students --- and more than a few shop owners --- think of it as "the slow bit",
without knowing which step is slow or why.

This drives a complete guest checkout over HTTP, with nothing but a cookie jar,
and measures each step twice over: wall-clock time, and the number of database
queries the server actually ran, read out of MariaDB's slow query log rather
than guessed.

The claim worth testing: the last step --- the one that writes the order --
costs more than all the steps before it put together.

    python m10_checkout.py              one checkout, step by step
    python m10_checkout.py --repeat 5   medians over five checkouts
    python m10_checkout.py --states     the order state machine, from the data
    python m10_checkout.py --clean      delete the orders this made

Every order it creates is a real order in your shop, with a recognisable
surname, so that --states can show you a state machine built from your own
data and --clean can take them away again.
"""
import argparse
import http.cookiejar
import json
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
MARK = "Harness"                         # the surname every test order carries

# The store's own country and zone, so that neither tax nor shipping is
# excluded by a geo zone and the measurement is of the checkout rather than
# of a rejection.
COUNTRY, ZONE = 222, 3563
GOODS = [(43, 1), (40, 2)]               # no required options; see Chapter 7


def rootsql(stmt: str) -> str:
    r = subprocess.run(
        ["docker", "compose", "exec", "-T", "db", "mariadb", "-uroot",
         "-prootpw"], cwd=CODE, input=stmt, text=True, capture_output=True)
    return r.stdout


def sql(stmt: str) -> str:
    return S.sql(stmt)


class Checkout:
    """One guest shopper, driven through the whole sequence."""

    def __init__(self, tag: int):
        self.cj = http.cookiejar.CookieJar()
        self.op = urllib.request.build_opener(
            urllib.request.HTTPCookieProcessor(self.cj))
        self.tag = tag

    def post(self, route: str, fields: dict) -> dict:
        data = urllib.parse.urlencode(fields, doseq=True).encode()
        req = urllib.request.Request(
            SHOP + "index.php?route=" + route, data=data,
            headers={"X-Requested-With": "XMLHttpRequest"})
        with self.op.open(req, timeout=120) as r:
            body = r.read().decode("utf-8", "replace")
        try:
            return json.loads(body)
        except ValueError:
            return {"_raw": body[:200]}

    def get(self, route: str = "") -> str:
        url = SHOP + ("index.php?route=" + route if route else "")
        with self.op.open(url, timeout=120) as r:
            return r.read().decode("utf-8", "replace")

    # -- the steps ---------------------------------------------------------
    def fill_basket(self) -> dict:
        for pid, qty in GOODS:
            self.post("checkout/cart.add",
                      {"product_id": pid, "quantity": qty})
        return {}

    def register(self) -> dict:
        return self.post("checkout/register.save", {
            "account": 0, "customer_group_id": 1,
            "firstname": "Test", "lastname": MARK,
            "email": f"harness{self.tag}@example.invalid",
            "telephone": "03001234567",
            "payment_address_1": "1 Example Road", "payment_city": "Lahore",
            "payment_postcode": "54000", "payment_country_id": COUNTRY,
            "payment_zone_id": ZONE, "address_match": 1, "agree": 1})

    def shipping_address(self) -> dict:
        return self.post("checkout/shipping_address.save", {
            "firstname": "Test", "lastname": MARK,
            "address_1": "1 Example Road", "city": "Lahore",
            "postcode": "54000", "country_id": COUNTRY, "zone_id": ZONE})

    def shipping_method(self) -> dict:
        self.post("checkout/shipping_method.quote", {})
        return self.post("checkout/shipping_method.save",
                         {"shipping_method": "flat.flat"})

    def payment_method(self) -> dict:
        self.post("checkout/payment_method.getMethods", {})
        return self.post("checkout/payment_method.save",
                         {"payment_method": "cod.cod"})

    def confirm(self) -> dict:
        """Writes the order row -- at status 0, which is not a status."""
        self.get("checkout/confirm")
        return self.post("checkout/confirm.confirm", {})

    def pay(self) -> dict:
        """The payment extension's own confirm. THIS is what moves the order
        off status 0; nothing before it has committed the customer."""
        return self.post("extension/opencart/payment/cod.confirm", {})

    def success(self) -> dict:
        self.get("checkout/success")
        return {}

    def steps(self):
        return [("fill basket", self.fill_basket),
                ("register (guest)", self.register),
                ("shipping address*", self.shipping_address),
                ("shipping method", self.shipping_method),
                ("payment method", self.payment_method),
                ("confirm", self.confirm),
                ("pay (cod.confirm)", self.pay),
                ("success", self.success)]


# ---------------------------------------------------------------- measuring
def log_on() -> None:
    rootsql("SET GLOBAL slow_query_log=0;"
            "SET GLOBAL log_output='TABLE';"
            "TRUNCATE TABLE mysql.slow_log;"
            "SET GLOBAL long_query_time=0;"
            "SET GLOBAL slow_query_log=1;")


def log_read() -> tuple[int, int, float]:
    """(queries, writes, db milliseconds) since the log was last truncated."""
    rootsql("SET GLOBAL slow_query_log=0;")
    out = rootsql(
        "SELECT COUNT(*), "
        "SUM(sql_text LIKE 'INSERT%' OR sql_text LIKE 'UPDATE%' "
        "OR sql_text LIKE 'DELETE%' OR sql_text LIKE 'REPLACE%'), "
        "COALESCE(SUM(TIME_TO_SEC(query_time)),0)*1000 FROM mysql.slow_log;")
    try:
        r = out.strip().splitlines()[1].split("\t")
        return int(r[0]), int(r[1] or 0), float(r[2])
    except (IndexError, ValueError):
        return 0, 0, 0.0


def one_checkout(tag: int) -> list[tuple[str, float, int, int, float]]:
    """[(step, ms, queries, writes, db ms)] for one complete checkout."""
    c = Checkout(tag)
    c.get()                                   # a session, before measuring
    rows = []
    for name, fn in c.steps():
        log_on()
        t0 = time.perf_counter()
        reply = fn()
        ms = (time.perf_counter() - t0) * 1000
        q, w, dbms = log_read()
        if isinstance(reply, dict) and reply.get("error"):
            print(f"  ! {name} refused: "
                  f"{json.dumps(reply['error'])[:150]}", file=sys.stderr)
        rows.append((name, ms, q, w, dbms))
    return rows


def report(runs: list[list[tuple]]) -> None:
    names = [r[0] for r in runs[0]]
    head = (f"  {'step':<20}{'ms':>8}{'queries':>9}{'writes':>8}"
            f"{'db ms':>8}{'share':>8}")
    print(f"\n{head}")
    print("  " + "-" * (len(head) - 2))

    med = []
    for i, name in enumerate(names):
        ms = statistics.median(r[i][1] for r in runs)
        q = statistics.median(r[i][2] for r in runs)
        w = statistics.median(r[i][3] for r in runs)
        d = statistics.median(r[i][4] for r in runs)
        med.append((name, ms, q, w, d))
    total_ms = sum(m[1] for m in med)
    total_q = sum(m[2] for m in med)
    for name, ms, q, w, d in med:
        print(f"  {name:<20}{ms:>8.0f}{q:>9.0f}{w:>8.0f}{d:>8.0f}"
              f"{ms / total_ms * 100:>7.0f}%")
    print("  " + "-" * (len(head) - 2))
    print(f"  {'TOTAL':<20}{total_ms:>8.0f}{total_q:>9.0f}")

    # Index by name rather than by position: a step added to steps() must not
    # silently re-point this at a different row, which is exactly what
    # happened once and made the headline claim read off the wrong line.
    by = {m[0]: m for m in med}
    tail = [by[k] for k in ("confirm", "pay (cod.confirm)", "success")]
    front = [m for m in med if m[0] not in {t[0] for t in tail}]

    def total(rows, i):
        return sum(r[i] for r in rows)

    print(f"\n  {'':<22}{'ms':>8}{'queries':>9}{'writes':>8}")
    print(f"  {'typing (5 steps)':<22}{total(front, 1):>8.0f}"
          f"{total(front, 2):>9.0f}{total(front, 3):>8.0f}")
    print(f"  {'ordering (3 steps)':<22}{total(tail, 1):>8.0f}"
          f"{total(tail, 2):>9.0f}{total(tail, 3):>8.0f}")
    print(f"  {'of which confirm':<22}{by['confirm'][1]:>8.0f}"
          f"{by['confirm'][2]:>9.0f}{by['confirm'][3]:>8.0f}")
    print("\n  * a guest never uses the shipping-address step: it redirects")
    print("    to the login page. Its cost is a round trip for nothing.")


def abandon() -> None:
    """Stop at confirm, the way a customer who changes their mind does.

    The order row is already written by then. It sits at status 0, which is
    not one of the fourteen statuses the shop defines, and the admin order
    list does not show it -- so the shop has a row for an order that, as far
    as anyone looking at the screens can tell, never happened.
    """
    print("\n=== A checkout abandoned at the payment button ===")
    before = int(sql("SELECT COUNT(*) FROM oc_order WHERE order_status_id=0;")
                 .strip().splitlines()[-1])
    c = Checkout(int(time.time()) % 100000)
    c.get()
    for name, fn in c.steps():
        fn()
        if name == "confirm":                 # the customer stops here
            break
    after = int(sql("SELECT COUNT(*) FROM oc_order WHERE order_status_id=0;")
                .strip().splitlines()[-1])
    print(f"  orders at status 0: {before} -> {after}")
    print(sql(f"SELECT order_id, order_status_id, total FROM oc_order "
              f"WHERE lastname='{MARK}' AND order_status_id=0;"))
    print("  The row exists. The customer never paid, the admin order list")
    print("  filters it out, and nothing will ever remove it.")


def states() -> None:
    """The order state machine, as the shop's own tables record it."""
    print("\n=== Order statuses this shop defines ===")
    print(sql("SELECT order_status_id, name FROM oc_order_status "
              "WHERE language_id=1 ORDER BY order_status_id;"))
    print("=== Orders made by this harness, and their history ===")
    print(sql(f"SELECT o.order_id, o.order_status_id, o.total, "
              f"(SELECT COUNT(*) FROM oc_order_history h "
              f"WHERE h.order_id=o.order_id) history_rows "
              f"FROM oc_order o WHERE o.lastname='{MARK}' "
              f"ORDER BY o.order_id;"))
    print("=== What status 0 means ===")
    print(sql("SELECT COUNT(*) status_0_orders FROM oc_order "
              "WHERE order_status_id=0;"))
    print("  Status 0 is not in the status table. OpenCart writes an order")
    print("  row BEFORE payment is confirmed and leaves it at 0 -- an order")
    print("  that exists and has not happened. Those are the rows a shop")
    print("  owner sees as 'missing orders'.")


def clean() -> None:
    out = sql(f"SELECT COUNT(*) FROM oc_order WHERE lastname='{MARK}';")
    n = out.strip().splitlines()[-1] if out.strip() else "0"
    sql(f"""
        DELETE FROM oc_order_history WHERE order_id IN
          (SELECT order_id FROM oc_order WHERE lastname='{MARK}');
        DELETE FROM oc_order_product WHERE order_id IN
          (SELECT order_id FROM oc_order WHERE lastname='{MARK}');
        DELETE FROM oc_order_total WHERE order_id IN
          (SELECT order_id FROM oc_order WHERE lastname='{MARK}');
        DELETE FROM oc_order WHERE lastname='{MARK}';
    """)
    print(f"removed {n} harness order(s)")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--repeat", type=int, default=1)
    ap.add_argument("--states", action="store_true")
    ap.add_argument("--abandon", action="store_true")
    ap.add_argument("--clean", action="store_true")
    a = ap.parse_args()

    wait_for_shop(quiet=True)
    if a.clean:
        clean()
        return 0
    if a.states:
        states()
        return 0
    if a.abandon:
        abandon()
        return 0

    before = sql("SELECT COUNT(*) FROM oc_order;").strip().splitlines()[-1]
    runs = [one_checkout(int(time.time() * 1000) % 100000 + i)
            for i in range(a.repeat)]
    after = sql("SELECT COUNT(*) FROM oc_order;").strip().splitlines()[-1]
    report(runs)
    print(f"\n  orders in the table: {before} -> {after} "
          f"for {a.repeat} checkout(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
