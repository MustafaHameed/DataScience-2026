"""Chapter 18 -- what changing an order's status actually does.

The admin panel does not edit orders itself. It signs a request and calls the
STOREFRONT's own API over HTTP, so that one code path handles stock,
notification and totals whether the change comes from a human clicking a
dropdown, from a payment provider's callback, or from a script like this one.
That is good design, and it has a consequence nobody mentions: the admin
panel cannot change an order unless the API accepts the call.

Three things measured here:

  1. Whether the API accepts a correctly signed call at all, out of the box.
  2. What a status change does to stock -- which is decided by the TRANSITION
     between sets of statuses, not by the status itself.
  3. That "Refunded" means two different things in two places, and only one of
     them touches anything.

    python m18_orders.py             all three
    python m18_orders.py --api       just the API check
    python m18_orders.py --stock     the stock transitions
    python m18_orders.py --refund    the two meanings of Refunded
    python m18_orders.py --clean     remove the orders it made

Everything runs against your own container. See Appendix A.
"""
import argparse
import base64
import hashlib
import hmac
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402
import m10_checkout as C                 # noqa: E402

HOST = "localhost:8090"
SHOP = f"http://{HOST}/"
WATCH = [43, 40]                         # the products m10's basket buys


def sql(stmt: str) -> str:
    return S.sql(stmt)


def one(stmt: str, default: str = "") -> str:
    # MariaDB prints an empty aggregate as the four characters NULL, which
    # int() does not accept. This caught the file once already.
    out = sql(stmt).strip().splitlines()
    if len(out) < 2:
        return default
    val = out[1].split("\t")[0]
    return default if val in ("NULL", "") else val


def api_key() -> tuple[str, str]:
    r = sql("SELECT username, `key` FROM oc_api WHERE status=1 LIMIT 1;")
    rows = r.strip().splitlines()
    if len(rows) < 2:
        return "", ""
    parts = rows[1].split("\t")
    return parts[0], parts[1]


def call(action: str, post: dict, order_id: int) -> dict:
    """One signed API call, built exactly as startup/api.php verifies it."""
    user, key = api_key()
    if not key:
        return {"error": "no API user"}
    now = int(time.time())
    get = {
        "route": "api/order", "call": action, "username": user,
        "store_id": 0, "language": "en-gb", "currency": "USD",
        "time": now,
    }
    # order_id travels in the POST body, not the query string -- and the
    # signature covers an md5 of that body, so the two must be built in
    # this order.
    body = urllib.parse.urlencode(dict(post, order_id=order_id))
    # The signed string is newline-joined, in this exact order. PHP_SELF's
    # directory is "/" here because index.php sits at the web root.
    s = "\n".join([
        get["route"], action, user, HOST, "/",
        str(get["store_id"]), get["language"], get["currency"],
        hashlib.md5(body.encode()).hexdigest(), str(now),
    ]) + "\n"
    sig = base64.b64encode(
        hmac.new(key.encode(), s.encode(), hashlib.sha1).digest()).decode()
    get["signature"] = sig
    url = SHOP + "index.php?" + urllib.parse.urlencode(get)
    req = urllib.request.Request(url, data=body.encode())
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read() or b"{}")
    except urllib.error.HTTPError as e:
        return {"http": e.code, "body": e.read().decode()[:160]}
    except OSError as exc:
        return {"error": str(exc)}


def allow_ip(on: bool) -> str:
    """The API checks the caller's IP against oc_api_ip, which ships EMPTY --
    and an empty list means nothing is allowed, not that everything is."""
    api_id = one("SELECT api_id FROM oc_api WHERE status=1 LIMIT 1;", "1")
    sql(f"DELETE FROM oc_api_ip WHERE api_id={api_id};")
    if not on:
        return "(none)"
    # Whatever address the shop container sees us arrive from.
    ip = one("SELECT 1;") and _seen_ip()
    sql(f"INSERT INTO oc_api_ip (api_id, ip) VALUES ({api_id}, '{ip}');")
    return ip


def _seen_ip() -> str:
    """The address our requests arrive from, as the shop sees it.

    Read from the container's stdout, not from /var/log/apache2/access.log:
    in this image that path is a symlink to /dev/stdout, and tailing it
    blocks forever. The first version of this function did exactly that and
    hung the harness.
    """
    import re as _re
    import subprocess
    here = os.path.dirname(os.path.abspath(__file__))
    urllib.request.urlopen(SHOP, timeout=30).read()
    r = subprocess.run(["docker", "compose", "logs", "--tail", "40", "shop"],
                       cwd=here, text=True, capture_output=True)
    hits = _re.findall(r"(\d{1,3}(?:\.\d{1,3}){3}) - -", r.stdout)
    return hits[-1] if hits else "172.17.0.1"


def stock() -> dict:
    ids = ",".join(str(i) for i in WATCH)
    rows = sql(f"SELECT product_id, quantity FROM oc_product "
               f"WHERE product_id IN ({ids}) ORDER BY product_id;")
    return {ln.split("\t")[0]: int(ln.split("\t")[1])
            for ln in rows.strip().splitlines()[1:]}


# ----------------------------------------------------------------- 1. API
def check_api() -> bool:
    print("\n=== 1. Can the admin panel change an order at all? ===")
    n = one("SELECT COUNT(*) FROM oc_api_ip;", "0")
    print(f"  rows in oc_api_ip: {n}")
    allow_ip(False)
    oid = int(one("SELECT MAX(order_id) FROM oc_order;", "0") or 0)
    if not oid:
        C.one_checkout(880001)
        oid = int(one("SELECT MAX(order_id) FROM oc_order;", "0"))
    r = call("history_add", {"order_status_id": 2, "comment": "", "notify": 0},
             oid)
    print(f"  signed call with an empty allow-list -> {json.dumps(r)[:90]}")

    ip = allow_ip(True)
    print(f"  added {ip} to oc_api_ip")
    r = call("history_add", {"order_status_id": 2, "comment": "", "notify": 0},
             oid)
    ok = "error" not in r and "http" not in r
    print(f"  the same call now                    -> {json.dumps(r)[:90]}")
    print("\n  The API user ships ENABLED and the allow-list ships EMPTY.")
    print("  An empty allow-list permits nothing, so every signed call is")
    print("  refused -- and the admin panel's own order editing is one.")
    return ok


# --------------------------------------------------------------- 2. stock
TRANSITIONS = [
    (1, "Pending", "in the processing set"),
    (3, "Shipped", "in the processing and complete sets"),
    (7, "Canceled", "OUTSIDE both sets"),
    (1, "Pending", "back inside"),
    (11, "Refunded", "OUTSIDE both sets"),
    (1, "Pending", "back inside"),
    (12, "Reversed", "in the processing set -- note the name"),
]


def stock_moves() -> None:
    print("\n=== 2. Which status changes move stock ===")
    allow_ip(True)
    before_all = stock()
    C.one_checkout(880002)
    oid = int(one("SELECT MAX(order_id) FROM oc_order;", "0"))
    print(f"  order {oid}: 1 x product 43, 2 x product 40")
    head = (f"  {'status':<22}{"where it sits":<38}"
            f"{'stock 43':>9}{'stock 40':>9}")
    print(head)
    print("  " + "-" * 76)
    print(f"  {'(before the order)':<22}{'':<36}"
          f"{before_all['43']:>9}{before_all['40']:>9}")
    s = stock()
    print(f"  {'1 Pending (paid)':<22}{'in the processing set':<36}"
          f"{s['43']:>9}{s['40']:>9}")
    for sid, name, where in TRANSITIONS[1:]:
        call("history_add",
             {"order_status_id": sid, "comment": "", "notify": 0}, oid)
        s = stock()
        print(f"  {f'{sid} {name}':<22}{where:<38}{s['43']:>9}{s['40']:>9}")
    print("\n  Stock is decided by the TRANSITION, not the status: moving")
    print("  INTO the processing-or-complete set subtracts, moving OUT of")
    print("  it adds back. Nothing else does anything.")
    print("  Note 11 Refunded and 12 Reversed. Both mean money went back;")
    print("  only one of them returns the goods to stock.")


# -------------------------------------------------------------- 3. refunds
def refund() -> None:
    print("\n=== 3. The two things called Refunded ===")
    print("  (a) An ORDER moved to status 11 -- see section 2: stock returns.")
    print("  (b) A RETURN recorded with action 'Refunded'.")
    print()
    print("  What (b) does, read from admin/model/sale/returns.php:")
    print("    - UPDATE oc_return SET return_status_id = ...")
    print("    - INSERT INTO oc_return_history ...")
    print("  and nothing else. It does not change the order's status, does")
    print("  not restore stock, does not create a customer transaction, and")
    print("  does not move any money.")
    print()
    for t in ("oc_return", "oc_return_history"):
        print(f"  {t:<22}{one(f'SELECT COUNT(*) FROM {t};', '0'):>6} rows")
    print("\n  A return marked Refunded is a NOTE. The money is sent back by")
    print("  hand through the payment provider, the stock is adjusted by")
    print("  hand, and the order still reads Complete to everyone who looks.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--api", action="store_true")
    ap.add_argument("--stock", action="store_true")
    ap.add_argument("--refund", action="store_true")
    ap.add_argument("--clean", action="store_true")
    a = ap.parse_args()
    picked = any((a.api, a.stock, a.refund))

    wait_for_shop(quiet=True)
    if a.clean:
        C.clean()
        allow_ip(False)
        print("orders removed, API allow-list emptied")
        return 0
    if a.api or not picked:
        check_api()
    if a.stock or not picked:
        stock_moves()
    if a.refund or not picked:
        refund()
    return 0


if __name__ == "__main__":
    sys.exit(main())
