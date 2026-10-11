"""Chapter 14 -- how long to wait for a gateway, and what waiting wrongly costs.

Integrating a payment gateway is three decisions, and only the first is about
code:

  1. How long do you wait before giving up?
  2. What do you believe about an order whose answer never arrived?
  3. What do you do when the provider tells you the same thing twice?

This measures all three against the provider in gateway/provider.py, which is
slow, irregular and occasionally silent on purpose.

    python m14_gateway.py                 all three
    python m14_gateway.py --latency 200   the distribution, over 200 calls
    python m14_gateway.py --timeout       what each candidate timeout costs
    python m14_gateway.py --idempotency   the same payment, sent twice

The headline the chapter is written around: a timeout chosen from the median
response time rejects roughly one payment in ten, and each rejection is a
customer who may already have been charged.
"""
import argparse
import json
import os
import statistics
import sys
import time
import urllib.error
import urllib.request
import uuid

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

GATEWAY = "http://localhost:8092"
# Deliberately not round numbers alone: 1.0 s is what people pick, 3.0 s is
# what the distribution actually needs, and the ones between show the shape.
CANDIDATES = [0.5, 1.0, 1.5, 2.0, 3.0, 5.0, 8.0]


def authorize(order_id: int, amount: str = "854.40",
              key: str | None = None, timeout: float = 30.0) -> tuple:
    """(milliseconds, reply-or-None). None means the client gave up."""
    body = json.dumps({"order_id": order_id, "amount": amount,
                       "idempotency_key": key}).encode()
    req = urllib.request.Request(
        GATEWAY + "/authorize", data=body,
        headers={"Content-Type": "application/json"})
    t0 = time.perf_counter()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            reply = json.loads(r.read())
        return (time.perf_counter() - t0) * 1000, reply
    except (urllib.error.URLError, TimeoutError, OSError):
        return (time.perf_counter() - t0) * 1000, None


def stats() -> dict:
    with urllib.request.urlopen(GATEWAY + "/stats", timeout=10) as r:
        return json.loads(r.read())


def reset() -> None:
    req = urllib.request.Request(GATEWAY + "/reset", data=b"",
                                 method="POST")
    urllib.request.urlopen(req, timeout=10).read()


def up() -> bool:
    try:
        with urllib.request.urlopen(GATEWAY + "/health", timeout=5) as r:
            print(f"  gateway: {r.read().decode()}")
        return True
    except OSError:
        print("  gateway is not running: docker compose up -d gateway",
              file=sys.stderr)
        return False


def pct(values: list[float], p: float) -> float:
    s = sorted(values)
    i = min(len(s) - 1, int(round(p / 100 * (len(s) - 1))))
    return s[i]


# ------------------------------------------------------- 1. the distribution
def latency(n: int = 200) -> list[float]:
    print(f"\n=== 1. {n} authorisations, waiting as long as it takes ===")
    reset()
    times, lost = [], 0
    for i in range(n):
        ms, reply = authorize(i, timeout=40)
        if reply is None:
            lost += 1
        else:
            times.append(ms)
    print(f"  {'median':<10}{pct(times, 50):>9.0f} ms")
    print(f"  {'mean':<10}{statistics.mean(times):>9.0f} ms")
    print(f"  {'p90':<10}{pct(times, 90):>9.0f} ms")
    print(f"  {'p95':<10}{pct(times, 95):>9.0f} ms")
    print(f"  {'p99':<10}{pct(times, 99):>9.0f} ms")
    print(f"  {'slowest':<10}{max(times):>9.0f} ms")
    print(f"  never answered at all: {lost} of {n}")
    print(f"  the slowest answer is {max(times) / pct(times, 50):.0f}x the "
          f"median: a mean would have hidden that")
    return times


# ------------------------------------------------------------ 2. the timeout
def timeout(times: list[float] | None = None, n: int = 200) -> None:
    print("\n=== 2. What each candidate timeout would have cost ===")
    if not times:
        times = latency(n)
    head = (f"  {'timeout':>9}{'answered':>10}{'cut off':>9}"
            f"{'% cut off':>11}   consequence")
    print(head)
    print("  " + "-" * (len(head) - 2))
    for t in CANDIDATES:
        ok = sum(1 for ms in times if ms <= t * 1000)
        cut = len(times) - ok
        share = cut / len(times) * 100
        note = ("every one of these may already have been charged"
                if share >= 5 else
                "customer waits, but the answer arrives" if share else
                "no payment lost to the clock")
        print(f"  {t:>7.1f}s{ok:>10}{cut:>9}{share:>10.1f}%   {note}")
    print("\n  A timeout is not a performance setting. Below the tail it is a")
    print("  decision to abandon payments that were about to succeed, and the")
    print("  shop cannot tell those from payments that failed.")


# -------------------------------------------------------- 3. the double send
def idempotency() -> None:
    print("\n=== 3. The same payment, sent twice ===")
    reset()

    print("\n  Without an idempotency key:")
    a_ms, a = authorize(9001, key=None)
    b_ms, b = authorize(9001, key=None)
    print(f"    first  -> {a['transaction_id']}  {a['status']}  "
          f"{a_ms:.0f} ms")
    print(f"    second -> {b['transaction_id']}  {b['status']}  "
          f"{b_ms:.0f} ms")
    print(f"    two transactions: "
          f"{'YES -- the customer paid twice' if a['transaction_id'] != b['transaction_id'] else 'no'}")

    print("\n  With one:")
    key = uuid.uuid4().hex
    c_ms, c = authorize(9002, key=key)
    d_ms, d = authorize(9002, key=key)
    print(f"    first  -> {c['transaction_id']}  {c['status']}  "
          f"{c_ms:.0f} ms")
    print(f"    second -> {d['transaction_id']}  {d['status']}  "
          f"{d_ms:.0f} ms  replayed={d['replayed']}")
    print(f"    two transactions: "
          f"{'YES' if c['transaction_id'] != d['transaction_id'] else 'no -- the second was answered from the first'}")

    s = stats()
    print(f"\n  the provider saw {s['authorize']} requests and charged "
          f"{s['approved'] + s['declined']} times "
          f"({s['replayed']} replayed from a key)")
    print("  The key has to come from the ORDER, not from the request: a")
    print("  retry generates a new request and must reuse the old key.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--latency", type=int, nargs="?", const=200)
    ap.add_argument("--timeout", action="store_true")
    ap.add_argument("--idempotency", action="store_true")
    a = ap.parse_args()
    picked = any((a.latency, a.timeout, a.idempotency))

    if not up():
        return 1
    if a.idempotency:
        idempotency()
        return 0
    times = latency(a.latency or 200) if (a.latency or not picked) else None
    if a.timeout or not picked:
        timeout(times)
    if not picked:
        idempotency()
    return 0


if __name__ == "__main__":
    sys.exit(main())
