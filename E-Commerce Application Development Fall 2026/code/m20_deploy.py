"""Chapter 20 -- what the shop does under load, and what a backup costs.

Two halves, because deployment is two jobs that get confused with each other:

  PERFORMANCE. How many customers at once can this shop serve, and what does
  each one experience? The answer is not one number: throughput and latency
  move in opposite directions as concurrency rises, and a shop that reports
  "120 requests per second" may be serving every one of them in four seconds.

  MAINTENANCE. What does it cost to take a backup, and -- the question nobody
  asks until the day it matters -- what does it cost to RESTORE one?

    python m20_deploy.py                 all of it
    python m20_deploy.py --load          1, 5 and 20 concurrent clients
    python m20_deploy.py --opcache       the same load with the compiler
                                         cache off, then on again
    python m20_deploy.py --backup        time a backup and a restore

The load figures are from one machine against another process on the same
machine, which flatters the network and punishes the CPU. Appendix A names
the machine. They are a comparison between configurations, not a capacity
plan for anybody's shop.
"""
import argparse
import concurrent.futures as cf
import os
import statistics
import subprocess
import sys
import time
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from waitfor import wait_for_shop        # noqa: E402

CODE = os.path.dirname(os.path.abspath(__file__))
SHOP = "http://localhost:8090/"
LEVELS = [1, 5, 20]
REQUESTS = 60                            # per level, shared across clients


def shell(cmd: str) -> str:
    r = subprocess.run(["docker", "compose", "exec", "-T", "shop", "sh", "-lc",
                        cmd], cwd=CODE, text=True, capture_output=True)
    return r.stdout + r.stderr


def fetch() -> float:
    t0 = time.perf_counter()
    try:
        with urllib.request.urlopen(SHOP, timeout=120) as r:
            r.read()
    except OSError:
        return -1.0
    return (time.perf_counter() - t0) * 1000


def pct(values: list[float], p: float) -> float:
    s = sorted(values)
    i = min(len(s) - 1, int(round(p / 100 * (len(s) - 1))))
    return s[i]


def run_level(clients: int, n: int = REQUESTS) -> dict:
    """n requests spread over `clients` threads, all starting together."""
    fetch()                              # warm, and not counted
    t0 = time.perf_counter()
    with cf.ThreadPoolExecutor(max_workers=clients) as pool:
        times = list(pool.map(lambda _: fetch(), range(n)))
    wall = time.perf_counter() - t0
    ok = [t for t in times if t >= 0]
    return {
        "clients": clients, "ok": len(ok), "failed": len(times) - len(ok),
        "rps": len(ok) / wall if wall else 0,
        "median": statistics.median(ok) if ok else 0,
        "p95": pct(ok, 95) if ok else 0,
        "worst": max(ok) if ok else 0,
    }


def table(rows: list[dict], title: str) -> None:
    head = (f"  {'clients':>8}{'req/sec':>10}{'median ms':>12}{'p95 ms':>10}"
            f"{'worst ms':>11}{'failed':>8}")
    print(f"\n  {title}")
    print(head)
    print("  " + "-" * (len(head) - 2))
    for r in rows:
        print(f"  {r['clients']:>8}{r['rps']:>10.1f}{r['median']:>12.0f}"
              f"{r['p95']:>10.0f}{r['worst']:>11.0f}{r['failed']:>8}")


def load() -> list[dict]:
    print("\n=== 1. The same page, by more people at once ===")
    rows = [run_level(c) for c in LEVELS]
    table(rows, "as deployed")
    base, last = rows[0], rows[-1]
    peak = max(rows, key=lambda r: r["rps"])
    print(f"\n  Throughput peaks at {peak['clients']} clients "
          f"({peak['rps']:.1f} req/sec).")
    if peak is not last:
        print(f"  At {last['clients']} it has FALLEN to {last['rps']:.1f} "
              f"-- more customers, fewer served.")
    print(f"  Meanwhile the median response grew "
          f"{last['median'] / base['median']:.1f}x, from "
          f"{base['median']:.0f} to {last['median']:.0f} ms.")
    print("  That turning point is the shop's saturation point, and it is")
    print("  the only number worth knowing before a sale day. Past it, each")
    print("  extra customer makes every other one slower and the shop")
    print("  serves fewer of them: the queue is doing the work, not")
    print("  the server.")
    return rows


def opcache(state: bool) -> None:
    """PHP's compiler cache. Without it every file is recompiled
    on every request."""
    val = "1" if state else "0"
    shell(f"printf 'opcache.enable={val}\\n' "
          f"> /usr/local/etc/php/conf.d/zzz-opcache.ini; "
          f"apachectl -k restart")
    wait_for_shop(quiet=True)


def opcache_status() -> str:
    return shell("php -r 'var_export(ini_get(\"opcache.enable\"));'").strip()


def compare_opcache() -> None:
    print("\n=== 2. With and without the compiler cache ===")
    opcache(False)
    off = [run_level(c) for c in LEVELS]
    table(off, "opcache OFF -- every PHP file recompiled on every request")

    opcache(True)
    on = [run_level(c) for c in LEVELS]
    table(on, "opcache ON -- as shipped in this stack")

    print()
    for a, b in zip(off, on):
        print(f"  {a['clients']:>2} clients: "
              f"{a['rps']:>6.1f} -> {b['rps']:>6.1f} req/sec "
              f"({b['rps'] / a['rps']:.1f}x), median "
              f"{a['median']:>5.0f} -> {b['median']:>5.0f} ms")
    print("\n  One line of configuration. No code, no caching layer, no CDN.")
    print("  Check it is on before optimising anything else; a shop running")
    print("  without it is being measured as a compiler, not as a shop.")


def backup() -> None:
    print("\n=== 3. What a backup costs, and what a restore costs ===")
    dump = "/tmp/ecad-backup.sql"

    t0 = time.perf_counter()
    r = subprocess.run(
        ["docker", "compose", "exec", "-T", "db", "sh", "-lc",
         f"mariadb-dump -uroot -prootpw --single-transaction opencart "
         f"> {dump} 2>/dev/null; wc -c < {dump}"],
        cwd=CODE, text=True, capture_output=True)
    dump_s = time.perf_counter() - t0
    size = (r.stdout.strip().splitlines() or ["0"])[-1]
    try:
        kb = int(size) / 1024
    except ValueError:
        kb = 0
    print(f"  database dump   {dump_s:>6.1f} s   {kb:>9,.0f} KB")

    t0 = time.perf_counter()
    subprocess.run(
        ["docker", "compose", "exec", "-T", "db", "sh", "-lc",
         f"mariadb -uroot -prootpw opencart < {dump}"],
        cwd=CODE, capture_output=True)
    restore_s = time.perf_counter() - t0
    print(f"  database restore{restore_s:>6.1f} s")

    t0 = time.perf_counter()
    files = shell("du -sk /var/www/html | cut -f1").strip()
    count = shell("find /var/www/html -type f | wc -l").strip()
    walk_s = time.perf_counter() - t0
    print(f"  shop files      {walk_s:>6.1f} s   {int(files or 0):>9,} KB "
          f"in {count} files")

    wait_for_shop(quiet=True)
    print(f"\n  A backup is {dump_s:.0f} s. A restore is {restore_s:.0f} s --")
    print("  and it is the only one of the two numbers that matters, because")
    print("  it is the one you will be timing while the shop is down.")
    print("  A backup nobody has restored is a file, not a backup.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--load", action="store_true")
    ap.add_argument("--opcache", action="store_true")
    ap.add_argument("--backup", action="store_true")
    a = ap.parse_args()
    picked = any((a.load, a.opcache, a.backup))

    wait_for_shop(quiet=True)
    print(f"  opcache.enable = {opcache_status()}")
    if a.load or not picked:
        load()
    if a.opcache or not picked:
        compare_opcache()
    if a.backup or not picked:
        backup()
    return 0


if __name__ == "__main__":
    sys.exit(main())
