"""Chapter 15 -- what a password costs, and what a login gives away.

Storing a password is the one place in a shop where you are deliberately
making your own software slow. The whole value of a modern password hash is
that it is expensive to compute, because an attacker who steals your database
has to compute it once per guess. The cost is a dial, and turning it is a
trade: security against how many people can log in per second.

Three questions:

  1. What does each bcrypt cost factor actually cost, in milliseconds and in
     logins per second?
  2. Which cost is this shop using, and is it the right one in 2026?
  3. What does the login form tell an attacker who has no password at all?

    python m15_accounts.py              all three
    python m15_accounts.py --cost       the cost factor against the clock
    python m15_accounts.py --enumerate  does a wrong email answer faster?
    python m15_accounts.py --lockout    what five wrong guesses do
    python m15_accounts.py --clean      remove the test customer

The hashing is timed inside the shop container by PHP itself, because timing
it from Python would measure a different library on a different machine. The
enumeration test is timed from outside, because that is where an attacker is.
"""
import argparse
import http.cookiejar
import os
import re
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
LOGIN = SHOP + "index.php?route=account/login.login"

KNOWN = "known.customer@example.invalid"
UNKNOWN = "no.such.person.at.all@example.invalid"
PASSWORD = "correct-horse-battery"
COSTS = [8, 9, 10, 11, 12, 13, 14]


def sql(stmt: str) -> str:
    return S.sql(stmt)


def php(code: str) -> str:
    """Run a PHP snippet inside the shop container."""
    r = subprocess.run(
        ["docker", "compose", "exec", "-T", "shop", "php", "-r", code],
        cwd=CODE, text=True, capture_output=True)
    if r.returncode != 0:
        print(r.stderr[:400], file=sys.stderr)
    return r.stdout


# ------------------------------------------------------- 1. the cost factor
def cost() -> None:
    print("\n=== 1. What each bcrypt cost factor costs ===")
    out = php(
        'foreach ([8,9,10,11,12,13,14] as $c) {'
        '  $t = microtime(true);'
        '  $h = password_hash("correct-horse-battery", PASSWORD_BCRYPT,'
        '                     ["cost" => $c]);'
        '  $ms = (microtime(true) - $t) * 1000;'
        '  $t2 = microtime(true);'
        '  password_verify("correct-horse-battery", $h);'
        '  $vms = (microtime(true) - $t2) * 1000;'
        '  printf("%d\t%.1f\t%.1f\n", $c, $ms, $vms);'
        '}')
    print(f"  {'cost':>5}{'hash ms':>10}{'verify ms':>12}"
          f"{'logins/sec':>13}{'vs cost 8':>11}")
    print("  " + "-" * 49)
    base = None
    for line in out.strip().splitlines():
        parts = line.split("\t")
        if len(parts) != 3:
            continue
        c, hms, vms = int(parts[0]), float(parts[1]), float(parts[2])
        base = base or vms
        print(f"  {c:>5}{hms:>10.1f}{vms:>12.1f}{1000 / vms:>13.0f}"
              f"{vms / base:>10.0f}x")

    actual = php(
        '$h = password_hash("x", PASSWORD_DEFAULT);'
        'echo $h . "\n" . PHP_VERSION . "\n";').strip().splitlines()
    if actual:
        h = actual[0]
        algo = {"$2y$": "bcrypt", "$argon": "argon2"}.get(h[:4], h[:4])
        print(f"\n  PASSWORD_DEFAULT here produces: {h[:7]}...  ({algo})")
        print(f"  which is cost {h[4:6]} -- the PHP default, unchanged since"
              f" 2013")
        if len(actual) > 1:
            print(f"  PHP {actual[1]}")
    print("\n  Each step up the cost factor DOUBLES the work. That is the")
    print("  point: it doubles for the attacker too, and the attacker is")
    print("  doing it a billion times.")


# ------------------------------------------------- 2. does a wrong email leak
def make_customer() -> None:
    h = php(f'echo password_hash("{PASSWORD}", PASSWORD_DEFAULT);').strip()
    sql(f"DELETE FROM oc_customer WHERE email='{KNOWN}';")
    sql(f"INSERT INTO oc_customer (customer_group_id, store_id, language_id,"
        f" firstname, lastname, email, telephone, custom_field, password,"
        f" newsletter, ip, status, safe, commenter, token, code, date_added)"
        f" VALUES (1,0,1,'Known','Customer','{KNOWN}','03001234567','',"
        f"'{h}',0,'127.0.0.1',1,0,0,'','',NOW());")
    sql(f"DELETE FROM oc_customer_login WHERE email IN "
        f"('{KNOWN}','{UNKNOWN}');")


def try_login(email: str, password: str) -> tuple[float, str]:
    """(milliseconds, reply) for one login attempt.

    The login page issues a one-time `login_token` and the action refuses
    any request that does not carry the matching one -- a CSRF defence, and
    the subject of Chapter 16. A harness that skips it is answered with a
    redirect in 14 ms and measures nothing at all, which is exactly what the
    first version of this function did: all three cases, including the
    correct password, came back identical. If a timing experiment shows no
    difference, check first that it reached the code it meant to time.
    """
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cj))
    with op.open(SHOP + "index.php?route=account/login", timeout=60) as r:
        page = r.read().decode("utf-8", "replace")
    m = re.search(r"login_token=([A-Za-z0-9]+)", page)
    token = m.group(1) if m else ""

    data = urllib.parse.urlencode(
        {"email": email, "password": password}).encode()
    req = urllib.request.Request(
        LOGIN + "&login_token=" + token, data=data,
        headers={"X-Requested-With": "XMLHttpRequest"})
    t0 = time.perf_counter()
    with op.open(req, timeout=60) as r:
        body = r.read().decode("utf-8", "replace")
    return (time.perf_counter() - t0) * 1000, body[:160]


def enumerate_(reps: int = 15) -> None:
    print("\n=== 2. Does the login form say which emails are registered? ===")
    make_customer()
    sets = {}
    for label, email in (("registered email", KNOWN),
                         ("unknown email", UNKNOWN)):
        times = []
        for _ in range(reps):
            sql(f"DELETE FROM oc_customer_login WHERE email='{email}';")
            ms, _body = try_login(email, "definitely-the-wrong-password")
            times.append(ms)
        sets[label] = times
        print(f"  {label:<20}median {statistics.median(times):>7.1f} ms   "
              f"min {min(times):>7.1f}   max {max(times):>7.1f}")

    a = statistics.median(sets["registered email"])
    b = statistics.median(sets["unknown email"])
    gap = a - b
    print(f"\n  difference: {gap:+.1f} ms "
          f"({abs(gap) / max(a, b) * 100:.0f}% of the slower)")
    verdict = ("the registered email is measurably slower -- an attacker can "
               "tell" if gap > 8 else
               "no usable difference at this sample size")
    print(f"  verdict: {verdict}")
    print("  A wrong password on a real account still has to be verified")
    print("  against a hash; a wrong email need not be. Whether that shows")
    print("  up over the network is an empirical question, not a theoretical")
    print("  one -- which is why it is measured rather than asserted.")


# --------------------------------------------------------- 3. the lockout
def lockout() -> None:
    print("\n=== 3. What five wrong guesses do ===")
    make_customer()
    limit = sql("SELECT value FROM oc_setting "
                "WHERE `key`='config_login_attempts';")
    limit = limit.strip().splitlines()[-1]
    print(f"  config_login_attempts = {limit}")
    for i in range(1, int(limit) + 3):
        ms, body = try_login(KNOWN, "wrong")
        total = sql(f"SELECT total FROM oc_customer_login "
                    f"WHERE email='{KNOWN}';").strip().splitlines()
        n = total[-1] if len(total) > 1 else "0"
        locked = "Exceeded" in body or "attempts" in body
        print(f"  attempt {i:>2}  {ms:>6.0f} ms  recorded={n:<4}"
              f"{'LOCKED' if locked else ''}")
    ms, body = try_login(KNOWN, PASSWORD)
    refused = "warning" in body or "Exceeded" in body
    print(f"  correct password now: "
          f"{'still refused' if refused else 'accepted'} ({ms:.0f} ms)")
    print("  The lock is on (email, IP). Note what that means for an attacker")
    print("  with many addresses, and for a shop behind one office NAT.")


def clean() -> None:
    sql(f"DELETE FROM oc_customer_login WHERE email IN "
        f"('{KNOWN}','{UNKNOWN}');")
    sql(f"DELETE FROM oc_customer WHERE email='{KNOWN}';")
    print("test customer removed")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--cost", action="store_true")
    ap.add_argument("--enumerate", dest="enum", action="store_true")
    ap.add_argument("--lockout", action="store_true")
    ap.add_argument("--clean", action="store_true")
    a = ap.parse_args()
    picked = any((a.cost, a.enum, a.lockout, a.clean))

    wait_for_shop(quiet=True)
    if a.clean:
        clean()
        return 0
    if a.cost or not picked:
        cost()
    if a.enum or not picked:
        enumerate_()
    if a.lockout or not picked:
        lockout()
    if not picked:
        clean()
    return 0


if __name__ == "__main__":
    sys.exit(main())
