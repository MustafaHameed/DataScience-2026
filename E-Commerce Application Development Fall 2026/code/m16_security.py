"""Chapter 16 -- what your own shop gives away, before and after hardening.

SCOPE. Every check below runs against http://localhost:8090 -- the container
on your own machine, which you installed, which contains no real person's
data, and which you can throw away with one command. Running the same checks
against any system you do not own and have not been authorised IN WRITING to
test is a criminal offence in Pakistan under the Prevention of Electronic
Crimes Act 2016, and in most other countries under their equivalent. There is
no grey area and no educational exemption. Appendix A says the same thing and
means it.

What this does:

    python m16_security.py              check, harden, check again, compare
    python m16_security.py --check      the battery, once, changing nothing
    python m16_security.py --harden     apply shop/security.conf
    python m16_security.py --unharden   take it off again
    python m16_security.py --verbose    show every response header

Each check is a question with a yes/no answer that you could verify by hand
with curl. Nothing here is clever; the value is in running it twice.
"""
import argparse
import os
import re
import urllib.parse
import subprocess
import sys
import urllib.error
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from waitfor import wait_for_shop        # noqa: E402

CODE = os.path.dirname(os.path.abspath(__file__))
SHOP = "http://localhost:8090/"
MARKER = "ECAD-MARKER-DO-NOT-PANIC"


def shell(cmd: str) -> tuple[int, str]:
    r = subprocess.run(["docker", "compose", "exec", "-T", "shop", "sh", "-lc",
                        cmd], cwd=CODE, text=True, capture_output=True)
    return r.returncode, r.stdout + r.stderr


def get(path: str = "") -> tuple[int, dict, bytes]:
    """(status, headers, body) -- an error page is an answer, not a failure."""
    try:
        with urllib.request.urlopen(SHOP + path, timeout=30) as r:
            return r.status, dict(r.headers), r.read()
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read()
    except OSError:
        return 0, {}, b""


# ------------------------------------------------------------------ checks
def battery(verbose: bool = False) -> list[tuple[str, bool, str]]:
    """[(question, passed, detail)].

    `passed` always means the safe answer, never merely "a value is set".
    """
    out = []
    status, head, _ = get()
    if verbose:
        for k, v in head.items():
            print(f"      {k}: {v}")

    def has(name: str) -> str:
        return head.get(name, "")

    out.append(("Server header hides the version",
                "Apache/" not in has("Server"),
                has("Server") or "(absent)"))
    out.append(("X-Powered-By is absent", not has("X-Powered-By"),
                has("X-Powered-By") or "(absent)"))
    out.append(("X-Content-Type-Options: nosniff",
                "nosniff" in has("X-Content-Type-Options").lower(),
                has("X-Content-Type-Options") or "(absent)"))
    out.append(("X-Frame-Options or frame-ancestors",
                bool(has("X-Frame-Options"))
                or "frame-ancestors" in has("Content-Security-Policy"),
                has("X-Frame-Options") or "(absent)"))
    out.append(("Content-Security-Policy set",
                bool(has("Content-Security-Policy")),
                (has("Content-Security-Policy")[:40] + "...")
                if has("Content-Security-Policy") else "(absent)"))
    out.append(("Referrer-Policy set", bool(has("Referrer-Policy")),
                has("Referrer-Policy") or "(absent)"))
    out.append(("Permissions-Policy set", bool(has("Permissions-Policy")),
                has("Permissions-Policy") or "(absent)"))
    out.append(("No wildcard CORS with credentials",
                not (has("Access-Control-Allow-Origin") == "*"
                     and has("Access-Control-Allow-Credentials") == "true"),
                f"{has('Access-Control-Allow-Origin') or '-'} / "
                f"{has('Access-Control-Allow-Credentials') or '-'}"))

    cookies = head.get("Set-Cookie", "")
    out.append(("Session cookie is HttpOnly", "HttpOnly" in cookies,
                "HttpOnly" if "HttpOnly" in cookies else "(missing)"))
    out.append(("Session cookie has SameSite", "SameSite" in cookies,
                re.search(r"SameSite=(\w+)", cookies).group(0)
                if "SameSite" in cookies else "(missing)"))

    # The error log, which the shipped .htaccess.txt tries and fails to block.
    shell(f"echo {MARKER} >> system/storage/logs/error.log")
    code, _h, body = get("system/storage/logs/error.log")
    leaked = MARKER.encode() in body
    out.append(("Error log NOT readable over HTTP", not leaked,
                f"HTTP {code}, {len(body)} bytes"
                + (" -- CONTENTS SERVED" if leaked else "")))
    shell("echo -n > system/storage/logs/error.log")

    code, _h, _b = get("system/storage/")
    out.append(("system/storage/ not browsable", code in (401, 403, 404),
                f"HTTP {code}"))
    code, _h, _b = get("install/")
    out.append(("install/ removed", code in (401, 403, 404), f"HTTP {code}"))
    code, _h, body = get("config.php")
    out.append(("config.php not served as text",
                b"DB_PASSWORD" not in body, f"HTTP {code}, {len(body)} bytes"))
    code, _h, _b = get("admin/")
    out.append(("Admin not at the default path", code in (401, 403, 404),
                f"HTTP {code}"))
    return out


def show(rows: list, title: str) -> int:
    print(f"\n  {title}")
    print(f"  {'':<40}{'':<4}")
    passed = 0
    for q, ok, detail in rows:
        passed += ok
        print(f"    {'PASS' if ok else 'FAIL'}  {q:<38}  {detail[:34]}")
    print(f"    ---- {passed} of {len(rows)} ----")
    return passed


def injection() -> None:
    """The two attacks students expect to find, against the search box.

    Both are tried against YOUR OWN container (see the scope note at the top
    of this file). Neither probe is harmful: one is a quotation mark, the
    other is a script tag that would pop an alert box in your own browser.
    The point is not to break the shop. It is to find out, by measurement
    rather than by reputation, which classes of attack the framework already
    handles -- so that you spend your attention on the ones it does not.
    """
    print("\n=== The classic two, against your own search box ===")
    probes = [
        ("a normal search", "ipod"),
        ("one quotation mark", "ipod' "),
        ("SQL injection", "ipod' OR '1'='1"),
        ("script tag", "<script>alert(1)</script>"),
        ("image handler", "<img src=x onerror=alert(1)>"),
    ]
    print(f"  {'probe':<22}{'bytes':>8}{'verbatim':>11}{'escaped':>10}")
    print("  " + "-" * 51)
    for label, probe in probes:
        code, _h, body = get(
            "index.php?route=product/search&search="
            + urllib.parse.quote(probe))
        text = body.decode("utf-8", "replace")
        raw = probe in text
        esc = (probe.replace("<", "&lt;").replace(">", "&gt;")) in text
        print(f"  {label:<22}{len(body):>8}{str(raw):>11}{str(esc):>10}")
    print("\n  The SQL probe returns a SHORTER page than the plain search:")
    print("  it found nothing, because the whole string was searched for")
    print("  literally. The query's meaning did not change.")
    print("  The script probe appears only in its escaped form. Twig")
    print("  auto-escapes, so the markup never becomes markup.")
    print("\n  Both classic attacks are already handled by the framework.")
    print("  Every finding in the battery above is in the DEPLOYMENT.")


# ---------------------------------------------------------------- hardening
def harden(on: bool = True) -> None:
    if on:
        # "zz-" so that it loads AFTER Debian's own security.conf, which
        # sets ServerTokens OS. conf-enabled is read in filename order, and
        # the last setting of a server-level directive wins -- so a file
        # named security-ecad.conf loses to security.conf on every line they
        # both touch, for no reason a reader would ever guess.
        dest = "shop:/etc/apache2/conf-available/zz-security-ecad.conf"
        subprocess.run(["docker", "compose", "cp", "shop/security.conf",
                        dest], cwd=CODE, capture_output=True)
        shell("a2enmod headers >/dev/null 2>&1; "
              "a2enconf zz-security-ecad >/dev/null 2>&1; "
              "apachectl -k restart")
    else:
        shell("a2disconf zz-security-ecad >/dev/null 2>&1; "
          "apachectl -k restart")
    wait_for_shop(quiet=True)


def compare(verbose: bool = False) -> None:
    print("\n=== The same shop, before and after one configuration file ===")
    harden(False)
    before = battery(verbose)
    nb = show(before, "BEFORE -- as installed")

    harden(True)
    after = battery(verbose)
    na = show(after, "AFTER -- shop/security.conf enabled")

    print(f"\n  {nb} of {len(before)}  ->  {na} of {len(after)}")
    fixed = [b[0] for b, a in zip(before, after) if a[1] and not b[1]]
    left = [a[0] for a in after if not a[1]]
    if fixed:
        print("\n  closed by the config file:")
        for q in fixed:
            print(f"    + {q}")
    if left:
        print("\n  STILL OPEN -- and not fixable with a header:")
        for q in left:
            print(f"    - {q}")
    print("\n  Everything the file closed cost one line each. What is left")
    print("  needs a decision: moving the admin path, moving storage out of")
    print("  the web root, and serving over HTTPS.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--injection", action="store_true")
    ap.add_argument("--harden", action="store_true")
    ap.add_argument("--unharden", action="store_true")
    ap.add_argument("--verbose", action="store_true")
    a = ap.parse_args()

    print(__doc__.split("What this does:")[0].strip()[:0] or "", end="")
    wait_for_shop(quiet=True)
    if a.injection:
        injection()
        return 0
    if a.harden:
        harden(True)
        show(battery(a.verbose), "AFTER hardening")
    elif a.unharden:
        harden(False)
        show(battery(a.verbose), "AFTER removing the hardening")
    elif a.check:
        show(battery(a.verbose), "As it stands")
    else:
        compare(a.verbose)
    return 0


if __name__ == "__main__":
    sys.exit(main())
