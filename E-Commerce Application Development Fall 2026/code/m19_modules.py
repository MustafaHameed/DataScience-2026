"""Chapter 19 -- install a module you wrote, and measure what it costs.

Everything up to here configured a platform somebody else wrote. This is the
chapter the course's title names: adding behaviour the shop did not have,
without editing a single file that OpenCart ships, so that the change survives
an upgrade.

Two extensions live in modules/ecadmod/:

    lowstock   a storefront module -- "nearly gone", one query per render
    zonerate   a shipping method   -- per-zone rates with a free threshold

and this script installs them the way the marketplace installer would, then
measures what a module costs a page.

    python m19_modules.py --install    copy in and register
    python m19_modules.py --place N    put N instances of lowstock on the
                                       home page's Content Top
    python m19_modules.py --measure    page time and queries at 0, 1, 3, 5
    python m19_modules.py --remove     take it all out again

The measurement answers a question students ask and nobody writes down: is a
module expensive? The answer is in the table, and it is not "yes" or "no".
"""
import argparse
import json
import os
import statistics
import subprocess
import sys
import time
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402

CODE = os.path.dirname(os.path.abspath(__file__))
SHOP = "http://localhost:8090/"
EXT = "ecadmod"
# "Content Top" of the Home layout, which is layout_id 1 in a stock install.
HOME_LAYOUT, POSITION = 1, "content_top"


def sql(stmt: str) -> str:
    return S.sql(stmt)


def one(stmt: str, default: str = "0") -> str:
    out = sql(stmt).strip().splitlines()
    if len(out) < 2:
        return default
    v = out[1].split("\t")[0]
    return default if v in ("NULL", "") else v


def shell(cmd: str) -> str:
    r = subprocess.run(["docker", "compose", "exec", "-T", "shop", "sh", "-lc",
                        cmd], cwd=CODE, text=True, capture_output=True)
    return r.stdout + r.stderr


def rootsql(stmt: str) -> str:
    r = subprocess.run(
        ["docker", "compose", "exec", "-T", "db", "mariadb", "-uroot",
         "-prootpw"], cwd=CODE, input=stmt, text=True, capture_output=True)
    return r.stdout


# ------------------------------------------------------------- installing
def install() -> None:
    """What the marketplace installer does, in three steps you can read."""
    print("\n=== Installing ecadmod ===")

    # 1. The files. An extension is a directory under extension/ whose shape
    #    mirrors the shop's own -- admin/, catalog/, and a manifest.
    subprocess.run(["docker", "compose", "cp", f"modules/{EXT}",
                    f"shop:/var/www/html/extension/{EXT}"],
                   cwd=CODE, capture_output=True)
    shell(f"chown -R www-data:www-data /var/www/html/extension/{EXT}")
    n = shell(f"find /var/www/html/extension/{EXT} -type f | wc -l").strip()
    print(f"  copied {n} files to extension/{EXT}/")

    # 2. The manifest row. Without it the admin panel does not know the
    #    extension exists, even though every file is in place.
    meta = json.load(open(os.path.join(CODE, "modules", EXT, "install.json"),
                          encoding="utf-8"))
    sql(f"DELETE FROM oc_extension_install WHERE code='{EXT}';")
    sql(f"INSERT INTO oc_extension_install "
        f"(extension_id, extension_download_id, name, description, code, "
        f" version, author, link, status, date_added) VALUES "
        f"(0, 0, '{meta['name']}', '{meta['description'][:200]}', '{EXT}', "
        f"'{meta['version']}', '{meta['author']}', "
        f"'{meta['link']}', 1, NOW());")

    # 3. One row per thing the extension provides, by type and code. This is
    #    what makes each appear in its own list in the admin panel.
    sql(f"DELETE FROM oc_extension WHERE extension='{EXT}';")
    sql(f"INSERT INTO oc_extension (extension, type, code) VALUES "
        f"('{EXT}', 'module', 'lowstock'), ('{EXT}', 'shipping', 'zonerate');")

    # The shipping method needs settings before it will quote anything.
    sql("DELETE FROM oc_setting WHERE `code`='shipping_zonerate';")
    for key, val in [("shipping_zonerate_status", "1"),
                     ("shipping_zonerate_sort_order", "2"),
                     ("shipping_zonerate_geo_zone_id", "0"),
                     ("shipping_zonerate_tax_class_id", "9"),
                     ("shipping_zonerate_free_over", "1000"),
                     ("shipping_zonerate_rates", "0:7.50\\n3563:4.00")]:
        sql(f"INSERT INTO oc_setting (store_id, `code`, `key`, value, "
            f"serialized) VALUES (0, 'shipping_zonerate', '{key}', "
            f"'{val}', 0);")

    listed = sql(f"SELECT type, code FROM oc_extension "
                 f"WHERE extension='{EXT}';").strip().splitlines()[1:]
    print(f"  registered: {', '.join(r.replace(chr(9), ' ') for r in listed)}")
    print("  files, a manifest row, and one row per thing provided.")
    print("  Nothing OpenCart ships was edited.")


def remove() -> None:
    place(0)
    sql(f"DELETE FROM oc_extension WHERE extension='{EXT}';")
    sql(f"DELETE FROM oc_extension_install WHERE code='{EXT}';")
    sql("DELETE FROM oc_setting WHERE `code`='shipping_zonerate';")
    sql("DELETE FROM oc_module WHERE `code` LIKE 'ecadmod.%';")
    shell(f"rm -rf /var/www/html/extension/{EXT}")
    print("ecadmod removed: files, rows, settings and placements")


# -------------------------------------------------------------- placing
def place(n: int) -> None:
    """Put exactly n instances of lowstock on the home page."""
    sql(f"DELETE FROM oc_layout_module WHERE `code` LIKE 'ecadmod.lowstock%' "
        f"AND layout_id={HOME_LAYOUT};")
    sql("DELETE FROM oc_module WHERE `code` LIKE 'ecadmod.lowstock';")
    for i in range(n):
        setting = json.dumps({"name": f"Nearly gone {i + 1}",
                              "threshold": 950, "limit": 4, "status": 1})
        # Both statements in ONE call. Each sql() opens its own connection,
        # so LAST_INSERT_ID() in a separate call returns 0 -- which inserts a
        # layout row pointing at module 0, and the storefront silently
        # renders nothing at all.
        sql(f"INSERT INTO oc_module (name, `code`, setting) VALUES "
            f"('Nearly gone {i + 1}', 'ecadmod.lowstock', '{setting}');"
            f"INSERT INTO oc_layout_module "
            f"(layout_id, `code`, position, sort_order) VALUES "
            f"({HOME_LAYOUT}, CONCAT('ecadmod.lowstock.', LAST_INSERT_ID()), "
            f"'{POSITION}', {i + 1});")


# ------------------------------------------------------------- measuring
def page_ms(reps: int = 11) -> float:
    times = []
    for i in range(reps + 1):
        t0 = time.perf_counter()
        with urllib.request.urlopen(SHOP, timeout=60) as r:
            r.read()
        if i:
            times.append((time.perf_counter() - t0) * 1000)
    return statistics.median(times)


def page_queries() -> int:
    rootsql("SET GLOBAL slow_query_log=0;"
            "SET GLOBAL log_output='TABLE';"
            "TRUNCATE TABLE mysql.slow_log;"
            "SET GLOBAL long_query_time=0;"
            "SET GLOBAL slow_query_log=1;")
    urllib.request.urlopen(SHOP, timeout=60).read()
    rootsql("SET GLOBAL slow_query_log=0;")
    out = rootsql("SELECT COUNT(*) FROM mysql.slow_log;")
    try:
        return int(out.strip().splitlines()[1].split("\t")[0])
    except (IndexError, ValueError):
        return 0


def measure() -> None:
    print("\n=== What does a module cost a page? ===")
    head = (f"  {'modules':>8}{'page ms':>10}{'queries':>9}"
            f"{'ms each':>10}{'queries each':>14}")
    print(head)
    print("  " + "-" * (len(head) - 2))
    base_ms = base_q = None
    for n in (0, 1, 3, 5):
        place(n)
        with urllib.request.urlopen(SHOP, timeout=60) as r:
            body = r.read().decode("utf-8", "replace")
        shown = body.count("ecad-lowstock")
        ms, q = page_ms(), page_queries()
        if base_ms is None:
            base_ms, base_q = ms, q
            print(f"  {n:>8}{ms:>10.0f}{q:>9}{'---':>10}{'---':>14}")
        else:
            print(f"  {n:>8}{ms:>10.0f}{q:>9}"
                  f"{(ms - base_ms) / n:>10.1f}{(q - base_q) / n:>14.1f}")
        if n and shown != n:
            print(f"      (rendered {shown}, expected {n})")
    place(0)
    print("\n  Each instance is a controller call, a query and a template")
    print("  render. The cost is additive and small -- which is the argument")
    print("  for building features as modules rather than editing the core.")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--install", action="store_true")
    ap.add_argument("--remove", action="store_true")
    ap.add_argument("--place", type=int)
    ap.add_argument("--measure", action="store_true")
    a = ap.parse_args()

    wait_for_shop(quiet=True)
    if a.remove:
        remove()
        return 0
    if a.install:
        install()
    if a.place is not None:
        place(a.place)
        print(f"placed {a.place} instance(s) on the home page")
    if a.measure or not any((a.install, a.place is not None)):
        measure()
    return 0


if __name__ == "__main__":
    sys.exit(main())
