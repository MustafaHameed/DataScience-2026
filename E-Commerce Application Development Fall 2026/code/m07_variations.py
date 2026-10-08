"""Chapter 7 -- what options cost, and what variants cost instead.

OpenCart can express "this shirt comes in three colours and four sizes" two
ways, and they are not the same system:

  OPTIONS   one product, with option values that MODIFY it. Red is +Rs.0,
            Large is +Rs.200. Twelve combinations are not stored anywhere;
            they are computed when the customer chooses.

  VARIANTS  twelve products, each a row of its own with master_id pointing
            at the parent. Every combination exists, can hold its own SKU,
            its own stock and its own price.

The first is cheap and cannot track stock per combination. The second tracks
everything and grows multiplicatively. This measures both, so the choice can
be made on evidence rather than on which one the admin panel shows first.

    python m07_variations.py             the full run
    python m07_variations.py --quick     the two smallest option sets

WARNING: replaces the catalogue. seed/catalogue.py --restore afterwards.
"""
import argparse
import os
import statistics
import sys
import time
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402

BASE = "http://localhost:8090/index.php?route=product/product&product_id="
PID = S.BASE_ID                          # the product we attach options to

# (colours, sizes, materials) -- 2, 3 and 4 option sets, growing.
SHAPES = [(3, 0, 0), (3, 4, 0), (3, 4, 5), (6, 8, 5)]


def sql(stmt: str) -> str:
    return S.sql(stmt)


def fetch(url: str, reps: int = 7) -> tuple[float, int]:
    """(median ms, bytes) for a page."""
    times, size = [], 0
    for i in range(reps + 1):
        t0 = time.perf_counter()
        with urllib.request.urlopen(url, timeout=120) as r:
            body = r.read()
        if i:
            times.append((time.perf_counter() - t0) * 1000)
            size = len(body)
    return statistics.median(times), size


def clear_options() -> None:
    sql(f"DELETE FROM oc_product_option_value WHERE product_id = {PID};"
        f"DELETE FROM oc_product_option       WHERE product_id = {PID};")


def add_options(counts: tuple[int, ...]) -> int:
    """Attach option sets of the given sizes. Returns combinations implied."""
    clear_options()
    combos = 1
    # option_id 11 is the demo "Size" select; 5 and 1 are the generic
    # "Select" and "Radio". Three real select options is enough to show the
    # shape, and reusing the demo options keeps this readable in the admin
    # panel afterwards, which is where students will look.
    for slot, (option_id, n) in enumerate(zip((11, 5, 1), counts)):
        if not n:
            continue
        combos *= n
        poid = 900000 + slot
        sql(f"INSERT INTO oc_product_option "
            f"(product_option_id, product_id, option_id, value, required) "
            f"VALUES ({poid}, {PID}, {option_id}, '', 1);")
        rows = ",".join(
            f"({poid * 100 + i}, {poid}, {PID}, {option_id}, "
            f"{(option_id * 10) + i}, 100, 1, {i * 50}.0000, '+', 0, '+', "
            f"0.0, '+')"
            for i in range(n))
        sql("INSERT INTO oc_product_option_value (product_option_value_id,"
            "product_option_id, product_id, option_id, option_value_id,"
            "quantity, subtract, price, price_prefix, points, points_prefix,"
            "weight, weight_prefix) VALUES " + rows + ";")
    return combos


def option_rows() -> tuple[int, int]:
    out = sql(f"SELECT (SELECT COUNT(*) FROM oc_product_option "
              f"WHERE product_id={PID}), "
              f"(SELECT COUNT(*) FROM oc_product_option_value "
              f"WHERE product_id={PID});")
    nums = [x for x in out.split() if x.isdigit()]
    return (int(nums[0]), int(nums[1])) if len(nums) >= 2 else (0, 0)


def variant_rows(combos: int) -> int:
    """How many oc_product rows the same catalogue needs AS VARIANTS."""
    return combos


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true")
    a = ap.parse_args()
    shapes = SHAPES[:2] if a.quick else SHAPES

    wait_for_shop(quiet=True)
    S.seed(1, keep=False)                # one product, id = S.BASE_ID
    url = BASE + str(PID)
    fetch(url, reps=3)                   # warm the opcache, not the result

    print("\n=== options: one product, values that modify it ===")
    print(f"{'option sets':>12}  {'combinations':>12}  {'option rows':>11}"
          f"  {'value rows':>10}  {'page ms':>8}  {'page KB':>8}")
    for counts in shapes:
        combos = add_options(counts)
        nopt, nval = option_rows()
        ms, size = fetch(url)
        label = "x".join(str(c) for c in counts if c)
        print(f"{label:>12}  {combos:>12,}  {nopt:>11}  {nval:>10}"
              f"  {ms:>8.0f}  {size / 1024:>8.1f}")

    print("\n=== the same catalogue expressed as variants ===")
    print(f"{'option sets':>12}  {'combinations':>12}  "
          f"{'oc_product rows':>16}  {'description rows':>17}")
    for counts in shapes:
        combos = 1
        for c in counts:
            if c:
                combos *= c
        label = "x".join(str(c) for c in counts if c)
        # One product row and one description row per combination, plus the
        # master. This is the number that explodes.
        print(f"{label:>12}  {combos:>12,}  {variant_rows(combos) + 1:>16,}"
              f"  {variant_rows(combos) + 1:>17,}")

    clear_options()
    return 0


if __name__ == "__main__":
    sys.exit(main())
