"""Chapter 12 -- a discount is a line, and the line has a position.

Every total a shop shows -- sub-total, shipping, discount, tax, grand total --
is produced by a separate extension, and each extension has a sort order. The
sort order is not cosmetic. It decides what each line is computed *from*, and
therefore how much the customer pays.

The demonstration shop puts the coupon at 4 and tax at 5, so a discount
reduces the amount tax is charged on. Swap them and the arithmetic is still
internally consistent, every page still works, and the customer pays a
different amount.

    python m12_discounts.py --order      one order with a coupon, broken down
    python m12_discounts.py --sortorder  the same basket, coupon before and
                                         after tax, side by side
    python m12_discounts.py --coupon on  enable the demo coupon (it ships
                                         expired and disabled)
    python m12_discounts.py --restore    put the sort order back and clean up

The coupon used is the one the shop ships with, code 2222, which is disabled
and expired out of the box -- a sensible thing for demonstration data to be,
and the first thing to check when a code "does not work".
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from seed import catalogue as S          # noqa: E402
from waitfor import wait_for_shop        # noqa: E402
import m10_checkout as C                 # noqa: E402

CODE = os.path.dirname(os.path.abspath(__file__))
COUPON = "2222"
DEFAULT_COUPON_SORT = 4                  # as shipped: before tax, which is 5


def sql(stmt: str) -> str:
    return S.sql(stmt)


def set_coupon(on: bool) -> None:
    if on:
        sql(f"UPDATE oc_coupon SET status=1, date_start='2020-01-01', "
            f"date_end='2030-12-31', uses_total=1000, uses_customer=1000 "
            f"WHERE code='{COUPON}';")
    else:
        sql(f"UPDATE oc_coupon SET status=0 WHERE code='{COUPON}';")


def coupon_sort(n: int) -> None:
    sql(f"UPDATE oc_setting SET value={n} "
        f"WHERE `key`='total_coupon_sort_order';")


def totals_of(order_id: str) -> list[tuple[str, str, float]]:
    out = sql(f"SELECT code, title, value FROM oc_order_total "
              f"WHERE order_id={order_id} "
              f"ORDER BY sort_order, order_total_id;")
    rows = []
    for line in out.strip().splitlines()[1:]:
        parts = line.split("\t")
        if len(parts) >= 3:
            rows.append((parts[0], parts[1], float(parts[2])))
    return rows


def place(tag: int) -> str:
    """One complete order with the coupon applied; returns its order_id."""
    c = C.Checkout(tag)
    c.get()
    c.fill_basket()
    c.register()
    # The coupon goes in BEFORE the methods are chosen, and the order is not
    # a matter of taste. Applying a coupon changes the order total, which
    # invalidates the payment method already stored in the session: the pay
    # step then answers "Payment method is incorrect!" and no order is
    # created. Nothing says so; the checkout simply stops.
    reply = c.post("extension/opencart/checkout/coupon.save",
                   {"coupon": COUPON})
    if not reply.get("success"):
        print(f"  coupon refused: {reply}", file=sys.stderr)
    c.shipping_method()
    c.payment_method()
    c.confirm()
    c.pay()
    c.success()
    out = sql("SELECT MAX(order_id) FROM oc_order;")
    return out.strip().splitlines()[-1]


def show(order_id: str, label: str) -> float:
    print(f"\n  {label}  (order {order_id})")
    rows = totals_of(order_id)
    for code, title, value in rows:
        print(f"    {title:<28}{value:>12,.2f}   [{code}]")
    return rows[-1][2] if rows else 0.0


def sortorder() -> None:
    """The same basket, with the discount line on either side of the tax."""
    print("\n=== A coupon before tax, and the same coupon after tax ===")
    set_coupon(True)

    coupon_sort(4)                       # before tax (sort 5): as shipped
    before = place(10001)
    a = show(before, "coupon at sort 4 -- BEFORE tax (shipped default)")

    coupon_sort(6)                       # after tax
    after = place(10002)
    b = show(after, "coupon at sort 6 -- AFTER tax")

    coupon_sort(DEFAULT_COUPON_SORT)
    print(f"\n  customer pays {a:,.2f} either way?  "
          f"{'yes' if abs(a - b) < 0.005 else f'NO -- {b - a:+,.2f}'}")
    print("  Both orders are internally consistent. Both add up. Only one of")
    print("  them charges tax on money the customer was never asked for.")
    print(f"\n  sort order restored to {DEFAULT_COUPON_SORT}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--order", action="store_true")
    ap.add_argument("--sortorder", action="store_true")
    ap.add_argument("--coupon", choices=("on", "off"))
    ap.add_argument("--restore", action="store_true")
    a = ap.parse_args()

    wait_for_shop(quiet=True)
    if a.coupon:
        set_coupon(a.coupon == "on")
        print(f"demo coupon {COUPON} is now {a.coupon}")
        return 0
    if a.restore:
        coupon_sort(DEFAULT_COUPON_SORT)
        set_coupon(False)
        C.clean()
        print(f"sort order {DEFAULT_COUPON_SORT}, coupon disabled")
        return 0
    if a.sortorder:
        sortorder()
        return 0

    set_coupon(True)
    show(place(10000), "one order, coupon applied")
    return 0


if __name__ == "__main__":
    sys.exit(main())
