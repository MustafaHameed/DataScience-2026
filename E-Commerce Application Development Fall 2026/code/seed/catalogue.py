"""Fill the catalogue with N products, reproducibly.

The demonstration catalogue has nineteen products, which is enough to learn
the admin panel and far too few to learn anything about what a catalogue
costs. Chapter 6 needs fifty thousand.

Products are inserted with SQL rather than through the admin panel, for the
obvious reason and for a less obvious one: going through the panel would
measure the panel. What Chapter 6 is asking about is the shape of the data,
so the data is put there directly and the storefront is then asked to serve
it, exactly as it would serve a catalogue somebody had typed in.

Everything is derived from a fixed seed, so two students with the same N get
the same catalogue and can compare numbers.

    python seed/catalogue.py 50000          replace the catalogue with 50,000
    python seed/catalogue.py 50000 --keep   add to what is there
    python seed/catalogue.py --restore      back to the 19 demo products
"""
import argparse
import os
import random
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
CODE = os.path.dirname(HERE)

# The category the seeded products are filed under. 20 ("Desktops") is the
# demo catalogue's largest, so it is the one whose page is worth timing.
CATEGORY = 20
LANGUAGE = 1

# The DEFAULT store is 0, not 1. Getting this wrong is silent and expensive:
# the storefront's every query filters on
# `p2s.store_id = config_store_id`, which is 0, so products written with
# store_id = 1 exist in the database, appear in the admin panel, and are
# invisible to every customer. The first version of this file used 1, and the
# Chapter 6 measurements were taken against a category page that was showing
# the nineteen demo products and none of the fifty thousand seeded ones.
# Nothing failed; the page just quietly meant something else.
STORE = 0

# Below this id the demo products live; at or above it, ours. Keeping them
# apart is what makes --restore a one-line DELETE rather than a guess.
BASE_ID = 100000

ADJ = ["Classic", "Modern", "Compact", "Deluxe", "Essential", "Premium",
       "Rugged", "Slim", "Portable", "Industrial", "Vintage", "Smart"]
NOUN = ["Keyboard", "Monitor", "Chair", "Lamp", "Cable", "Adapter", "Case",
        "Stand", "Mouse", "Hub", "Charger", "Headset", "Webcam", "Router"]


def sql(statements: str, quiet=True) -> str:
    """Run SQL in the database container, reading from stdin.

    Piped in rather than passed with -e: a 50,000-row insert is megabytes of
    SQL and will not fit in a command line on any platform.
    """
    cmd = ["docker", "compose", "exec", "-T", "db",
           "mariadb", "-ushop", "-pshoppw", "opencart"]
    r = subprocess.run(cmd, cwd=CODE, input=statements, text=True,
                       capture_output=True)
    if r.returncode != 0 and not quiet:
        print(r.stderr[:2000], file=sys.stderr)
    return r.stdout


def restore() -> None:
    print(f"Removing every product with id >= {BASE_ID}")
    sql(f"""
        DELETE FROM oc_product_to_category WHERE product_id >= {BASE_ID};
        DELETE FROM oc_product_to_store    WHERE product_id >= {BASE_ID};
        DELETE FROM oc_product_description WHERE product_id >= {BASE_ID};
        DELETE FROM oc_product             WHERE product_id >= {BASE_ID};
    """)
    print(counts())


def counts() -> str:
    out = sql("SELECT COUNT(*) FROM oc_product;"
              "SELECT COUNT(*) FROM oc_product_to_category "
              f"WHERE category_id = {CATEGORY};")
    nums = [ln for ln in out.split() if ln.isdigit()]
    if len(nums) >= 2:
        return (f"  catalogue: {int(nums[0]):,} products; "
                f"{int(nums[1]):,} in category {CATEGORY}")
    return "  (counts unavailable)"


def seed(n: int, keep: bool) -> None:
    if not keep:
        restore()
    rng = random.Random(20260601)
    t0 = time.perf_counter()

    # One INSERT per 2,000 rows. One statement for all 50,000 exceeds the
    # server's max_allowed_packet; one statement per row takes minutes,
    # because the cost is per round trip and not per row.
    chunk = 2000
    done = 0
    while done < n:
        k = min(chunk, n - done)
        prod, desc, store, cat = [], [], [], []
        for i in range(done, done + k):
            pid = BASE_ID + i
            name = (f"{rng.choice(ADJ)} {rng.choice(NOUN)} "
                    f"{rng.randint(100, 999)}")
            price = round(rng.uniform(150, 48000), 2)
            qty = rng.randint(0, 500)
            sort = rng.randint(0, 999)
            prod.append(
                f"({pid},'SEED-{i}',{qty},7,'',0,1,{price},0,9,"
                f"'2026-01-01',0.5,1,0,0,0,1,1,1,{sort},1,NOW(),NOW())")
            safe = name.replace("'", "")
            desc.append(
                f"({pid},{LANGUAGE},'{safe}','<p>A seeded product.</p>',"
                f"'','{safe}','Seeded product {i}','')")
            store.append(f"({pid},{STORE})")
            cat.append(f"({pid},{CATEGORY})")

        sql(
            "INSERT INTO oc_product (product_id,model,quantity,"
            "stock_status_id,image,manufacturer_id,shipping,price,points,"
            "tax_class_id,date_available,weight,weight_class_id,length,"
            "width,height,length_class_id,subtract,minimum,sort_order,"
            "status,date_added,date_modified) VALUES "
            + ",".join(prod) + ";\n"
            "INSERT INTO oc_product_description (product_id,language_id,"
            "name,description,tag,meta_title,meta_description,meta_keyword)"
            " VALUES " + ",".join(desc) + ";\n"
            "INSERT INTO oc_product_to_store (product_id,store_id) VALUES "
            + ",".join(store) + ";\n"
            "INSERT INTO oc_product_to_category (product_id,category_id)"
            " VALUES " + ",".join(cat) + ";\n")

        done += k
        print(f"\r  seeded {done:,} / {n:,}", end="", flush=True)

    print(f"\n  {n:,} products in {time.perf_counter() - t0:.1f} s")
    print(counts())


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("n", nargs="?", type=int, default=0)
    ap.add_argument("--keep", action="store_true")
    ap.add_argument("--restore", action="store_true")
    a = ap.parse_args()
    if a.restore:
        restore()
    elif a.n:
        seed(a.n, a.keep)
    else:
        print(counts())
