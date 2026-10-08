"""Chapter 5 -- what the stack costs to stand up.

Four numbers the chapter quotes:

  0. from scratch -- no images at all: pull, build, install.  This is what
                     Week 3 feels like, and it is the only one of the four
                     that includes downloading 1.5 GB;
  1. cold start   -- empty volumes but the images already built, so OpenCart
                     installs itself but nothing is downloaded;
  2. warm start   -- the shop already exists, which is what a student does
                     every morning;
  3. footprint    -- how much disk the three images and two volumes occupy.

Keeping 0 and 1 apart is the point of the measurement.  They differ by more
than an order of magnitude, students meet 0 once and 2 about sixty times, and
a course that quoted a single "startup time" would be describing an experience
nobody actually has.

    python m05_coldstart.py             cold + warm (DESTROYS the shop)
    python m05_coldstart.py --scratch   also removes the images first
    python m05_coldstart.py --warm      warm start and footprint only

WARNING: without --warm this runs "docker compose down -v", which deletes the
database volume and therefore your shop.  Do it in week 3, not in week 12.
"""
import json
import subprocess
import sys
import time

from waitfor import wait_for_shop

COMPOSE = ["docker", "compose"]


def run(args, check=True):
    return subprocess.run(COMPOSE + args, capture_output=True, text=True,
                          check=check)


def timed_up(label):
    """Bring the stack up and time it to the first 200."""
    print(f"\n=== {label} ===")
    t0 = time.perf_counter()
    run(["up", "-d"])
    up_returned = time.perf_counter() - t0
    waited = wait_for_shop()
    total = time.perf_counter() - t0
    print(f"  compose up returned after   {up_returned:7.1f} s")
    print(f"  first HTTP 200 after        {total:7.1f} s")
    print(f"  spent waiting on the shop   {waited:7.1f} s "
          f"({waited / total * 100:.0f}% of the wall clock)")
    return total


def footprint():
    """Image and volume sizes, straight from the daemon."""
    print("\n=== footprint ===")
    images = subprocess.run(
        ["docker", "images", "--format", "{{.Repository}}:{{.Tag}} {{.Size}}"],
        capture_output=True, text=True, check=True).stdout
    wanted = ("ecad-shop", "mariadb", "adminer")
    rows = [ln for ln in images.splitlines()
            if any(w in ln for w in wanted)]
    for ln in sorted(rows):
        print("  image   " + ln)

    df = subprocess.run(["docker", "system", "df", "-v", "--format", "json"],
                        capture_output=True, text=True, check=False)
    try:
        vols = json.loads(df.stdout).get("Volumes", [])
        for v in vols:
            if "ecad-shop" in v.get("Name", ""):
                print(f"  volume  {v['Name']:<28} {v.get('Size', '?')}")
    except Exception:
        # Older daemons do not support --format json here; the image sizes
        # are the headline number anyway, so this is not worth failing over.
        print("  volume  (sizes unavailable on this daemon)")


def from_scratch():
    """Remove every image this stack uses, then time the whole thing."""
    print("\n=== from scratch -- no images, no volumes ===")
    run(["down", "-v", "--rmi", "all"], check=False)
    t0 = time.perf_counter()
    run(["build", "--no-cache"])
    built = time.perf_counter() - t0
    run(["up", "-d"])
    wait_for_shop()
    total = time.perf_counter() - t0
    print(f"  build (pull + compile + download)  {built:7.1f} s")
    print(f"  everything, to first HTTP 200      {total:7.1f} s")
    return total


if __name__ == "__main__":
    warm_only = "--warm" in sys.argv

    if "--scratch" in sys.argv:
        from_scratch()

    if not warm_only:
        print("Destroying the shop (down -v) so the cold start is really cold.")
        run(["down", "-v"], check=False)
        cold = timed_up("cold start -- empty volumes, images already built")
        run(["down"], check=False)
        warm = timed_up("warm start -- the shop already exists")
        print(f"\n  cold/warm ratio             {cold / warm:7.1f}x")
    else:
        run(["down"], check=False)
        timed_up("warm start -- the shop already exists")

    footprint()
