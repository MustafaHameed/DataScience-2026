#!/usr/bin/env python3
"""Compile and run every experiment, and check its figures against the handout.

This is what makes the measured-reality strand honest rather than decorative:
a number quoted in a worked example should still be produced by the program it
came from. The script does three things.

  1. Compiles every chNN_*.cpp with the flags Appendix A records.
  2. Runs each one and captures its output.
  3. For each chapter, extracts the numbers the handout quotes in its worked
     examples and looks for them in that chapter's output.

Step 3 is deliberately loose. Timings are machine-specific and will not match;
what must match are the EXACT quantities -- comparison counts, copy counts,
ratios that theory predicts, table sizes. Those are listed per chapter in
EXACT below, and a mismatch there is a real failure.

    python run_all.py              compile, run and check everything
    python run_all.py --ch 9 12    only those chapters
    python run_all.py --list       show what is checked, run nothing
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
VCVARS = (r"C:\Program Files (x86)\Microsoft Visual Studio\2022"
          r"\BuildTools\VC\Auxiliary\Build\vcvars64.bat")

# Exact quantities the handout quotes, which do not depend on the machine.
# Each entry: (description, regex that must match the program's output).
EXACT: dict[int, list[tuple[str, str]]] = {
    2: [("trial division ratios near 2",
         r"\n\s*34\s+17179869143\s+[\d.]+\s+(?:2\.0\d|1\.9\d)")],
    4: [("invariant never violated", r"invariant violated\s*:\s*NO"),
        ("no wrong answers", r"wrong answers, A\+B\s*:\s*0"),
        ("1,048,576 input-space row", r"n=20, values 0\.\.99 : 2\.4\de\+22")],
    5: [("balanced tree: 11 levels at n=1024", r"levels = 11,"),
        ("every level does 1024 work", r"\n\s*10\s+1024\s+1024\s+1\.000")],
    6: [("case 1 quotient -> 2.0", r"\n\s*1024\s+2146435072\s+\d+\s+1\.999"),
        ("case 3 quotient -> 2.0", r"\n\s*1048576\s+2199022206976\s+\d+\s+2\.0000")],
    7: [("all three agree", r"n= 500   cubic=    1362  quad=    1362"
                           r"  linear=    1362  agree")],
    8: [("sorted comparisons exact", r"sorted\s+[\d.]+\s+20769984"),
        ("random comparisons exact", r"random\s+[\d.]+\s+39348350"),
        ("sorts correctly", r"top-down yes, branchless yes")],
    9: [("n(n-1)/2 on sorted input", r"\n\s*16000\s+127992000\s+0\.9999"),
        ("depth is n-1", r"\n\s*16000\s+15999\s+\d+\s+14\.0"),
        ("64 bytes per frame", r"bytes per frame\s*:\s*64\.0")],
    10: [("merge sort within 2% of the floor",
          r"\n\s*1000000\s+18672689\s+18488885\s+1\.0%"),
         ("all three sorted the same", r"same answer: yes")],
    11: [("naive is exactly 3n^2/8", r"\n\s*32000\s+383992000\s+0\.7500")],
    12: [("heap height is floor(lg n)", r"100000000\s+26\s+26\s+yes"),
         ("sorted BST height is exactly n", r"256000\s+18\.0\s+\d+\s+[\d.]+\s+256000\s+1\.000")],
    13: [("chain lengths match Poisson", r"\n\s*0\s+36\d{4}\s+367879\s+1\.00"),
         ("probes = 1 + alpha/2", r"\n\s*8\.00\s+5\.00\d\s+5\.000")],
    14: [("earliest-finish is optimal on every trial", r"20000\s+100\.0%"),
         ("greedy coin change fails at 6",
          r"first at 6 \(greedy 3, optimal 2\)"),
         ("Huffman within [H, H+1)", r"uniform bytes   H =  7\.995\d <=  8\.0000")],
    15: [("naive calls are exactly 2^n", r"\n\s*26\s+67108864\s+352"),
         ("all three agree", r"n= 25  naive=   421  memo=   421  bottom-up=   421  agree")],
    16: [("DP agrees with brute force", r"disagreements: 0"),
         ("LCS of ALGORITHM and ALTRUISTIC", r"= 5, \"ALRIT\""),
         ("one more bit doubles the time", r"\n\s*16777216\s+25\s+")],
    17: [("doubling copies ~ n", r"\n\s*64000\s+65535\s+1\.024"),
         ("grow-by-one is exactly n^2/2", r"\n\s*64000\s+65535\s+1\.024\s+2047968000\s+1\.0000"),
         ("worst single append", r"MOST EXPENSIVE append  : 524288 copies")],
    19: [("Kruskal and Prim agree", r"disagreements: 0"),
         ("union-find stays near 1.8", r"\n\s*5120000\s+3\.\d+\s+1\.8\d\s+4")],
    20: [("array Dijkstra is wrong on vertex 3",
          r"\n\s*3\s+4\s+3\s+3\s+3\s+<-- array WRONG"),
         ("negative cycle detected", r"negative cycle added: detected = yes")],
    21: [("all three agree", r"disagreements: 0"),
         ("naive hits exactly n*m", r"200000\s+1000\s+199001000\s+0\.9950"),
         ("Rabin-Karp worst case is n*m", r"200000\s+1000\s+199001000\s+0\.9950")],
    22: [("reduction preserves answers", r"3000 random graphs: answers differ in 0")],
    23: [("harness refuses to time a wrong algorithm",
          r"2000 instances across 5 shapes: all agree")],
}


def sources() -> list[tuple[int, str]]:
    out = []
    for f in sorted(os.listdir(HERE)):
        m = re.match(r"ch(\d\d)_.*\.cpp$", f)
        if m:
            out.append((int(m.group(1)), f))
    return out


def build_and_run(src: str) -> tuple[bool, str]:
    """Compile with Appendix A's flags and run. Returns (ok, output)."""
    stem = src[:-4]
    bat = os.path.join(HERE, "_runall.bat")
    with open(bat, "w", encoding="ascii") as fh:
        fh.write("@echo off\n")
        fh.write('call "%s" >nul 2>&1\n' % VCVARS)
        fh.write('cd /d "%s"\n' % HERE)
        fh.write("cl /nologo /O2 /EHsc /std:c++17 %s.cpp >nul\n" % stem)
        fh.write("if errorlevel 1 exit /b 1\n")
        fh.write('"%s\\%s.exe"\n' % (HERE, stem))
    try:
        r = subprocess.run([bat], capture_output=True, text=True, timeout=3600)
        return r.returncode == 0, r.stdout
    except subprocess.TimeoutExpired:
        return False, "*** TIMED OUT ***"
    finally:
        if os.path.exists(bat):
            os.remove(bat)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--ch", type=int, nargs="*", help="only these chapters")
    ap.add_argument("--list", action="store_true", help="show checks only")
    a = ap.parse_args()

    srcs = sources()
    if a.ch:
        srcs = [(n, f) for n, f in srcs if n in a.ch]

    if a.list:
        for n, f in srcs:
            checks = EXACT.get(n, [])
            print(f"  ch {n:2d}  {f:24s} {len(checks)} exact check(s)")
            for d, _ in checks:
                print(f"          - {d}")
        return 0

    print(f"Running {len(srcs)} experiment(s) with /O2 /EHsc /std:c++17\n")
    failed, checked, passed = [], 0, 0
    for n, f in srcs:
        t0 = time.time()
        ok, out = build_and_run(f)
        dt = time.time() - t0
        if not ok:
            print(f"  ch {n:2d}  {f:24s} *** BUILD OR RUN FAILED ***")
            failed.append(f)
            continue
        bad = []
        for desc, pat in EXACT.get(n, []):
            checked += 1
            if re.search(pat, out):
                passed += 1
            else:
                bad.append(desc)
        mark = "ok" if not bad else "MISMATCH"
        print(f"  ch {n:2d}  {f:24s} {dt:7.1f}s  "
              f"{len(EXACT.get(n, []))} check(s)  {mark}")
        for d in bad:
            print(f"          - did not find: {d}")
            failed.append(f"{f}: {d}")

    print(f"\n  {passed}/{checked} exact checks passed")
    if failed:
        print(f"  {len(failed)} failure(s)")
        return 1
    print("  every experiment reproduces the figures the handout quotes")
    return 0


if __name__ == "__main__":
    sys.exit(main())
