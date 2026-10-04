# The measured-reality strand

One program per chapter. Each prints the tables that appear in that chapter's
worked examples, and `run_all.py` checks that it still does.

```powershell
python run_all.py            # compile, run and check all 23
python run_all.py --ch 9 12  # just those chapters
python run_all.py --list     # what is checked, run nothing
```

---

## The machine these numbers came from

| | |
|---|---|
| OS | Windows 11 Pro, 64-bit |
| Compiler | MSVC 19.44 (VS 2022 Build Tools, toolset 14.44) |
| Flags | `/O2 /EHsc /std:c++17` |

**Your absolute times will differ. Your ratios should not.** Appendix A of the
handout has the setup, the portability notes for GCC/Clang, and a
symptom-to-cause table for when the numbers look wrong.

---

## What each program measures

| File | Chapter | Measures |
|---|---|---|
| `ch01_growth.cpp` | 1 | Insertion vs merge sort over a doubling sequence; the crossover |
| `ch02_cases.cpp` | 2 | Three input shapes; trial division, where size = bit length |
| `ch03_fit.cpp` | 3 | $T/g(n)$ against three candidate $g$; excluding the rival exponents |
| `ch04_invariant.cpp` | 4 | A binary search whose invariant holds and which never terminates |
| `ch05_rectree.cpp` | 5 | Recursion trees counted level by level; balanced vs lopsided |
| `ch06_master.cpp` | 6 | The master theorem's three cases, counted exactly |
| `ch07_brute.cpp` | 7 | Maximum subarray in $n^3$, $n^2$ and $n$ |
| `ch08_merge.cpp` | 8 | Input shape, the cutoff, and a branchless merge |
| `ch09_quick.cpp` | 9 | Quicksort's worst case, its depth, and the stack it exhausts |
| `ch10_lowerbound.cpp` | 10 | $\lg(n!)$; merge sort against the floor; counting and radix sort |
| `ch11_select.cpp` | 11 | Quickselect against a full sort; cost against $k$ |
| `ch12_structures.cpp` | 12 | Heap height vs BST height; what a degenerate tree costs |
| `ch13_hashing.cpp` | 13 | Chain lengths vs Poisson; load factor; a stride that defeats a hash |
| `ch14_greedy.cpp` | 14 | Three greedy rules against an oracle; Huffman vs entropy |
| `ch15_dp1.cpp` | 15 | How many times naive recursion re-solves each subproblem |
| `ch16_dp2.cpp` | 16 | LCS, edit distance; knapsack's pseudo-polynomial wall |
| `ch17_amortised.cpp` | 17 | Doubling vs grow-by-one; the rare expensive append |
| `ch18_graphs.cpp` | 18 | Matrix vs list across densities; bit-packing |
| `ch19_mst.cpp` | 19 | Three union–find variants; Kruskal against Prim |
| `ch20_sp.cpp` | 20 | Dijkstra, Bellman–Ford, Floyd–Warshall; a negative edge |
| `ch21_strings.cpp` | 21 | Naive, Rabin–Karp, KMP, and the inputs that break two of them |
| `ch22_complexity.cpp` | 22 | Verifying against finding; meet-in-the-middle; a reduction |
| `ch23_harness.cpp` | 23 | The reusable measurement harness for the semester project |

---

## Which results are machine-specific

This matters for marking: a student reproducing these will match some columns
exactly and none of the timings.

### Exact — identical on any machine, any compiler

These are **counts**, not times. `run_all.py` checks all of them.

- Ch 4: zero invariant violations in 132,311,650 iterations; zero wrong answers
- Ch 5: 11 levels at $n=1024$, each doing exactly 1024 units
- Ch 6: case 1 and case 3 quotients converging to exactly 2.0000
- Ch 8: 20,769,984 comparisons sorted, 39,348,350 random
- Ch 9: $n(n-1)/2$ comparisons; depth exactly $n-1$
- Ch 10: merge sort 1.0% above $\lg(n!)$ at $n=10^6$
- Ch 11: naive selection exactly $3n^2/8$
- Ch 12: heap height exactly $\lfloor\lg n\rfloor$; sorted BST height exactly $n$
- Ch 13: probes $= 1 + \alpha/2$ to three decimals
- Ch 14: earliest-finish optimal on 20,000/20,000; coin change fails first at 6
- Ch 15: naive calls exactly $2^n$
- Ch 17: grow-by-one copies exactly $n^2/2$; worst append 524,288 copies
- Ch 19: Kruskal and Prim agree on 2,000/2,000 graphs
- Ch 21: naive and Rabin–Karp worst cases exactly $nm$
- Ch 22: the reduction preserves answers on 3,000/3,000 graphs

### Reproducible in shape, not in value

Ratios should match within a few per cent on any reasonable machine:

- insertion sort's $\approx 4$ and merge sort's $\approx 2.2$ on doubling (Ch 1)
- trial division's $\approx 2$ per two bits (Ch 2)
- knapsack's $\approx 2$ per bit of $W$ (Ch 16)
- union–find's flat $\approx 1.8$ hops (Ch 19)

### Machine-specific — do not expect to match

- Every millisecond figure
- Ch 8's best cutoff (32 here) — a property of this cache and element size
- Ch 9's 64 bytes per stack frame — this compiler, these flags
- Ch 18's crossover density (0.26) — moves with the implementation
- Ch 20's finding that the heap beat the array at every density

---

## Two things that bit, recorded so they do not bite again

**Timing below the clock's resolution.** Ch 22's verification takes ~15 ns
against a 100 ns clock; timing it per call reported `0.0000` and a ratio of
infinity. Time a million iterations and divide. Ch 17's first timing table
reported ratios of 0.92, 1.19, 1.81 and 4.04 for a *linear* algorithm, all of
it noise, for the same reason.

**Sampling in insertion order.** Ch 13's first load-factor table probed the
first-inserted keys, which sit at the front of their chains, and reported 1.096
probes at *every* load factor — a perfectly flat column that looked like a
result. Sample uniformly at random.

Both lessons, and eight more, are collected in Chapter 23 and implemented in
`ch23_harness.cpp`.

---

## Conventions

- Step comments `// --- n.` match the numbered excerpt in the handout's lab
  box, contiguously from 1. `checklabs.py` enforces this.
- Listings in the handout are **abridged**: the excerpt shows the parts worth
  reading and elides the printing boilerplate. The `.cpp` file is the complete,
  runnable version.
- No program needs input files or arguments.
- All output goes to stdout, unbuffered, with its own headings.
- Everything is portable C++17 except `_BitScanForward64` in `ch18_graphs.cpp`;
  Appendix A gives the GCC/Clang replacement.
