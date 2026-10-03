// ch09_quick.cpp -- quicksort, and why randomising is not superstition.
//
// Quicksort is Theta(n^2) in the worst case and Theta(n lg n) expected.
// Both numbers are true, and which one you get depends on something you
// may not control: the ORDER THE DATA ARRIVES IN.
//
// This program shows the gap, and then shows randomisation closing it:
//
//   A. a last-element pivot on already-sorted input is the worst case,
//      and it is not a contrived input -- it is the commonest input
//      shape in practice;
//   B. the same sort, same input, with a random pivot;
//   C. comparisons counted rather than timed, so the n^2 and the n lg n
//      can be seen directly;
//   D. recursion depth, which is what actually crashes a program.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch09_quick.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

static long long comps = 0;
static int maxdepth = 0;
static std::mt19937 rng(20261003);

// --- 1. Lomuto partition, the textbook one ---------------------------
// Everything <= the pivot is swept to the left. Returns where the pivot
// ended up. Linear, and it is where every comparison happens.
static size_t partition_last(std::vector<int>& a, size_t lo, size_t hi) {
    int pivot = a[hi - 1];
    size_t i = lo;
    for (size_t j = lo; j + 1 < hi; ++j) {
        ++comps;
        if (a[j] <= pivot) std::swap(a[i++], a[j]);
    }
    std::swap(a[i], a[hi - 1]);
    return i;
}

// --- 2. The same partition, after moving a random element to the end -
// One extra swap. That is the entire cost of the change, and it buys an
// expected-time bound that holds for EVERY input -- no assumption about
// the data at all (Chapter 2's distinction between an average over
// inputs and an expectation over the algorithm's own coins).
static size_t partition_random(std::vector<int>& a, size_t lo, size_t hi) {
    std::uniform_int_distribution<size_t> d(lo, hi - 1);
    std::swap(a[d(rng)], a[hi - 1]);
    return partition_last(a, lo, hi);
}

// The two stack marks let part D measure the size of one frame: record
// where a local lives at depth 0 and at the deepest depth reached, and
// the distance between them, divided by the depth, is bytes per frame.
static char* stack_top = nullptr;
static char* stack_deep = nullptr;

static void qsort_last(std::vector<int>& a, size_t lo, size_t hi, int d) {
    char here;
    if (d == 0) stack_top = &here;
    if (d > maxdepth) { maxdepth = d; stack_deep = &here; }
    if (hi - lo < 2) return;
    size_t p = partition_last(a, lo, hi);
    qsort_last(a, lo, p, d + 1);
    qsort_last(a, p + 1, hi, d + 1);
}

static void qsort_rand(std::vector<int>& a, size_t lo, size_t hi, int d) {
    maxdepth = std::max(maxdepth, d);
    if (hi - lo < 2) return;
    size_t p = partition_random(a, lo, hi);
    qsort_rand(a, lo, p, d + 1);
    qsort_rand(a, p + 1, hi, d + 1);
}

template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    for (int r = 0; r < reps; ++r) {
        auto t0 = std::chrono::steady_clock::now();
        f();
        auto t1 = std::chrono::steady_clock::now();
        t.push_back(
            std::chrono::duration<double, std::milli>(t1 - t0).count());
    }
    std::sort(t.begin(), t.end());
    return t[t.size() / 2];
}

static std::vector<int> sorted_in(size_t n) {
    std::vector<int> a(n);
    std::iota(a.begin(), a.end(), 0);
    return a;
}
static std::vector<int> random_in(size_t n) {
    std::vector<int> a = sorted_in(n);
    std::mt19937 r(777);
    std::shuffle(a.begin(), a.end(), r);
    return a;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);

    // --- 3. Comparisons counted: the two bounds, seen directly -------
    // On sorted input the last-element pivot splits n into 0 and n-1
    // every single time, so the count is 1+2+...+(n-1) = n(n-1)/2
    // exactly. Compare that with the random pivot on the same input.
    std::printf("A. comparisons on ALREADY-SORTED input\n\n");
    std::printf("%8s %16s %12s %16s %12s %12s\n", "n",
                "last-pivot", "/(n^2/2)", "random-pivot", "/(n lg n)",
                "speedup");
    for (size_t n = 1000; n <= 16000; n *= 2) {
        auto a = sorted_in(n);
        comps = 0; maxdepth = 0;
        qsort_last(a, 0, n, 0);
        long long c_last = comps;
        auto b = sorted_in(n);
        comps = 0; maxdepth = 0;
        qsort_rand(b, 0, n, 0);
        long long c_rand = comps;
        double nn = (double)n;
        std::printf("%8zu %16lld %12.4f %16lld %12.4f %12.1f\n",
                    n, c_last, c_last / (nn * nn / 2.0),
                    c_rand, c_rand / (nn * std::log2(nn)),
                    (double)c_last / c_rand);
    }

    // --- 4. Recursion depth, which is what actually crashes ----------
    // The worst case is not only slow. Depth n means n stack frames, and
    // a few hundred thousand of those is a crash, not a slow program.
    std::printf("\nB. recursion depth on ALREADY-SORTED input\n\n");
    std::printf("%8s %14s %14s %14s\n",
                "n", "last-pivot", "random-pivot", "lg n");
    for (size_t n = 1000; n <= 16000; n *= 2) {
        auto a = sorted_in(n);
        comps = 0; maxdepth = 0;
        qsort_last(a, 0, n, 0);
        int d_last = maxdepth;
        auto b = sorted_in(n);
        comps = 0; maxdepth = 0;
        qsort_rand(b, 0, n, 0);
        std::printf("%8zu %14d %14d %14.1f\n",
                    n, d_last, maxdepth, std::log2((double)n));
    }

    // --- 5. And the same thing in seconds ----------------------------
    // The counts above are exact; these are what they cost.
    std::printf("\nC. time on sorted input (ms)\n\n");
    std::printf("%8s %14s %7s %14s %7s\n",
                "n", "last-pivot", "ratio", "random", "ratio");
    // Capped at 16,000 deliberately: see part D for why 32,000 is not an
    // option on a default stack.
    double pl = 0, pr = 0;
    for (size_t n = 2000; n <= 16000; n *= 2) {
        auto base = sorted_in(n);
        std::vector<int> w(n);
        double tl = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            maxdepth = 0; qsort_last(w, 0, n, 0);
        }, 5);
        double tr = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            maxdepth = 0; qsort_rand(w, 0, n, 0);
        }, 5);
        std::printf("%8zu %14.3f", n, tl);
        if (pl) std::printf(" %7.2f", tl / pl); else std::printf(" %7s", "-");
        std::printf(" %14.3f", tr);
        if (pr) std::printf(" %7.2f", tr / pr); else std::printf(" %7s", "-");
        std::printf("\n");
        pl = tl; pr = tr;
    }

    // --- 6. The worst case does not just run slowly; it crashes ------
    // Depth n-1 means n-1 live stack frames. Measure how big one frame
    // actually is by looking at where a local variable sits at depth 0
    // and at the deepest depth, then work out how many fit in the
    // default 1 MB stack. This is a prediction that can be checked, and
    // it was: at n = 32,000 this program died with 0xC00000FD,
    // STATUS_STACK_OVERFLOW, which is why part C stops at 16,000.
    std::printf("\nD. the stack, measured\n\n");
    {
        size_t n = 16000;
        auto a = sorted_in(n);
        comps = 0; maxdepth = 0;
        qsort_last(a, 0, n, 0);
        double span = (double)(stack_top - stack_deep);
        double per = span / maxdepth;
        const double LIMIT = 1024.0 * 1024.0;      // MSVC default stack
        std::printf("   depth reached        : %d\n", maxdepth);
        std::printf("   stack span           : %.0f bytes\n", span);
        std::printf("   bytes per frame      : %.1f\n", per);
        std::printf("   default stack        : %.0f bytes (1 MB)\n", LIMIT);
        std::printf("   frames that fit      : %.0f\n", LIMIT / per);
        std::printf("   so sorted input over : n = %.0f crashes\n",
                    LIMIT / per);
        std::printf("   observed             : n = 32000 crashed"
                    " (0xC00000FD)\n");
    }

    // --- 7. On RANDOM input the naive pivot is fine ------------------
    // Which is the trap: the defect is invisible until the day someone
    // feeds your sort data that is already in order, and that day comes.
    std::printf("\nE. the same two sorts on RANDOM input (ms)\n\n");
    std::printf("%8s %14s %14s %12s\n",
                "n", "last-pivot", "random-pivot", "ratio");
    for (size_t n = 100000; n <= 800000; n *= 2) {
        auto base = random_in(n);
        std::vector<int> w(n);
        double tl = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            maxdepth = 0; qsort_last(w, 0, n, 0);
        }, 5);
        double tr = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            maxdepth = 0; qsort_rand(w, 0, n, 0);
        }, 5);
        std::printf("%8zu %14.2f %14.2f %12.2f\n", n, tl, tr, tl / tr);
    }
    return 0;
}
