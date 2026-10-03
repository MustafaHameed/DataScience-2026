// ch03_fit.cpp -- what Theta(g) actually asserts, tested.
//
// Theta(g(n)) says T(n)/g(n) settles on a positive constant. That is a
// testable claim, and it is testable against NEARBY exponents too: if
// insertion sort really is Theta(n^2), then
//
//     T(n)/n^2.0   should be roughly CONSTANT
//     T(n)/n^1.9   should GROW without bound
//     T(n)/n^2.1   should SHRINK towards zero
//
// So the experiment does not merely confirm the right exponent; it rules
// out the two exponents either side of it. Same for merge sort against
// n*lg(n), with n and n^1.2 as the rival hypotheses.
//
// Each column is normalised to 1.000 at the smallest n, because the
// constant itself is machine-specific and uninteresting -- the question
// is only whether the column stays put.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch03_fit.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

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

static void insertion_sort(std::vector<int>& a) {
    for (size_t j = 1; j < a.size(); ++j) {
        int key = a[j];
        size_t i = j;
        while (i > 0 && a[i - 1] > key) { a[i] = a[i - 1]; --i; }
        a[i] = key;
    }
}

static void merge_sort(std::vector<int>& a, std::vector<int>& buf,
                       size_t lo, size_t hi) {
    if (hi - lo < 2) return;
    size_t mid = lo + (hi - lo) / 2;
    merge_sort(a, buf, lo, mid);
    merge_sort(a, buf, mid, hi);
    size_t i = lo, j = mid, k = lo;
    while (i < mid && j < hi) buf[k++] = (a[i] <= a[j]) ? a[i++] : a[j++];
    while (i < mid) buf[k++] = a[i++];
    while (j < hi)  buf[k++] = a[j++];
    std::copy(buf.begin() + lo, buf.begin() + hi, a.begin() + lo);
}

static std::vector<int> random_input(size_t n, unsigned seed) {
    std::mt19937 rng(seed);
    std::vector<int> a(n);
    std::iota(a.begin(), a.end(), 0);
    std::shuffle(a.begin(), a.end(), rng);
    return a;
}

int main() {
    // --- 1. Insertion sort against three candidate exponents ---------
    // Only one of the three columns can stay constant, and which one
    // does is the empirical content of "Theta(n^2)".
    // Normalising to the FIRST row makes that row's noise everybody's
    // noise, so the first row must not be a sub-millisecond timing:
    // start at n = 4000 and take the median of nine runs.
    std::printf("A. insertion sort: T(n) divided by each candidate g(n),\n");
    std::printf("   each column normalised to 1.000 at n = 4000\n\n");
    std::printf("%8s %10s %10s %10s %10s\n",
                "n", "T (ms)", "T/n^1.9", "T/n^2.0", "T/n^2.1");
    double b19 = 0, b20 = 0, b21 = 0;
    for (size_t n = 4000; n <= 128000; n *= 2) {
        auto base = random_input(n, 2024);
        std::vector<int> work(n);
        double t = time_ms([&] {
            std::copy(base.begin(), base.end(), work.begin());
            insertion_sort(work);
        }, 9);
        double x = static_cast<double>(n);
        double r19 = t / std::pow(x, 1.9);
        double r20 = t / std::pow(x, 2.0);
        double r21 = t / std::pow(x, 2.1);
        if (b19 == 0) { b19 = r19; b20 = r20; b21 = r21; }
        std::printf("%8zu %10.2f %10.3f %10.3f %10.3f\n",
                    n, t, r19 / b19, r20 / b20, r21 / b21);
    }

    // --- 2. Merge sort against n*lg(n) and its two rivals ------------
    // n*lg(n) is only a whisker above n, so this is a far harder test
    // than part A, and it needs better hygiene to pass.
    //
    // The first draft of this program allocated a fresh vector inside
    // the timed region. At these sizes that allocation and copy cost
    // more than it measures -- the T/(n lg n) column came out as
    // 1.000, 1.134, 0.804, 1.342, 0.839, which says nothing at all.
    //
    // The fix is to allocate ONCE outside the timing and refill with
    // std::copy, then measure the refill on its own and subtract it.
    // What is left is the sort.
    std::printf("\nB. merge sort: same test against n, n*lg(n), n^1.2\n");
    std::printf("   (the O(n) refill is measured and subtracted)\n\n");
    std::printf("%9s %9s %9s %9s %11s %9s\n",
                "n", "copy", "total", "T=sort", "T/n", "T/(n lg n)");
    double c1 = 0, c2 = 0, c3 = 0;
    std::vector<double> col3;
    for (size_t n = 100000; n <= 6400000; n *= 2) {
        auto base = random_input(n, 2024);
        std::vector<int> work(n), buf(n);

        double tc = time_ms([&] {
            std::copy(base.begin(), base.end(), work.begin());
        }, 11);
        double tt = time_ms([&] {
            std::copy(base.begin(), base.end(), work.begin());
            merge_sort(work, buf, 0, n);
        }, 11);
        double t = tt - tc;

        double x = static_cast<double>(n);
        double r1 = t / x;
        double r2 = t / (x * std::log2(x));
        double r3 = t / std::pow(x, 1.2);
        if (c1 == 0) { c1 = r1; c2 = r2; c3 = r3; }
        col3.push_back(r3 / c3);
        std::printf("%9zu %9.2f %9.2f %9.2f %11.3f %9.3f\n",
                    n, tc, tt, t, r1 / c1, r2 / c2);
    }
    std::printf("\n   T/n^1.2 for comparison:");
    for (double v : col3) std::printf(" %.3f", v);
    std::printf("\n");

    // --- 3. Where an asymptotically worse function is still smaller --
    // Theta is a statement about a limit, and the limit can be a long
    // way off. Report the crossover rather than pretending there is
    // none: 100 n lg n < n^2 requires 100 lg n < n.
    std::printf("\nC. 100 n lg n versus n^2: where is the crossover?\n\n");
    std::printf("%8s %16s %16s %s\n", "n", "100 n lg n", "n^2", "");
    for (size_t n : {200u, 400u, 800u, 900u, 950u, 1000u, 1100u, 1600u}) {
        double x = static_cast<double>(n);
        double f = 100.0 * x * std::log2(x), g = x * x;
        std::printf("%8zu %16.0f %16.0f %s\n", n, f, g,
                    (f < g) ? "n lg n ahead" : "n^2 still ahead");
    }
    for (size_t n = 2; n < 100000; ++n) {
        double x = static_cast<double>(n);
        if (100.0 * x * std::log2(x) < x * x) {
            std::printf("\n   crossover: n lg n first wins at n = %zu\n", n);
            break;
        }
    }
    return 0;
}
