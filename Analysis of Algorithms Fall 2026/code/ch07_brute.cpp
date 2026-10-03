// ch07_brute.cpp -- brute force, and the two things that beat it.
//
// The maximum-subarray problem: given an array of signed numbers, find the
// contiguous block with the largest sum. It has a famous ladder of
// solutions, and the ladder is the whole of algorithm design in miniature:
//
//   cubic      try every (i, j) and re-add the block        Theta(n^3)
//   quadratic  try every (i, j), but carry a running sum    Theta(n^2)
//   linear     Kadane: one pass, one invariant              Theta(n)
//
// Each step removes redundant work rather than making the loop faster.
// That is the difference between optimising a program and choosing a
// better algorithm, and the timings below show what each is worth.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch07_brute.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <limits>
#include <random>
#include <vector>

// --- 1. Brute force, as literally as it can be written ---------------
// For every start i and end j, add up a[i..j] from scratch. The inner
// re-addition is pure waste, and it is the n^3.
static long long max_sub_cubic(const std::vector<int>& a) {
    long long best = std::numeric_limits<long long>::min();
    size_t n = a.size();
    for (size_t i = 0; i < n; ++i)
        for (size_t j = i; j < n; ++j) {
            long long s = 0;
            for (size_t k = i; k <= j; ++k) s += a[k];
            best = std::max(best, s);
        }
    return best;
}

// --- 2. The same search, with the waste removed ----------------------
// Extending the block from j to j+1 costs one addition, not a fresh
// traversal. The SEARCH is unchanged -- still every (i, j) pair -- so
// this is not yet a better idea, only a better accountant.
static long long max_sub_quadratic(const std::vector<int>& a) {
    long long best = std::numeric_limits<long long>::min();
    size_t n = a.size();
    for (size_t i = 0; i < n; ++i) {
        long long s = 0;
        for (size_t j = i; j < n; ++j) {
            s += a[j];
            best = std::max(best, s);
        }
    }
    return best;
}

// --- 3. Kadane: a different idea, not a faster loop ------------------
// Invariant: after processing a[0..j], "here" is the largest sum of a
// block ENDING at j, and "best" is the largest sum anywhere in a[0..j].
// Maintenance: a block ending at j+1 either extends the best block
// ending at j, or starts fresh at j+1. That one observation removes a
// whole dimension from the search.
static long long max_sub_linear(const std::vector<int>& a) {
    long long best = std::numeric_limits<long long>::min(), here = 0;
    for (int x : a) {
        here = std::max((long long)x, here + x);
        best = std::max(best, here);
    }
    return best;
}

template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    for (int r = 0; r < reps; ++r) {
        auto t0 = std::chrono::steady_clock::now();
        volatile long long v = f();
        (void)v;
        auto t1 = std::chrono::steady_clock::now();
        t.push_back(
            std::chrono::duration<double, std::milli>(t1 - t0).count());
    }
    std::sort(t.begin(), t.end());
    return t[t.size() / 2];
}

static std::vector<int> input(size_t n, unsigned seed) {
    std::mt19937 rng(seed);
    std::uniform_int_distribution<int> d(-100, 100);
    std::vector<int> a(n);
    for (auto& x : a) x = d(rng);
    return a;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);

    // --- 4. All three agree, which is what makes the race fair -------
    // A faster algorithm that gets a different answer is not faster.
    std::printf("A. do the three agree?\n\n");
    for (size_t n : {1u, 2u, 10u, 100u, 500u}) {
        auto a = input(n, 99);
        long long x = max_sub_cubic(a), y = max_sub_quadratic(a),
                  z = max_sub_linear(a);
        std::printf("   n=%4zu   cubic=%8lld  quad=%8lld  linear=%8lld  %s\n",
                    n, x, y, z,
                    (x == y && y == z) ? "agree" : "*** DISAGREE ***");
    }

    // --- 5. The ratios each order of growth predicts -----------------
    // Doubling n should multiply the times by 8, 4 and 2.
    std::printf("\nB. doubling n: predicted ratios 8, 4, 2\n\n");
    std::printf("%8s %11s %7s %11s %7s %11s %7s\n",
                "n", "cubic", "ratio", "quad", "ratio", "linear", "ratio");
    double pc = 0, pq = 0, pl = 0;
    for (size_t n = 500; n <= 8000; n *= 2) {
        auto a = input(n, 7);
        double tc = (n <= 4000) ? time_ms([&] {
            return max_sub_cubic(a); }, 3) : 0.0;
        double tq = time_ms([&] { return max_sub_quadratic(a); }, 15);
        double tl = time_ms([&] { return max_sub_linear(a); }, 101);
        std::printf("%8zu", n);
        if (tc > 0) std::printf(" %11.2f", tc); else std::printf(" %11s", "-");
        if (pc && tc) std::printf(" %7.2f", tc / pc);
        else std::printf(" %7s", "-");
        std::printf(" %11.3f", tq);
        if (pq) std::printf(" %7.2f", tq / pq); else std::printf(" %7s", "-");
        std::printf(" %11.4f", tl);
        if (pl) std::printf(" %7.2f", tl / pl); else std::printf(" %7s", "-");
        std::printf("\n");
        pc = tc; pq = tq; pl = tl;
    }

    // --- 6. Where brute force is the right answer --------------------
    // The cubic version is slower by a factor that grows without bound.
    // That is not the same as being unusable: at the small end its
    // ABSOLUTE cost is a few microseconds, and it is four lines you can
    // get right first time and never have to debug. So report the
    // absolute cost, not a ratio against a linear time too small to
    // measure -- and say where the cost crosses thresholds a person
    // would actually notice.
    std::printf("\nC. what brute force costs in absolute terms\n\n");
    std::printf("%8s %16s %22s\n", "n", "cubic", "verdict");
    for (size_t n : {10u, 30u, 100u, 300u, 1000u, 3000u}) {
        auto a = input(n, 7);
        double us = time_ms([&] { return max_sub_cubic(a); },
                            n <= 300 ? 201 : 5) * 1000.0;
        const char* v = us < 1000.0        ? "imperceptible"
                        : us < 100000.0    ? "a visible pause"
                                           : "too slow to ship";
        if (us < 1000.0)
            std::printf("%8zu %13.1f us %22s\n", n, us, v);
        else
            std::printf("%8zu %13.1f ms %22s\n", n, us / 1000.0, v);
    }
    return 0;
}
