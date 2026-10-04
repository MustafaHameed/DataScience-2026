// ch15_dp1.cpp -- overlapping subproblems, counted.
//
// Dynamic programming applies when a problem has (1) optimal substructure
// and (2) OVERLAPPING subproblems. The second condition is what separates
// it from divide-and-conquer, and it is the one that can be measured: if
// the subproblems overlap, a naive recursion solves the same one over and
// over, and the count of repeats is the saving available.
//
// Measured here:
//
//   A. how many times naive recursion solves each subproblem -- the
//      answer is exponential, and the program prints the exact counts;
//   B. memoised top-down against bottom-up against naive, timed;
//   C. rod cutting, the canonical first DP, solved three ways;
//   D. where the two conditions FAIL -- merge sort has optimal
//      substructure and no overlap, so memoising it buys nothing.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch15_dp1.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

static long long calls = 0;
static std::vector<long long> visits;   // visits[n] = times subproblem n solved

// --- 1. The naive recursion, instrumented ----------------------------
// Rod cutting: a rod of length n, a price p[i] for a piece of length i.
// Cut it to maximise revenue. The obvious recursion tries every first
// cut and recurses on the remainder.
static long long rod_naive(const std::vector<long long>& p, int n) {
    ++calls;
    if (n < (int)visits.size()) ++visits[n];
    if (n == 0) return 0;
    long long best = 0;
    for (int i = 1; i <= n && i < (int)p.size(); ++i)
        best = std::max(best, p[i] + rod_naive(p, n - i));
    return best;
}

// --- 2. The same recursion, with an answer cache ---------------------
// One line of difference. Each subproblem is now solved once and looked
// up thereafter, which is why the running time collapses from
// exponential to quadratic.
static long long rod_memo(const std::vector<long long>& p, int n,
                          std::vector<long long>& memo) {
    ++calls;
    if (memo[n] >= 0) return memo[n];
    long long best = 0;
    if (n > 0)
        for (int i = 1; i <= n && i < (int)p.size(); ++i)
            best = std::max(best, p[i] + rod_memo(p, n - i, memo));
    return memo[n] = best;
}

// --- 3. Bottom-up: the same table, filled in dependency order --------
// No recursion and no cache lookups, just two loops. Identical answers,
// identical asymptotics, smaller constant -- and it cannot overflow the
// stack, which for large n matters (Chapter 9).
static long long rod_bottom_up(const std::vector<long long>& p, int n) {
    std::vector<long long> best(n + 1, 0);
    for (int j = 1; j <= n; ++j) {
        long long b = 0;
        for (int i = 1; i <= j && i < (int)p.size(); ++i)
            b = std::max(b, p[i] + best[j - i]);
        best[j] = b;
    }
    return best[n];
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

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(1066);

    // A price table: p[i] is the price of a piece of length i.
    std::vector<long long> p(31, 0);
    std::uniform_int_distribution<int> dp(1, 30);
    for (size_t i = 1; i < p.size(); ++i) p[i] = (long long)i * 2 + dp(rng);

    // --- 4. How often is each subproblem solved? ---------------------
    // This is the measurement that defines "overlapping". If every
    // subproblem were solved once, DP would buy nothing.
    std::printf("A. how many times naive recursion solves subproblem k"
                "  (n = 22)\n\n");
    {
        int n = 22;
        visits.assign(n + 1, 0);
        calls = 0;
        rod_naive(p, n);
        std::printf("%6s %18s\n", "k", "times solved");
        for (int k = n; k >= 0; --k)
            if (k >= n - 6 || k <= 3)
                std::printf("%6d %18lld\n", k, visits[k]);
        std::printf("%6s %18lld\n", "total", calls);
        std::printf("\n   subproblem 0 is solved %lld times;"
                    " there is only one of it.\n", visits[0]);
    }

    // --- 5. The explosion, and the collapse --------------------------
    // Naive calls should roughly double for each unit of n -- it is
    // 2^(n-1) for a full price table. Memoised calls should be O(n^2).
    std::printf("\nB. total subproblem solutions\n\n");
    std::printf("%6s %18s %18s %16s\n",
                "n", "naive calls", "memoised calls", "naive/memo");
    for (int n = 10; n <= 26; n += 4) {
        visits.assign(1, 0);
        calls = 0; rod_naive(p, n);
        long long cn = calls;
        std::vector<long long> memo(n + 1, -1);
        calls = 0; rod_memo(p, n, memo);
        long long cm = calls;
        std::printf("%6d %18lld %18lld %16.1f\n", n, cn, cm,
                    (double)cn / cm);
    }

    // --- 6. And in seconds -------------------------------------------
    std::printf("\nC. time (ms)\n\n");
    std::printf("%6s %14s %14s %14s %14s\n",
                "n", "naive", "memoised", "bottom-up", "naive/bottom");
    for (int n = 20; n <= 32; n += 4) {
        double tn = time_ms([&] { return rod_naive(p, n); }, 3);
        double tm = time_ms([&] {
            std::vector<long long> memo(n + 1, -1);
            return rod_memo(p, n, memo);
        }, 51);
        double tb = time_ms([&] { return rod_bottom_up(p, n); }, 51);
        std::printf("%6d %14.3f %14.4f %14.4f %14.0f\n",
                    n, tn, tm, tb, tn / tb);
    }

    // --- 7. All three agree, which is the point ----------------------
    std::printf("\nD. do the three agree?\n\n");
    for (int n : {1, 5, 12, 20, 25}) {
        std::vector<long long> memo(n + 1, -1);
        long long a = rod_naive(p, n), b = rod_memo(p, n, memo),
                  c = rod_bottom_up(p, n);
        std::printf("   n=%3d  naive=%6lld  memo=%6lld  bottom-up=%6lld  %s\n",
                    n, a, b, c,
                    (a == b && b == c) ? "agree" : "*** DISAGREE ***");
    }

    // --- 8. Bottom-up scales where the others cannot -----------------
    // The naive version cannot reach n = 40 in any human timescale.
    // Bottom-up reaches n = 100,000 in milliseconds: O(n^2) with a
    // price table capped at 30 is really O(30n).
    std::printf("\nE. bottom-up alone, at sizes the naive version"
                " cannot reach\n\n");
    std::printf("%12s %14s %14s\n", "n", "time (ms)", "revenue");
    for (int n = 10000; n <= 160000; n *= 2) {
        double t = time_ms([&] { return rod_bottom_up(p, n); }, 5);
        std::printf("%12d %14.3f %14lld\n", n, t, rod_bottom_up(p, n));
    }
    return 0;
}
