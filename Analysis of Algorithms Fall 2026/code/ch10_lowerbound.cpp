// ch10_lowerbound.cpp -- the comparison lower bound, and sorting past it.
//
// Theorem: any sorting algorithm that learns about its input ONLY by
// comparing pairs must make Omega(n lg n) comparisons in the worst case.
// The proof is a decision tree with n! leaves and height >= lg(n!).
//
// Three things are measured here:
//
//   A. lg(n!) against n lg n, so the bound's shape is visible, and
//      Stirling's n lg n - 1.44n as the sharper form;
//   B. how close a real comparison sort gets to that floor -- merge sort
//      is within a few per cent of information-theoretic optimality;
//   C. counting sort and radix sort, which beat the bound outright, and
//      the assumption each one had to buy to do it.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch10_lowerbound.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

static long long comps = 0;

// --- 1. lg(n!) exactly, by summing logs ------------------------------
// n! overflows at n = 21, so the factorial is never formed: the log of
// a product is the sum of the logs, and that stays in range for ever.
static double lg_factorial(long long n) {
    double s = 0;
    for (long long k = 2; k <= n; ++k) s += std::log2((double)k);
    return s;
}

// --- 2. A comparison sort, counting every comparison -----------------
// Merge sort, because Chapter 8 proved it Theta(n lg n) in every case.
// The question here is how close its CONSTANT is to the floor.
static void merge_c(std::vector<int>& a, std::vector<int>& buf,
                    size_t lo, size_t mid, size_t hi) {
    size_t i = lo, j = mid, k = lo;
    while (i < mid && j < hi) {
        ++comps;
        buf[k++] = (a[i] <= a[j]) ? a[i++] : a[j++];
    }
    while (i < mid) buf[k++] = a[i++];
    while (j < hi)  buf[k++] = a[j++];
    std::copy(buf.begin() + lo, buf.begin() + hi, a.begin() + lo);
}
static void msort_c(std::vector<int>& a, std::vector<int>& buf,
                    size_t lo, size_t hi) {
    if (hi - lo < 2) return;
    size_t mid = lo + (hi - lo) / 2;
    msort_c(a, buf, lo, mid);
    msort_c(a, buf, mid, hi);
    merge_c(a, buf, lo, mid, hi);
}

// --- 3. Counting sort: no comparisons at all -------------------------
// It never asks "is x < y". It asks "how many keys equal v", which is a
// different question and not covered by the theorem. The price is the
// assumption that keys are integers in a known, small range [0, K).
static void counting_sort(const std::vector<int>& in, std::vector<int>& out,
                          int K) {
    std::vector<int> cnt(K, 0);
    for (int x : in) ++cnt[x];
    for (int v = 1; v < K; ++v) cnt[v] += cnt[v - 1];
    for (size_t i = in.size(); i-- > 0;) out[--cnt[in[i]]] = in[i];
}

// --- 4. Radix sort: counting sort, one digit at a time ---------------
// Lifts counting sort's "small range" assumption to "fixed width": 32-bit
// keys in four passes of an 8-bit counting sort. Still no comparisons,
// so still outside the theorem.
static void radix_sort(std::vector<int>& a, std::vector<int>& tmp) {
    const int BITS = 8, R = 1 << BITS;
    size_t n = a.size();
    for (int shift = 0; shift < 32; shift += BITS) {
        std::vector<size_t> cnt(R + 1, 0);
        for (size_t i = 0; i < n; ++i)
            ++cnt[((unsigned)a[i] >> shift & (R - 1)) + 1];
        for (int v = 0; v < R; ++v) cnt[v + 1] += cnt[v];
        for (size_t i = 0; i < n; ++i)
            tmp[cnt[(unsigned)a[i] >> shift & (R - 1)]++] = a[i];
        a.swap(tmp);
    }
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

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);

    // --- 5. The bound itself, and Stirling's sharper form ------------
    // lg(n!) >= n lg n - n lg e = n lg n - 1.4427n. The second column
    // is the crude bound, the third the accurate one.
    std::printf("A. the information-theoretic floor\n\n");
    std::printf("%10s %16s %16s %16s %10s\n",
                "n", "lg(n!)", "n lg n", "n lg n - 1.443n", "lg(n!)/n lg n");
    for (long long n = 1000; n <= 1000000; n *= 10) {
        double lgf = lg_factorial(n);
        double nlg = n * std::log2((double)n);
        std::printf("%10lld %16.0f %16.0f %16.0f %10.4f\n",
                    n, lgf, nlg, nlg - 1.442695 * n, lgf / nlg);
    }

    // --- 6. How close does a real sort get to the floor? -------------
    // Merge sort is Theta(n lg n); so is the bound. The question is the
    // constant, and the answer is "within a few per cent".
    std::printf("\nB. merge sort against the floor\n\n");
    std::printf("%10s %16s %16s %12s\n",
                "n", "comparisons", "lg(n!)", "excess");
    std::mt19937 rng(31337);
    for (size_t n = 1000; n <= 1000000; n *= 10) {
        std::vector<int> a(n);
        std::iota(a.begin(), a.end(), 0);
        std::shuffle(a.begin(), a.end(), rng);
        std::vector<int> buf(n);
        comps = 0;
        msort_c(a, buf, 0, n);
        double lgf = lg_factorial((long long)n);
        std::printf("%10zu %16lld %16.0f %11.1f%%\n",
                    n, comps, lgf, 100.0 * (comps - lgf) / lgf);
    }

    // --- 7. Beating the bound, by not comparing ----------------------
    // Counting and radix sort are not counter-examples to the theorem.
    // They are outside it: they never compare two keys. What they pay
    // instead is an assumption about what the keys ARE.
    std::printf("\nC. sorting without comparing (ms)\n\n");
    std::printf("%10s %12s %12s %12s %12s\n",
                "n", "std::sort", "counting", "radix", "best speedup");
    const int K = 1000;
    for (size_t n = 100000; n <= 6400000; n *= 2) {
        std::uniform_int_distribution<int> d(0, K - 1);
        std::vector<int> base(n);
        for (auto& x : base) x = d(rng);
        std::vector<int> w(n), out(n), tmp(n);

        double ts = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            std::sort(w.begin(), w.end());
        }, 5);
        double tc = time_ms([&] {
            counting_sort(base, out, K);
        }, 5);
        double tr = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            radix_sort(w, tmp);
        }, 5);
        std::printf("%10zu %12.1f %12.1f %12.1f %12.2f\n",
                    n, ts, tc, tr, ts / std::min(tc, tr));
    }

    // --- 8. And they really did sort ---------------------------------
    {
        std::uniform_int_distribution<int> d(0, K - 1);
        std::vector<int> base(100000), out(100000), tmp(100000);
        for (auto& x : base) x = d(rng);
        counting_sort(base, out, K);
        std::vector<int> r = base;
        radix_sort(r, tmp);
        std::printf("\n   counting sorted: %s;  radix sorted: %s;"
                    "  same answer: %s\n",
                    std::is_sorted(out.begin(), out.end()) ? "yes" : "NO",
                    std::is_sorted(r.begin(), r.end()) ? "yes" : "NO",
                    (out == r) ? "yes" : "NO");
    }
    return 0;
}
