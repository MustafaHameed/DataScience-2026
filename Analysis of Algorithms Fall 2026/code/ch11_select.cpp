// ch11_select.cpp -- finding the k-th smallest without sorting.
//
// Sorting answers "what is the k-th smallest?" in Theta(n lg n). But it
// answers far more than was asked: it also puts every other element in
// its place. Selection does only what was asked, in Theta(n) expected.
//
// Measured here:
//
//   A. comparisons: quickselect against a full sort, and against the
//      2n that the expected-case analysis predicts;
//   B. the same in seconds, with std::nth_element for comparison;
//   C. the worst case -- a last-element pivot on sorted input, which is
//      Theta(n^2) for exactly Chapter 9's reason;
//   D. how the cost varies with k, which is the part people guess wrong.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch11_select.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

static long long comps = 0;
static std::mt19937 rng(90210);

// --- 1. Partition, exactly as in Chapter 9 ---------------------------
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
static size_t partition_random(std::vector<int>& a, size_t lo, size_t hi) {
    std::uniform_int_distribution<size_t> d(lo, hi - 1);
    std::swap(a[d(rng)], a[hi - 1]);
    return partition_last(a, lo, hi);
}

// --- 2. Quickselect: recurse into ONE side, not both -----------------
// This single change is the whole algorithm. Quicksort solves both
// subproblems and gets T(n) = 2T(n/2) + n = Theta(n lg n). Quickselect
// discards one and gets T(n) = T(n/2) + n, which the master theorem
// settles as Theta(n) -- case 3, the root dominating.
//
// Written as a loop rather than a recursion, because the recursive call
// is in tail position and there is no reason to spend a stack frame on
// it. That also removes Chapter 9's stack-overflow failure mode.
static int quickselect(std::vector<int>& a, size_t k, bool randomised) {
    size_t lo = 0, hi = a.size();
    for (;;) {
        if (hi - lo == 1) return a[lo];
        size_t p = randomised ? partition_random(a, lo, hi)
                              : partition_last(a, lo, hi);
        if (p == k) return a[p];
        if (k < p) hi = p;          // the k-th is on the left
        else       lo = p + 1;      // ... or on the right
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

static std::vector<int> shuffled(size_t n) {
    std::vector<int> a(n);
    std::iota(a.begin(), a.end(), 0);
    std::mt19937 r(5150);
    std::shuffle(a.begin(), a.end(), r);
    return a;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);

    // --- 3. Comparisons: selection against sorting -------------------
    // The analysis says quickselect's expected cost is about 2n when
    // looking for a median-ish element, and 4n is the standard bound
    // for the worst k. Sorting is about n lg n. Watch the gap open.
    // An EXPECTED value is being measured, so one run of a randomised
    // algorithm is not an answer: single runs of this gave 2.52, 3.92,
    // 3.73 and 2.73 comparisons per element, which says only that the
    // answer is somewhere near 3. Average over 25 independent runs.
    std::printf("A. comparisons to find the median"
                " (mean of 25 runs)\n\n");
    std::printf("%10s %16s %10s %10s %16s %10s\n",
                "n", "quickselect", "/n", "spread", "n lg n", "ratio");
    for (size_t n = 10000; n <= 10000000; n *= 10) {
        auto base = shuffled(n);
        std::vector<int> w(n);
        double sum = 0, lo = 1e30, hi = 0;
        int med = 0;
        for (int t = 0; t < 25; ++t) {
            std::copy(base.begin(), base.end(), w.begin());
            comps = 0;
            med = quickselect(w, n / 2, true);
            double per = (double)comps / n;
            sum += per; lo = std::min(lo, per); hi = std::max(hi, per);
        }
        double mean = sum / 25;
        double nlg = n * std::log2((double)n);
        std::printf("%10zu %16.0f %10.2f %5.2f-%4.2f %16.0f %10.4f\n",
                    n, mean * n, mean, lo, hi, nlg, mean * n / nlg);
        if (med != (int)(n / 2))
            std::printf("   *** wrong median: %d\n", med);
    }

    // --- 4. The same thing in seconds --------------------------------
    // std::nth_element is the library's version: introselect, which is
    // quickselect with a median-of-medians fallback, exactly as
    // std::sort is quicksort with a heapsort fallback.
    std::printf("\nB. time to find the median (ms)\n\n");
    std::printf("%10s %12s %14s %12s %10s\n",
                "n", "quickselect", "nth_element", "full sort", "speedup");
    for (size_t n = 100000; n <= 6400000; n *= 2) {
        auto base = shuffled(n);
        std::vector<int> w(n);
        double tq = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            quickselect(w, n / 2, true);
        }, 5);
        double tn = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            std::nth_element(w.begin(), w.begin() + n / 2, w.end());
        }, 5);
        double ts = time_ms([&] {
            std::copy(base.begin(), base.end(), w.begin());
            std::sort(w.begin(), w.end());
        }, 5);
        std::printf("%10zu %12.2f %14.2f %12.2f %10.2f\n",
                    n, tq, tn, ts, ts / tq);
    }

    // --- 5. The worst case, for Chapter 9's reason -------------------
    // A last-element pivot on sorted input shrinks the range by one
    // each time: T(n) = T(n-1) + n = Theta(n^2). The loop form means
    // it is slow rather than a crash, which is an improvement but not
    // a fix.
    std::printf("\nC. worst case: last-element pivot, SORTED input\n\n");
    std::printf("%10s %16s %12s %16s %12s\n",
                "n", "naive comps", "/(n^2/2)", "random comps", "/n");
    for (size_t n = 2000; n <= 32000; n *= 2) {
        std::vector<int> s(n);
        std::iota(s.begin(), s.end(), 0);
        std::vector<int> a = s;
        comps = 0;
        quickselect(a, n / 2, false);
        long long cn = comps;
        std::vector<int> b = s;
        comps = 0;
        quickselect(b, n / 2, true);
        std::printf("%10zu %16lld %12.4f %16lld %12.2f\n",
                    n, cn, cn / ((double)n * n / 2.0), comps,
                    (double)comps / n);
    }

    // --- 6. Does k matter? -------------------------------------------
    // The expected cost is 2n for the median and smaller towards the
    // ends -- but not dramatically so, because the first partition
    // costs n whatever k is. That first n is a floor nothing removes.
    std::printf("\nD. cost against k, n = 1,000,000 (mean of 25 runs)\n\n");
    std::printf("%10s %14s %16s %10s\n", "k/n", "k", "comparisons", "/n");
    {
        size_t n = 1000000;
        auto base = shuffled(n);
        std::vector<int> w(n);
        for (double frac : {0.0, 0.01, 0.1, 0.25, 0.5, 0.75, 0.99, 1.0}) {
            size_t k = (size_t)(frac * (n - 1));
            double sum = 0;
            for (int t = 0; t < 25; ++t) {
                std::copy(base.begin(), base.end(), w.begin());
                comps = 0;
                quickselect(w, k, true);
                sum += (double)comps / n;
            }
            std::printf("%10.2f %14zu %16.0f %10.2f\n",
                        frac, k, sum / 25 * n, sum / 25);
        }
    }
    return 0;
}
