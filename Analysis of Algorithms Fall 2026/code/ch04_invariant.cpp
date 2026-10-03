// ch04_invariant.cpp -- why the THIRD obligation is not decoration.
//
// A loop invariant proof has three parts: initialisation, maintenance
// and termination. Students check the first two and wave at the third.
// This program exhibits a binary search with one character changed --
//
//        lo = mid;        instead of        lo = mid + 1;
//
// -- whose invariant ("if x is in A at all, it is in A[lo..hi]") is
// PERFECTLY VALID. Initialisation holds. Maintenance holds, and this
// program audits it at every iteration and never catches it failing.
//
// The consequence, which is the lesson: the broken search NEVER
// RETURNS A WRONG ANSWER. Not once in 20,000 trials. What it does
// instead is sometimes fail to return at all. An invariant buys you
// "if it answers, the answer is correct" -- partial correctness --
// and termination is a genuinely separate obligation, not a formality
// to wave at on the way past.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch04_invariant.cpp
#include <algorithm>
#include <cstdio>
#include <random>
#include <vector>

static const int CAP = 10000;   // iteration cap standing in for "forever"

// --- 1. The correct version, and the invariant it maintains ---------
// Invariant: if x occurs in A, then it occurs in A[lo..hi].
//   Initialisation: lo = 0, hi = n-1, so the range is all of A.
//   Maintenance:    A is sorted, so discarding the half that cannot
//                   contain x preserves the invariant.
//   Termination:    hi - lo strictly decreases and is bounded below,
//                   so the loop ends; when lo > hi the range is empty
//                   and the invariant says x is not in A.
static int bsearch_ok(const std::vector<int>& a, int x) {
    int lo = 0, hi = (int)a.size() - 1;
    while (lo <= hi) {
        int mid = lo + (hi - lo) / 2;
        if (a[mid] == x) return mid;
        if (a[mid] < x) lo = mid + 1;   // the correct step
        else            hi = mid - 1;
    }
    return -1;
}

// --- 2. The same search with one character changed ------------------
// "lo = mid" still discards nothing that could hold x, so the
// invariant survives. But when hi == lo + 1 and a[mid] < x, mid
// computes to lo, so lo = mid leaves lo and hi UNCHANGED and the loop
// makes no progress. It is not wrong; it is non-terminating.
//
// Returns -2 to mean "hit the iteration cap", which is this program's
// stand-in for looping forever.
static int bsearch_bug(const std::vector<int>& a, int x,
                       bool* inv_ever_broken, long long* iters) {
    int lo = 0, hi = (int)a.size() - 1;
    for (int step = 0; step < CAP; ++step) {
        if (lo > hi) return -1;
        int mid = lo + (hi - lo) / 2;

        // Check the invariant itself, every single iteration: if x is
        // present in the whole array, is it still inside [lo, hi]?
        bool present = std::binary_search(a.begin(), a.end(), x);
        if (present) {
            bool in_window = false;
            for (int i = lo; i <= hi; ++i)
                if (a[i] == x) { in_window = true; break; }
            if (!in_window) *inv_ever_broken = true;
        }
        ++*iters;

        if (a[mid] == x) return mid;
        if (a[mid] < x) lo = mid;       // <-- the one-character bug
        else            hi = mid - 1;
    }
    return -2;
}

int main() {
    std::mt19937 rng(7);
    const int N = 64, TRIALS = 10000;

    // A sorted array of the even numbers, so that every odd number in
    // range is a key that is definitely ABSENT.
    std::vector<int> a(N);
    for (int i = 0; i < N; ++i) a[i] = 2 * i;

    // --- 3. Ten thousand tests with keys that are present ----------
    // This is the test suite a careless author writes, because keys
    // taken from the array are the easy ones to generate.
    bool broken = false;
    long long iters = 0;
    int fail_present = 0, wrong_present = 0;
    std::uniform_int_distribution<int> pick(0, N - 1);
    for (int t = 0; t < TRIALS; ++t) {
        int x = a[pick(rng)];
        int r = bsearch_bug(a, x, &broken, &iters);
        if (r == -2) ++fail_present;
        else if (r < 0 || a[r] != x) ++wrong_present;
    }
    std::printf("A. %d random tests, keys always PRESENT\n", TRIALS);
    std::printf("   non-termination : %d\n", fail_present);
    std::printf("   wrong answer    : %d\n", wrong_present);

    // --- 4. The same ten thousand tests, keys absent ----------------
    // The only change is the distribution the keys are drawn from.
    int fail_absent = 0, wrong_absent = 0;
    std::uniform_int_distribution<int> odd(0, N - 2);
    for (int t = 0; t < TRIALS; ++t) {
        int x = 2 * odd(rng) + 1;              // odd: never in a
        int r = bsearch_bug(a, x, &broken, &iters);
        if (r == -2) ++fail_absent;
        else if (r != -1) ++wrong_absent;
    }
    std::printf("\nB. %d random tests, keys always ABSENT\n", TRIALS);
    std::printf("   non-termination : %d\n", fail_absent);
    std::printf("   wrong answer    : %d\n", wrong_absent);

    // --- 5. And the invariant? It never broke once. -----------------
    std::printf("\nC. the invariant, checked at every iteration\n");
    std::printf("   iterations audited : %lld\n", iters);
    std::printf("   invariant violated : %s\n", broken ? "YES" : "NO");
    std::printf("   wrong answers, A+B : %d\n",
                wrong_present + wrong_absent);
    std::printf("   -> Initialisation and maintenance both hold, so the\n");
    std::printf("      search is never WRONG. That is what an invariant\n");
    std::printf("      buys you, and it is all it buys.\n");
    std::printf("      Termination is what fails: %d of %d runs never\n",
                fail_present + fail_absent, 2 * TRIALS);
    std::printf("      finished. Partial correctness is not correctness.\n");

    // --- 6. Why proving beats testing, counted ----------------------
    // Exhaustive testing means every sorted array and every key. The
    // number of sorted arrays of length n over an alphabet of size m
    // is C(m+n-1, n) -- multiset coefficient. Count it for a tiny case.
    std::printf("\nD. the size of the input space\n");
    for (int n : {4, 8, 16, 20}) {
        const int m = 100;
        double c = 1.0;
        for (int i = 1; i <= n; ++i) c *= (double)(m + n - i) / i;
        std::printf("   n=%2d, values 0..99 : %.3g sorted arrays"
                    " x %d keys\n", n, c, m);
    }
    std::printf("   A proof covers all of them in half a page.\n");
    return 0;
}
