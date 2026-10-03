// ch05_rectree.cpp -- a recursion tree, counted rather than drawn.
//
// A recursion tree is a bookkeeping device: it claims that the work of a
// recursive algorithm can be summed level by level. That claim is checkable.
// This program instruments two recursions and prints, for each level, how
// many calls it holds and how much work they do.
//
//   A. T(n) = 2T(n/2) + n   -- merge sort's shape. Every level should do
//      about the same total work, and there should be about lg(n) levels.
//
//   B. T(n) = T(n/3) + T(2n/3) + n  -- a LOPSIDED split. The tree is no
//      longer uniform: the shallowest leaf is at depth log_3(n) and the
//      deepest at log_{3/2}(n), and the levels stop being full in between.
//      The total is still O(n lg n), and this is where you can see why.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch05_rectree.cpp
#include <algorithm>
#include <cmath>
#include <cstdio>
#include <vector>

static std::vector<long long> work;    // work[d] = work done at depth d
static std::vector<long long> calls;   // calls[d] = calls made at depth d

static void note(int d, long long w) {
    if ((int)work.size() <= d) { work.resize(d + 1, 0); calls.resize(d + 1, 0); }
    work[d] += w;
    calls[d] += 1;
}

// --- 1. The balanced recursion, T(n) = 2T(n/2) + n -------------------
// The body does n units of work and then halves. Nothing is sorted here:
// the point is the SHAPE, so the work is just counted, not performed.
static void balanced(long long n, int d) {
    if (n <= 1) { note(d, 1); return; }
    note(d, n);
    balanced(n / 2, d + 1);
    balanced(n - n / 2, d + 1);
}

// --- 2. The lopsided recursion, T(n) = T(n/3) + T(2n/3) + n ----------
// The same n units of work per call, but the split is one-third to
// two-thirds. The left spine bottoms out at log_3(n); the right spine
// runs on to log_{3/2}(n), which is where the depth of the tree comes
// from -- and the bound with it.
// The max(1, ...) is not decoration. Written as lopsided(n/3) and
// lopsided(n - n/3), the case n = 2 gives 2/3 = 0 and 2 - 0 = 2, so the
// larger branch recurses on 2 for ever: the measure n never decreases and
// the recursion does not terminate. That is Chapter 4's third obligation,
// failing in a program four lines long.
static void lopsided(long long n, int d) {
    if (n <= 1) { note(d, 1); return; }
    note(d, n);
    long long small = std::max(1LL, n / 3);
    lopsided(small, d + 1);
    lopsided(n - small, d + 1);
}

static void report(const char* title, long long n,
                   void (*f)(long long, int)) {
    work.assign(1, 0);
    calls.assign(1, 0);
    f(n, 0);
    long long total = 0;
    for (long long w : work) total += w;

    std::printf("\n%s   n = %lld\n\n", title, n);
    std::printf("%6s %10s %14s %10s\n",
                "depth", "calls", "work", "work/n");
    for (size_t d = 0; d < work.size(); ++d)
        std::printf("%6zu %10lld %14lld %10.3f\n",
                    d, calls[d], work[d], (double)work[d] / n);
    std::printf("%6s %10s %14lld %10.3f\n",
                "TOTAL", "", total, (double)total / n);
    std::printf("   levels = %zu,  n lg n = %.0f,  total/(n lg n) = %.3f\n",
                work.size(), n * std::log2((double)n),
                total / (n * std::log2((double)n)));
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    // --- 3. Each level of a balanced tree does the same work ---------
    // The "work/n" column is the one to read: if the recursion tree
    // argument is sound, it is 1.000 all the way down until the leaves.
    report("A. balanced:  T(n) = 2T(n/2) + n", 1024, balanced);

    // --- 4. A lopsided tree, where the levels stop being full --------
    // Now "work/n" holds at 1.000 only while every branch is still
    // alive -- down to depth log_3(n) -- and then decays. The tree is
    // deeper than the balanced one and still sums to O(n lg n).
    report("B. lopsided:  T(n) = T(n/3) + T(2n/3) + n", 1024, lopsided);

    // --- 5. Does the depth grow the way the algebra says? ------------
    // Balanced depth should be lg(n); lopsided depth should be
    // log_{3/2}(n) = lg(n)/lg(1.5), about 1.71 times as deep.
    std::printf("\nC. depth against the prediction\n\n");
    std::printf("%10s %8s %8s %8s %8s %8s\n",
                "n", "bal", "lg n", "lop", "log1.5", "lop/bal");
    std::fflush(stdout);
    for (long long n = 1024; n <= 262144; n *= 4) {
        work.assign(1, 0); calls.assign(1, 0);
        balanced(n, 0);
        size_t db = work.size();
        work.assign(1, 0); calls.assign(1, 0);
        lopsided(n, 0);
        size_t dl = work.size();
        double lg = std::log2((double)n);
        std::printf("%10lld %8zu %8.1f %8zu %8.1f %8.2f\n",
                    n, db, lg, dl, lg / std::log2(1.5),
                    (double)dl / db);
    }
    return 0;
}
