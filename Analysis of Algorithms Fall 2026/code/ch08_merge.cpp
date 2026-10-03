// ch08_merge.cpp -- the divide-and-conquer pattern, and what tuning buys.
//
// Merge sort is Theta(n lg n) in every case, which is the point of it.
// This program measures three things the asymptotics cannot tell you:
//
//   A. that the bound really is case-independent -- sorted, reversed and
//      random input all cost the same, unlike insertion sort (Chapter 2);
//   B. what the insertion-sort CUTOFF is worth, and where it sits on this
//      machine -- the classic hybrid every production sort uses;
//   C. top-down against bottom-up, which have the same recurrence and
//      different constants.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch08_merge.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

static long long comparisons = 0;

// --- 1. The merge, which is the whole idea ---------------------------
// Two sorted runs become one in a single pass: look at both fronts, take
// the smaller. Linear in the total length, and it is the only place the
// algorithm compares anything.
static void merge(std::vector<int>& a, std::vector<int>& buf,
                  size_t lo, size_t mid, size_t hi, bool count) {
    size_t i = lo, j = mid, k = lo;
    while (i < mid && j < hi) {
        if (count) ++comparisons;
        buf[k++] = (a[i] <= a[j]) ? a[i++] : a[j++];
    }
    while (i < mid) buf[k++] = a[i++];
    while (j < hi)  buf[k++] = a[j++];
    std::copy(buf.begin() + lo, buf.begin() + hi, a.begin() + lo);
}

// --- 2. Top-down: the recurrence written out as code -----------------
// T(n) = 2T(n/2) + Theta(n), which Chapter 6 settles as Theta(n lg n).
// The cutoff is the one tuning knob: below it, insertion sort's smaller
// constant wins, exactly as Chapter 1 measured.
static void insertion(std::vector<int>& a, size_t lo, size_t hi) {
    for (size_t j = lo + 1; j < hi; ++j) {
        int key = a[j];
        size_t i = j;
        while (i > lo && a[i - 1] > key) { a[i] = a[i - 1]; --i; }
        a[i] = key;
    }
}

static void msort(std::vector<int>& a, std::vector<int>& buf,
                  size_t lo, size_t hi, size_t cutoff, bool count) {
    if (hi - lo < 2) return;
    if (hi - lo <= cutoff) { insertion(a, lo, hi); return; }
    size_t mid = lo + (hi - lo) / 2;
    msort(a, buf, lo, mid, cutoff, count);
    msort(a, buf, mid, hi, cutoff, count);
    merge(a, buf, lo, mid, hi, count);
}

// --- 3a. The same merge with the branch removed ----------------------
// Same comparisons and same output; the choice becomes arithmetic that
// compiles to a conditional move, so there is no branch to mispredict.
// Used only by part D, to test what the part A gap is made of.
static void merge_bl(std::vector<int>& a, std::vector<int>& buf,
                     size_t lo, size_t mid, size_t hi) {
    size_t i = lo, j = mid, k = lo;
    while (i < mid && j < hi) {
        bool take_left = a[i] <= a[j];
        buf[k++] = take_left ? a[i] : a[j];
        i += take_left ? 1 : 0;
        j += take_left ? 0 : 1;
    }
    while (i < mid) buf[k++] = a[i++];
    while (j < hi)  buf[k++] = a[j++];
    std::copy(buf.begin() + lo, buf.begin() + hi, a.begin() + lo);
}

static void msort_branchless(std::vector<int>& a, std::vector<int>& buf,
                             size_t lo, size_t hi) {
    if (hi - lo < 2) return;
    size_t mid = lo + (hi - lo) / 2;
    msort_branchless(a, buf, lo, mid);
    msort_branchless(a, buf, mid, hi);
    merge_bl(a, buf, lo, mid, hi);
}

// --- 3. Bottom-up: the same work, no recursion -----------------------
// Merge runs of 1, then 2, then 4 ... Same recurrence, same bound, no
// call stack. Whether it is faster is a question about this machine.
static void msort_bu(std::vector<int>& a, std::vector<int>& buf) {
    size_t n = a.size();
    for (size_t w = 1; w < n; w *= 2)
        for (size_t lo = 0; lo < n - w; lo += 2 * w)
            merge(a, buf, lo, lo + w, std::min(lo + 2 * w, n), false);
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

enum Shape { SORTED, REVERSED, RANDOM };
static std::vector<int> make(size_t n, Shape s) {
    std::vector<int> a(n);
    std::iota(a.begin(), a.end(), 0);
    if (s == REVERSED) std::reverse(a.begin(), a.end());
    if (s == RANDOM) { std::mt19937 r(4242); std::shuffle(a.begin(), a.end(), r); }
    return a;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    const size_t N = 2000000;

    // --- 4. The bound does not care what the input looks like --------
    // Chapter 2 found insertion sort varying by a factor of 6,600 across
    // these same three shapes. Merge sort should vary by almost nothing.
    std::printf("A. merge sort on three input shapes, n = %zu\n\n", N);
    std::printf("%12s %12s %16s\n", "shape", "time (ms)", "comparisons");
    const char* names[] = {"sorted", "reversed", "random"};
    double t0 = 0;
    for (int s = 0; s < 3; ++s) {
        auto base = make(N, (Shape)s);
        std::vector<int> work(N), buf(N);
        comparisons = 0;
        std::copy(base.begin(), base.end(), work.begin());
        msort(work, buf, 0, N, 0, true);
        long long cmp = comparisons;
        double t = time_ms([&] {
            std::copy(base.begin(), base.end(), work.begin());
            msort(work, buf, 0, N, 0, false);
        }, 5);
        if (s == 0) t0 = t;
        std::printf("%12s %12.1f %16lld", names[s], t, cmp);
        if (s) std::printf("   (x%.3f of sorted)", t / t0);
        std::printf("\n");
    }
    std::printf("\n   n lg n = %.0f\n", N * std::log2((double)N));

    // --- 5. What is the cutoff worth, and where is it? ---------------
    // Below the cutoff, insertion sort's smaller constant wins. The best
    // value is a property of this machine, not of the algorithm, which
    // is exactly why it has to be measured rather than looked up.
    std::printf("\nB. the insertion-sort cutoff, n = %zu\n\n", N);
    std::printf("%10s %12s %12s\n", "cutoff", "time (ms)", "vs none");
    auto base = make(N, RANDOM);
    std::vector<int> work(N), buf(N);
    double none = 0;
    for (size_t c : {size_t(0), size_t(4), size_t(8), size_t(16),
                     size_t(32), size_t(64), size_t(128), size_t(256)}) {
        double t = time_ms([&] {
            std::copy(base.begin(), base.end(), work.begin());
            msort(work, buf, 0, N, c, false);
        }, 7);
        if (c == 0) none = t;
        std::printf("%10zu %12.1f %12.3f\n", c, t, t / none);
    }

    // --- 6. Top-down against bottom-up, and against the library ------
    // Same recurrence, same Theta. Everything that differs here is
    // constant factors -- recursion overhead, memory order, and in
    // std::sort's case a different algorithm entirely (Chapter 9).
    std::printf("\nC. three implementations of the same bound\n\n");
    std::printf("%22s %12s\n", "implementation", "time (ms)");
    double td = time_ms([&] {
        std::copy(base.begin(), base.end(), work.begin());
        msort(work, buf, 0, N, 0, false);
    }, 7);
    std::printf("%22s %12.1f\n", "top-down, no cutoff", td);
    double td32 = time_ms([&] {
        std::copy(base.begin(), base.end(), work.begin());
        msort(work, buf, 0, N, 32, false);
    }, 7);
    std::printf("%22s %12.1f\n", "top-down, cutoff 32", td32);
    double bu = time_ms([&] {
        std::copy(base.begin(), base.end(), work.begin());
        msort_bu(work, buf);
    }, 7);
    std::printf("%22s %12.1f\n", "bottom-up", bu);
    double st = time_ms([&] {
        std::copy(base.begin(), base.end(), work.begin());
        std::sort(work.begin(), work.end());
    }, 7);
    std::printf("%22s %12.1f\n", "std::sort", st);

    // --- 7. Why is random input 4x slower than sorted input? ---------
    // Part A found merge sort 3.9x slower on random input than on sorted
    // -- but the comparison counts differ by only 1.9x. So roughly half
    // the slowdown is NOT accounted for by work done.
    //
    // Hypothesis: the branch "a[i] <= a[j]" in the merge loop. On sorted
    // input it goes the same way for a whole run and the processor
    // predicts it perfectly; on random input it is a coin flip and
    // mispredicts about half the time, at roughly 15 cycles a go.
    //
    // That hypothesis is testable. Replace the branch with arithmetic --
    // same comparisons, same answer, no branch to mispredict -- and see
    // whether the gap closes. If it does, the branch was the cause.
    std::printf("\nD. is the gap branch misprediction?\n\n");
    std::printf("%34s %12s\n", "merge variant (random input)", "time (ms)");
    double br = time_ms([&] {
        std::copy(base.begin(), base.end(), work.begin());
        msort(work, buf, 0, N, 0, false);
    }, 7);
    std::printf("%34s %12.1f\n", "branchy (as above)", br);
    double bl = time_ms([&] {
        std::copy(base.begin(), base.end(), work.begin());
        msort_branchless(work, buf, 0, N);
    }, 7);
    std::printf("%34s %12.1f\n", "branchless (cmov)", bl);
    std::printf("%34s %12.3f\n", "branchless / branchy", bl / br);

    auto sb = make(N, SORTED);
    double so = time_ms([&] {
        std::copy(sb.begin(), sb.end(), work.begin());
        msort(work, buf, 0, N, 0, false);
    }, 7);
    std::printf("%34s %12.1f\n", "branchy, SORTED input", so);
    std::printf("\n   random/sorted was %.2fx with a branch;"
                " %.2fx without.\n", br / so, bl / so);

    // Confirm the sorts actually sorted, so the race was real.
    std::copy(base.begin(), base.end(), work.begin());
    msort(work, buf, 0, N, 32, false);
    bool ok1 = std::is_sorted(work.begin(), work.end());
    std::copy(base.begin(), base.end(), work.begin());
    msort_branchless(work, buf, 0, N);
    bool ok2 = std::is_sorted(work.begin(), work.end());
    std::printf("\n   sorted correctly: top-down %s, branchless %s\n",
                ok1 ? "yes" : "NO", ok2 ? "yes" : "NO");
    return 0;
}
