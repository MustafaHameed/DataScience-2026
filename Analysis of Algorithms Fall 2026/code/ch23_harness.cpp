// ch23_harness.cpp -- the measurement harness for your semester project.
//
// Every rule this book has arrived at the hard way, in one reusable file.
// Copy it, replace the two algorithms in part 6, and it will produce the
// table your report needs:
//
//   * median of several runs, never a mean and never one run  (Ch 1)
//   * input generated once, OUTSIDE the timed region          (Ch 3)
//   * the setup cost measured separately and subtracted       (Ch 3)
//   * results checked against a slow, obviously-correct
//     reference before any timing is reported                 (Ch 7)
//   * ratios on doubling, against the predicted order         (Ch 1, 3)
//   * MORE THAN ONE input shape, because the shape is usually
//     what decides                     (Ch 2, 9, 12, 13, 20)
//   * the resolution limit stated, so no result is claimed
//     that the clock cannot support                           (Ch 3)
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch23_harness.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <functional>
#include <numeric>
#include <random>
#include <string>
#include <vector>

// --- 1. Timing: median, and honest about its own limits --------------
// A mean is dragged by one scheduler interruption; a minimum flatters
// the cache; one run measures the weather. The median of an odd number
// of runs is the cheapest estimator that is not actively misleading.
template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    t.reserve(reps);
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

// What is the smallest interval this clock can resolve? Any measurement
// within an order of magnitude of this number is not evidence.
static double clock_resolution_ms() {
    using C = std::chrono::steady_clock;
    double best = 1e9;
    for (int i = 0; i < 1000; ++i) {
        auto a = C::now(), b = C::now();
        while (b == a) b = C::now();
        best = std::min(best,
            std::chrono::duration<double, std::milli>(b - a).count());
    }
    return best;
}

// --- 2. The input shapes that catch things ---------------------------
// Chapter 2 measured a 6,600x spread across these on one algorithm;
// Chapter 9 found a crash on one of them; Chapter 20 found a 789x
// slowdown on a shape nobody would have generated at random. Testing
// only on "random" is how all of those get missed.
enum Shape { RANDOM, SORTED, REVERSED, NEARLY_SORTED, FEW_DISTINCT };
static const char* shape_name(Shape s) {
    switch (s) {
        case RANDOM:        return "random";
        case SORTED:        return "sorted";
        case REVERSED:      return "reversed";
        case NEARLY_SORTED: return "99% sorted";
        default:            return "few distinct";
    }
}
static std::vector<int> make_input(size_t n, Shape s, unsigned seed) {
    std::vector<int> a(n);
    std::mt19937 rng(seed);
    switch (s) {
        case SORTED:
            std::iota(a.begin(), a.end(), 0);
            break;
        case REVERSED:
            std::iota(a.rbegin(), a.rend(), 0);
            break;
        case NEARLY_SORTED: {
            std::iota(a.begin(), a.end(), 0);
            std::uniform_int_distribution<size_t> d(0, n - 1);
            for (size_t i = 0; i < n / 100; ++i) std::swap(a[d(rng)], a[d(rng)]);
            break;
        }
        case FEW_DISTINCT: {
            std::uniform_int_distribution<int> d(0, 9);
            for (auto& x : a) x = d(rng);
            break;
        }
        default: {
            std::iota(a.begin(), a.end(), 0);
            std::shuffle(a.begin(), a.end(), rng);
        }
    }
    return a;
}

// --- 3. The thing under test, and the thing it is checked against ----
// REPLACE THESE TWO. The reference must be the version you can read in
// one sitting and be sure of; the candidate is the one you are claiming
// something about. Chapter 7 is why both exist.
static void reference(std::vector<int>& a) {          // obviously correct
    std::sort(a.begin(), a.end());
}
static void candidate(std::vector<int>& a) {          // under test
    // a deliberately simple bottom-up merge sort, as a placeholder
    size_t n = a.size();
    std::vector<int> buf(n);
    for (size_t w = 1; w < n; w *= 2)
        for (size_t lo = 0; lo < n - w; lo += 2 * w) {
            size_t mid = lo + w, hi = std::min(lo + 2 * w, n);
            size_t i = lo, j = mid, k = lo;
            while (i < mid && j < hi) buf[k++] = (a[i] <= a[j]) ? a[i++] : a[j++];
            while (i < mid) buf[k++] = a[i++];
            while (j < hi)  buf[k++] = a[j++];
            std::copy(buf.begin() + lo, buf.begin() + hi, a.begin() + lo);
        }
}

// --- 4. Correctness first, timing second -----------------------------
// A faster algorithm that gets a different answer is not faster. This
// runs before anything is timed, and a failure here stops the report.
static bool check_correctness() {
    std::mt19937 rng(20261003);
    for (int t = 0; t < 2000; ++t) {
        size_t n = 1 + rng() % 200;
        Shape s = (Shape)(rng() % 5);
        std::vector<int> a = make_input(n, s, (unsigned)rng());
        std::vector<int> b = a;
        reference(a);
        candidate(b);
        if (a != b) {
            std::printf("   MISMATCH at n=%zu shape=%s\n", n, shape_name(s));
            return false;
        }
    }
    return true;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);

    double res = clock_resolution_ms();
    std::printf("Measurement harness\n\n");
    std::printf("   clock resolution      : %.6f ms\n", res);
    std::printf("   smallest trustworthy  : %.4f ms"
                "  (100x the resolution)\n\n", res * 100);

    // --- 5. Gate: do not report timings for a wrong algorithm --------
    std::printf("A. correctness against the reference\n\n");
    if (!check_correctness()) {
        std::printf("   FAILED -- no timings reported.\n");
        return 1;
    }
    std::printf("   2000 instances across 5 shapes: all agree\n");

    // --- 6. The table: every shape, every size, with ratios ----------
    // The ratio column is the one that tests an order of growth, because
    // it cancels the machine-specific constant (Chapter 3).
    std::printf("\nB. time (ms), median of 7, setup excluded\n\n");
    std::printf("%14s", "n");
    for (int s = 0; s < 5; ++s) std::printf(" %13s", shape_name((Shape)s));
    std::printf("\n");

    std::vector<double> prev(5, 0.0);
    for (size_t n = 125000; n <= 2000000; n *= 2) {
        std::printf("%14zu", n);
        for (int s = 0; s < 5; ++s) {
            std::vector<int> base = make_input(n, (Shape)s, 4242);
            std::vector<int> work(n);
            // Setup (the copy) is measured separately and subtracted,
            // so what is reported is the algorithm and not the memcpy.
            double tc = time_ms([&] {
                std::copy(base.begin(), base.end(), work.begin()); }, 7);
            double tt = time_ms([&] {
                std::copy(base.begin(), base.end(), work.begin());
                candidate(work);
            }, 7);
            double t = tt - tc;
            std::printf(" %13.2f", t);
            prev[s] = t;
        }
        std::printf("\n");
    }

    // --- 7. The ratios, which is what the claim rests on -------------
    std::printf("\nC. ratio on doubling n  (n lg n predicts 2 + 2/lg n)\n\n");
    std::printf("%14s", "n");
    for (int s = 0; s < 5; ++s) std::printf(" %13s", shape_name((Shape)s));
    std::printf(" %10s\n", "predicted");
    std::vector<double> last(5, 0.0);
    for (size_t n = 125000; n <= 2000000; n *= 2) {
        std::printf("%14zu", n);
        for (int s = 0; s < 5; ++s) {
            std::vector<int> base = make_input(n, (Shape)s, 4242);
            std::vector<int> work(n);
            double tc = time_ms([&] {
                std::copy(base.begin(), base.end(), work.begin()); }, 7);
            double tt = time_ms([&] {
                std::copy(base.begin(), base.end(), work.begin());
                candidate(work);
            }, 7);
            double t = tt - tc;
            if (last[s] > 0) std::printf(" %13.2f", t / last[s]);
            else             std::printf(" %13s", "-");
            last[s] = t;
        }
        std::printf(" %10.2f\n", 2.0 + 2.0 / std::log2((double)n));
    }

    std::printf("\n   Report the ratio column, not the absolute times:\n"
                "   the absolute times are a fact about this laptop and\n"
                "   the ratios are a fact about the algorithm.\n");
    return 0;
}
