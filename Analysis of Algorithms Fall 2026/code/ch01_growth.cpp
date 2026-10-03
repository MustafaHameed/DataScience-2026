// ch01_growth.cpp -- what "the question is asymptotic" actually looks like.
//
// Times two sorts over a doubling sequence of input sizes and prints the
// ratio of each time to the one before it. A quadratic algorithm should show
// a ratio near 4 when n doubles; an n log n one should show a ratio a little
// above 2. That is the whole claim of Chapter 1, measured.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch01_growth.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

// --- 1. The two algorithms, written plainly ---------------------------
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

// --- 2. Timing that is honest -----------------------------------------
// Each size is run several times and the MEDIAN reported: a mean is pulled
// about by a single scheduler hiccup, and the minimum flatters the cache.
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

static std::vector<int> random_input(size_t n, unsigned seed) {
    std::mt19937 rng(seed);
    std::vector<int> a(n);
    std::iota(a.begin(), a.end(), 0);
    std::shuffle(a.begin(), a.end(), rng);
    return a;
}

int main() {
    std::printf("%8s %12s %7s %12s %7s\n",
                "n", "insertion", "ratio", "merge", "ratio");
    double prev_i = 0, prev_m = 0;

    // --- 3. Double n each time, and watch the ratios ------------------
    for (size_t n = 1000; n <= 32000; n *= 2) {
        auto base = random_input(n, 12345);
        std::vector<int> buf(n);

        double ti = time_ms([&] {
            auto a = base;
            insertion_sort(a);
        }, 5);
        double tm = time_ms([&] {
            auto a = base;
            merge_sort(a, buf, 0, n);
        }, 5);

        if (prev_i == 0)
            std::printf("%8zu %12.2f %7s %12.2f %7s\n", n, ti, "-", tm, "-");
        else
            std::printf("%8zu %12.2f %7.2f %12.2f %7.2f\n",
                        n, ti, ti / prev_i, tm, tm / prev_m);
        prev_i = ti;
        prev_m = tm;
    }

    // --- 4. Where they cross, which is the other half of the point ----
    std::printf("\n%8s %12s %12s\n", "n", "insertion", "merge");
    for (size_t n = 10; n <= 200; n += 10) {
        auto base = random_input(n, 999);
        std::vector<int> buf(n);
        double ti = time_ms([&] { auto a = base; insertion_sort(a); }, 2001);
        double tm = time_ms([&] { auto a = base;
                                  merge_sort(a, buf, 0, n); }, 2001);
        std::printf("%8zu %12.5f %12.5f%s\n", n, ti, tm,
                    (ti < tm) ? "   <- insertion still wins" : "");
    }
    return 0;
}
