// ch17_amortised.cpp -- a sequence of operations, costed as a sequence.
//
// Appending to a dynamic array is usually O(1) and occasionally O(n),
// when the buffer is full and everything must be copied. Charging every
// append the worst case gives O(n) each and O(n^2) for n appends, which
// is wildly wrong: n appends really cost O(n) in total.
//
// Amortised analysis is the machinery for saying that properly, and this
// program measures the three things it claims:
//
//   A. the total number of element copies for n appends, which should be
//      under 2n for doubling growth and n^2/2 for grow-by-one;
//   B. the individual append costs, so the rare expensive ones are
//      visible -- amortised O(1) does not mean every operation is cheap;
//   C. growth factor against total copies and against wasted space,
//      which is the real engineering trade;
//   D. the same thing in seconds, including std::vector.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch17_amortised.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <vector>

static long long copies = 0;
static long long reallocs = 0;

// --- 1. A dynamic array with an adjustable growth policy -------------
// Deliberately hand-rolled so every copy can be counted. "factor" is the
// multiplier on resize; a factor of 0 means "grow by one", which is the
// policy the quadratic analysis describes.
struct Dyn {
    std::vector<int> buf;
    size_t len = 0, cap = 0;
    double factor;
    explicit Dyn(double f) : factor(f) {}

    void push(int x) {
        if (len == cap) {
            size_t ncap = (factor <= 1.0) ? cap + 1
                        : std::max<size_t>(1, (size_t)(cap * factor));
            if (ncap == cap) ncap = cap + 1;
            std::vector<int> nb(ncap);
            for (size_t i = 0; i < len; ++i) { nb[i] = buf[i]; ++copies; }
            buf.swap(nb);
            cap = ncap;
            ++reallocs;
        }
        buf[len++] = x;
    }
};

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

    // --- 2. Total copies: the aggregate method, measured -------------
    // Doubling: the copies are 1 + 2 + 4 + ... + n/2 < n, so the total
    // work for n appends is under 2n and the amortised cost per append
    // is under 2. Grow-by-one: 1 + 2 + ... + (n-1) = n(n-1)/2.
    std::printf("A. total element copies for n appends\n\n");
    std::printf("%10s %16s %10s %18s %12s\n",
                "n", "doubling", "/n", "grow-by-one", "/(n^2/2)");
    for (size_t n = 1000; n <= 64000; n *= 2) {
        copies = 0; reallocs = 0;
        { Dyn d(2.0); for (size_t i = 0; i < n; ++i) d.push((int)i); }
        long long cd = copies;
        copies = 0;
        { Dyn d(0.0); for (size_t i = 0; i < n; ++i) d.push((int)i); }
        long long c1 = copies;
        std::printf("%10zu %16lld %10.3f %18lld %12.4f\n",
                    n, cd, (double)cd / n, c1,
                    c1 / ((double)n * n / 2.0));
    }

    // --- 3. Amortised O(1) does not mean every operation is cheap ----
    // Print the cost of the single most expensive append, and how many
    // appends are expensive at all. This is the distinction between
    // amortised and worst-case that the chapter exists for.
    std::printf("\nB. the expensive appends, n = 1,000,000 (doubling)\n\n");
    {
        size_t n = 1000000;
        copies = 0; reallocs = 0;
        std::vector<long long> cost;
        Dyn d(2.0);
        for (size_t i = 0; i < n; ++i) {
            long long before = copies;
            d.push((int)i);
            cost.push_back(copies - before);
        }
        long long worst = *std::max_element(cost.begin(), cost.end());
        size_t expensive = 0;
        for (long long c : cost) if (c > 0) ++expensive;
        std::printf("   appends                : %zu\n", n);
        std::printf("   reallocations          : %lld\n", reallocs);
        std::printf("   appends that copied    : %zu  (%.4f%%)\n",
                    expensive, 100.0 * expensive / n);
        std::printf("   total copies           : %lld  (%.3f per append)\n",
                    copies, (double)copies / n);
        std::printf("   MOST EXPENSIVE append  : %lld copies\n", worst);
        std::printf("   worst case x n         : %lld  (the bound that\n",
                    worst * (long long)n);
        std::printf("                            amortised analysis"
                    " replaces)\n");
    }

    // --- 4. The growth factor is a real engineering choice -----------
    // Larger factor: fewer copies, more wasted space. Smaller factor:
    // the reverse. 1.5 is what several standard libraries use; 2.0 is
    // what most textbooks describe.
    std::printf("\nC. growth factor against copies and waste"
                " (n = 1,000,000)\n\n");
    std::printf("%10s %16s %12s %16s %12s\n",
                "factor", "total copies", "per append", "final capacity",
                "waste");
    for (double f : {1.1, 1.25, 1.5, 2.0, 3.0, 4.0}) {
        size_t n = 1000000;
        copies = 0; reallocs = 0;
        Dyn d(f);
        for (size_t i = 0; i < n; ++i) d.push((int)i);
        std::printf("%10.2f %16lld %12.3f %16zu %11.1f%%\n",
                    f, copies, (double)copies / n, d.cap,
                    100.0 * (d.cap - n) / n);
    }

    // --- 5. And in seconds, against the real thing -------------------
    // --- 5. And in seconds, against the real thing -------------------
    // The two policies need different ranges. Doubling at n = 160,000
    // takes under a millisecond, which is too small to measure -- the
    // first version of this table reported ratios of 0.92, 1.19, 1.81
    // and 4.04 for a linear algorithm, all of it noise. Grow-by-one at
    // n = 1,600,000 would take about ten minutes. So: a large range for
    // the linear policies, a small one for the quadratic policy.
    std::printf("\nD. time for n appends, doubling (ms)\n\n");
    std::printf("%12s %14s %7s %14s %7s\n",
                "n", "doubling", "ratio", "std::vector", "ratio");
    double pd = 0, pv = 0;
    for (size_t n = 1000000; n <= 32000000; n *= 2) {
        double td = time_ms([&] {
            Dyn d(2.0); for (size_t i = 0; i < n; ++i) d.push((int)i);
        }, 5);
        double tv = time_ms([&] {
            std::vector<int> v;
            for (size_t i = 0; i < n; ++i) v.push_back((int)i);
        }, 5);
        std::printf("%12zu %14.1f", n, td);
        if (pd) std::printf(" %7.2f", td / pd); else std::printf(" %7s", "-");
        std::printf(" %14.1f", tv);
        if (pv) std::printf(" %7.2f", tv / pv); else std::printf(" %7s", "-");
        std::printf("\n");
        pd = td; pv = tv;
    }

    std::printf("\nE. time for n appends, grow-by-one (ms)\n\n");
    std::printf("%12s %14s %7s\n", "n", "grow-by-one", "ratio");
    double p1 = 0;
    // Stops at 160,000: the next doubling takes 39 seconds, and its
    // ratio comes out at 6.8 rather than 4 because allocating a third
    // of a million successively larger buffers starts to cost more than
    // the copying does. Quadratic is established by then.
    for (size_t n = 20000; n <= 160000; n *= 2) {
        double t1 = time_ms([&] {
            Dyn d(0.0); for (size_t i = 0; i < n; ++i) d.push((int)i);
        }, 3);
        std::printf("%12zu %14.1f", n, t1);
        if (p1) std::printf(" %7.2f", t1 / p1); else std::printf(" %7s", "-");
        std::printf("\n");
        p1 = t1;
    }
    std::printf("\n   doubling: ratio ~2 (linear)."
                "  grow-by-one: ratio ~4 (quadratic).\n");
    return 0;
}
