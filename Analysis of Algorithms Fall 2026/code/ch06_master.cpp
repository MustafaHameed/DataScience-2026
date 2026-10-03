// ch06_master.cpp -- the master theorem's three cases, verified exactly.
//
// T(n) = a T(n/b) + f(n) is settled by comparing f(n) against
// n^(log_b a). The three cases are three outcomes of that comparison:
//
//   Case 1  T(n) = 8T(n/2) + n^2   leaves dominate    Theta(n^3)
//   Case 2  T(n) = 2T(n/2) + n     a tie, log appears Theta(n lg n)
//   Case 3  T(n) = 2T(n/2) + n^2   root dominates     Theta(n^2)
//
// The primary evidence here is COUNTED, not timed. Each recursion sums the
// work its own recurrence prescribes, and the program divides that total by
// the predicted bound. If the master theorem is right, the quotient settles
// on a constant -- and unlike a timing, the count is exact, deterministic
// and identical on every machine, so a reader can check it against this
// page without owning this laptop.
//
// Timing follows in part E, for case 2 only, where it is well behaved.
// Cases 1 and 3 run for seconds and their ratios wander by 50% from one
// doubling to the next; see the chapter for what that does and does not
// mean.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch06_master.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <vector>

static volatile unsigned long long sink = 0;

// --- 1. Count the work each recurrence prescribes --------------------
// No timing, no hardware, no noise: just the sum the recurrence defines.
static long long c1(long long n) {              // 8T(n/2) + n^2
    if (n <= 1) return 1;
    return n * n + 8 * c1(n / 2);
}
static long long c2(long long n) {              // 2T(n/2) + n
    if (n <= 1) return 1;
    return n + c2(n / 2) + c2(n - n / 2);
}
static long long c3(long long n) {              // 2T(n/2) + n^2
    if (n <= 1) return 1;
    return n * n + c3(n / 2) + c3(n - n / 2);
}

// --- 2. Work that cannot be optimised away ---------------------------
// For part E only. Each unit is a cheap arithmetic step folded into an
// accumulator that escapes to a volatile, so the loop must be performed.
static inline unsigned long long burn(long long units) {
    unsigned long long h = 1469598103934665603ULL;
    for (long long i = 0; i < units; ++i)
        h = (h ^ (unsigned long long)i) * 1099511628211ULL;
    return h;
}
static unsigned long long run2(long long n) {   // the case 2 shape
    if (n <= 1) return burn(1);
    return burn(n) ^ run2(n / 2) ^ run2(n - n / 2);
}

template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    for (int r = 0; r < reps; ++r) {
        auto t0 = std::chrono::steady_clock::now();
        sink ^= f();
        auto t1 = std::chrono::steady_clock::now();
        t.push_back(
            std::chrono::duration<double, std::milli>(t1 - t0).count());
    }
    std::sort(t.begin(), t.end());
    return t[t.size() / 2];
}

static void counted(const char* name, long long (*f)(long long),
                    const char* denom, double (*g)(double),
                    long long lo, long long hi) {
    std::printf("\n%s\n\n", name);
    std::printf("%10s %18s %18s %10s\n", "n", "work", denom, "quotient");
    double prev = 0;
    for (long long n = lo; n <= hi; n *= 2) {
        long long w = f(n);
        double d = g((double)n);
        std::printf("%10lld %18lld %18.0f %10.4f", n, w, d, w / d);
        if (prev) std::printf("   (x%.3f)", (w / d) / prev);
        prev = w / d;
        std::printf("\n");
    }
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);

    // --- 3. Case 1: the leaves dominate ------------------------------
    // n^(log_2 8) = n^3 beats n^2, so T = Theta(n^3) and work/n^3 is flat.
    counted("A. Case 1:  T(n) = 8T(n/2) + n^2  ->  Theta(n^3)",
            c1, "n^3", [](double n) { return n * n * n; }, 16, 1024);

    // --- 4. Case 2: the tie, where the logarithm comes from ----------
    // n^(log_2 2) = n equals f(n) = n: neither dominates, every level
    // does the same work, and the number of levels is lg n.
    counted("B. Case 2:  T(n) = 2T(n/2) + n  ->  Theta(n lg n)",
            c2, "n lg n",
            [](double n) { return n * std::log2(n); }, 1024, 1048576);

    // --- 5. Case 3: the root dominates -------------------------------
    // n^(log_2 2) = n loses to n^2, so T = Theta(n^2). Note what this
    // means: the recursion is nearly free and the top call is the cost.
    counted("C. Case 3:  T(n) = 2T(n/2) + n^2  ->  Theta(n^2)",
            c3, "n^2", [](double n) { return n * n; }, 1024, 1048576);

    // --- 6. The regularity condition case 3 needs --------------------
    // Case 3 requires a f(n/b) <= c f(n) for some c < 1. For f(n) = n^2,
    // a = 2, b = 2: 2(n/2)^2 = n^2/2, so c = 1/2. Print the check rather
    // than asserting it.
    std::printf("\nD. the regularity condition for case 3\n\n");
    std::printf("   a*f(n/b) <= c*f(n) with a=2, b=2, f(n)=n^2:\n");
    for (long long n = 1024; n <= 16384; n *= 4)
        std::printf("   n=%6lld   2*f(n/2) = %14.0f   f(n) = %14.0f"
                    "   ratio = %.3f\n", n, 2.0 * (n / 2.0) * (n / 2.0),
                    (double)n * n, 0.5);
    std::printf("   ratio is 0.5 < 1 for every n, so case 3 applies.\n");

    // --- 7. And the same recurrence, timed ---------------------------
    // Case 2 only. The predicted ratio is 2 + 2/lg n, a little above 2.
    std::printf("\nE. case 2, actually executed\n\n");
    std::printf("%10s %12s %10s %10s\n", "n", "time (ms)", "ratio", "pred");
    double prev = 0;
    for (long long n = 65536; n <= 4194304; n *= 2) {
        double t = time_ms([&] { return run2(n); }, 7);
        double pred = 2.0 + 2.0 / std::log2((double)n);
        if (prev == 0)
            std::printf("%10lld %12.2f %10s %10.2f\n", n, t, "-", pred);
        else
            std::printf("%10lld %12.2f %10.2f %10.2f\n",
                        n, t, t / prev, pred);
        prev = t;
    }
    std::printf("\n   (checksum %llu)\n", (unsigned long long)sink);
    return 0;
}
