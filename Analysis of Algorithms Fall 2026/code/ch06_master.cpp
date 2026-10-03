// ch06_master.cpp -- the master theorem's three cases, each one timed.
//
// T(n) = a T(n/b) + f(n) is settled by comparing f(n) against
// n^(log_b a). The three cases are three outcomes of that comparison, and
// each predicts a different ratio when n doubles:
//
//   Case 1  T(n) = 8T(n/2) + n^2   leaves dominate   Theta(n^3)   ratio 8
//   Case 2  T(n) = 2T(n/2) + n     tie, log appears  Theta(n lg n) ratio ~2.2
//   Case 3  T(n) = 2T(n/2) + n^2   root dominates    Theta(n^2)   ratio 4
//
// Three different predicted ratios from one theorem. The work is real
// arithmetic (summed into a checksum the compiler cannot discard), not a
// sleep, so the timings are honest.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch06_master.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <vector>

static volatile unsigned long long sink = 0;

// --- 1. Work that cannot be optimised away ---------------------------
// Each "unit" is a cheap arithmetic step folded into an accumulator. The
// accumulator escapes to a volatile at the end of the run, so the
// compiler must actually perform the loop.
static inline unsigned long long burn(long long units) {
    unsigned long long h = 1469598103934665603ULL;
    for (long long i = 0; i < units; ++i)
        h = (h ^ (unsigned long long)i) * 1099511628211ULL;
    return h;
}

// --- 2. Case 1: a = 8, b = 2, f(n) = n^2 -----------------------------
// n^(log_2 8) = n^3 beats n^2, so the LEAVES dominate and T = Theta(n^3).
static unsigned long long case1(long long n) {
    if (n <= 1) return burn(1);
    unsigned long long h = burn(n * n);
    for (int k = 0; k < 8; ++k) h ^= case1(n / 2);
    return h;
}

// --- 3. Case 2: a = 2, b = 2, f(n) = n -------------------------------
// n^(log_2 2) = n equals f(n): neither dominates, a log appears, and
// T = Theta(n lg n). This is merge sort's recurrence.
static unsigned long long case2(long long n) {
    if (n <= 1) return burn(1);
    unsigned long long h = burn(n);
    h ^= case2(n / 2);
    h ^= case2(n - n / 2);
    return h;
}

// --- 4. Case 3: a = 2, b = 2, f(n) = n^2 -----------------------------
// n^(log_2 2) = n loses to n^2, so the ROOT dominates and T = Theta(n^2).
// Note what this means: the recursion is almost free, and nearly all the
// time is spent in the top call.
static unsigned long long case3(long long n) {
    if (n <= 1) return burn(1);
    unsigned long long h = burn(n * n);
    h ^= case3(n / 2);
    h ^= case3(n - n / 2);
    return h;
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

static void run(const char* name, unsigned long long (*f)(long long),
                long long lo, long long hi, double predicted,
                const char* pred_text) {
    std::printf("\n%s\n   predicted ratio when n doubles: %s\n\n",
                name, pred_text);
    std::printf("%10s %12s %10s %10s\n", "n", "time (ms)", "ratio", "pred");
    double prev = 0;
    for (long long n = lo; n <= hi; n *= 2) {
        double t = time_ms([&] { return f(n); }, 5);
        if (prev == 0)
            std::printf("%10lld %12.2f %10s %10.2f\n", n, t, "-", predicted);
        else
            std::printf("%10lld %12.2f %10.2f %10.2f\n",
                        n, t, t / prev, predicted);
        prev = t;
    }
}

int main() {
    // Each case is given a range where its times are milliseconds rather
    // than microseconds -- case 1 grows as n^3, so it needs a small n.
    run("A. Case 1:  T(n) = 8T(n/2) + n^2  ->  Theta(n^3)",
        case1, 32, 512, 8.0, "8");
    run("B. Case 2:  T(n) = 2T(n/2) + n    ->  Theta(n lg n)",
        case2, 65536, 4194304, 2.1, "2 + 2/lg n, about 2.1");
    run("C. Case 3:  T(n) = 2T(n/2) + n^2  ->  Theta(n^2)",
        case3, 1024, 32768, 4.0, "4");

    // --- 5. The regularity condition, and a recurrence that defeats it
    // Case 3 needs a f(n/b) <= c f(n) for some c < 1. For f(n) = n^2,
    // a = 2, b = 2: 2(n/2)^2 = n^2/2, so c = 1/2 and the condition holds
    // comfortably. Print the check rather than asserting it.
    std::printf("\nD. the regularity condition for case 3\n\n");
    std::printf("   a*f(n/b) <= c*f(n) with a=2, b=2, f(n)=n^2:\n");
    for (long long n = 1024; n <= 16384; n *= 2) {
        double lhs = 2.0 * (n / 2.0) * (n / 2.0);
        double rhs = (double)n * n;
        std::printf("   n=%6lld   2*f(n/2) = %12.0f   f(n) = %12.0f"
                    "   ratio = %.3f\n", n, lhs, rhs, lhs / rhs);
    }
    std::printf("   ratio is 0.5 < 1 for every n, so case 3 applies.\n");
    std::printf("   (checksum %llu)\n", (unsigned long long)sink);
    return 0;
}
