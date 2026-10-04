// ch22_complexity.cpp -- checking against finding, measured.
//
// NP is the class of problems whose solutions can be VERIFIED quickly.
// P is the class that can be SOLVED quickly. Whether they are the same
// is the largest open question in the subject, and this program makes
// the gap between the two activities concrete:
//
//   A. subset-sum -- verifying a certificate is linear; finding one by
//      search is 2^n, and the table shows both on the same instances;
//   B. the ratio, which grows without bound;
//   C. meet-in-the-middle, 2^(n/2) -- exponential algorithms are not all
//      equally bad, and this one doubles the n you can reach;
//   D. a reduction, run: vertex cover and independent set are the same
//      problem in different clothes, and the program checks that the
//      transformation preserves answers.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch22_complexity.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <random>
#include <unordered_set>
#include <vector>

// --- 1. Verifying a certificate: linear, and that is the point -------
// Given a claimed subset (as a bitmask), check it sums to the target.
// No search, no cleverness: one pass. This is what "in NP" means.
static bool verify(const std::vector<long long>& a, unsigned long long cert,
                   long long target) {
    long long s = 0;
    for (size_t i = 0; i < a.size(); ++i)
        if (cert >> i & 1ULL) s += a[i];
    return s == target;
}

// --- 2. Finding a certificate by exhaustive search: 2^n --------------
// Every subset. Nothing is wasted and nothing is clever; this is the
// only method known that works for every instance.
static unsigned long long solve_brute(const std::vector<long long>& a,
                                      long long target, bool* found) {
    size_t n = a.size();
    for (unsigned long long mask = 0; mask < (1ULL << n); ++mask) {
        long long s = 0;
        for (size_t i = 0; i < n; ++i)
            if (mask >> i & 1ULL) s += a[i];
        if (s == target) { *found = true; return mask; }
    }
    *found = false;
    return 0;
}

// --- 3. Meet in the middle: 2^(n/2) ----------------------------------
// Split the set in half, enumerate all subset sums of each half, and
// look for a pair that adds to the target. Still exponential -- but the
// exponent is halved, which doubles the largest n you can handle. A
// reminder that "exponential" is not a single speed.
static bool solve_mitm(const std::vector<long long>& a, long long target) {
    size_t n = a.size(), h = n / 2;
    std::unordered_set<long long> left;
    for (unsigned long long m = 0; m < (1ULL << h); ++m) {
        long long s = 0;
        for (size_t i = 0; i < h; ++i) if (m >> i & 1ULL) s += a[i];
        left.insert(s);
    }
    size_t r = n - h;
    for (unsigned long long m = 0; m < (1ULL << r); ++m) {
        long long s = 0;
        for (size_t i = 0; i < r; ++i) if (m >> i & 1ULL) s += a[h + i];
        if (left.count(target - s)) return true;
    }
    return false;
}

template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    for (int r = 0; r < reps; ++r) {
        auto t0 = std::chrono::steady_clock::now();
        volatile auto v = f();
        (void)v;
        auto t1 = std::chrono::steady_clock::now();
        t.push_back(
            std::chrono::duration<double, std::milli>(t1 - t0).count());
    }
    std::sort(t.begin(), t.end());
    return t[t.size() / 2];
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937_64 rng(2718281828ULL);

    // --- 4. Checking against finding ---------------------------------
    // Each instance has a planted solution, so the answer is always yes
    // and the search cannot exit early by luck. The two columns are the
    // two activities that P vs NP is about.
    // Verification takes tens of NANOseconds, which is below the clock's
    // resolution -- timing it one call at a time reports 0.0000 and a
    // ratio of infinity, which is what the first version of this program
    // printed. Time a million verifications and divide.
    std::printf("A. subset-sum: verifying against finding\n\n");
    std::printf("%6s %18s %18s %16s\n",
                "n", "verify (ns)", "brute force (ms)", "ratio");
    for (size_t n = 16; n <= 28; n += 2) {
        std::vector<long long> a(n);
        std::uniform_int_distribution<long long> d(1, 1000000000LL);
        for (auto& x : a) x = d(rng);
        // plant a solution so the instance is satisfiable
        unsigned long long plant = rng() & ((1ULL << n) - 1);
        long long target = 0;
        for (size_t i = 0; i < n; ++i) if (plant >> i & 1ULL) target += a[i];

        const int REPS = 1000000;
        double tv_ns = time_ms([&] {
            int ok = 0;
            for (int r = 0; r < REPS; ++r) ok += verify(a, plant, target);
            return ok;
        }, 3) * 1e6 / REPS;              // ms for REPS -> ns per call
        bool found = false;
        double tb = time_ms([&] {
            bool f = false; return solve_brute(a, target, &f); }, 3);
        solve_brute(a, target, &found);
        std::printf("%6zu %18.1f %18.2f %16.3g%s\n",
                    n, tv_ns, tb, (tb * 1e6) / tv_ns,
                    found ? "" : "   *** no solution found");
    }

    // --- 5. Where the wall is ----------------------------------------
    // Extrapolate the brute-force column. The point is not that the
    // numbers are large; it is how FAST they become large.
    std::printf("\nB. extrapolating the brute force\n\n");
    std::printf("%8s %22s\n", "n", "estimated time");
    {
        // calibrate from n = 28 measured above
        std::vector<long long> a(28);
        std::uniform_int_distribution<long long> d(1, 1000000000LL);
        for (auto& x : a) x = d(rng);
        long long target = -1;                   // unsatisfiable: full search
        double ms28 = time_ms([&] {
            bool f = false; return solve_brute(a, target, &f); }, 3);
        double per = ms28 / 268435456.0;         // ms per subset
        for (int n : {30, 40, 50, 60, 80, 100}) {
            double ms = per * std::pow(2.0, n);
            const char* unit = "ms";
            double v = ms;
            if (v > 1000) { v /= 1000; unit = "s"; }
            if (v > 3600 && unit[0] == 's') { v /= 3600; unit = "hours"; }
            if (v > 24 && unit[0] == 'h') { v /= 24; unit = "days"; }
            if (v > 365 && unit[0] == 'd') { v /= 365; unit = "years"; }
            std::printf("%8d %18.3g %s\n", n, v, unit);
        }
    }

    // --- 6. Exponential is not one speed -----------------------------
    // Meet-in-the-middle is 2^(n/2). It is still exponential and it is
    // still useless at n = 200 -- but it reaches n = 40 comfortably,
    // where brute force does not.
    std::printf("\nC. brute force against meet-in-the-middle\n\n");
    std::printf("%6s %18s %22s %14s\n",
                "n", "brute force (ms)", "meet-in-middle (ms)", "speedup");
    for (size_t n = 20; n <= 32; n += 4) {
        std::vector<long long> a(n);
        std::uniform_int_distribution<long long> d(1, 1000000000LL);
        for (auto& x : a) x = d(rng);
        unsigned long long plant = rng() & ((1ULL << n) - 1);
        long long target = 0;
        for (size_t i = 0; i < n; ++i) if (plant >> i & 1ULL) target += a[i];
        double tb = (n <= 28) ? time_ms([&] {
            bool f = false; return solve_brute(a, target, &f); }, 3) : 0.0;
        double tm = time_ms([&] { return solve_mitm(a, target); }, 3);
        if (tb > 0)
            std::printf("%6zu %18.1f %22.2f %14.0f\n", n, tb, tm, tb / tm);
        else
            std::printf("%6zu %18s %22.2f %14s\n", n, "(too slow)", tm, "-");
    }

    // --- 7. A reduction, actually run --------------------------------
    // S is an independent set of G exactly when V \ S is a vertex cover.
    // So "is there an independent set of size k?" and "is there a vertex
    // cover of size |V| - k?" are the same question. A reduction is a
    // translation, and translations can be tested.
    std::printf("\nD. a reduction: independent set <-> vertex cover\n\n");
    {
        int trials = 3000, bad = 0;
        std::uniform_int_distribution<int> dn(4, 14);
        for (int t = 0; t < trials; ++t) {
            int V = dn(rng);
            std::vector<std::vector<char>> adj(V, std::vector<char>(V, 0));
            std::uniform_int_distribution<int> de(0, 2);
            for (int i = 0; i < V; ++i)
                for (int j = i + 1; j < V; ++j)
                    if (de(rng) == 0) adj[i][j] = adj[j][i] = 1;
            int k = 1 + (int)(rng() % (unsigned)V);

            bool has_is = false, has_vc = false;
            for (int mask = 0; mask < (1 << V); ++mask) {
                int pop = 0;
                for (int i = 0; i < V; ++i) if (mask >> i & 1) ++pop;
                // independent set of size >= k ?
                if (pop >= k) {
                    bool ok = true;
                    for (int i = 0; i < V && ok; ++i)
                        if (mask >> i & 1)
                            for (int j = i + 1; j < V; ++j)
                                if ((mask >> j & 1) && adj[i][j]) { ok = false; break; }
                    if (ok) has_is = true;
                }
                // vertex cover of size <= V - k ?
                if (pop <= V - k) {
                    bool ok = true;
                    for (int i = 0; i < V && ok; ++i)
                        for (int j = i + 1; j < V; ++j)
                            if (adj[i][j] && !(mask >> i & 1) && !(mask >> j & 1)) {
                                ok = false; break;
                            }
                    if (ok) has_vc = true;
                }
            }
            if (has_is != has_vc) ++bad;
        }
        std::printf("   %d random graphs: answers differ in %d\n",
                    trials, bad);
        std::printf("   S independent  <=>  V \\ S is a vertex cover.\n");
        std::printf("   So a fast algorithm for either gives one for both,\n");
        std::printf("   and that is what a reduction is FOR.\n");
    }
    return 0;
}
