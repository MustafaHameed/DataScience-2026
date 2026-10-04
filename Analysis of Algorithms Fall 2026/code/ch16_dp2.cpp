// ch16_dp2.cpp -- the classic dynamic programs, and one nasty surprise.
//
//   A. longest common subsequence -- Theta(mn) time and space, and the
//      Theta(min(m,n)) space version when the sequence itself is not
//      wanted;
//   B. edit distance, the same table with three operations;
//   C. 0/1 knapsack -- Theta(nW), which LOOKS polynomial and is not.
//      W is a VALUE, and Chapter 2 established that an input's size is
//      its bit length. Doubling W doubles the work while adding one bit
//      of input: that is exponential in the input size, and the name
//      for it is PSEUDO-POLYNOMIAL. Chapter 22 explains why it matters.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch16_dp2.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <random>
#include <string>
#include <vector>

// --- 1. LCS: the full table, so the subsequence can be recovered -----
static int lcs_full(const std::string& a, const std::string& b,
                    std::string* out) {
    size_t m = a.size(), n = b.size();
    std::vector<std::vector<int>> c(m + 1, std::vector<int>(n + 1, 0));
    for (size_t i = 1; i <= m; ++i)
        for (size_t j = 1; j <= n; ++j)
            c[i][j] = (a[i - 1] == b[j - 1]) ? c[i - 1][j - 1] + 1
                                             : std::max(c[i - 1][j], c[i][j - 1]);
    if (out) {                       // walk back through the table
        std::string s;
        size_t i = m, j = n;
        while (i && j) {
            if (a[i - 1] == b[j - 1]) { s += a[i - 1]; --i; --j; }
            else if (c[i - 1][j] >= c[i][j - 1]) --i;
            else --j;
        }
        std::reverse(s.begin(), s.end());
        *out = s;
    }
    return c[m][n];
}

// --- 2. LCS in two rows ----------------------------------------------
// Row i depends only on row i-1, so only two rows need to exist. The
// LENGTH survives; the subsequence does not, because reconstructing it
// needs the whole table. That is the trade, and it is typical of DP.
static int lcs_two_rows(const std::string& a, const std::string& b) {
    size_t m = a.size(), n = b.size();
    std::vector<int> prev(n + 1, 0), cur(n + 1, 0);
    for (size_t i = 1; i <= m; ++i) {
        for (size_t j = 1; j <= n; ++j)
            cur[j] = (a[i - 1] == b[j - 1]) ? prev[j - 1] + 1
                                            : std::max(prev[j], cur[j - 1]);
        prev.swap(cur);
    }
    return prev[n];
}

// --- 3. Edit distance: the same shape, three choices per cell --------
// Insert, delete or substitute. The recurrence differs from LCS only in
// what the three predecessors cost.
static int edit_distance(const std::string& a, const std::string& b) {
    size_t m = a.size(), n = b.size();
    std::vector<int> prev(n + 1), cur(n + 1);
    for (size_t j = 0; j <= n; ++j) prev[j] = (int)j;
    for (size_t i = 1; i <= m; ++i) {
        cur[0] = (int)i;
        for (size_t j = 1; j <= n; ++j)
            cur[j] = std::min({prev[j] + 1,              // delete
                               cur[j - 1] + 1,           // insert
                               prev[j - 1] + (a[i - 1] != b[j - 1])});
        prev.swap(cur);
    }
    return prev[n];
}

// --- 4. 0/1 knapsack: Theta(nW), and W is a number not a count -------
static long long knapsack(const std::vector<int>& w,
                          const std::vector<long long>& v, int W) {
    std::vector<long long> best(W + 1, 0);
    for (size_t i = 0; i < w.size(); ++i)
        for (int cap = W; cap >= w[i]; --cap)
            best[cap] = std::max(best[cap], best[cap - w[i]] + v[i]);
    return best[W];
}

// --- 5. The same problem by brute force, as an oracle ----------------
// 2^n subsets. Only usable for n <= 22, which is exactly why DP exists.
static long long knapsack_brute(const std::vector<int>& w,
                                const std::vector<long long>& v, int W) {
    int n = (int)w.size();
    long long best = 0;
    for (int mask = 0; mask < (1 << n); ++mask) {
        long long tw = 0, tv = 0;
        for (int i = 0; i < n; ++i)
            if (mask >> i & 1) { tw += w[i]; tv += v[i]; }
        if (tw <= W) best = std::max(best, tv);
    }
    return best;
}

template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    for (int r = 0; r < reps; ++r) {
        auto t0 = std::chrono::steady_clock::now();
        volatile long long v = f();
        (void)v;
        auto t1 = std::chrono::steady_clock::now();
        t.push_back(
            std::chrono::duration<double, std::milli>(t1 - t0).count());
    }
    std::sort(t.begin(), t.end());
    return t[t.size() / 2];
}

static std::string random_string(size_t n, int alphabet, std::mt19937& r) {
    std::uniform_int_distribution<int> d(0, alphabet - 1);
    std::string s(n, 'a');
    for (auto& c : s) c = (char)('a' + d(r));
    return s;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(1492);

    // --- 6. LCS and edit distance are Theta(mn) ----------------------
    // Doubling both lengths should quadruple the time.
    std::printf("A. LCS and edit distance: Theta(mn)\n\n");
    std::printf("%8s %12s %7s %12s %7s %12s\n",
                "n", "LCS full", "ratio", "LCS 2-row", "ratio", "edit dist");
    double pf = 0, pr = 0;
    for (size_t n = 500; n <= 8000; n *= 2) {
        std::string a = random_string(n, 4, rng), b = random_string(n, 4, rng);
        double tf = time_ms([&] { return lcs_full(a, b, nullptr); }, 3);
        double tr = time_ms([&] { return lcs_two_rows(a, b); }, 3);
        double te = time_ms([&] { return edit_distance(a, b); }, 3);
        std::printf("%8zu %12.1f", n, tf);
        if (pf) std::printf(" %7.2f", tf / pf); else std::printf(" %7s", "-");
        std::printf(" %12.1f", tr);
        if (pr) std::printf(" %7.2f", tr / pr); else std::printf(" %7s", "-");
        std::printf(" %12.1f\n", te);
        pf = tf; pr = tr;
    }

    // --- 7. A worked instance, small enough to check by eye ----------
    std::printf("\nB. a small instance, to check the answers\n\n");
    {
        std::string a = "ALGORITHM", b = "ALTRUISTIC", out;
        int L = lcs_full(a, b, &out);
        std::printf("   LCS(\"%s\", \"%s\") = %d, \"%s\"\n",
                    a.c_str(), b.c_str(), L, out.c_str());
        std::printf("   edit(\"%s\", \"%s\") = %d\n",
                    a.c_str(), b.c_str(), edit_distance(a, b));
        std::printf("   two-row LCS agrees: %s\n",
                    (lcs_two_rows(a, b) == L) ? "yes" : "NO");
    }

    // --- 8. Knapsack: DP against brute force -------------------------
    // Chapter 7's oracle again. If they disagree, the DP is wrong.
    std::printf("\nC. knapsack: DP against brute force (n = 18)\n\n");
    {
        int n = 18, bad = 0;
        std::uniform_int_distribution<int> dw(1, 40), dv(1, 100);
        for (int t = 0; t < 300; ++t) {
            std::vector<int> w(n);
            std::vector<long long> v(n);
            for (int i = 0; i < n; ++i) { w[i] = dw(rng); v[i] = dv(rng); }
            int W = 120;
            if (knapsack(w, v, W) != knapsack_brute(w, v, W)) ++bad;
        }
        std::printf("   300 random instances, disagreements: %d\n", bad);
    }

    // --- 9. The surprise: doubling W doubles the time ----------------
    // n is FIXED at 200 here. Only the capacity changes, and the
    // capacity is one number. Writing 2W instead of W adds ONE BIT to
    // the input and doubles the running time.
    std::printf("\nD. knapsack with n fixed at 200, W doubling\n\n");
    std::printf("%12s %8s %14s %8s %16s\n",
                "W", "bits", "time (ms)", "ratio", "table entries");
    {
        int n = 200;
        std::vector<int> w(n);
        std::vector<long long> v(n);
        std::uniform_int_distribution<int> dv(1, 1000);
        for (int i = 0; i < n; ++i) { w[i] = 1 + (i % 50); v[i] = dv(rng); }
        double prev = 0;
        for (int W = 65536; W <= 16777216; W *= 2) {
            double t = time_ms([&] { return knapsack(w, v, W); }, 3);
            int bits = 0;
            for (int x = W; x; x >>= 1) ++bits;
            std::printf("%12d %8d %14.1f", W, bits, t);
            if (prev) std::printf(" %8.2f", t / prev);
            else std::printf(" %8s", "-");
            std::printf(" %16lld\n", (long long)n * W);
            prev = t;
        }
        std::printf("\n   Each row adds ONE BIT to the input and doubles"
                    " the work.\n   Theta(nW) is polynomial in the VALUE of W"
                    " and exponential\n   in its LENGTH: pseudo-polynomial."
                    " See Chapter 22.\n");
    }
    return 0;
}
