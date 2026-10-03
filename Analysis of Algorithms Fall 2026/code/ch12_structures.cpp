// ch12_structures.cpp -- heaps and search trees, analysed not rebuilt.
//
// Data Structures built these. This program asks the question that course
// did not: WHY are the heights what they are, and what does that cost?
//
//   A. a heap's height is exactly floor(lg n) -- it is complete by
//      construction, so there is nothing to go wrong;
//   B. a BST's height is NOT. Built from random keys it is about
//      3 lg n; built from sorted keys it is exactly n, and the "tree"
//      is a linked list;
//   C. what that does to search time, measured;
//   D. the constant in the random-BST height, against the theoretical
//      4.311 ln n = 2.988 lg n.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch12_structures.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

// --- 1. A BST as three parallel arrays -------------------------------
// No pointers and no allocation: left[i], right[i] and key[i]. The
// shape is identical to a pointer tree, and the measurement is about
// shape, not about allocators.
struct BST {
    std::vector<int> key, left, right;
    int root = -1;
    explicit BST(size_t cap) {
        key.reserve(cap); left.reserve(cap); right.reserve(cap);
    }
    int make(int k) {
        key.push_back(k); left.push_back(-1); right.push_back(-1);
        return (int)key.size() - 1;
    }
    // Insert without balancing: the ordinary unbalanced BST insert,
    // which is the point -- a balanced tree would hide the effect.
    void insert(int k) {
        if (root < 0) { root = make(k); return; }
        int cur = root;
        for (;;) {
            if (k < key[cur]) {
                if (left[cur] < 0) { left[cur] = make(k); return; }
                cur = left[cur];
            } else {
                if (right[cur] < 0) { right[cur] = make(k); return; }
                cur = right[cur];
            }
        }
    }
    // Iterative, so a degenerate tree does not overflow the stack the
    // way Chapter 9's quicksort did.
    int height() const {
        if (root < 0) return 0;
        std::vector<std::pair<int, int>> st{{root, 1}};
        int best = 0;
        while (!st.empty()) {
            auto [v, d] = st.back(); st.pop_back();
            best = std::max(best, d);
            if (left[v] >= 0) st.push_back({left[v], d + 1});
            if (right[v] >= 0) st.push_back({right[v], d + 1});
        }
        return best;
    }
    // Total depth over all nodes: the average successful-search cost.
    double mean_depth() const {
        if (root < 0) return 0;
        std::vector<std::pair<int, int>> st{{root, 1}};
        long long total = 0, cnt = 0;
        while (!st.empty()) {
            auto [v, d] = st.back(); st.pop_back();
            total += d; ++cnt;
            if (left[v] >= 0) st.push_back({left[v], d + 1});
            if (right[v] >= 0) st.push_back({right[v], d + 1});
        }
        return (double)total / cnt;
    }
    bool contains(int k) const {
        int cur = root;
        while (cur >= 0) {
            if (k == key[cur]) return true;
            cur = (k < key[cur]) ? left[cur] : right[cur];
        }
        return false;
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

    // --- 2. A heap's height is not a hope, it is a guarantee ---------
    // A binary heap is a COMPLETE tree stored in an array: node i has
    // children 2i+1 and 2i+2. Completeness is not maintained by
    // clever rebalancing -- it is maintained by the array layout, and
    // cannot fail. Hence height = floor(lg n), exactly, always.
    std::printf("A. heap height is exactly floor(lg n)\n\n");
    std::printf("%12s %14s %14s %10s\n",
                "n", "height", "floor(lg n)", "equal?");
    for (size_t n = 1000; n <= 100000000; n *= 10) {
        int h = 0;
        for (size_t i = n; i > 1; i /= 2) ++h;      // walk to the root
        int fl = (int)std::floor(std::log2((double)n));
        std::printf("%12zu %14d %14d %10s\n",
                    n, h, fl, (h == fl) ? "yes" : "NO");
    }

    // --- 3. A BST's height depends entirely on the insertion order --
    // Same keys, same code, two orders. This is Chapter 2's lesson --
    // the nature of the input, not its size -- applied to a structure
    // rather than to a sort.
    std::printf("\nB. BST height: random keys against sorted keys\n\n");
    std::printf("%10s %10s %10s %12s %12s %12s\n",
                "n", "lg n", "random h", "h/lg n", "sorted h", "sorted/n");
    std::mt19937 rng(8675309);
    for (size_t n = 1000; n <= 256000; n *= 2) {
        std::vector<int> keys(n);
        std::iota(keys.begin(), keys.end(), 0);

        std::vector<int> shuf = keys;
        std::shuffle(shuf.begin(), shuf.end(), rng);
        BST a(n);
        for (int k : shuf) a.insert(k);

        BST b(n);
        for (int k : keys) b.insert(k);          // in order: worst case

        double lg = std::log2((double)n);
        std::printf("%10zu %10.1f %10d %12.2f %12d %12.3f\n",
                    n, lg, a.height(), a.height() / lg,
                    b.height(), (double)b.height() / n);
    }

    // --- 4. The constant, averaged -----------------------------------
    // One tree is one sample. The theoretical result is that a random
    // BST's expected height approaches 4.311 ln n = 2.988 lg n, and
    // its mean node depth approaches 2 ln n = 1.386 lg n. Both are
    // measurable; average over 25 trees.
    // Quoting only the leading constants (2.988 and 1.386) would make
    // the measurement look wrong. Both have lower-order corrections
    // that are still large at these sizes:
    //   height     ~ 4.311 ln n - 1.953 ln ln n + O(1)   (Reed, 2003)
    //   mean depth ~ 2 ln n - 3 + O(1)
    // So the honest comparison is against the corrected forms, with
    // the bare asymptote shown beside them to make the gap visible.
    std::printf("\nC. the random-BST constants (mean of 25 trees),\n");
    std::printf("   all divided by lg n\n\n");
    std::printf("%9s %9s %9s %9s %10s %9s %9s\n",
                "n", "height", "corrected", "asympt.",
                "mean dep", "corrected", "asympt.");
    for (size_t n = 2000; n <= 128000; n *= 2) {
        double hsum = 0, dsum = 0;
        for (int t = 0; t < 25; ++t) {
            std::vector<int> keys(n);
            std::iota(keys.begin(), keys.end(), 0);
            std::shuffle(keys.begin(), keys.end(), rng);
            BST a(n);
            for (int k : keys) a.insert(k);
            double lg = std::log2((double)n);
            hsum += a.height() / lg;
            dsum += a.mean_depth() / lg;
        }
        double ln = std::log((double)n), lg = std::log2((double)n);
        double h_pred = (4.31107 * ln - 1.9533 * std::log(ln)) / lg;
        double d_pred = (2.0 * ln - 3.0) / lg;
        std::printf("%9zu %9.3f %9.3f %9.3f %10.3f %9.3f %9.3f\n",
                    n, hsum / 25, h_pred, 2.988,
                    dsum / 25, d_pred, 1.386);
    }

    // --- 5. And what the height costs, in seconds --------------------
    // Searching a degenerate BST is a linear scan with pointer chasing,
    // which is worse than a linear scan of an array: same number of
    // steps, none of them predictable to the prefetcher.
    std::printf("\nD. 100,000 successful searches (ms)\n\n");
    std::printf("%10s %14s %14s %12s\n",
                "n", "random BST", "sorted BST", "ratio");
    for (size_t n = 2000; n <= 32000; n *= 2) {
        std::vector<int> keys(n);
        std::iota(keys.begin(), keys.end(), 0);
        std::vector<int> shuf = keys;
        std::shuffle(shuf.begin(), shuf.end(), rng);
        BST a(n); for (int k : shuf) a.insert(k);
        BST b(n); for (int k : keys) b.insert(k);

        std::vector<int> probe(100000);
        std::uniform_int_distribution<int> d(0, (int)n - 1);
        for (auto& x : probe) x = d(rng);

        volatile int hits = 0;
        double ta = time_ms([&] {
            int h = 0; for (int x : probe) h += a.contains(x); hits = h;
        }, 3);
        double tb = time_ms([&] {
            int h = 0; for (int x : probe) h += b.contains(x); hits = h;
        }, 3);
        std::printf("%10zu %14.2f %14.2f %12.1f\n", n, ta, tb, tb / ta);
    }
    return 0;
}
