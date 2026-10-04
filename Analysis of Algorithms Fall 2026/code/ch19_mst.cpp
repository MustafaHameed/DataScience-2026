// ch19_mst.cpp -- Kruskal, Prim, and where Chapter 17 pays off.
//
// Both algorithms are greedy and both are optimal, for the same reason:
// the cut property says the lightest edge crossing any cut is safe. They
// differ in which cut they take, and therefore in which data structure
// they need:
//
//   Kruskal  sorts the edges and adds any that joins two components
//            -> needs UNION-FIND, and its cost is an amortised bound
//   Prim     grows one tree, always taking the lightest edge leaving it
//            -> needs a PRIORITY QUEUE
//
// Measured here:
//   A. the three union-find variants, which differ by a factor that
//      grows -- naive, union by rank, and rank + path compression;
//   B. Kruskal against Prim across densities;
//   C. both against a brute-force check on small graphs;
//   D. the inverse-Ackermann bound, seen as a flat line.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch19_mst.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <numeric>
#include <queue>
#include <random>
#include <vector>

static long long steps = 0;      // pointer hops inside find()

// --- 1. Union-find, three ways ---------------------------------------
// The only differences are whether union balances the trees and whether
// find flattens the path behind it. Both are two lines, and together
// they take the amortised cost from O(lg n) to O(alpha(n)).
struct DSU {
    std::vector<int> parent, rank_;
    bool by_rank, compress;
    DSU(size_t n, bool r, bool c) : parent(n), rank_(n, 0),
                                    by_rank(r), compress(c) {
        std::iota(parent.begin(), parent.end(), 0);
    }
    int find(int x) {
        if (!compress) {
            while (parent[x] != x) { x = parent[x]; ++steps; }
            return x;
        }
        int root = x;
        while (parent[root] != root) { root = parent[root]; ++steps; }
        while (parent[x] != root) {          // second pass: flatten
            int next = parent[x];
            parent[x] = root;
            x = next;
        }
        return root;
    }
    bool unite(int a, int b) {
        a = find(a); b = find(b);
        if (a == b) return false;
        if (!by_rank) { parent[a] = b; return true; }
        if (rank_[a] < rank_[b]) std::swap(a, b);
        parent[b] = a;
        if (rank_[a] == rank_[b]) ++rank_[a];
        return true;
    }
};

struct Edge { unsigned u, v; int w; };

// --- 2. Kruskal: sort, then add anything that is safe ----------------
// The sort is Theta(E lg E) and dominates; the union-find work is
// E * alpha(V), which is why the structure's amortised cost matters.
static long long kruskal(std::vector<Edge> e, size_t V,
                         bool by_rank, bool compress) {
    std::sort(e.begin(), e.end(),
              [](const Edge& a, const Edge& b) { return a.w < b.w; });
    DSU d(V, by_rank, compress);
    long long total = 0;
    size_t used = 0;
    for (const Edge& x : e) {
        if (d.unite(x.u, x.v)) { total += x.w; if (++used == V - 1) break; }
    }
    return total;
}

// --- 3. Prim: grow one tree from a root ------------------------------
// With a binary heap this is O(E lg V). The "lazy" form below pushes
// every edge and discards stale entries when popped, which is simpler
// than decrease-key and no worse asymptotically.
static long long prim(const std::vector<std::vector<std::pair<unsigned, int>>>& g,
                      size_t V) {
    std::vector<char> in(V, 0);
    using P = std::pair<int, unsigned>;              // (weight, vertex)
    std::priority_queue<P, std::vector<P>, std::greater<P>> pq;
    pq.push({0, 0});
    long long total = 0;
    size_t count = 0;
    while (!pq.empty() && count < V) {
        auto [w, u] = pq.top(); pq.pop();
        if (in[u]) continue;                          // stale entry
        in[u] = 1; total += w; ++count;
        for (auto& [v, wt] : g[u])
            if (!in[v]) pq.push({wt, v});
    }
    return count == V ? total : -1;                   // -1 if disconnected
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

// A connected random graph: a random spanning path first, then extra
// edges. Connectivity matters -- an MST of a disconnected graph does
// not exist, and comparing -1 against a number teaches nothing.
static std::vector<Edge> random_connected(size_t V, double avg_deg,
                                          std::mt19937& rng) {
    std::uniform_int_distribution<int> dw(1, 1000000);
    std::uniform_int_distribution<unsigned> dv(0, (unsigned)V - 1);
    std::vector<unsigned> perm(V);
    std::iota(perm.begin(), perm.end(), 0u);
    std::shuffle(perm.begin(), perm.end(), rng);
    std::vector<Edge> e;
    for (size_t i = 1; i < V; ++i)
        e.push_back({perm[i - 1], perm[i], dw(rng)});
    size_t target = (size_t)(avg_deg * V / 2.0);
    while (e.size() < target) {
        unsigned u = dv(rng), v = dv(rng);
        if (u != v) e.push_back({u, v, dw(rng)});
    }
    return e;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(161803);

    // --- 4. Do the two algorithms agree? -----------------------------
    // They are different greedy rules. If the cut property is right,
    // they must produce the same total weight.
    std::printf("A. Kruskal against Prim, 2000 random graphs\n\n");
    {
        int T = 2000, bad = 0;
        for (int t = 0; t < T; ++t) {
            size_t V = 40;
            auto e = random_connected(V, 6.0, rng);
            std::vector<std::vector<std::pair<unsigned, int>>> g(V);
            for (auto& x : e) {
                g[x.u].push_back({x.v, x.w});
                g[x.v].push_back({x.u, x.w});
            }
            if (kruskal(e, V, true, true) != prim(g, V)) ++bad;
        }
        std::printf("   graphs tested: %d   disagreements: %d\n", T, bad);
    }

    // --- 5. The three union-find variants ----------------------------
    // "steps" counts pointer hops inside find(). This is the measurement
    // Chapter 17 predicts: naive grows, rank alone is much better, and
    // rank + compression is essentially flat.
    // The naive variant's find() is LINEAR in the component size, so
    // the whole run is quadratic and V = 320,000 does not finish in any
    // reasonable time. It stops at 160,000; the other two carry on, and
    // that asymmetry is itself the result.
    std::printf("\nB. union-find: pointer hops per find()\n\n");
    std::printf("%10s %14s %14s %16s %12s\n",
                "V", "naive", "by rank", "rank+compress", "naive/both");
    for (size_t V = 10000; V <= 160000; V *= 2) {
        auto e = random_connected(V, 8.0, rng);
        steps = 0; kruskal(e, V, false, false); long long s1 = steps;
        steps = 0; kruskal(e, V, true, false);  long long s2 = steps;
        steps = 0; kruskal(e, V, true, true);   long long s3 = steps;
        std::printf("%10zu %14.2f %14.2f %16.2f %12.1f\n", V,
                    (double)s1 / e.size(), (double)s2 / e.size(),
                    (double)s3 / e.size(), (double)s1 / s3);
    }
    std::printf("\n   and the two good variants alone, further out:\n\n");
    std::printf("%10s %14s %16s %14s\n",
                "V", "by rank", "rank+compress", "alpha(V) <= ");
    for (size_t V = 320000; V <= 5120000; V *= 2) {
        auto e = random_connected(V, 8.0, rng);
        steps = 0; kruskal(e, V, true, false);  long long s2 = steps;
        steps = 0; kruskal(e, V, true, true);   long long s3 = steps;
        std::printf("%10zu %14.2f %16.2f %14d\n", V,
                    (double)s2 / e.size(), (double)s3 / e.size(), 4);
    }

    // --- 6. And the same thing in seconds ----------------------------
    std::printf("\nC. Kruskal with each union-find variant (ms)\n\n");
    std::printf("%10s %14s %14s %16s %12s\n",
                "V", "naive", "by rank", "rank+compress", "speedup");
    for (size_t V = 25000; V <= 100000; V *= 2) {
        auto e = random_connected(V, 8.0, rng);
        double t1 = time_ms([&] { return kruskal(e, V, false, false); }, 3);
        double t2 = time_ms([&] { return kruskal(e, V, true, false); }, 3);
        double t3 = time_ms([&] { return kruskal(e, V, true, true); }, 3);
        std::printf("%10zu %14.1f %14.1f %16.1f %12.2f\n",
                    V, t1, t2, t3, t1 / t3);
    }

    // --- 7. Kruskal against Prim, by density -------------------------
    // Kruskal sorts E edges: Theta(E lg E), which on a dense graph is
    // Theta(V^2 lg V). Prim with a heap is Theta(E lg V). They are the
    // same order; the constants and the sort are what differ.
    std::printf("\nD. Kruskal against Prim, V = 100000\n\n");
    std::printf("%10s %12s %14s %14s %10s\n",
                "avg deg", "|E|", "Kruskal (ms)", "Prim (ms)", "winner");
    {
        size_t V = 100000;
        for (double deg : {2.0, 4.0, 8.0, 16.0, 32.0, 64.0}) {
            auto e = random_connected(V, deg, rng);
            std::vector<std::vector<std::pair<unsigned, int>>> g(V);
            for (auto& x : e) {
                g[x.u].push_back({x.v, x.w});
                g[x.v].push_back({x.u, x.w});
            }
            double tk = time_ms([&] { return kruskal(e, V, true, true); }, 3);
            double tp = time_ms([&] { return prim(g, V); }, 3);
            std::printf("%10.0f %12zu %14.1f %14.1f %10s\n",
                        deg, e.size(), tk, tp,
                        (tk < tp) ? "Kruskal" : "Prim");
        }
    }
    return 0;
}
