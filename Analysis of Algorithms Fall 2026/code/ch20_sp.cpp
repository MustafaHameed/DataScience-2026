// ch20_sp.cpp -- three shortest-path algorithms, and when each one wins.
//
//   Dijkstra       single source, non-negative weights   O((V+E) lg V)
//   Bellman-Ford   single source, negative weights ok    O(VE)
//   Floyd-Warshall all pairs                             Theta(V^3)
//
// They are not interchangeable. Dijkstra is much the fastest and is
// WRONG on negative edges -- which this program demonstrates rather than
// asserts, because "you must not use it there" is more convincing with a
// counter-example attached.
//
// Measured here:
//   A. Dijkstra with a binary heap against the O(V^2) array version,
//      across densities -- the same algorithm, two priority queues;
//   B. Dijkstra against Bellman-Ford on the same graphs;
//   C. a graph with one negative edge, where Dijkstra returns a wrong
//      answer and Bellman-Ford does not;
//   D. all-pairs: Floyd-Warshall against running Dijkstra V times.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch20_sp.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <limits>
#include <numeric>
#include <queue>
#include <random>
#include <vector>

static const long long INF = (1LL << 60);
struct Arc { unsigned to; int w; };
using Graph = std::vector<std::vector<Arc>>;

// --- 1. Dijkstra with a binary heap ----------------------------------
// The "lazy" form: push a new entry on every improvement and skip stale
// ones when popped. O((V+E) lg V), and simpler than decrease-key.
//
// The invariant is the whole correctness argument: when a vertex is
// popped with the smallest tentative distance, that distance is final.
// It holds only because every edge weight is >= 0 -- no later path can
// come back and be shorter. Part C is what happens when that fails.
static std::vector<long long> dijkstra_heap(const Graph& g, unsigned s) {
    std::vector<long long> d(g.size(), INF);
    using P = std::pair<long long, unsigned>;
    std::priority_queue<P, std::vector<P>, std::greater<P>> pq;
    d[s] = 0;
    pq.push({0, s});
    while (!pq.empty()) {
        auto [du, u] = pq.top(); pq.pop();
        if (du > d[u]) continue;                     // stale
        for (const Arc& a : g[u])
            if (du + a.w < d[a.to]) {
                d[a.to] = du + a.w;
                pq.push({d[a.to], a.to});
            }
    }
    return d;
}

// --- 2. Dijkstra with a linear scan ----------------------------------
// Same algorithm; the priority queue is an unsorted array, so
// extract-min is O(V) and the total is O(V^2 + E). On a DENSE graph
// that beats the heap, because V^2 ~ E and the heap's lg V is pure
// overhead. This is the classic case of a "worse" bound winning.
static std::vector<long long> dijkstra_array(const Graph& g, unsigned s) {
    size_t V = g.size();
    std::vector<long long> d(V, INF);
    std::vector<char> done(V, 0);
    d[s] = 0;
    for (size_t it = 0; it < V; ++it) {
        long long best = INF;
        size_t u = V;
        for (size_t i = 0; i < V; ++i)
            if (!done[i] && d[i] < best) { best = d[i]; u = i; }
        if (u == V) break;
        done[u] = 1;
        for (const Arc& a : g[u])
            if (d[u] + a.w < d[a.to]) d[a.to] = d[u] + a.w;
    }
    return d;
}

// --- 3. Bellman-Ford: relax every edge, V-1 times --------------------
// No priority queue and no assumption about signs. The bound is O(VE),
// which is far worse -- and it is the price of handling negative edges
// and of being able to DETECT a negative cycle, which Dijkstra cannot.
static std::vector<long long> bellman_ford(const Graph& g, unsigned s,
                                           bool* neg_cycle) {
    size_t V = g.size();
    std::vector<long long> d(V, INF);
    d[s] = 0;
    for (size_t pass = 0; pass + 1 < V; ++pass) {
        bool changed = false;
        for (size_t u = 0; u < V; ++u) {
            if (d[u] == INF) continue;
            for (const Arc& a : g[u])
                if (d[u] + a.w < d[a.to]) { d[a.to] = d[u] + a.w;
                                            changed = true; }
        }
        if (!changed) break;                 // early exit: often much faster
    }
    if (neg_cycle) {                          // one more pass detects it
        *neg_cycle = false;
        for (size_t u = 0; u < V && !*neg_cycle; ++u) {
            if (d[u] == INF) continue;
            for (const Arc& a : g[u])
                if (d[u] + a.w < d[a.to]) { *neg_cycle = true; break; }
        }
    }
    return d;
}

// --- 4. Floyd-Warshall: all pairs, three nested loops ----------------
// The k loop must be OUTERMOST. That is not a style choice: k indexes
// the subproblem ("paths using only intermediate vertices < k"), and
// putting it inside would compute something that is not a shortest path.
static std::vector<long long> floyd_warshall(const Graph& g) {
    size_t V = g.size();
    std::vector<long long> d(V * V, INF);
    for (size_t i = 0; i < V; ++i) d[i * V + i] = 0;
    for (size_t u = 0; u < V; ++u)
        for (const Arc& a : g[u])
            d[u * V + a.to] = std::min(d[u * V + a.to], (long long)a.w);
    for (size_t k = 0; k < V; ++k)
        for (size_t i = 0; i < V; ++i) {
            long long dik = d[i * V + k];
            if (dik == INF) continue;
            for (size_t j = 0; j < V; ++j)
                if (dik + d[k * V + j] < d[i * V + j])
                    d[i * V + j] = dik + d[k * V + j];
        }
    return d;
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

static Graph random_graph(size_t V, double avg_deg, std::mt19937& rng,
                          int wmin = 1, int wmax = 1000) {
    Graph g(V);
    std::uniform_int_distribution<unsigned> dv(0, (unsigned)V - 1);
    std::uniform_int_distribution<int> dw(wmin, wmax);
    size_t E = (size_t)(avg_deg * V);
    for (size_t i = 0; i < E; ++i) {
        unsigned u = dv(rng), v = dv(rng);
        if (u != v) g[u].push_back({v, dw(rng)});
    }
    return g;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(271828);

    // --- 5. Heap against array, across densities ---------------------
    // V is fixed; only the edge count changes. The array version does
    // V^2 work regardless; the heap version grows with E.
    std::printf("A. Dijkstra: binary heap against linear scan, V = 4000\n\n");
    std::printf("%10s %12s %14s %14s %10s\n",
                "avg deg", "|E|", "heap (ms)", "array (ms)", "winner");
    {
        size_t V = 4000;
        for (double deg : {2.0, 8.0, 32.0, 128.0, 512.0, 2000.0}) {
            Graph g = random_graph(V, deg, rng);
            size_t E = 0; for (auto& a : g) E += a.size();
            double th = time_ms([&] { dijkstra_heap(g, 0); }, 3);
            double ta = time_ms([&] { dijkstra_array(g, 0); }, 3);
            std::printf("%10.0f %12zu %14.2f %14.2f %10s\n",
                        deg, E, th, ta, (th < ta) ? "heap" : "array");
        }
    }

    // --- 6. Dijkstra against Bellman-Ford ----------------------------
    // Same answers on non-negative graphs -- checked, not assumed --
    // and very different costs.
    std::printf("\nB. Dijkstra against Bellman-Ford (non-negative weights)\n\n");
    std::printf("%10s %12s %14s %16s %10s %10s\n",
                "V", "|E|", "Dijkstra (ms)", "Bellman-Ford (ms)",
                "ratio", "agree?");
    for (size_t V = 2000; V <= 32000; V *= 2) {
        Graph g = random_graph(V, 8.0, rng);
        size_t E = 0; for (auto& a : g) E += a.size();
        auto d1 = dijkstra_heap(g, 0);
        bool nc = false;
        auto d2 = bellman_ford(g, 0, &nc);
        double t1 = time_ms([&] { dijkstra_heap(g, 0); }, 3);
        double t2 = time_ms([&] { bellman_ford(g, 0, nullptr); }, 3);
        std::printf("%10zu %12zu %14.2f %16.1f %10.1f %10s\n",
                    V, E, t1, t2, t2 / t1, (d1 == d2) ? "yes" : "NO");
    }
    std::printf("\n   Bellman-Ford keeps up only because of the early\n"
                "   exit: on a random graph the distances stop changing\n"
                "   after a handful of passes. Its O(VE) bound needs a\n"
                "   graph that forces all V-1 passes, so here is one.\n\n");

    // --- 6b. The graph that forces every pass ------------------------
    // A single path 0 -> 1 -> ... -> V-1, with the adjacency lists
    // visited in an order that propagates exactly one vertex per pass.
    // Nothing adversarial about the weights; only the ORDER matters.
    std::printf("%10s %12s %16s %12s %12s\n",
                "V", "Dijkstra", "Bellman-Ford", "ratio", "passes used");
    for (size_t V = 1000; V <= 8000; V *= 2) {
        Graph g(V);
        // edge i -> i+1 stored at vertex i, but we relax vertices in
        // increasing order, so a path laid out BACKWARDS is the bad case
        for (size_t i = V - 1; i > 0; --i) g[i].push_back({(unsigned)(i - 1), 1});
        double t1 = time_ms([&] { dijkstra_heap(g, (unsigned)(V - 1)); }, 3);
        double t2 = time_ms([&] {
            bellman_ford(g, (unsigned)(V - 1), nullptr); }, 3);
        std::printf("%10zu %12.2f %16.1f %12.1f %12zu\n",
                    V, t1, t2, t2 / t1, V - 1);
    }

    // --- 7. One negative edge, and Dijkstra is wrong -----------------
    // A tiny graph. 0->1 costs 5, 0->2 costs 2, 2->1 costs -4. The true
    // distance to 1 is 2 + (-4) = -2, but Dijkstra finalises vertex 1
    // at 5 before ever looking at the edge from 2.
    // Which Dijkstra matters here. The ARRAY version marks a vertex
    // done and never looks at it again -- that is textbook Dijkstra, and
    // it is the one that fails. The lazy HEAP version never finalises
    // anything; it just re-pushes on every improvement, so on this graph
    // it happens to recover. That is luck, not correctness: it is really
    // behaving like Bellman-Ford, and on adversarial graphs with
    // negative edges it can take exponential time. Both are shown,
    // because "Dijkstra is wrong on negative edges" turns out to depend
    // on which implementation you mean.
    std::printf("\nC. a negative edge, on four vertices\n\n");
    {
        // 0 -> 1 costs 3, 0 -> 2 costs 4, 2 -> 1 costs -2, 1 -> 3 costs 1.
        // True distances: d(1) = 2 via vertex 2, so d(3) = 3.
        //
        // Array Dijkstra pops vertex 1 at distance 3 and finalises it,
        // relaxing 1 -> 3 to give d(3) = 4. It then pops vertex 2 and
        // lowers d(1) to 2 -- but vertex 1 is done, so its edge to 3 is
        // never re-examined and d(3) stays at 4. The error has
        // PROPAGATED, which is why this graph fails where the obvious
        // three-vertex example does not: there the wrong value was
        // overwritten before anything depended on it.
        Graph g(4);
        g[0].push_back({1, 3});
        g[0].push_back({2, 4});
        g[2].push_back({1, -2});
        g[1].push_back({3, 1});
        auto da = dijkstra_array(g, 0);
        auto dd = dijkstra_heap(g, 0);
        bool nc = false;
        auto db = bellman_ford(g, 0, &nc);
        std::printf("%8s %14s %12s %10s %9s\n",
                    "vertex", "Dijkstra(arr)", "Dijkstra(heap)",
                    "Bellman", "correct");
        const long long truth[] = {0, 2, 4, 3};
        for (int v = 0; v < 4; ++v)
            std::printf("%8d %14lld %12lld %10lld %9lld %s\n",
                        v, da[v], dd[v], db[v], truth[v],
                        (da[v] != truth[v]) ? "  <-- array WRONG" : "");
        std::printf("\n   negative cycle detected: %s\n", nc ? "yes" : "no");

        // And now a graph that DOES have a negative cycle.
        Graph h(3);
        h[0].push_back({1, 1});
        h[1].push_back({2, -3});
        h[2].push_back({0, 1});
        bool nc2 = false;
        bellman_ford(h, 0, &nc2);
        std::printf("   with a negative cycle added: detected = %s\n",
                    nc2 ? "yes" : "no");
    }

    // --- 8. All pairs: Floyd-Warshall against V x Dijkstra -----------
    // Theta(V^3) regardless of density, against O(V(V+E) lg V). On a
    // sparse graph repeated Dijkstra wins easily; on a dense one
    // Floyd-Warshall's tight triple loop and perfect locality win.
    std::printf("\nD. all pairs, V = 600\n\n");
    std::printf("%10s %12s %18s %18s %10s\n",
                "avg deg", "|E|", "Floyd-Warshall", "V x Dijkstra",
                "winner");
    {
        size_t V = 600;
        for (double deg : {2.0, 8.0, 32.0, 128.0, 300.0}) {
            Graph g = random_graph(V, deg, rng);
            size_t E = 0; for (auto& a : g) E += a.size();
            double tf = time_ms([&] { floyd_warshall(g); }, 3);
            double td = time_ms([&] {
                for (unsigned s = 0; s < V; ++s) dijkstra_heap(g, s);
            }, 3);
            std::printf("%10.0f %12zu %18.1f %18.1f %10s\n",
                        deg, E, tf, td, (tf < td) ? "Floyd" : "Dijkstra");
        }
    }
    return 0;
}
