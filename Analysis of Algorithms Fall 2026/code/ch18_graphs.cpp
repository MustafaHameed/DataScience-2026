// ch18_graphs.cpp -- the representation decides the bound.
//
// Every bound in Part V is quoted in TWO parameters, |V| and |E|, and the
// reason is that a graph algorithm's cost depends on how the graph is
// stored as much as on what the algorithm does.
//
//   adjacency matrix:  O(V^2) space, O(1) edge test, O(V) to list a
//                      vertex's neighbours -- whether it has 2 or 2000
//   adjacency list:    O(V+E) space, O(deg) edge test, O(deg) to list
//
// So BFS is Theta(V^2) from a matrix and Theta(V+E) from a list. On a
// dense graph those are the same; on a sparse one they are not, and this
// program finds the crossover rather than asserting it.
//
// Measured here:
//   A. memory, which is the first thing that makes the choice for you;
//   B. BFS from each representation across a range of densities;
//   C. the crossover in density, located;
//   D. one edge test versus one neighbour sweep, which is the whole
//      trade stated in two numbers.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch18_graphs.cpp
#include <algorithm>
#include <chrono>
#include <intrin.h>
#include <cstdio>
#include <numeric>
#include <queue>
#include <random>
#include <vector>

// --- 1. The two representations --------------------------------------
// The matrix is a bitset-free vector<char> so that one entry is one
// byte: V^2 bytes, which at V = 20,000 is already 400 MB and makes the
// point about space without any timing at all.
struct Matrix {
    size_t V;
    std::vector<char> a;
    explicit Matrix(size_t v) : V(v), a(v * v, 0) {}
    void add(size_t u, size_t w) { a[u * V + w] = a[w * V + u] = 1; }
    bool has(size_t u, size_t w) const { return a[u * V + w] != 0; }
    size_t bytes() const { return a.size(); }
};

// --- 1b. The matrix, bit-packed --------------------------------------
// One BIT per potential edge rather than one byte, so V^2/8 bytes, and
// a whole 64-bit word of candidate neighbours can be tested at once.
// This is the real reason adjacency matrices are recommended for dense
// graphs, and without it the comparison in part C is unfair to them.
struct BitMatrix {
    size_t V, W;                       // W = words per row
    std::vector<unsigned long long> a;
    explicit BitMatrix(size_t v) : V(v), W((v + 63) / 64), a(v * W, 0) {}
    void add(size_t u, size_t w) {
        a[u * W + (w >> 6)] |= 1ULL << (w & 63);
        a[w * W + (u >> 6)] |= 1ULL << (u & 63);
    }
    size_t bytes() const { return a.size() * sizeof(unsigned long long); }
};

// Compressed adjacency list: one flat array of neighbours plus an index
// of where each vertex's run begins. Same information as vector<vector>,
// one allocation instead of V of them, and contiguous in memory.
struct List {
    size_t V;
    std::vector<size_t> start;     // V+1 entries
    std::vector<unsigned> nbr;
    size_t bytes() const {
        return start.size() * sizeof(size_t) + nbr.size() * sizeof(unsigned);
    }
};

// --- 2. A random graph of a given average degree ---------------------
// Edges are drawn uniformly, so the density is controlled directly:
// |E| = avg_deg * V / 2.
static std::vector<std::pair<unsigned, unsigned>>
random_edges(size_t V, double avg_deg, std::mt19937& rng) {
    size_t E = (size_t)(avg_deg * V / 2.0);
    std::vector<std::pair<unsigned, unsigned>> e;
    e.reserve(E);
    std::uniform_int_distribution<unsigned> d(0, (unsigned)V - 1);
    while (e.size() < E) {
        unsigned u = d(rng), w = d(rng);
        if (u != w) e.push_back({u, w});
    }
    return e;
}

static List build_list(size_t V,
                       const std::vector<std::pair<unsigned, unsigned>>& e) {
    List g;
    g.V = V;
    g.start.assign(V + 1, 0);
    for (auto& p : e) { ++g.start[p.first + 1]; ++g.start[p.second + 1]; }
    for (size_t i = 0; i < V; ++i) g.start[i + 1] += g.start[i];
    g.nbr.resize(g.start[V]);
    std::vector<size_t> fill(g.start.begin(), g.start.end() - 1);
    for (auto& p : e) {
        g.nbr[fill[p.first]++] = p.second;
        g.nbr[fill[p.second]++] = p.first;
    }
    return g;
}

// --- 3. BFS from each representation ---------------------------------
// The algorithms are identical. The only difference is how line "for
// each neighbour of u" is implemented, and that difference is the whole
// of the asymptotic gap.
static long long bfs_matrix(const Matrix& g, size_t s) {
    std::vector<char> seen(g.V, 0);
    std::vector<unsigned> q;
    q.reserve(g.V);
    q.push_back((unsigned)s);
    seen[s] = 1;
    long long reached = 1;
    for (size_t h = 0; h < q.size(); ++h) {
        size_t u = q[h];
        for (size_t w = 0; w < g.V; ++w)        // <-- Theta(V) per vertex
            if (g.has(u, w) && !seen[w]) {
                seen[w] = 1; ++reached; q.push_back((unsigned)w);
            }
    }
    return reached;
}
// The bit-packed scan: skip 64 non-neighbours at a time, and within a
// word walk only the bits that are set. No branch per candidate.
static long long bfs_bitmatrix(const BitMatrix& g, size_t s) {
    std::vector<char> seen(g.V, 0);
    std::vector<unsigned> q;
    q.reserve(g.V);
    q.push_back((unsigned)s);
    seen[s] = 1;
    long long reached = 1;
    for (size_t h = 0; h < q.size(); ++h) {
        size_t u = q[h];
        const unsigned long long* row = &g.a[u * g.W];
        for (size_t wi = 0; wi < g.W; ++wi) {
            unsigned long long bits = row[wi];
            while (bits) {
                // The index of the lowest set bit, as ONE instruction.
                // Finding it with a shift loop instead costs up to 64
                // iterations per edge and makes the bit-packed matrix
                // slower than the byte one -- which is what the first
                // version of this program measured.
                unsigned long idx;
                _BitScanForward64(&idx, bits);
                size_t w = wi * 64 + idx;
                if (w < g.V && !seen[w]) {
                    seen[w] = 1; ++reached; q.push_back((unsigned)w);
                }
                bits &= bits - 1;   // clear it
            }
        }
    }
    return reached;
}

static long long bfs_list(const List& g, size_t s) {
    std::vector<char> seen(g.V, 0);
    std::vector<unsigned> q;
    q.reserve(g.V);
    q.push_back((unsigned)s);
    seen[s] = 1;
    long long reached = 1;
    for (size_t h = 0; h < q.size(); ++h) {
        size_t u = q[h];
        for (size_t i = g.start[u]; i < g.start[u + 1]; ++i) {  // Theta(deg)
            unsigned w = g.nbr[i];
            if (!seen[w]) { seen[w] = 1; ++reached; q.push_back(w); }
        }
    }
    return reached;
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

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(314159);

    // --- 4. Space, which usually decides before speed does -----------
    std::printf("A. memory for the two representations\n\n");
    std::printf("%10s %10s %16s %16s %12s\n",
                "V", "avg deg", "matrix (MB)", "list (MB)", "ratio");
    for (size_t V : {1000u, 4000u, 16000u, 64000u}) {
        double deg = 8;
        size_t E = (size_t)(deg * V / 2);
        double mb_m = (double)V * V / 1048576.0;
        double mb_l = ((V + 1) * sizeof(size_t) + 2 * E * sizeof(unsigned))
                      / 1048576.0;
        std::printf("%10zu %10.0f %16.1f %16.2f %12.0f\n",
                    V, deg, mb_m, mb_l, mb_m / mb_l);
    }

    // --- 5. BFS across a range of densities --------------------------
    // V is FIXED. Only the number of edges changes, so the matrix
    // version should take the same time throughout -- it reads all V^2
    // entries regardless -- while the list version grows with E.
    std::printf("\nB. BFS, V = 8000 fixed, density varying\n\n");
    std::printf("%10s %12s %16s %14s %12s\n",
                "avg deg", "|E|", "matrix (ms)", "list (ms)", "speedup");
    {
        size_t V = 8000;
        for (double deg : {2.0, 4.0, 8.0, 16.0, 64.0, 256.0, 1024.0}) {
            auto e = random_edges(V, deg, rng);
            Matrix m(V);
            for (auto& p : e) m.add(p.first, p.second);
            List l = build_list(V, e);
            double tm = time_ms([&] { return bfs_matrix(m, 0); }, 3);
            double tl = time_ms([&] { return bfs_list(l, 0); }, 3);
            std::printf("%10.0f %12zu %16.2f %14.2f %12.1f\n",
                        deg, e.size(), tm, tl, tm / tl);
        }
    }

    // --- 6. Where is the crossover? ----------------------------------
    // The matrix wins only when the graph is so dense that V^2 is no
    // worse than V+E -- that is, when E is close to V^2. Find the
    // density at which the list version stops winning.
    // A byte-per-entry matrix would lose this unfairly, so the
    // bit-packed version is included: V^2/8 bytes, and 64 candidate
    // neighbours tested per word.
    std::printf("\nC. the crossover, V = 4000\n\n");
    std::printf("%9s %8s %12s %12s %12s %9s\n",
                "avg deg", "density", "byte mtx", "bit mtx", "list",
                "winner");
    {
        size_t V = 4000;
        for (double deg : {512.0, 1024.0, 2048.0, 3000.0, 3990.0}) {
            auto e = random_edges(V, deg, rng);
            Matrix m(V);
            BitMatrix bm(V);
            for (auto& p : e) { m.add(p.first, p.second);
                                bm.add(p.first, p.second); }
            List l = build_list(V, e);
            double tm = time_ms([&] { return bfs_matrix(m, 0); }, 3);
            double tb = time_ms([&] { return bfs_bitmatrix(bm, 0); }, 3);
            double tl = time_ms([&] { return bfs_list(l, 0); }, 3);
            const char* win = (tb < tl) ? "bit mtx" : "list";
            std::printf("%9.0f %8.3f %12.2f %12.2f %12.2f %9s\n",
                        deg, 2.0 * e.size() / ((double)V * V),
                        tm, tb, tl, win);
        }
    }

    // --- 7. The trade, in two numbers --------------------------------
    // The matrix answers "is there an edge u-w?" in one memory access.
    // The list must scan u's neighbours. That is the one thing the
    // matrix is better at, and it is worth measuring rather than
    // waving at.
    std::printf("\nD. one edge test, 10 million times\n\n");
    std::printf("%10s %16s %16s %12s\n",
                "avg deg", "matrix (ms)", "list (ms)", "list/matrix");
    {
        size_t V = 4000;
        std::uniform_int_distribution<unsigned> d(0, (unsigned)V - 1);
        std::vector<std::pair<unsigned, unsigned>> probe(10000000);
        for (auto& p : probe) p = {d(rng), d(rng)};
        for (double deg : {4.0, 32.0, 256.0}) {
            auto e = random_edges(V, deg, rng);
            Matrix m(V);
            for (auto& p : e) m.add(p.first, p.second);
            List l = build_list(V, e);
            double tm = time_ms([&] {
                long long c = 0;
                for (auto& p : probe) c += m.has(p.first, p.second);
                return c;
            }, 3);
            double tl = time_ms([&] {
                long long c = 0;
                for (auto& p : probe) {
                    for (size_t i = l.start[p.first];
                         i < l.start[p.first + 1]; ++i)
                        if (l.nbr[i] == p.second) { ++c; break; }
                }
                return c;
            }, 3);
            std::printf("%10.0f %16.1f %16.1f %12.1f\n",
                        deg, tm, tl, tl / tm);
        }
    }
    return 0;
}
