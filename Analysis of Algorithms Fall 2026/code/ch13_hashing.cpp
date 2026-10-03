// ch13_hashing.cpp -- what O(1) actually promises, and what voids it.
//
// A hash table's famous O(1) is an EXPECTED cost, and the expectation is
// taken over an assumption: that the hash function scatters the keys
// evenly. That assumption is about the relationship between your hash and
// your data, and it can fail -- by accident, or because somebody made it.
//
// Measured here:
//
//   A. the chain-length distribution under a good hash, against the
//      Poisson prediction that the analysis is built on;
//   B. the SAME table and the SAME keys with a hash that is perfectly
//      respectable and happens to collide with the data's structure;
//   C. load factor against probe length, which is the one knob;
//   D. an adversarial key set, built deliberately, which turns the
//      table into a linked list.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch13_hashing.cpp
#include <algorithm>
#include <chrono>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

// --- 1. A chaining hash table, as plainly as possible ----------------
// Buckets of vectors. No resizing, no tricks: the table size is fixed
// so that the load factor is exactly what the experiment says it is.
struct Table {
    std::vector<std::vector<unsigned>> bucket;
    unsigned (*h)(unsigned, unsigned);
    size_t m;
    Table(size_t slots, unsigned (*hash)(unsigned, unsigned))
        : bucket(slots), h(hash), m(slots) {}
    void insert(unsigned k) { bucket[h(k, (unsigned)m)].push_back(k); }
    // Probes to find k: the position reached within its chain.
    size_t probes(unsigned k) const {
        const auto& b = bucket[h(k, (unsigned)m)];
        for (size_t i = 0; i < b.size(); ++i)
            if (b[i] == k) return i + 1;
        return b.size() + 1;
    }
    size_t longest() const {
        size_t L = 0;
        for (const auto& b : bucket) L = std::max(L, b.size());
        return L;
    }
    double mean_probe(const std::vector<unsigned>& keys) const {
        double s = 0;
        for (unsigned k : keys) s += (double)probes(k);
        return s / keys.size();
    }
};

// --- 2. Two hash functions, both entirely reasonable -----------------
// "Modulo the table size" is the textbook division method and is what
// most people write first. It is fine when the table size is prime and
// the keys are unstructured, and catastrophic when the keys share a
// factor with the table size.
static unsigned h_mod(unsigned k, unsigned m) { return k % m; }

// A multiplicative mix (the finaliser from MurmurHash3). It destroys
// structure in the key before taking the modulus, which is the whole
// job of a hash function.
static unsigned h_mix(unsigned k, unsigned m) {
    k ^= k >> 16; k *= 0x85ebca6bu;
    k ^= k >> 13; k *= 0xc2b2ae35u;
    k ^= k >> 16;
    return k % m;
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

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(13131);

    // --- 3. The distribution the analysis assumes --------------------
    // With n keys in m slots and perfect scattering, a slot's load is
    // Binomial(n, 1/m), which for large m is Poisson(alpha). The
    // analysis of "expected O(1)" rests on that shape, so check it.
    std::printf("A. chain lengths under a good hash,"
                " n = m = 1,000,000 (alpha = 1)\n\n");
    {
        size_t n = 1000000, m = 1000000;
        Table t(m, h_mix);
        std::vector<unsigned> keys(n);
        std::uniform_int_distribution<unsigned> d(0, 0xFFFFFFFFu);
        for (auto& k : keys) k = d(rng);
        for (unsigned k : keys) t.insert(k);

        std::vector<size_t> hist(10, 0);
        for (const auto& b : t.bucket)
            ++hist[std::min<size_t>(b.size(), 9)];
        std::printf("%8s %14s %14s %12s\n",
                    "length", "slots", "Poisson(1)", "ratio");
        double e = std::exp(-1.0), fact = 1;
        for (size_t L = 0; L < 7; ++L) {
            if (L > 0) fact *= L;
            double pred = m * e / fact;
            std::printf("%8zu %14zu %14.0f %12.3f\n",
                        L, hist[L], pred, hist[L] / pred);
        }
        std::printf("   longest chain: %zu   (lg n / lg lg n = %.1f)\n",
                    t.longest(),
                    std::log2((double)n) / std::log2(std::log2((double)n)));
    }

    // --- 4. The same keys, a hash that meets the data's structure ----
    // Real keys are not random. They are IDs, timestamps, pointers,
    // struct offsets -- and they very often share a stride. Here every
    // key is a multiple of 64, which is what you get from an array of
    // 64-byte records, and the table size is a power of two, which is
    // what you get from every "round up to the next power of 2".
    std::printf("\nB. the same table, keys with a stride of 64\n\n");
    std::printf("%22s %14s %14s %14s\n",
                "hash", "mean probes", "longest chain", "slots used");
    {
        size_t n = 1000000, m = 1u << 20;       // 1,048,576: power of two
        std::vector<unsigned> keys(n);
        for (size_t i = 0; i < n; ++i) keys[i] = (unsigned)(i * 64);
        for (auto hf : {h_mod, h_mix}) {
            Table t(m, hf);
            for (unsigned k : keys) t.insert(k);
            size_t used = 0;
            for (const auto& b : t.bucket) if (!b.empty()) ++used;
            std::printf("%22s %14.1f %14zu %14zu\n",
                        (hf == h_mod) ? "k % m" : "mix then % m",
                        t.mean_probe(keys), t.longest(), used);
        }
    }

    // --- 5. Load factor is the one knob you actually turn ------------
    // The expected probe length for a successful search in a chaining
    // table is 1 + alpha/2. That is a prediction with no fitted
    // constant in it, so it is worth checking directly.
    std::printf("\nC. load factor against probe length (good hash)\n\n");
    std::printf("%10s %14s %16s %12s\n",
                "alpha", "mean probes", "1 + alpha/2", "lookup (ms)");
    {
        size_t m = 1 << 20;
        std::uniform_int_distribution<unsigned> d(0, 0xFFFFFFFFu);
        for (double alpha : {0.25, 0.5, 1.0, 2.0, 4.0, 8.0}) {
            size_t n = (size_t)(alpha * m);
            Table t(m, h_mix);
            std::vector<unsigned> keys(n);
            for (auto& k : keys) k = d(rng);
            for (unsigned k : keys) t.insert(k);
            // Sample the keys at RANDOM, not the first 200,000. The
            // first-inserted keys sit at the front of their chains, so
            // probing them reports 1.096 probes whatever the load
            // factor is -- which is what the first draft of this
            // program did, and the flat column was the clue.
            std::vector<unsigned> probe(200000);
            std::uniform_int_distribution<size_t> pick(0, n - 1);
            for (auto& x : probe) x = keys[pick(rng)];
            volatile size_t sink = 0;
            double ms = time_ms([&] {
                size_t s = 0;
                for (unsigned k : probe) s += t.probes(k);
                sink = s;
            }, 3);
            std::printf("%10.2f %14.3f %16.3f %12.2f\n",
                        alpha, t.mean_probe(probe), 1 + alpha / 2, ms);
        }
    }

    // --- 6. An adversary, who does not need to be clever -------------
    // If the hash is fixed and public, the keys that land in one slot
    // can be enumerated. Here they are simply the multiples of m. No
    // cryptography, no timing attack: arithmetic.
    std::printf("\nD. adversarial keys: every one in the same slot\n\n");
    std::printf("%10s %16s %16s %14s\n",
                "n", "random keys (ms)", "chosen keys (ms)", "ratio");
    {
        size_t m = 1 << 16;
        std::uniform_int_distribution<unsigned> d(0, 0xFFFFFFFFu);
        for (size_t n = 2000; n <= 32000; n *= 2) {
            Table good(m, h_mod), bad(m, h_mod);
            std::vector<unsigned> rk(n), ak(n);
            for (auto& k : rk) k = d(rng);
            for (size_t i = 0; i < n; ++i)
                ak[i] = (unsigned)((i + 1) * m);   // all hash to 0
            for (unsigned k : rk) good.insert(k);
            for (unsigned k : ak) bad.insert(k);
            volatile size_t sink = 0;
            double tg = time_ms([&] {
                size_t s = 0; for (unsigned k : rk) s += good.probes(k);
                sink = s;
            }, 3);
            double tb = time_ms([&] {
                size_t s = 0; for (unsigned k : ak) s += bad.probes(k);
                sink = s;
            }, 3);
            std::printf("%10zu %16.2f %16.2f %14.1f\n",
                        n, tg, tb, tb / tg);
        }
        std::printf("\n   The chosen keys all satisfy k %% m == 0, so every\n"
                    "   one lands in slot 0 and the table is a linked list.\n");
    }
    return 0;
}
