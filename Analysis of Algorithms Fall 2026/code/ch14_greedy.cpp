// ch14_greedy.cpp -- when a local choice can be trusted, and when it cannot.
//
// A greedy algorithm commits to the choice that looks best right now and
// never reconsiders. That is either optimal or badly wrong, and which one
// it is depends on a property of the problem, not on the care taken over
// the implementation.
//
// Measured here:
//
//   A. activity selection -- greedy by earliest finish time, against
//      brute force on small instances, to show it really is optimal;
//   B. three greedy rules for the SAME problem, two of which are wrong,
//      so that "greedy" is clearly not one algorithm;
//   C. coin change, where greedy is optimal for one coin system and
//      provably wrong for another -- the counter-example is tiny;
//   D. Huffman coding on real text, against the Shannon entropy that
//      bounds every prefix code.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch14_greedy.cpp
#include <algorithm>
#include <cmath>
#include <cstdio>
#include <numeric>
#include <queue>
#include <random>
#include <string>
#include <vector>

struct Act { int start, finish; };

// --- 1. Activity selection, greedy by earliest finish -----------------
// Sort by finishing time and take anything compatible. The claim is that
// the activity finishing first is in SOME optimal solution -- which is
// the greedy-choice property, and it is what the chapter proves.
static int greedy_finish(std::vector<Act> a) {
    std::sort(a.begin(), a.end(),
              [](const Act& x, const Act& y) { return x.finish < y.finish; });
    int count = 0, last = -1;
    for (const Act& x : a)
        if (x.start >= last) { ++count; last = x.finish; }
    return count;
}

// --- 2. Two plausible rules that are simply wrong --------------------
// Both "take the shortest" and "take the earliest starting" sound just as
// reasonable as rule 1. Neither is optimal, and the program finds the
// counter-examples by brute force rather than asserting it.
static int greedy_shortest(std::vector<Act> a) {
    std::sort(a.begin(), a.end(), [](const Act& x, const Act& y) {
        return (x.finish - x.start) < (y.finish - y.start);
    });
    std::vector<Act> chosen;
    for (const Act& x : a) {
        bool ok = true;
        for (const Act& c : chosen)
            if (x.start < c.finish && c.start < x.finish) { ok = false; break; }
        if (ok) chosen.push_back(x);
    }
    return (int)chosen.size();
}
static int greedy_earliest_start(std::vector<Act> a) {
    std::sort(a.begin(), a.end(),
              [](const Act& x, const Act& y) { return x.start < y.start; });
    int count = 0, last = -1;
    for (const Act& x : a)
        if (x.start >= last) { ++count; last = x.finish; }
    return count;
}

// --- 3. The oracle: every subset, checked ----------------------------
// Chapter 7's lesson. n <= 18 so 2^n is affordable, and it is the only
// way to be sure the greedy answers are right rather than plausible.
static int brute_force(const std::vector<Act>& a) {
    int n = (int)a.size(), best = 0;
    for (int mask = 0; mask < (1 << n); ++mask) {
        std::vector<Act> s;
        for (int i = 0; i < n; ++i) if (mask >> i & 1) s.push_back(a[i]);
        std::sort(s.begin(), s.end(),
                  [](const Act& x, const Act& y) { return x.start < y.start; });
        bool ok = true;
        for (size_t i = 1; i < s.size(); ++i)
            if (s[i].start < s[i - 1].finish) { ok = false; break; }
        if (ok) best = std::max(best, (int)s.size());
    }
    return best;
}

// --- 4. Coin change: greedy, and the optimum it sometimes misses -----
// Greedy takes the largest coin that fits. For {1, 5, 10, 25} that is
// optimal; for {1, 3, 4} it is not, and the smallest counter-example is
// six. Dynamic programming (Chapter 15) gets it right for any system.
static int coins_greedy(std::vector<int> c, int amount) {
    std::sort(c.rbegin(), c.rend());
    int used = 0;
    for (int v : c) while (amount >= v) { amount -= v; ++used; }
    return amount == 0 ? used : -1;
}
static int coins_optimal(const std::vector<int>& c, int amount) {
    const int INF = 1 << 29;
    std::vector<int> best(amount + 1, INF);
    best[0] = 0;
    for (int a = 1; a <= amount; ++a)
        for (int v : c)
            if (v <= a && best[a - v] + 1 < best[a]) best[a] = best[a - v] + 1;
    return best[amount] >= INF ? -1 : best[amount];
}

// --- 5. Huffman coding -----------------------------------------------
// Repeatedly merge the two least frequent symbols. The greedy choice is
// that the two rarest symbols can be put deepest, as siblings.
struct Node { long long freq; int left, right; };
static double huffman_bits(const std::vector<long long>& freq) {
    std::vector<Node> t;
    auto cmp = [&](int x, int y) { return t[x].freq > t[y].freq; };
    std::priority_queue<int, std::vector<int>, decltype(cmp)> pq(cmp);
    for (long long f : freq)
        if (f > 0) { t.push_back({f, -1, -1}); pq.push((int)t.size() - 1); }
    if (pq.size() == 1) return 1.0;                 // one symbol: 1 bit
    while (pq.size() > 1) {
        int x = pq.top(); pq.pop();
        int y = pq.top(); pq.pop();
        t.push_back({t[x].freq + t[y].freq, x, y});
        pq.push((int)t.size() - 1);
    }
    // Total encoded length = sum over leaves of freq * depth, which is
    // also the sum of every internal node's frequency.
    long long total = 0, bits = 0;
    for (size_t i = 0; i < t.size(); ++i)
        if (t[i].left >= 0) bits += t[i].freq;
    for (long long f : freq) total += f;
    return (double)bits / total;
}
static double entropy(const std::vector<long long>& freq) {
    long long total = std::accumulate(freq.begin(), freq.end(), 0LL);
    double h = 0;
    for (long long f : freq)
        if (f > 0) { double p = (double)f / total; h -= p * std::log2(p); }
    return h;
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(1848);

    // --- 6. Is the greedy rule actually optimal? ---------------------
    std::printf("A. activity selection: three greedy rules vs brute force\n\n");
    std::printf("%10s %12s %12s %14s %14s\n", "trials", "by finish",
                "by length", "by start", "optimal");
    {
        int T = 20000, okf = 0, oks = 0, oke = 0;
        std::uniform_int_distribution<int> ds(0, 40), dd(1, 12);
        for (int t = 0; t < T; ++t) {
            int n = 10;
            std::vector<Act> a(n);
            for (auto& x : a) { x.start = ds(rng); x.finish = x.start + dd(rng); }
            int opt = brute_force(a);
            okf += (greedy_finish(a) == opt);
            oks += (greedy_shortest(a) == opt);
            oke += (greedy_earliest_start(a) == opt);
        }
        std::printf("%10d %11.1f%% %11.1f%% %13.1f%% %14s\n", T,
                    100.0 * okf / T, 100.0 * oks / T, 100.0 * oke / T,
                    "(by definition)");
    }

    // --- 7. Coin change: the same algorithm, two coin systems --------
    std::printf("\nB. coin change: greedy against optimal\n\n");
    for (std::vector<int> sys : {std::vector<int>{1, 5, 10, 25},
                                 std::vector<int>{1, 3, 4}}) {
        int bad = 0, first = -1;
        for (int amt = 1; amt <= 200; ++amt) {
            int g = coins_greedy(sys, amt), o = coins_optimal(sys, amt);
            if (g != o) { ++bad; if (first < 0) first = amt; }
        }
        std::printf("   coins {");
        for (size_t i = 0; i < sys.size(); ++i)
            std::printf("%s%d", i ? ", " : "", sys[i]);
        std::printf("}: greedy wrong on %d of 200 amounts", bad);
        if (first > 0) {
            std::printf(", first at %d (greedy %d, optimal %d)",
                        first, coins_greedy(sys, first),
                        coins_optimal(sys, first));
        }
        std::printf("\n");
    }

    // --- 8. Huffman against the entropy bound ------------------------
    // Shannon: no prefix code can average fewer than H bits per symbol,
    // and Huffman is always within 1 bit of it. Both are checkable.
    std::printf("\nC. Huffman coding against the entropy bound\n\n");
    std::printf("%26s %10s %10s %10s %12s %10s\n", "text", "symbols",
                "entropy", "Huffman", "fixed 8-bit", "saving");
    struct Corpus { const char* name; std::string text; };
    std::vector<Corpus> corpora;
    {
        // English-like prose: the chapter's own opening sentence,
        // repeated so the frequencies are stable.
        std::string en =
            "a greedy algorithm commits to the choice that looks best "
            "right now and never reconsiders that is either optimal or "
            "badly wrong and which one it is depends on a property of "
            "the problem not on the care taken over the implementation ";
        std::string rep;
        for (int i = 0; i < 40; ++i) rep += en;
        corpora.push_back({"English prose", rep});

        // DNA: four symbols, nearly uniform. Entropy should be near 2.
        std::string dna;
        std::uniform_int_distribution<int> d4(0, 3);
        for (int i = 0; i < 40000; ++i) dna += "ACGT"[d4(rng)];
        corpora.push_back({"DNA (4 symbols)", dna});

        // Highly skewed: one symbol dominates. Entropy well below 1.
        std::string skew;
        std::uniform_int_distribution<int> d100(0, 99);
        for (int i = 0; i < 40000; ++i)
            skew += (d100(rng) < 95) ? 'a' : (char)('b' + d100(rng) % 5);
        corpora.push_back({"95% one symbol", skew});

        // Uniform over 256 byte values: entropy 8, so nothing to save.
        std::string uni;
        std::uniform_int_distribution<int> d256(0, 255);
        for (int i = 0; i < 40000; ++i) uni += (char)d256(rng);
        corpora.push_back({"uniform bytes", uni});
    }
    for (const Corpus& c : corpora) {
        std::vector<long long> freq(256, 0);
        for (unsigned char ch : c.text) ++freq[ch];
        int distinct = 0;
        for (long long f : freq) if (f) ++distinct;
        double H = entropy(freq), hb = huffman_bits(freq);
        std::printf("%26s %10d %10.4f %10.4f %12d %9.1f%%\n",
                    c.name, distinct, H, hb, 8,
                    100.0 * (1.0 - hb / 8.0));
    }
    std::printf("\n   Huffman is always >= entropy and < entropy + 1:\n");
    for (const Corpus& c : corpora) {
        std::vector<long long> freq(256, 0);
        for (unsigned char ch : c.text) ++freq[ch];
        double H = entropy(freq), hb = huffman_bits(freq);
        std::printf("%26s   H = %7.4f <= %7.4f < %7.4f   %s\n",
                    c.name, H, hb, H + 1.0,
                    (hb >= H - 1e-9 && hb < H + 1.0) ? "ok" : "VIOLATED");
    }
    return 0;
}
