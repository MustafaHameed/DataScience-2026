// ch21_strings.cpp -- naive, Rabin-Karp and KMP, on text and on traps.
//
// All three find every occurrence of a pattern of length m in a text of
// length n. They differ in what they do after a mismatch:
//
//   naive        shift by 1 and start again      O(nm) worst, O(n) typical
//   Rabin-Karp   compare hashes, verify on hit   O(n+m) expected, O(nm) worst
//   KMP          shift by as much as the prefix  O(n+m) ALWAYS
//                function allows, never re-read
//
// The interesting measurement is not on English text, where all three are
// fast. It is on the inputs that separate them: a highly repetitive text,
// which destroys the naive algorithm, and a crafted key set, which
// destroys Rabin-Karp.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch21_strings.cpp
#include <algorithm>
#include <chrono>
#include <cstdio>
#include <random>
#include <string>
#include <vector>

static long long cmps = 0;

// --- 1. Naive matching: try every alignment --------------------------
// After a mismatch at offset s it restarts at s+1, re-reading characters
// it has already looked at. That re-reading is the O(nm).
static std::vector<size_t> naive(const std::string& t, const std::string& p) {
    std::vector<size_t> hits;
    size_t n = t.size(), m = p.size();
    if (m > n) return hits;
    for (size_t s = 0; s + m <= n; ++s) {
        size_t i = 0;
        while (i < m) { ++cmps; if (t[s + i] != p[i]) break; ++i; }
        if (i == m) hits.push_back(s);
    }
    return hits;
}

// --- 2. Rabin-Karp: compare numbers, not strings ---------------------
// A rolling hash of each window. Equal hashes mean a POSSIBLE match that
// must be verified; unequal hashes mean a certain mismatch, decided in
// one comparison. Expected O(n+m); the worst case is every window
// colliding, which is Chapter 13's problem in a new costume.
static std::vector<size_t> rabin_karp(const std::string& t,
                                      const std::string& p,
                                      unsigned long long base,
                                      unsigned long long mod,
                                      long long* spurious) {
    std::vector<size_t> hits;
    size_t n = t.size(), m = p.size();
    if (m > n) return hits;
    unsigned long long hp = 0, ht = 0, high = 1;
    for (size_t i = 0; i + 1 < m; ++i) high = (high * base) % mod;
    for (size_t i = 0; i < m; ++i) {
        hp = (hp * base + (unsigned char)p[i]) % mod;
        ht = (ht * base + (unsigned char)t[i]) % mod;
    }
    for (size_t s = 0;; ++s) {
        if (hp == ht) {                       // candidate -- verify it
            size_t i = 0;
            while (i < m) { ++cmps; if (t[s + i] != p[i]) break; ++i; }
            if (i == m) hits.push_back(s);
            else if (spurious) ++*spurious;   // hash agreed, text did not
        }
        if (s + m >= n) break;
        ht = (ht + mod - (high * (unsigned char)t[s]) % mod) % mod;
        ht = (ht * base + (unsigned char)t[s + m]) % mod;
    }
    return hits;
}

// --- 3. KMP: never re-read a character of the text -------------------
// The prefix function pi[i] is the length of the longest proper prefix
// of p[0..i] that is also a suffix of it. On a mismatch, that is exactly
// how far the pattern can be shifted without missing an occurrence --
// so the text pointer only ever moves forward, giving O(n+m) with no
// hypothesis about the input at all.
static std::vector<size_t> prefix_function(const std::string& p) {
    std::vector<size_t> pi(p.size(), 0);
    for (size_t i = 1; i < p.size(); ++i) {
        size_t k = pi[i - 1];
        while (k > 0 && p[i] != p[k]) k = pi[k - 1];
        if (p[i] == p[k]) ++k;
        pi[i] = k;
    }
    return pi;
}
static std::vector<size_t> kmp(const std::string& t, const std::string& p) {
    std::vector<size_t> hits;
    if (p.empty() || p.size() > t.size()) return hits;
    std::vector<size_t> pi = prefix_function(p);
    size_t k = 0;
    for (size_t i = 0; i < t.size(); ++i) {
        while (k > 0 && t[i] != p[k]) { ++cmps; k = pi[k - 1]; }
        ++cmps;
        if (t[i] == p[k]) ++k;
        if (k == p.size()) { hits.push_back(i + 1 - p.size()); k = pi[k - 1]; }
    }
    return hits;
}

template <class F>
static double time_ms(F&& f, int reps) {
    std::vector<double> t;
    for (int r = 0; r < reps; ++r) {
        auto t0 = std::chrono::steady_clock::now();
        auto v = f();
        (void)v.size();
        auto t1 = std::chrono::steady_clock::now();
        t.push_back(
            std::chrono::duration<double, std::milli>(t1 - t0).count());
    }
    std::sort(t.begin(), t.end());
    return t[t.size() / 2];
}

int main() {
    std::setvbuf(stdout, nullptr, _IONBF, 0);
    std::mt19937 rng(65537);

    // --- 4. All three agree, on every input used below ---------------
    std::printf("A. do the three agree?\n\n");
    {
        int bad = 0;
        std::uniform_int_distribution<int> da(0, 3);
        for (int t = 0; t < 2000; ++t) {
            std::string text, pat;
            for (int i = 0; i < 200; ++i) text += (char)('a' + da(rng));
            int m = 1 + da(rng);
            for (int i = 0; i < m; ++i) pat += (char)('a' + da(rng));
            long long sp = 0;
            auto h1 = naive(text, pat);
            auto h2 = rabin_karp(text, pat, 256, 1000000007ULL, &sp);
            auto h3 = kmp(text, pat);
            if (h1 != h2 || h2 != h3) ++bad;
        }
        std::printf("   2000 random texts, disagreements: %d\n", bad);
    }

    // --- 5. Ordinary text, where nothing much happens ----------------
    // English-like prose and a word-length pattern. All three are fast,
    // and the comparison counts are all close to n. This is the case
    // that makes people think the choice does not matter.
    std::printf("\nB. ordinary text (4-letter alphabet), m = 8\n\n");
    std::printf("%10s %14s %14s %14s %12s\n",
                "n", "naive cmps/n", "RK cmps/n", "KMP cmps/n", "naive (ms)");
    {
        std::uniform_int_distribution<int> da(0, 3);
        for (size_t n = 1000000; n <= 8000000; n *= 2) {
            std::string text(n, 'a');
            for (auto& c : text) c = (char)('a' + da(rng));
            std::string pat = text.substr(n / 2, 8);
            long long sp = 0;
            cmps = 0; naive(text, pat);          long long c1 = cmps;
            cmps = 0; rabin_karp(text, pat, 256, 1000000007ULL, &sp);
            long long c2 = cmps;
            cmps = 0; kmp(text, pat);            long long c3 = cmps;
            double t1 = time_ms([&] { return naive(text, pat); }, 3);
            std::printf("%10zu %14.3f %14.3f %14.3f %12.1f\n",
                        n, (double)c1 / n, (double)c2 / n,
                        (double)c3 / n, t1);
        }
    }

    // --- 6. The text that destroys the naive algorithm ---------------
    // text = aaaa...a, pattern = aaa...ab. Every alignment matches m-1
    // characters and then fails, so the naive version does n*m
    // comparisons -- the worst case, reached exactly.
    std::printf("\nC. repetitive text: t = a^n, p = a^(m-1)b\n\n");
    std::printf("%8s %6s %14s %12s %12s %12s %10s\n",
                "n", "m", "naive cmps", "/(n*m)", "KMP cmps",
                "naive (ms)", "KMP (ms)");
    for (size_t m : {10u, 100u, 1000u}) {
        size_t n = 200000;
        std::string text(n, 'a');
        std::string pat(m - 1, 'a');
        pat += 'b';
        cmps = 0; naive(text, pat);  long long c1 = cmps;
        cmps = 0; kmp(text, pat);    long long c3 = cmps;
        double t1 = time_ms([&] { return naive(text, pat); }, 3);
        double t3 = time_ms([&] { return kmp(text, pat); }, 3);
        std::printf("%8zu %6zu %14lld %12.4f %12lld %12.2f %10.2f\n",
                    n, m, c1, (double)c1 / ((double)n * m), c3, t1, t3);
    }

    // --- 7. The input that destroys Rabin-Karp -----------------------
    // Rabin-Karp's expected O(n+m) assumes hash collisions are rare. Use
    // a tiny modulus -- or let an adversary pick the text for a known
    // modulus -- and every window becomes a candidate that must be
    // verified character by character. This is Chapter 13's adversarial
    // hashing, in a new costume.
    std::printf("\nD. Rabin-Karp with a weak modulus\n\n");
    std::printf("%14s %16s %16s %14s\n",
                "modulus", "spurious hits", "comparisons", "time (ms)");
    {
        size_t n = 2000000;
        std::uniform_int_distribution<int> da(0, 3);
        std::string text(n, 'a');
        for (auto& c : text) c = (char)('a' + da(rng));
        std::string pat = text.substr(n / 2, 8);
        for (unsigned long long mod : {13ULL, 1009ULL, 1000003ULL,
                                       1000000007ULL}) {
            long long sp = 0;
            cmps = 0;
            auto h = rabin_karp(text, pat, 256, mod, &sp);
            long long c = cmps;
            double t = time_ms([&] {
                long long s2 = 0;
                return rabin_karp(text, pat, 256, mod, &s2);
            }, 3);
            std::printf("%14llu %16lld %16lld %14.1f\n", mod, sp, c, t);
        }
        std::printf("\n   The verification count explodes 1000-fold and the\n"
                    "   clock barely moves: on RANDOM text a spurious\n"
                    "   candidate is rejected after about 1.3 characters.\n"
                    "   The bound needs BOTH many collisions and expensive\n"
                    "   verification, so here is the input that has both.\n");
    }

    // --- 8. Rabin-Karp's actual worst case ---------------------------
    // The pattern a^(m-1)b does NOT do it: its hash differs from every
    // window of a^n, so Rabin-Karp rejects all of them in one comparison
    // each and never verifies anything. (The first version of this
    // program measured exactly zero comparisons, which was the clue.)
    //
    // The bound needs every window to be a CANDIDATE. Take p = a^m: now
    // every window hashes equal, every verification runs the full m
    // characters, and the total is n*m. The matches are real, so this is
    // not even adversarial -- it is what happens when a pattern occurs
    // everywhere, and KMP still finishes in O(n+m).
    std::printf("\nE. Rabin-Karp on repetitive text, t = a^n, p = a^m\n\n");
    std::printf("%8s %6s %16s %12s %14s %12s\n",
                "n", "m", "RK comparisons", "/(n*m)", "RK (ms)", "KMP (ms)");
    for (size_t m : {10u, 100u, 1000u}) {
        size_t n = 200000;
        std::string text(n, 'a');
        std::string pat(m, 'a');
        long long sp = 0;
        cmps = 0;
        rabin_karp(text, pat, 256, 1000000007ULL, &sp);
        long long c = cmps;
        double t1 = time_ms([&] {
            long long s2 = 0;
            return rabin_karp(text, pat, 256, 1000000007ULL, &s2);
        }, 3);
        double t2 = time_ms([&] { return kmp(text, pat); }, 3);
        std::printf("%8zu %6zu %16lld %12.4f %14.2f %12.2f\n",
                    n, m, c, (double)c / ((double)n * m), t1, t2);
    }
    std::printf("\n   KMP has no bad case at all: its text pointer never\n"
                "   moves backwards, so O(n+m) holds for every input.\n");
    return 0;
}
