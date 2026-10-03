// ch02_cases.cpp -- the nature of the input, and the size of the input.
//
// Two experiments, one for each half of the chapter.
//
//   A. The SHAPE of the input decides which case you get. The same
//      insertion sort is linear on sorted input and quadratic on reversed
//      input -- not by a constant factor, but by an order of growth.
//
//   B. The SIZE of a number is the count of its bits, not its value.
//      Trial division is linear in sqrt(N), which is exponential in the
//      number of bits. Add two bits and the time doubles.
//
// Build: see Appendix A.  cl /O2 /EHsc /std:c++17 ch02_cases.cpp
#include <algorithm>
#include <chrono>
#include <cstdint>
#include <cstdio>
#include <numeric>
#include <random>
#include <vector>

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

static void insertion_sort(std::vector<int>& a) {
    for (size_t j = 1; j < a.size(); ++j) {
        int key = a[j];
        size_t i = j;
        while (i > 0 && a[i - 1] > key) { a[i] = a[i - 1]; --i; }
        a[i] = key;
    }
}

// --- 1. Three shapes of input, all of the same size ------------------
// Best case: already sorted, so the inner while never executes.
// Worst case: reversed, so the inner while runs to the front every time.
// Average case: a random permutation, which moves each element about
// half way -- half of a quadratic is still a quadratic.
static std::vector<int> sorted_input(size_t n) {
    std::vector<int> a(n);
    std::iota(a.begin(), a.end(), 0);
    return a;
}
static std::vector<int> reversed_input(size_t n) {
    std::vector<int> a = sorted_input(n);
    std::reverse(a.begin(), a.end());
    return a;
}
static std::vector<int> random_input(size_t n, unsigned seed) {
    std::vector<int> a = sorted_input(n);
    std::mt19937 rng(seed);
    std::shuffle(a.begin(), a.end(), rng);
    return a;
}

// --- 2. Trial division, whose cost is governed by sqrt(N) ------------
// The loop runs about sqrt(N)/2 times when N is prime, which is the
// worst case for this algorithm. N prime is therefore what we time.
static bool is_prime_trial(uint64_t n) {
    if (n < 2) return false;
    if (n % 2 == 0) return n == 2;
    for (uint64_t d = 3; d * d <= n; d += 2)
        if (n % d == 0) return false;
    return true;
}

int main() {
    // --- 3. The same algorithm, three orders of growth ---------------
    std::printf("A. insertion sort, by input shape (ms)\n\n");
    std::printf("%8s %10s %7s %10s %7s %10s %7s\n",
                "n", "sorted", "ratio", "random", "ratio",
                "reversed", "ratio");
    double ps = 0, pr = 0, pv = 0;
    for (size_t n = 2000; n <= 32000; n *= 2) {
        auto s = sorted_input(n);
        auto r = random_input(n, 12345);
        auto v = reversed_input(n);
        double ts = time_ms([&] { auto a = s; insertion_sort(a); }, 5);
        double tr = time_ms([&] { auto a = r; insertion_sort(a); }, 5);
        double tv = time_ms([&] { auto a = v; insertion_sort(a); }, 5);
        if (ps == 0)
            std::printf("%8zu %10.3f %7s %10.3f %7s %10.3f %7s\n",
                        n, ts, "-", tr, "-", tv, "-");
        else
            std::printf("%8zu %10.3f %7.2f %10.3f %7.2f %10.3f %7.2f\n",
                        n, ts, ts / ps, tr, tr / pr, tv, tv / pv);
        ps = ts; pr = tr; pv = tv;
    }

    // --- 4. Why the size of a number is its number of bits -----------
    // Each row adds TWO bits to the input -- four extra characters of
    // nothing, if you measure the input by its value. Watch the time.
    std::printf("\nB. trial division on a prime of b bits\n\n");
    std::printf("%4s %22s %12s %7s\n", "bits", "N", "time (ms)", "ratio");
    const uint64_t primes[] = {
        4294967291ULL,      17179869143ULL,     68719476731ULL,
        274877906899ULL,    1099511627689ULL,   4398046511093ULL,
        17592186044399ULL,  70368744177643ULL,  281474976710597ULL};
    int bits = 32;
    double prev = 0;
    for (uint64_t p : primes) {
        bool r = false;
        double t = time_ms([&] { r = is_prime_trial(p); }, 3);
        if (!r) { std::printf("  BUG: %llu is not prime\n", p); return 1; }
        if (prev == 0)
            std::printf("%4d %22llu %12.2f %7s\n", bits, p, t, "-");
        else
            std::printf("%4d %22llu %12.2f %7.2f\n", bits, p, t, t / prev);
        prev = t;
        bits += 2;
    }
    std::printf("\n  Each row is two more bits of input and twice the\n"
                "  work: the algorithm is exponential in the input size.\n");
    return 0;
}
