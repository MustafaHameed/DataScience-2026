# Chapter 6 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 6 lab -- local search on the staffing problem.

Nine developers, four workstreams. Reproduces Table 6.1. Every method
gets exactly the same number of objective-function evaluations, which
is the only comparison that means anything.
"""

import itertools
import math
import random

DEVS = ("Ayesha", "Bilal", "Chen", "Dina", "Emre",
        "Farah", "Gul", "Hassan", "Iqra")
STREAMS = ("api", "ui", "data", "infra")

# skill of each developer on each stream, 1 (weak) to 5 (strong)
SKILL = {
    "Ayesha": (5, 2, 3, 1),
    "Bilal":  (4, 3, 2, 2),
    "Chen":   (2, 5, 1, 2),
    "Dina":   (1, 4, 5, 2),
    "Emre":   (3, 1, 4, 5),
    "Farah":  (2, 2, 5, 3),
    "Gul":    (1, 5, 2, 3),
    "Hassan": (4, 1, 2, 5),
    "Iqra":   (3, 3, 3, 3),
}
TARGET = (3, 2, 2, 2)      # headcount wanted per stream
PENALTY = 4                # cost per developer away from target
NS = len(STREAMS)
BUDGET = 4000              # objective evaluations allowed per run


def raw_score(state):
    fit = sum(SKILL[DEVS[i]][s] for i, s in enumerate(state))
    counts = [0] * NS
    for s in state:
        counts[s] += 1
    off = sum(abs(counts[j] - TARGET[j]) for j in range(NS))
    return fit - PENALTY * off


class Budget:
    """Counts objective evaluations, so no method can cheat."""

    def __init__(self, cap):
        self.cap, self.n = cap, 0

    def score(self, state):
        self.n += 1
        return raw_score(state)

    def left(self):
        return self.n < self.cap


def neighbours(state):
    for i in range(len(state)):
        for s in range(NS):
            if s != state[i]:
                yield state[:i] + (s,) + state[i + 1:]


def random_state(rng):
    return tuple(rng.randrange(NS) for _ in DEVS)


def one_neighbour(state, rng):
    i = rng.randrange(len(state))
    s = rng.randrange(NS)
    while s == state[i]:
        s = rng.randrange(NS)
    return state[:i] + (s,) + state[i + 1:]


# --- 1. steepest-ascent hill climbing ------------------------------
def hill_climb(b, rng, start=None):
    st = start if start is not None else random_state(rng)
    while b.left():
        cur = b.score(st)
        best, bestv = None, cur
        for n in neighbours(st):
            if not b.left():
                break
            v = b.score(n)
            if v > bestv:
                best, bestv = n, v
        if best is None:
            return st, cur          # no better neighbour: stuck
        st = best
    return st, raw_score(st)


# --- 2. random-restart hill climbing -------------------------------
def random_restart(b, rng):
    best, bestv = None, None
    while b.left():
        st, v = hill_climb(b, rng)
        if bestv is None or v > bestv:
            best, bestv = st, v
    return best, bestv


# --- 3. simulated annealing ----------------------------------------
def anneal(b, rng, t0=8.0):
    st = random_state(rng)
    cur = b.score(st)
    best, bestv = st, cur
    while b.left():
        T = t0 * (1 - b.n / b.cap) + 1e-9      # linear cooling
        cand = one_neighbour(st, rng)
        v = b.score(cand)
        d = v - cur
        if d > 0 or rng.random() < math.exp(d / T):
            st, cur = cand, v
            if cur > bestv:
                best, bestv = st, cur
    return best, bestv


# --- 4. genetic algorithm ------------------------------------------
def genetic(b, rng, pop=30, mut=0.12):
    """Standard fitness accounting: every individual is evaluated once
    per generation and the value is then cached for every selection
    decision in that generation. So one generation costs exactly `pop`
    evaluations -- no selection is ever free.

    This matters more than it looks. Charging each tournament
    comparison separately instead costs about six times as much per
    generation, and the method drops from first place to last. See the
    alertbox in Section 6.6.
    """
    P = [random_state(rng) for _ in range(pop)]
    bestx, bestv = None, None
    while b.n + pop <= b.cap:
        fit = [b.score(x) for x in P]          # pop evaluations, once
        order = sorted(range(pop), key=lambda i: -fit[i])
        if bestv is None or fit[order[0]] > bestv:
            bestv, bestx = fit[order[0]], P[order[0]]
        elite = [P[i] for i in order[: pop // 5]]
        nxt = list(elite)

        def tournament():
            """Best of three, judged on the cached fitness."""
            picked = rng.sample(range(pop), 3)
            return P[max(picked, key=lambda i: fit[i])]

        while len(nxt) < pop:
            a, c = tournament(), tournament()
            cut = rng.randrange(1, len(DEVS))
            child = a[:cut] + c[cut:]
            child = tuple(rng.randrange(NS) if rng.random() < mut else g
                          for g in child)
            nxt.append(child)
        P = nxt
    return bestx, bestv


def exhaustive():
    """4^9 = 262144 states: small enough to know the truth."""
    best, arg = None, None
    for st in itertools.product(range(NS), repeat=len(DEVS)):
        v = raw_score(st)
        if best is None or v > best:
            best, arg = v, st
    return best, arg


def show(state):
    by = {j: [] for j in range(NS)}
    for i, s in enumerate(state):
        by[s].append(DEVS[i])
    return "  ".join("%s: %s" % (STREAMS[j], ", ".join(by[j]))
                     for j in range(NS))


if __name__ == "__main__":
    opt, arg = exhaustive()
    print("states: %d      exhaustive optimum: %d" % (NS ** len(DEVS), opt))
    print(show(arg))
    print()
    print("budget: %d evaluations per run, 200 runs, seeds 0-199" % BUDGET)
    print()
    print("method                        optimum    mean   worst")
    print("-" * 56)

    methods = (("hill climbing, one start", hill_climb),
               ("random restart          ", random_restart),
               ("simulated annealing     ", anneal),
               ("genetic algorithm       ", genetic))
    table = {}
    for name, fn in methods:
        hits, total, worst = 0, 0, None
        for seed in range(200):
            rng = random.Random(seed)
            _, v = fn(Budget(BUDGET), rng)
            total += v
            worst = v if worst is None else min(worst, v)
            if v == opt:
                hits += 1
        table[name.strip()] = (hits, total / 200, worst)
        print("%s  %4d/200 %7.2f %6d"
              % (name, hits, total / 200, worst))

    hc = table["hill climbing, one start"]
    rr = table["random restart"]
    sa = table["simulated annealing"]
    ga = table["genetic algorithm"]
    escapes = (rr, sa, ga)

    # 1. plain hill climbing is hopeless here: the landscape is a ridge
    assert hc[0] <= 5
    # 2. every escape strategy is far better than no escape
    for m in escapes:
        assert m[1] > hc[1] + 5
    # 3. the gap to no-escape dwarfs the gap between the escapes
    spread = max(m[1] for m in escapes) - min(m[1] for m in escapes)
    gain = min(m[1] for m in escapes) - hc[1]
    assert spread < gain
    # 4. even the winner is unreliable on any single run
    assert max(m[0] for m in escapes) < 200

    print()
    print("escaping local maxima is worth %.1f points of mean score;"
          % gain)
    print("choosing the best escape is worth only %.1f more." % spread)
    print()
    print("random restart has no parameters at all and reaches %.2f."
          % rr[1])
    print("measure it before reaching for anything more interesting.")
    print()
    best = max(escapes, key=lambda m: m[1])
    print("the best method still missed the optimum in %d of 200 runs."
          % (200 - best[0]))
