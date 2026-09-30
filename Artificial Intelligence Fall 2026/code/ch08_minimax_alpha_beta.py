# Chapter 8 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 8 lab -- minimax and alpha-beta on a release negotiation.

Five plies, three options each, 243 leaves. All four configurations
return the same minimax value; they differ only in how many leaves
they have to look at. The perfect-ordering row hits the theoretical
bound b^ceil(d/2) + b^floor(d/2) - 1 exactly.
"""

import random

DEPTH = 5          # plies of negotiation
BRANCH = 3         # hold, concede, swap
OPTIONS = ("hold", "concede", "swap")


def build(seed=7):
    """Leaf utilities to us, in days of delay avoided."""
    rng = random.Random(seed)
    leaves = {}

    def rec(path):
        if len(path) == DEPTH:
            leaves[path] = rng.randrange(-20, 21)
            return
        for m in range(BRANCH):
            rec(path + (m,))

    rec(())
    return leaves


LEAVES = build()


class Counter:
    def __init__(self):
        self.evals = 0          # leaf evaluations: the expensive op
        self.nodes = 0


def minimax(path, maximising, c):
    c.nodes += 1
    if len(path) == DEPTH:
        c.evals += 1
        return LEAVES[path]
    vals = [minimax(path + (m,), not maximising, c)
            for m in range(BRANCH)]
    return max(vals) if maximising else min(vals)


def true_value(path, maximising):
    """An oracle, used only to build the best/worst orderings."""
    return minimax(path, maximising, Counter())


def ordered(path, maximising, how):
    ms = list(range(BRANCH))
    if how == "natural":
        return ms
    # 'best' puts the strongest move first, 'worst' puts it last
    rev = maximising if how == "best" else not maximising
    return sorted(ms, key=lambda m: true_value(path + (m,), not maximising),
                  reverse=rev)


def alphabeta(path, maximising, alpha, beta, c, how="natural"):
    c.nodes += 1
    if len(path) == DEPTH:
        c.evals += 1
        return LEAVES[path]
    if maximising:
        v = -10 ** 9
        for m in ordered(path, maximising, how):
            v = max(v, alphabeta(path + (m,), False, alpha, beta, c, how))
            alpha = max(alpha, v)
            if alpha >= beta:
                break                      # beta cutoff
        return v
    v = 10 ** 9
    for m in ordered(path, maximising, how):
        v = min(v, alphabeta(path + (m,), True, alpha, beta, c, how))
        beta = min(beta, v)
        if beta <= alpha:
            break                          # alpha cutoff
    return v


def shallow(path, maximising, limit, c):
    """A depth-limited probe, used to ORDER a deeper search.

    Every leaf it touches is charged, because in a real engine the
    evaluation function is what costs.
    """
    if len(path) == DEPTH:
        c.evals += 1
        return LEAVES[path]
    if limit == 0:
        p = path
        while len(p) < DEPTH:
            p = p + (0,)
        c.evals += 1
        return LEAVES[p]
    vals = [shallow(path + (m,), not maximising, limit - 1, c)
            for m in range(BRANCH)]
    return max(vals) if maximising else min(vals)


def ab_probe_ordered(path, maximising, alpha, beta, c, probe=1):
    """Alpha-beta ordering its moves by a shallow probe."""
    c.nodes += 1
    if len(path) == DEPTH:
        c.evals += 1
        return LEAVES[path]
    ms = sorted(range(BRANCH),
                key=lambda m: shallow(path + (m,), not maximising,
                                      probe, c),
                reverse=maximising)
    if maximising:
        v = -10 ** 9
        for m in ms:
            v = max(v, ab_probe_ordered(path + (m,), False, alpha, beta,
                                        c, probe))
            alpha = max(alpha, v)
            if alpha >= beta:
                break
        return v
    v = 10 ** 9
    for m in ms:
        v = min(v, ab_probe_ordered(path + (m,), True, alpha, beta,
                                    c, probe))
        beta = min(beta, v)
        if beta <= alpha:
            break
    return v


if __name__ == "__main__":
    print("depth %d, branching %d, leaves %d"
          % (DEPTH, BRANCH, BRANCH ** DEPTH))

    c0 = Counter()
    value = minimax((), True, c0)
    print("minimax value: %d" % value)
    print()
    print("algorithm                      leaf evals   nodes  value")
    print("-" * 60)
    print("%-30s %10d %7d %6d"
          % ("minimax, no pruning", c0.evals, c0.nodes, value))

    results = {"minimax": c0.evals}
    for how, label in (("worst", "alpha-beta, worst ordering  "),
                       ("natural", "alpha-beta, natural ordering"),
                       ("best", "alpha-beta, perfect ordering")):
        c = Counter()
        v = alphabeta((), True, -10 ** 9, 10 ** 9, c, how)
        assert v == value, how          # pruning is EXACT
        results[how] = c.evals
        print("%-30s %10d %7d %6d" % (label, c.evals, c.nodes, v))

    # the theoretical bound, reached exactly by perfect ordering
    bound = BRANCH ** ((DEPTH + 1) // 2) + BRANCH ** (DEPTH // 2) - 1
    print()
    print("b^ceil(d/2) + b^floor(d/2) - 1 = %d^%d + %d^%d - 1 = %d"
          % (BRANCH, (DEPTH + 1) // 2, BRANCH, DEPTH // 2, bound))
    assert results["best"] == bound
    print("perfect ordering evaluated %d leaves: the bound, exactly."
          % results["best"])

    # pruning never loses, and ordering decides how much it wins
    assert results["worst"] < results["minimax"]
    assert results["best"] < results["natural"] < results["worst"]
    print()
    print("saving against minimax:  worst %.1fx   natural %.1fx   "
          "perfect %.1fx"
          % (results["minimax"] / results["worst"],
             results["minimax"] / results["natural"],
             results["minimax"] / results["best"]))

    # does a shallow probe pay for the ordering it buys? Not here.
    c = Counter()
    v = ab_probe_ordered((), True, -10 ** 9, 10 ** 9, c, probe=1)
    assert v == value
    print()
    print("ordering by a depth-1 probe: %d leaf evals against natural"
          % c.evals)
    print("ordering's %d. The probes cost more than they save at this"
          % results["natural"])
    print("depth. Ordering pays in chess, at twenty ply with a cache.")
    assert c.evals > results["natural"]

    # how much does the instance matter?
    totals = []
    for seed in range(20):
        globals()["LEAVES"] = build(seed)
        c = Counter()
        alphabeta((), True, -10 ** 9, 10 ** 9, c, "natural")
        totals.append(c.evals)
    print()
    print("across 20 different negotiations, natural ordering:")
    print("   leaf evals  min %d   mean %.1f   max %d"
          % (min(totals), sum(totals) / len(totals), max(totals)))
