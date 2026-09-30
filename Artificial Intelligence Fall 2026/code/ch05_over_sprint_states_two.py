# Chapter 5 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 5 lab -- informed search on the sprint network.

Reproduces Table 5.1. The headline result is the last row: greedy
search with the mandatory-work heuristic returns a 29-day plan when a
28-day plan exists, because a test harness reduces its estimate by
nothing at all.
"""

import heapq
import itertools
from collections import deque

# duration, then a conjunction of disjunctions: a task is ready when
# every group has at least one member already finished.
TASKS = {
    "spec":        (3, []),
    "schema":      (2, [{"spec"}]),
    "api":         (5, [{"schema"}]),
    "ui":          (4, [{"spec"}]),
    "auth":        (3, [{"schema"}]),
    "harness":     (3, [{"schema"}]),
    "auto_test":   (2, [{"api"}, {"auth"}, {"harness"}]),
    "manual_test": (6, [{"api"}, {"auth"}]),
    "docs":        (2, [{"api"}]),
    "uat":         (3, [{"ui"}, {"auto_test", "manual_test"}]),
    "release":     (1, [{"uat"}, {"docs"}]),
}
GOAL = "release"

# The eight tasks that appear in every plan. Everything except the two
# testing routes and the harness that one of them needs.
MANDATORY = ("spec", "schema", "api", "auth", "ui", "docs", "uat",
             "release")


def dur(t):
    return TASKS[t][0]


def ready(state):
    """Legal actions: the successor function, as a generator."""
    for t, (_, groups) in TASKS.items():
        if t not in state and all(g & state for g in groups):
            yield t


# --- the two heuristics, both relaxations --------------------------
def h_zero(state):
    return 0


def h_mandatory(state):
    """Relax away the testing requirement: just add up the work that
    cannot be avoided. 23 days at the root."""
    return sum(dur(t) for t in MANDATORY if t not in state)


def h_chain(state):
    """Relax away the one-task-at-a-time constraint: the longest
    remaining dependency chain. 16 days at the root."""
    memo = {}

    def chain(t):
        if t in state:
            return 0
        if t in memo:
            return memo[t]
        need = 0
        for g in TASKS[t][1]:
            need = max(need, min(chain(x) for x in g))
        memo[t] = need + dur(t)
        return memo[t]

    return chain(GOAL)


def h_max(state):
    """Admissible, and dominates both by construction."""
    return max(h_mandatory(state), h_chain(state))


# --- best-first search; the priority decides the algorithm ---------
def best_first(h, use_g=True):
    """use_g=True  -> A*      (f = g + h)
       use_g=False -> greedy  (f = h)

    The heap entry is (priority, counter, g, state, plan). The counter
    sits immediately after the priority so that ties are broken by
    insertion order alone. Putting g there instead would quietly let
    the cost leak into greedy search's ordering -- and greedy search
    would then find the optimal plan and the point of the lab would
    be lost.
    """
    start = frozenset()
    tie = itertools.count()
    frontier = [(h(start), next(tie), 0, start, ())]
    best = {start: 0}
    expanded = 0
    while frontier:
        _, _, g, state, plan = heapq.heappop(frontier)
        if GOAL in state:
            return g, list(plan), expanded
        if use_g and g > best.get(state, 1 << 30):
            continue                    # a cheaper path already went
        expanded += 1
        for t in ready(state):
            nxt, ng = state | {t}, g + dur(t)
            if use_g:
                if ng >= best.get(nxt, 1 << 30):
                    continue
                best[nxt] = ng
                pri = ng + h(nxt)
            else:
                if nxt in best:
                    continue
                best[nxt] = ng
                pri = h(nxt)
            heapq.heappush(frontier,
                           (pri, next(tie), ng, nxt, plan + (t,)))
    return None, None, expanded


def reachable():
    seen, q = {frozenset()}, deque([frozenset()])
    while q:
        s = q.popleft()
        for t in ready(s):
            n = s | {t}
            if n not in seen:
                seen.add(n)
                q.append(n)
    return seen


def cheapest_from(state):
    """True remaining cost h*, by uniform-cost search from state."""
    tie = itertools.count()
    fr = [(0, next(tie), state)]
    best = {state: 0}
    while fr:
        g, _, st = heapq.heappop(fr)
        if GOAL in st:
            return g
        if g > best.get(st, 1 << 30):
            continue
        for t in ready(st):
            n, ng = st | {t}, g + dur(t)
            if ng < best.get(n, 1 << 30):
                best[n] = ng
                heapq.heappush(fr, (ng, next(tie), n))
    return None


def route(plan):
    return "automated" if "auto_test" in plan else "manual"


if __name__ == "__main__":
    print("h at the root:  zero=%d  chain=%d  mandatory=%d"
          % (h_zero(frozenset()), h_chain(frozenset()),
             h_mandatory(frozenset())))
    print()
    print("algorithm                cost  expanded  route")
    print("-" * 52)
    runs = [
        ("uniform-cost (h=0)   ", h_zero, True),
        ("A* critical path     ", h_chain, True),
        ("A* mandatory work    ", h_mandatory, True),
        ("A* max of the two    ", h_max, True),
        ("greedy critical path ", h_chain, False),
        ("greedy mandatory work", h_mandatory, False),
    ]
    results = {}
    for name, h, use_g in runs:
        c, plan, exp = best_first(h, use_g)
        results[name.strip()] = (c, plan, exp)
        print("%s %5d %9d  %s" % (name, c, exp, route(plan)))

    opt = results["uniform-cost (h=0)"][0]
    print()
    print("optimal cost:", opt)

    # 1. A* is optimal with every heuristic here
    astars = ("A* critical path", "A* mandatory work",
              "A* max of the two")
    for n in astars:
        assert results[n][0] == opt, n
    # 2. and expands strictly fewer nodes than uniform-cost search
    ucs_exp = results["uniform-cost (h=0)"][2]
    for n in astars:
        assert results[n][2] < ucs_exp, n

    # 3. NEITHER heuristic dominates the other -- the interesting part
    states = reachable()
    mand_bigger = [s for s in states if h_mandatory(s) > h_chain(s)]
    chain_bigger = [s for s in states if h_chain(s) > h_mandatory(s)]
    print()
    print("reachable states: %d" % len(states))
    print("  mandatory work is the larger estimate on %d"
          % len(mand_bigger))
    print("  critical path  is the larger estimate on %d"
          % len(chain_bigger))
    assert mand_bigger and chain_bigger, "expected them to cross"
    # so max dominates both, by construction
    assert all(h_max(s) >= h_mandatory(s) for s in states)
    assert all(h_max(s) >= h_chain(s) for s in states)
    assert results["A* max of the two"][2] <= min(
        results["A* critical path"][2],
        results["A* mandatory work"][2])
    ex = min(chain_bigger, key=len)
    print("  e.g. %s: mandatory=%d, critical path=%d"
          % (" ".join(sorted(ex)), h_mandatory(ex), h_chain(ex)))

    # 4. greedy with the mandatory heuristic is NOT optimal
    assert results["greedy mandatory work"][0] > opt
    print()
    print("greedy with the mandatory heuristic returns %d, not %d"
          % (results["greedy mandatory work"][0], opt))

    # 5. all three heuristics are admissible on every reachable state
    bad = []
    for s in states:
        if GOAL in s:
            continue
        true = cheapest_from(s)
        if true is None:
            continue
        for nm, h in (("mandatory", h_mandatory), ("chain", h_chain),
                      ("max", h_max)):
            if h(s) > true:
                bad.append((nm, sorted(s), h(s), true))
    print("admissibility violations:", len(bad))
    assert not bad
    print()
    print("A* kept the guarantee and did %d%% less work than"
          % round(100 * (ucs_exp - results["A* mandatory work"][2])
                  / ucs_exp))
    print("uniform-cost search. Greedy did a quarter of the work and")
    print("got the answer wrong.")
