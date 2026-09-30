# Chapter 9 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 9 lab -- planning a release from STRIPS operators.

Components go through code, review, test and deploy, sharing a single
staging environment. Testing seizes staging and deletes it -- that one
delete effect is what makes this planning rather than search.
"""

import heapq
import itertools
from collections import deque

COMPONENTS = ("A", "B")


def operators(components):
    """(name, preconditions, add, delete), all frozensets."""
    ops = []
    for c in components:
        ops.append(("code %s" % c, frozenset(),
                    frozenset({"coded %s" % c}), frozenset()))
        ops.append(("review %s" % c, frozenset({"coded %s" % c}),
                    frozenset({"reviewed %s" % c}), frozenset()))
        ops.append(("test %s" % c,
                    frozenset({"reviewed %s" % c, "staging free"}),
                    frozenset({"tested %s" % c, "staging busy"}),
                    frozenset({"staging free"})))
        ops.append(("deploy %s" % c, frozenset({"tested %s" % c}),
                    frozenset({"deployed %s" % c}), frozenset()))
    ops.append(("release staging", frozenset({"staging busy"}),
                frozenset({"staging free"}),
                frozenset({"staging busy"})))
    return ops


def applicable(state, ops):
    for op in ops:
        if op[1] <= state:
            yield op


def apply_op(state, op):
    _, _, add, dele = op
    return (state - dele) | add


def reachable(init, ops):
    seen, q = {init}, deque([init])
    while q:
        s = q.popleft()
        for op in applicable(s, ops):
            n = apply_op(s, op)
            if n not in seen:
                seen.add(n)
                q.append(n)
    return seen


# --- 1. progression search: Chapter 4's frontier, unchanged ---------
def progression_bfs(init, goal, ops):
    seen = {init}
    q = deque([(init, ())])
    expanded = 0
    while q:
        state, plan = q.popleft()
        if goal <= state:
            return list(plan), expanded
        expanded += 1
        for op in applicable(state, ops):
            nxt = apply_op(state, op)
            if nxt not in seen:
                seen.add(nxt)
                q.append((nxt, plan + (op[0],)))
    return None, expanded


def h_remaining(state, components):
    """Actions each component still needs. They are specific to that
    component, so the sum is a lower bound: admissible."""
    total = 0
    for c in components:
        if "deployed %s" % c in state:
            total += 0
        elif "tested %s" % c in state:
            total += 1
        elif "reviewed %s" % c in state:
            total += 2
        elif "coded %s" % c in state:
            total += 3
        else:
            total += 4
    return total


def progression_astar(init, goal, ops, components):
    tie = itertools.count()
    frontier = [(h_remaining(init, components), next(tie), 0, init, ())]
    best = {init: 0}
    expanded = 0
    while frontier:
        _, _, g, state, plan = heapq.heappop(frontier)
        if goal <= state:
            return list(plan), expanded
        if g > best.get(state, 1 << 30):
            continue
        expanded += 1
        for op in applicable(state, ops):
            nxt, ng = apply_op(state, op), g + 1
            if ng < best.get(nxt, 1 << 30):
                best[nxt] = ng
                heapq.heappush(frontier,
                               (ng + h_remaining(nxt, components),
                                next(tie), ng, nxt, plan + (op[0],)))
    return None, expanded


# --- 2. means-ends analysis: the General Problem Solver -------------
class Work:
    def __init__(self):
        self.ops = 0


def means_ends(state, goals, ops, work, depth=0):
    """Achieve each goal in turn, recursing on preconditions.

    Returns (state, plan) or None. Note it never enumerates states.
    """
    if depth > 40:
        return None
    plan = []
    for g in goals:
        if g in state:
            continue
        achieved = False
        for op in ops:
            work.ops += 1
            name, pre, add, dele = op
            if g not in add:
                continue
            got = means_ends(state, sorted(pre), ops, work, depth + 1)
            if got is None:
                continue
            state, sub = got
            if not pre <= state:
                continue
            state = apply_op(state, op)
            plan += sub + [name]
            achieved = True
            break
        if not achieved:
            return None            # clobbered, or simply impossible
    return state, plan


def legal(plan, init, goal, ops):
    by_name = {o[0]: o for o in ops}
    s = init
    for step in plan:
        op = by_name.get(step)
        if op is None or not op[1] <= s:
            return False
        s = apply_op(s, op)
    return goal <= s


if __name__ == "__main__":
    print("components  states  plan  BFS exp  A* exp  GPS ops")
    print("-" * 54)
    rows = {}
    for n in (2, 3, 4):
        comps = ("A", "B", "C", "D")[:n]
        ops = operators(comps)
        init = frozenset({"staging free"})
        goal = frozenset({"deployed %s" % c for c in comps})

        states = len(reachable(init, ops))
        bplan, bexp = progression_bfs(init, goal, ops)
        aplan, aexp = progression_astar(init, goal, ops, comps)
        work = Work()
        got = means_ends(init, sorted(goal), ops, work)
        gplan = got[1]

        for p in (bplan, aplan, gplan):
            assert legal(p, init, goal, ops)
        # all three find an OPTIMAL plan on this domain
        assert len(aplan) == len(bplan) == len(gplan)

        rows[n] = (states, len(bplan), bexp, aexp, work.ops)
        print("%10d %7d %5d %8d %7d %8d"
              % (n, states, len(bplan), bexp, aexp, work.ops))

    # the heuristic is strong and still buys almost nothing
    st, ln, bexp, aexp, gops = rows[4]
    comps = ("A", "B", "C", "D")
    print()
    print("h at the initial state, 4 components: %d against a true"
          % h_remaining(frozenset({"staging free"}), comps))
    print("optimum of %d -- a strong admissible heuristic." % ln)
    print("A* expanded %d where BFS expanded %d: %.1f%% saved."
          % (aexp, bexp, 100 * (bexp - aexp) / bexp))
    print("Almost every reachable state lies on some optimal plan,")
    print("so there is nothing for f < C* to rule out.")
    assert (bexp - aexp) / bexp < 0.05

    # means-ends scales quite differently
    print()
    print("means-ends considered %d operators where BFS expanded %d"
          % (gops, bexp))
    assert gops * 5 < bexp

    # the nine-step plan, and why release staging is in it
    ops = operators(("A", "B"))
    init = frozenset({"staging free"})
    goal = frozenset({"deployed A", "deployed B"})
    plan, _ = progression_bfs(init, goal, ops)
    print()
    print("an optimal two-component plan:")
    for i, s in enumerate(plan, 1):
        print("   %d. %s" % (i, s))
    assert "release staging" in plan
    print("'release staging' appears in no goal and is indispensable.")

    # --- 3. goal clobbering, in two operators ----------------------
    print()
    print("goal clobbering:")
    clob = [
        ("lock staging", frozenset({"staging free"}),
         frozenset({"locked"}), frozenset({"staging free"})),
        ("run smoke test", frozenset({"staging free"}),
         frozenset({"smoke ok"}), frozenset()),
    ]
    c_init = frozenset({"staging free"})
    c_goal = frozenset({"locked", "smoke ok"})
    for order in (["locked", "smoke ok"], ["smoke ok", "locked"]):
        got = means_ends(c_init, order, clob, Work())
        if got is None:
            print("   order %-22s -> FAILS" % (" then ".join(order)))
        else:
            print("   order %-22s -> %s"
                  % (" then ".join(order), " then ".join(got[1])))
    # one order fails, the other succeeds, on the same solvable problem
    assert means_ends(c_init, ["locked", "smoke ok"], clob, Work()) is None
    assert means_ends(c_init, ["smoke ok", "locked"], clob,
                      Work()) is not None
    plan, _ = progression_bfs(c_init, c_goal, clob)
    print("   progression search finds it either way:", plan)
