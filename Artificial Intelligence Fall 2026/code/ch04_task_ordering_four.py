# Chapter 4 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 4 lab -- uninformed search on the project graph.

Breadth-first, depth-first, iterative deepening and uniform-cost over
the same task-ordering problem. The search body never changes: the
frontier decides the strategy, exactly as Chapter 3 claimed.
"""

import heapq
import itertools
from collections import deque, namedtuple

# duration in days, and the tasks that must finish first
TASKS = {
    "spec":    (3, set()),
    "schema":  (2, {"spec"}),
    "api":     (5, {"schema"}),
    "ui":      (4, {"spec"}),
    "auth":    (3, {"schema"}),
    "tests":   (4, {"api", "auth"}),
    "docs":    (2, {"api"}),
    "uat":     (3, {"tests", "ui"}),
    "release": (1, {"uat", "docs"}),
}
ALL = frozenset(TASKS)
Node = namedtuple("Node", "state parent action cost depth")


def actions(state):
    for t, (_, deps) in TASKS.items():
        if t not in state and deps <= state:
            yield t


def result(state, action):
    return state | {action}


def cost_of(action):
    return TASKS[action][0]


def path(node):
    out = []
    while node.parent is not None:
        out.append(node.action)
        node = node.parent
    return list(reversed(out))


def graph_search(pop_which, explored=True):
    """One body. pop_which selects the strategy.

    'fifo'  -> breadth-first      'lifo' -> depth-first
    'cheap' -> uniform-cost
    """
    start = Node(frozenset(), None, None, 0, 0)
    tie = itertools.count()
    if pop_which == "cheap":
        frontier = [(0, next(tie), start)]
    else:
        frontier = deque([start])
    seen = {start.state} if explored else None
    expanded = 0
    peak = 1
    while frontier:
        if pop_which == "cheap":
            node = heapq.heappop(frontier)[2]
        elif pop_which == "fifo":
            node = frontier.popleft()
        else:
            node = frontier.pop()

        if node.state == ALL:           # goal test on POPPING
            return path(node), node.cost, expanded, peak
        expanded += 1
        for a in actions(node.state):
            nxt = result(node.state, a)
            if explored:
                if nxt in seen:
                    continue
                seen.add(nxt)
            child = Node(nxt, node, a, node.cost + cost_of(a),
                         node.depth + 1)
            if pop_which == "cheap":
                heapq.heappush(frontier, (child.cost, next(tie), child))
            else:
                frontier.append(child)
        peak = max(peak, len(frontier))
    return None, None, expanded, peak


def depth_limited(node, limit, seen, stats):
    """Recursive depth-limited search. Returns a node or None."""
    if node.state == ALL:
        return node
    if node.depth >= limit:
        return None
    stats["expanded"] += 1
    for a in actions(node.state):
        nxt = result(node.state, a)
        if nxt in seen:
            continue
        seen.add(nxt)
        child = Node(nxt, node, a, node.cost + cost_of(a),
                     node.depth + 1)
        found = depth_limited(child, limit, seen, stats)
        if found is not None:
            return found
        seen.discard(nxt)          # allow other branches to reuse it
    return None


def iterative_deepening():
    stats = {"expanded": 0}
    for limit in range(1, len(TASKS) + 2):
        start = Node(frozenset(), None, None, 0, 0)
        found = depth_limited(start, limit, {start.state}, stats)
        if found is not None:
            return path(found), found.cost, stats["expanded"], limit
    return None, None, stats["expanded"], None


def legal(plan):
    done = set()
    for t in plan:
        if not TASKS[t][1] <= done:
            return False
        done.add(t)
    return done == set(TASKS)


if __name__ == "__main__":
    print("tasks: %d   total duration: %d days"
          % (len(TASKS), sum(d for d, _ in TASKS.values())))
    print()
    print("strategy             len  days  expanded  peak")
    print("-" * 49)
    rows = {}
    for name, key in (("breadth-first ", "fifo"),
                      ("depth-first   ", "lifo"),
                      ("uniform-cost  ", "cheap")):
        plan, cost, exp, peak = graph_search(key)
        rows[name.strip()] = (plan, cost, exp, peak)
        assert legal(plan), name
        print("%s %6d %5d %9d %5d"
              % (name, len(plan), cost, exp, peak))
    plan, cost, exp, depth = iterative_deepening()
    assert legal(plan)
    print("%s %6d %5d %9d %5s"
          % ("iter. deepening", len(plan), cost, exp, "-"))

    # every task must be done, so every complete plan costs the same
    costs = {c for (_, c, _, _) in rows.values()}
    assert costs == {sum(d for d, _ in TASKS.values())}
    print()
    print("Every plan costs %d days: all nine tasks are mandatory, so"
          % costs.pop())
    print("cost cannot discriminate. Chapter 5 gives it something to")
    print("discriminate on.")

    # what the explored set is worth
    _, _, with_set, _ = graph_search("fifo", explored=True)
    _, _, without, _ = graph_search("fifo", explored=False)
    print()
    print("breadth-first expansions with explored set:    %d" % with_set)
    print("breadth-first expansions without explored set: %d" % without)
    print("factor: %.1fx more work" % (without / with_set))
    assert without > with_set
