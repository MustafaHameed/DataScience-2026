# Chapter 3 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 3 lab -- the course toolkit.

The Problem interface, three frontiers, the project graph, and the
measurement that justifies using frozensets for states. Every later
chapter builds on exactly these shapes.
"""

import heapq
import itertools
from collections import deque, namedtuple

# --- the project, as data ------------------------------------------
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


# --- the interface every later chapter subclasses -------------------
class Problem:
    """A search problem: five things, and no search strategy."""

    def __init__(self, initial, goal=None):
        self.initial, self.goal = initial, goal

    def actions(self, state):
        raise NotImplementedError

    def result(self, state, action):
        raise NotImplementedError

    def goal_test(self, state):
        return state == self.goal

    def step_cost(self, state, action, nxt):
        return 1


class TaskOrder(Problem):
    """Finish every task, respecting dependencies.

    A state is the frozenset of finished tasks -- nothing else. The
    order they were finished in is deliberately not recorded.
    """

    def __init__(self):
        super().__init__(frozenset(), ALL)

    def actions(self, state):
        for t, (_, deps) in TASKS.items():
            if t not in state and deps <= state:
                yield t                      # a generator, not a list

    def result(self, state, action):
        return state | {action}              # new value, never mutated

    def step_cost(self, state, action, nxt):
        return TASKS[action][0]              # duration in days


Node = namedtuple("Node", "state parent action cost")


def path(node):
    """Walk parent pointers back to the root."""
    out = []
    while node.parent is not None:
        out.append(node.action)
        node = node.parent
    return list(reversed(out))


# --- the three frontiers, behind one small interface ----------------
class FIFO:
    def __init__(self):
        self.q = deque()

    def push(self, node):
        self.q.append(node)

    def pop(self):
        return self.q.popleft()          # O(1); list.pop(0) is O(n)

    def __len__(self):
        return len(self.q)


class LIFO:
    def __init__(self):
        self.q = []

    def push(self, node):
        self.q.append(node)

    def pop(self):
        return self.q.pop()

    def __len__(self):
        return len(self.q)


class Cheapest:
    """A heap of (priority, tie, node). The counter keeps the heap
    from ever comparing two nodes."""

    def __init__(self):
        self.h, self.tie = [], itertools.count()

    def push(self, node):
        heapq.heappush(self.h, (node.cost, next(self.tie), node))

    def pop(self):
        return heapq.heappop(self.h)[2]

    def __len__(self):
        return len(self.h)


def search(problem, Frontier):
    """One body; the container decides the strategy. Chapter 4 proves
    this by calling it three times."""
    start = Node(problem.initial, None, None, 0)
    frontier, seen = Frontier(), {problem.initial}
    frontier.push(start)
    expanded, peak = 0, 1
    while len(frontier):
        node = frontier.pop()
        if problem.goal_test(node.state):
            return path(node), node.cost, expanded, peak
        expanded += 1
        for a in problem.actions(node.state):
            nxt = problem.result(node.state, a)
            if nxt not in seen:
                seen.add(nxt)
                frontier.push(Node(nxt, node, a,
                                   node.cost + problem.step_cost(
                                       node.state, a, nxt)))
        peak = max(peak, len(frontier))
    return None, None, expanded, peak


# --- the measurement of the worked example -------------------------
def count_states():
    """Reachable situations, when a state is a set of finished tasks."""
    seen, stack = {frozenset()}, [frozenset()]
    while stack:
        s = stack.pop()
        for t, (_, deps) in TASKS.items():
            if t not in s and deps <= s:
                n = s | {t}
                if n not in seen:
                    seen.add(n)
                    stack.append(n)
    return len(seen)


def count_orderings():
    """Reachable sequences, when order is (wrongly) part of the
    state."""
    total = 0
    stack = [frozenset()]
    while stack:
        s = stack.pop()
        total += 1
        for t, (_, deps) in TASKS.items():
            if t not in s and deps <= s:
                stack.append(s | {t})
    return total


if __name__ == "__main__":
    # 1. the representation test of Figure 3.1
    a = frozenset({"spec", "schema", "ui"})
    b = frozenset({"spec", "ui", "schema"})
    assert a == b and len({a, b}) == 1
    print("states equal by two routes:", a == b)

    # 2. the three frontiers on the same problem
    p = TaskOrder()
    print()
    print("frontier      plan len   days   expanded   peak")
    print("-" * 50)
    for name, F in (("FIFO     ", FIFO), ("LIFO     ", LIFO),
                    ("Cheapest ", Cheapest)):
        plan, cost, exp, peak = search(p, F)
        print("%s %9d %6d %10d %6d" % (name, len(plan), cost, exp, peak))

    # 3. the factor the representation buys
    s, o = count_states(), count_orderings()
    print()
    print("situations (frozenset states):", s)
    print("sequences  (order in state)  :", o)
    print("factor saved                 : %.0fx" % (o / s))
    assert o > 5 * s, "sets should collapse many sequences into one"

    # 4. every plan is a legal ordering
    plan, _, _, _ = search(p, FIFO)
    done = set()
    for t in plan:
        assert TASKS[t][1] <= done, "dependency violated: " + t
        done.add(t)
    assert done == set(TASKS)
    print()
    print("plan is a legal topological order:", " -> ".join(plan))
