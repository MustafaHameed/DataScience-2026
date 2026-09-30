# Chapter 7 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 7 lab -- sprint scheduling as a constraint problem.

Twelve stories into four sprints under precedence, same-developer and
capacity constraints. The headline result is the second row: adding
MRV to plain backtracking changes nothing at all, because nothing has
shrunk the domains for it to compare.
"""

import itertools

# story -> (points, owner)
STORIES = {
    "S1": (3, "Ayesha"),  "S2": (5, "Bilal"), "S3": (2, "Ayesha"),
    "S4": (8, "Chen"),    "S5": (3, "Bilal"), "S6": (2, "Chen"),
    "S7": (5, "Ayesha"),  "S8": (3, "Bilal"), "S9": (5, "Dina"),
    "S10": (3, "Dina"),   "S11": (2, "Chen"), "S12": (4, "Ayesha"),
}
NAMES = list(STORIES)
SLOTS = (1, 2, 3, 4)
CAP = 12

# binary: a must finish in an earlier sprint than b
BEFORE = [("S1", "S4"), ("S2", "S5"), ("S3", "S7"), ("S6", "S8"),
          ("S9", "S12"), ("S10", "S11"), ("S4", "S12")]
# binary: same developer, too much combined work for one sprint
APART = [("S1", "S7"), ("S2", "S8"), ("S4", "S6"), ("S9", "S10"),
         ("S7", "S12")]


class Counter:
    def __init__(self):
        self.checks = 0
        self.nodes = 0


def binary_ok(a, va, b, vb, c):
    """Is the pair (a=va, b=vb) consistent? One consistency check."""
    c.checks += 1
    if (a, b) in BEFORE and not va < vb:
        return False
    if (b, a) in BEFORE and not vb < va:
        return False
    if ((a, b) in APART or (b, a) in APART) and va == vb:
        return False
    return True


def capacity_ok(assign, c):
    """The n-ary constraint. No arc can express this."""
    c.checks += 1
    load = dict.fromkeys(SLOTS, 0)
    for n, v in assign.items():
        load[v] += STORIES[n][0]
    return all(load[s] <= CAP for s in SLOTS)


def consistent(assign, var, val, c):
    for other, oval in assign.items():
        if not binary_ok(var, val, other, oval, c):
            return False
    trial = dict(assign)
    trial[var] = val
    return capacity_ok(trial, c)


def select_var(assign, domains, mrv, degree):
    un = [v for v in NAMES if v not in assign]
    if not mrv:
        return un[0]
    fewest = min(len(domains[v]) for v in un)
    tied = [v for v in un if len(domains[v]) == fewest]
    if len(tied) == 1 or not degree:
        return tied[0]

    def deg(v):
        return sum(1 for (a, b) in BEFORE + APART
                   if (a == v and b not in assign)
                   or (b == v and a not in assign))

    return max(tied, key=deg)


def forward_check(var, val, domains, c):
    """Prune the neighbours. None means a domain emptied."""
    nd = {k: list(v) for k, v in domains.items()}
    nd[var] = [val]
    for other in NAMES:
        if other == var:
            continue
        keep = [x for x in nd[other] if binary_ok(var, val, other, x, c)]
        if not keep:
            return None
        nd[other] = keep
    return nd


def ac3(domains, c):
    """Arc consistency over the BINARY constraints only."""
    nd = {k: list(v) for k, v in domains.items()}
    arcs = []
    for (a, b) in BEFORE + APART:
        arcs += [(a, b), (b, a)]
    queue = list(arcs)
    while queue:
        xi, xj = queue.pop(0)
        keep = [vi for vi in nd[xi]
                if any(binary_ok(xi, vi, xj, vj, c) for vj in nd[xj])]
        if len(keep) != len(nd[xi]):
            nd[xi] = keep
            if not nd[xi]:
                return None
            queue += [(a, b) for (a, b) in arcs if b == xi]
    return nd


def solve(mrv=False, degree=False, fc=False, use_ac3=False):
    c = Counter()
    domains = {v: list(SLOTS) for v in NAMES}
    if use_ac3:
        domains = ac3(domains, c)
        if domains is None:
            return None, c
    found = []

    def bt(assign, domains):
        c.nodes += 1
        if len(assign) == len(NAMES):
            found.append(dict(assign))
            return True
        var = select_var(assign, domains, mrv, degree)
        for val in domains[var]:
            if consistent(assign, var, val, c):
                assign[var] = val
                if fc:
                    nd = forward_check(var, val, domains, c)
                    if nd is not None and bt(assign, nd):
                        return True
                elif bt(assign, domains):
                    return True
                del assign[var]
        return False

    bt({}, domains)
    return (found[0] if found else None), c


def count_solutions():
    """Brute force over all 4^12 assignments, for the record."""
    c, n = Counter(), 0
    for combo in itertools.product(SLOTS, repeat=len(NAMES)):
        a = dict(zip(NAMES, combo))
        if all(a[x] < a[y] for x, y in BEFORE) \
           and all(a[x] != a[y] for x, y in APART) \
           and capacity_ok(a, c):
            n += 1
    return n


def show(sol):
    by = {s: [] for s in SLOTS}
    for n, v in sol.items():
        by[v].append(n)
    out = []
    for s in SLOTS:
        pts = sum(STORIES[n][0] for n in by[s])
        out.append("sprint %d: %-18s %2d pts"
                   % (s, ",".join(sorted(by[s], key=lambda x: int(x[1:]))),
                      pts))
    return "\n".join(out)


if __name__ == "__main__":
    total = sum(p for p, _ in STORIES.values())
    print("stories %d   sprints %d   assignments %d"
          % (len(NAMES), len(SLOTS), len(SLOTS) ** len(NAMES)))
    print("points %d against capacity %d -- almost no slack"
          % (total, CAP * len(SLOTS)))
    print("solutions:", count_solutions())
    print()
    print("configuration                    checks  nodes")
    print("-" * 50)
    runs = [
        ("plain backtracking           ", dict()),
        ("+ MRV                        ", dict(mrv=True)),
        ("+ MRV + degree               ", dict(mrv=True, degree=True)),
        ("+ MRV + degree + FC          ", dict(mrv=True, degree=True,
                                                fc=True)),
        ("+ MRV + degree + FC + AC-3   ", dict(mrv=True, degree=True,
                                                fc=True, use_ac3=True)),
    ]
    res = {}
    for name, kw in runs:
        sol, c = solve(**kw)
        assert sol is not None, name
        res[name.strip()] = (c.checks, c.nodes)
        print("%s %7d %6d" % (name, c.checks, c.nodes))

    plain = res["plain backtracking"]
    mrv_only = res["+ MRV"]
    fc = res["+ MRV + degree + FC"]
    ac = res["+ MRV + degree + FC + AC-3"]

    # 1. MRV alone changes NOTHING: no propagation, no information
    assert mrv_only == plain
    print()
    print("MRV alone: %d checks, exactly as without it. Nothing has"
          % mrv_only[0])
    print("shrunk a domain, so every variable ties and MRV cannot")
    print("express a preference.")

    # 2. forward checking is the win, and it is large
    assert fc[0] * 5 < plain[0]
    print()
    print("forward checking: %d -> %d checks (%.0fx), %d -> %d nodes"
          % (plain[0], fc[0], plain[0] / fc[0], plain[1], fc[1]))

    # 3. AC-3 cannot see the capacity constraint, so it buys no nodes
    assert ac[1] == fc[1] and ac[0] > fc[0]
    print()
    print("AC-3 preprocessing: %d extra checks, %d fewer nodes."
          % (ac[0] - fc[0], fc[1] - ac[1]))
    print("It prunes the binary structure and never sees capacity,")
    print("which is the constraint that actually binds.")

    # what AC-3 removes before any search
    c = Counter()
    d0 = {v: list(SLOTS) for v in NAMES}
    d1 = ac3(d0, c)
    narrowed = [v for v in NAMES if d1[v] != d0[v]]
    print()
    print("AC-3 narrowed %d of %d domains before any assignment:"
          % (len(narrowed), len(NAMES)))
    for v in narrowed[:4]:
        print("   %-4s %s -> %s" % (v, d0[v], d1[v]))
    print("   ...")

    print()
    sol, _ = solve(mrv=True, degree=True, fc=True)
    print(show(sol))
