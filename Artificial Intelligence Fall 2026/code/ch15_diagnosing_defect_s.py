# Chapter 15 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 15 lab -- inference in a Bayesian network.

Four ways to compute P(Rushed | Rollback): two exact, two by
sampling. Then explaining away, and the scaling comparison that
justifies variable elimination.
"""

import itertools
import random

# node -> (parents, table mapping parent values to P(node = True))
NET = {
    "Rushed":   ([], {(): 0.30}),
    "Junior":   ([], {(): 0.40}),
    "WeakTest": ([], {(): 0.25}),
    "Defect":   (["Rushed", "Junior", "WeakTest"], {
        (True, True, True): 0.95,   (True, True, False): 0.80,
        (True, False, True): 0.75,  (True, False, False): 0.50,
        (False, True, True): 0.70,  (False, True, False): 0.40,
        (False, False, True): 0.30, (False, False, False): 0.05,
    }),
    "Hotfix":   (["Defect"], {(True,): 0.90, (False,): 0.10}),
    "Rollback": (["Defect"], {(True,): 0.70, (False,): 0.05}),
}
ORDER = ["Rushed", "Junior", "WeakTest", "Defect", "Hotfix", "Rollback"]


def p_true(net, node, a):
    parents, table = net[node]
    return table[tuple(a[p] for p in parents)]


def prob(net, node, value, a):
    p = p_true(net, node, a)
    return p if value else 1 - p


# --- 1. exact: enumeration -----------------------------------------
def enumerate_query(net, order, query, evidence):
    hidden = [n for n in order if n != query and n not in evidence]
    totals, terms = {True: 0.0, False: 0.0}, 0
    for qv in (True, False):
        for combo in itertools.product([True, False],
                                       repeat=len(hidden)):
            a = dict(evidence)
            a[query] = qv
            a.update(dict(zip(hidden, combo)))
            pr = 1.0
            for n in order:
                pr *= prob(net, n, a[n], a)
            totals[qv] += pr
            terms += 1
    z = totals[True] + totals[False]
    return totals[True] / z, terms


# --- 2. exact: variable elimination --------------------------------
class Factor:
    def __init__(self, vars_, table):
        self.vars, self.table = list(vars_), table


def make_factor(net, node, evidence):
    allv = net[node][0] + [node]
    keep = [v for v in allv if v not in evidence]
    table = {}
    for combo in itertools.product([True, False], repeat=len(allv)):
        a = dict(zip(allv, combo))
        if any(v in evidence and a[v] != evidence[v] for v in allv):
            continue
        table[tuple(a[v] for v in keep)] = prob(net, node, a[node], a)
    return Factor(keep, table)


def multiply(f1, f2, ops):
    vars_ = f1.vars + [v for v in f2.vars if v not in f1.vars]
    table = {}
    for combo in itertools.product([True, False], repeat=len(vars_)):
        a = dict(zip(vars_, combo))
        table[combo] = (f1.table[tuple(a[v] for v in f1.vars)]
                        * f2.table[tuple(a[v] for v in f2.vars)])
        ops[0] += 1
    return Factor(vars_, table)


def sum_out(f, var, ops):
    i = f.vars.index(var)
    keep = [v for v in f.vars if v != var]
    table = {}
    for combo, val in f.table.items():
        k = combo[:i] + combo[i + 1:]
        table[k] = table.get(k, 0.0) + val
        ops[0] += 1
    return Factor(keep, table)


def var_elim(net, order, query, evidence):
    ops = [0]
    factors = [f for f in (make_factor(net, n, evidence) for n in order)
               if f.vars]
    for h in [n for n in order if n != query and n not in evidence]:
        involved = [f for f in factors if h in f.vars]
        if not involved:
            continue
        factors = [f for f in factors if h not in f.vars]
        prod = involved[0]
        for f in involved[1:]:
            prod = multiply(prod, f, ops)
        factors.append(sum_out(prod, h, ops))
    prod = factors[0]
    for f in factors[1:]:
        prod = multiply(prod, f, ops)
    i = prod.vars.index(query)
    totals = {True: 0.0, False: 0.0}
    for combo, val in prod.table.items():
        totals[combo[i]] += val
    z = totals[True] + totals[False]
    return totals[True] / z, ops[0]


# --- 3. rejection sampling -----------------------------------------
def rejection(net, order, query, evidence, n, rng):
    kept = hits = 0
    for _ in range(n):
        a = {}
        for node in order:
            a[node] = rng.random() < p_true(net, node, a)
        if all(a[k] == v for k, v in evidence.items()):
            kept += 1
            hits += a[query]
    return (hits / kept if kept else None), kept


# --- 4. likelihood weighting ---------------------------------------
def likelihood_weighting(net, order, query, evidence, n, rng):
    num = den = 0.0
    for _ in range(n):
        a, w = {}, 1.0
        for node in order:
            if node in evidence:
                a[node] = evidence[node]
                w *= prob(net, node, a[node], a)     # weight, not sample
            else:
                a[node] = rng.random() < p_true(net, node, a)
        den += w
        if a[query]:
            num += w
    return num / den, den


def make_chain(n):
    """X1 -> X2 -> ... -> Xn: a service dependency chain."""
    net = {"X1": ([], {(): 0.3})}
    for i in range(2, n + 1):
        net["X%d" % i] = (["X%d" % (i - 1)],
                          {(True,): 0.8, (False,): 0.1})
    return net, ["X%d" % i for i in range(1, n + 1)]


if __name__ == "__main__":
    print("variables: %d   full joint entries: %d"
          % (len(ORDER), 2 ** len(ORDER)))
    print("numbers the network stores: %d"
          % sum(len(NET[n][1]) for n in ORDER))
    print()

    ev = {"Rollback": True}
    exact, terms = enumerate_query(NET, ORDER, "Rushed", ev)
    ve, ops = var_elim(NET, ORDER, "Rushed", ev)
    print("P(Rushed)                = %.4f   (the prior)" % 0.30)
    print("P(Rushed | Rollback)     = %.5f" % exact)
    assert abs(exact - ve) < 1e-9
    print()
    print("method                          answer    work")
    print("-" * 48)
    print("enumeration (exact)           %9.5f  %4d terms" % (exact, terms))
    print("variable elimination (exact)  %9.5f  %4d ops" % (ve, ops))

    rng = random.Random(0)
    rej, kept = rejection(NET, ORDER, "Rushed", ev, 20000, rng)
    rng = random.Random(0)
    lw, _ = likelihood_weighting(NET, ORDER, "Rushed", ev, 20000, rng)
    print("rejection, 20000 samples      %9.4f  %4d kept (%.0f%%)"
          % (rej, kept, 100.0 * kept / 20000))
    print("likelihood weighting, 20000   %9.4f  all samples used" % lw)
    print()
    print("   errors: rejection %.4f, weighting %.4f"
          % (abs(rej - exact), abs(lw - exact)))
    assert abs(lw - exact) < 0.01 and abs(rej - exact) < 0.02

    # variable elimination LOSES on a network this small
    assert ops > terms
    print()
    print("Note that variable elimination used MORE work than")
    print("enumeration here (%d against %d). Its advantage is" % (ops, terms))
    print("asymptotic, and six variables is not asymptotic.")

    # --- explaining away -------------------------------------------
    print()
    print("=== explaining away ===")
    prior = 0.30
    given_d, _ = enumerate_query(NET, ORDER, "Rushed", {"Defect": True})
    given_dj, _ = enumerate_query(NET, ORDER, "Rushed",
                                  {"Defect": True, "Junior": True})
    print("   P(Rushed)                       = %.4f" % prior)
    print("   P(Rushed | Defect)              = %.4f   (rises)" % given_d)
    print("   P(Rushed | Defect, Junior)      = %.4f   (FALLS)"
          % given_dj)
    assert given_d > prior > 0 and given_dj < given_d
    print()
    print("   The junior developer already accounts for the defect,")
    print("   so less explaining is left for the rushed spec to do.")
    print("   Two causes of one effect are independent until the")
    print("   effect is seen, and negatively dependent afterwards.")
    print("   No production rule can express a conclusion that gets")
    print("   WEAKER when a condition is added.")

    # --- and where variable elimination earns its keep -------------
    print()
    print("=== scaling, on a chain of services ===")
    print("  n   enumeration   VE ops")
    print("-" * 30)
    for n in (6, 10, 14, 18):
        net, order = make_chain(n)
        e = {order[-1]: True}
        v, vops = var_elim(net, order, "X1", e)
        if n <= 14:
            ex, et = enumerate_query(net, order, "X1", e)
            assert abs(ex - v) < 1e-9
            print("%3d %13d %8d" % (n, et, vops))
        else:
            print("%3d %13s %8d" % (n, "infeasible", vops))
    print()
    print("Enumeration doubles with every service; variable")
    print("elimination grows linearly on a chain. At n=14 that is")
    print("8192 terms against 186 operations.")
