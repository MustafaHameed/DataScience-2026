# Chapter 12 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 12 lab -- first-order resolution on a delay chain.

Representation: a constant is a string; a variable is ("?", name);
a compound term is (functor, arg, ...); a literal is
(sign, predicate, args).
"""


def var(n):
    return ("?", n)


def is_var(t):
    return isinstance(t, tuple) and len(t) == 2 and t[0] == "?"


def is_compound(t):
    return isinstance(t, tuple) and not is_var(t)


def walk(t, s):
    while is_var(t) and t in s:
        t = s[t]
    return t


def occurs(v, t, s):
    """The occurs check: is v inside t? Omit this and unify(x, f(x))
    succeeds with a term that does not exist."""
    t = walk(t, s)
    if t == v:
        return True
    if is_compound(t):
        return any(occurs(v, a, s) for a in t[1:])
    return False


def unify(x, y, s):
    """Most general unifier, or None."""
    if s is None:
        return None
    x, y = walk(x, s), walk(y, s)
    if x == y:
        return s
    if is_var(x):
        if occurs(x, y, s):
            return None
        s = dict(s)
        s[x] = y
        return s
    if is_var(y):
        return unify(y, x, s)
    if is_compound(x) and is_compound(y):
        if x[0] != y[0] or len(x) != len(y):
            return None
        for a, b in zip(x[1:], y[1:]):
            s = unify(a, b, s)
            if s is None:
                return None
        return s
    return None


def subst(t, s):
    t = walk(t, s)
    if is_compound(t):
        return (t[0],) + tuple(subst(a, s) for a in t[1:])
    return t


def lit_subst(lit, s):
    sign, pred, args = lit
    return (sign, pred, tuple(subst(a, s) for a in args))


def standardise(clause, n):
    """Rename every variable, so two clauses never share one."""
    out = []
    for (sign, pred, args) in clause:
        out.append((sign, pred,
                    tuple(("?", "%s#%d" % (a[1], n)) if is_var(a) else a
                          for a in args)))
    return frozenset(out)


def resolve(ci, cj, n):
    ci, cj = standardise(ci, n), standardise(cj, n + 500000)
    out = []
    for li in ci:
        for lj in cj:
            if li[0] == lj[0] or li[1] != lj[1]:
                continue
            if len(li[2]) != len(lj[2]):
                continue
            s, ok = {}, True
            for a, b in zip(li[2], lj[2]):
                s = unify(a, b, s)
                if s is None:
                    ok = False
                    break
            if ok:
                out.append(frozenset(
                    lit_subst(l, s) for l in (ci - {li}) | (cj - {lj})))
    return out


def show_term(a):
    if is_var(a):
        return a[1]
    if is_compound(a):
        return "%s(%s)" % (a[0], ",".join(show_term(x) for x in a[1:]))
    return a


def show(clause):
    if not clause:
        return "[]  (the empty clause)"
    parts = []
    for (sign, pred, args) in clause:
        parts.append(("" if sign else "-") +
                     "%s(%s)" % (pred,
                                 ",".join(show_term(a) for a in args)))
    return " OR ".join(sorted(parts))


# --- the knowledge base -------------------------------------------
x, y = var("x"), var("y")
R1 = frozenset({(False, "depends", (x, y)),
                (False, "delayed", (y,)),
                (True, "delayed", (x,))})
F1 = frozenset({(True, "depends", ("Release47", "Payments"))})
F2 = frozenset({(True, "depends", ("Payments", "Auth"))})
F3 = frozenset({(True, "delayed", ("Auth",))})
KB = [R1, F1, F2, F3]
QUERY = (True, "delayed", ("Release47",))


def refute_naive(kb, query, rounds=6):
    """Resolve every pair, keep everything new, repeat."""
    clauses = list(kb) + [frozenset({(not query[0], query[1], query[2])})]
    steps = 0
    for _ in range(rounds):
        new = []
        for i in range(len(clauses)):
            for j in range(len(clauses)):
                if i == j:
                    continue
                for r in resolve(clauses[i], clauses[j], steps):
                    steps += 1
                    if not r:
                        return True, steps
                    if r not in clauses and r not in new:
                        new.append(r)
        if not new:
            return False, steps
        clauses += new
    return False, steps


def refute_sos(kb, query, rounds=8):
    """Set of support: every resolution must involve a clause
    descended from the negated query. Nothing else is relevant to
    proving THIS conclusion, and completeness is preserved."""
    neg_q = frozenset({(not query[0], query[1], query[2])})
    base, sos = list(kb), [neg_q]
    seen = set(base) | {neg_q}
    steps = 0
    for _ in range(rounds):
        new = []
        for s in sos:
            for b in base + sos:
                if b is s:
                    continue
                for r in resolve(s, b, steps):
                    steps += 1
                    if not r:
                        return True, steps
                    if r not in seen:
                        seen.add(r)
                        new.append(r)
        if not new:
            return False, steps
        sos += new
    return False, steps


if __name__ == "__main__":
    print("=== unification ===")
    cases = [
        ("blocks(x, Auth) vs blocks(Release47, y)",
         ("blocks", x, "Auth"), ("blocks", "Release47", y)),
        ("blocks(x, x)   vs blocks(Auth, Payments)",
         ("blocks", x, x), ("blocks", "Auth", "Payments")),
        ("x              vs f(x)   [occurs check]",
         x, ("f", x)),
    ]
    for label, a, b in cases:
        s = unify(a, b, {})
        if s is None:
            print("   %-42s FAIL" % label)
        else:
            binding = ", ".join("%s/%s" % (k[1], show_term(v))
                                for k, v in sorted(s.items()))
            print("   %-42s {%s}" % (label, binding))

    # the three results the chapter claims
    assert unify(("blocks", x, "Auth"),
                 ("blocks", "Release47", y), {}) is not None
    assert unify(("blocks", x, x),
                 ("blocks", "Auth", "Payments"), {}) is None
    assert unify(x, ("f", x), {}) is None

    print()
    print("=== the proof, by hand ===")
    # step 6: R1 with F3
    s6 = unify(("delayed", y), ("delayed", "Auth"), {})
    c6 = frozenset(lit_subst(l, s6)
                   for l in R1 - {(False, "delayed", (y,))})
    print("   6.", show(c6), "   theta = {y/Auth}")
    # step 7: that with F2
    got = resolve(c6, F2, 1)
    c7 = min(got, key=len)
    print("   7.", show(c7), "   theta = {x/Payments}")
    assert c7 == frozenset({(True, "delayed", ("Payments",))})
    # step 8: R1 with step 7
    c8 = min((c for c in resolve(R1, c7, 2) if len(c) == 2), key=len)
    print("   8.", show(c8), "   theta = {y/Payments}")
    # step 9: that with F1
    c9 = min(resolve(c8, F1, 3), key=len)
    print("   9.", show(c9), "   theta = {x/Release47}")
    assert c9 == frozenset({(True, "delayed", ("Release47",))})
    print("  10. [] -- contradicts the assumed -delayed(Release47)")
    print()
    print("   R1 was used three times, with three different unifiers.")
    print("   That is exactly what propositional logic cannot do.")

    print()
    print("=== finding it automatically ===")
    ok1, n1 = refute_naive(KB, QUERY)
    ok2, n2 = refute_sos(KB, QUERY)
    print("   naive saturation : proved %s in %4d resolution steps"
          % (ok1, n1))
    print("   set of support   : proved %s in %4d resolution steps"
          % (ok2, n2))
    print("   set of support is %.1fx cheaper" % (n1 / n2))
    assert ok1 and ok2 and n2 < n1
    print()
    print("   Both find a five-line proof. Checking a proof is easy;")
    print("   finding one is the hard part, and real provers add")
    print("   subsumption and indexing on top of this.")

    # and a query that does NOT follow
    print()
    print("=== a query that does not follow ===")
    ok3, n3 = refute_sos(KB, (True, "delayed", ("Billing",)), rounds=4)
    print("   delayed(Billing): proved %s after %d steps" % (ok3, n3))
    assert not ok3
    print("   No proof was found -- but first-order entailment is only")
    print("   semi-decidable, so 'no proof yet' is not the same as")
    print("   'no proof exists'. Chapter 11 could hand back a")
    print("   countermodel here; in general this chapter cannot.")
