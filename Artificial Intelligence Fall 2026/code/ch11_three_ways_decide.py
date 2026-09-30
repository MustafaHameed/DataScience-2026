# Chapter 11 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 11 lab -- deciding a release gate three ways.

A clause is a frozenset of literals; a literal is "X" or "-X". The
three procedures agree on every question, and differ enormously in
what they cost and in what they hand back.
"""

import itertools

SYMBOLS = ["Reviewed", "Tested", "SignedOff", "Blocker", "Staged",
           "Deploy"]

# The knowledge base, already in conjunctive normal form.
KB = [
    frozenset({"-Reviewed", "-Tested", "Staged"}),      # R1
    frozenset({"-Staged", "-SignedOff", "Deploy"}),     # R2
    frozenset({"-Blocker", "-Deploy"}),                 # R3
    frozenset({"Reviewed"}),                            # F1
    frozenset({"Tested"}),                              # F2
    frozenset({"SignedOff"}),                           # F3
    frozenset({"-Blocker"}),                            # F4
]
LABEL = {KB[0]: "R1", KB[1]: "R2", KB[2]: "R3", KB[3]: "F1",
         KB[4]: "F2", KB[5]: "F3", KB[6]: "F4"}


def neg(lit):
    return lit[1:] if lit.startswith("-") else "-" + lit


def holds(lit, model):
    return (not model[lit[1:]]) if lit.startswith("-") else model[lit]


def clause_true(clause, model):
    return any(holds(l, model) for l in clause)


# --- 1. model checking: the definition, made executable ------------
def model_check(kb, query, symbols):
    """Returns (entailed, models of kb, models enumerated)."""
    models, seen = [], 0
    for combo in itertools.product([False, True], repeat=len(symbols)):
        m = dict(zip(symbols, combo))
        seen += 1
        if all(clause_true(c, m) for c in kb):
            models.append(m)
    entailed = bool(models) and all(m[query] for m in models)
    return entailed, models, seen


# --- 2. DPLL: backtracking with propagation ------------------------
class Work:
    def __init__(self):
        self.calls = 0
        self.units = 0
        self.branches = 0


def simplify(kb, lit):
    out = []
    for c in kb:
        if lit in c:
            continue                       # clause already satisfied
        if neg(lit) in c:
            out.append(c - {neg(lit)})     # clause shortens
        else:
            out.append(c)
    return out


def dpll(kb, assign, w):
    w.calls += 1
    if not kb:
        return True, assign
    if any(len(c) == 0 for c in kb):
        return False, None
    for c in kb:                           # unit propagation
        if len(c) == 1:
            lit = next(iter(c))
            w.units += 1
            a = dict(assign)
            a[lit.lstrip("-")] = not lit.startswith("-")
            return dpll(simplify(kb, lit), a, w)
    lits = {l for c in kb for l in c}
    for l in lits:                         # pure literal
        if neg(l) not in lits:
            a = dict(assign)
            a[l.lstrip("-")] = not l.startswith("-")
            return dpll(simplify(kb, l), a, w)
    l = next(iter(next(iter(kb))))         # branch
    w.branches += 1
    for choice in (l, neg(l)):
        a = dict(assign)
        a[choice.lstrip("-")] = not choice.startswith("-")
        ok, res = dpll(simplify(kb, choice), a, w)
        if ok:
            return True, res
    return False, None


def entails_dpll(kb, query):
    """KB |= q  iff  KB and not-q is unsatisfiable."""
    w = Work()
    sat, _ = dpll(list(kb) + [frozenset({neg(query)})], {}, w)
    return (not sat), w


# --- 3. resolution refutation --------------------------------------
def resolve(ci, cj):
    out = []
    for l in ci:
        if neg(l) in cj:
            out.append((ci - {l}) | (cj - {neg(l)}))
    return out


def resolution(kb, query):
    """Breadth-first saturation. Honest, complete, and slow."""
    clauses = list(kb) + [frozenset({neg(query)})]
    steps = 0
    while True:
        new = set()
        for i in range(len(clauses)):
            for j in range(i + 1, len(clauses)):
                for r in resolve(clauses[i], clauses[j]):
                    steps += 1
                    if not r:
                        return True, steps
                    if r not in clauses:
                        new.add(r)
        if not new - set(clauses):
            return False, steps
        clauses += sorted(new, key=lambda c: (len(c), sorted(c)))


def show(model):
    return "  ".join("%s=%d" % (s, model[s]) for s in SYMBOLS)


if __name__ == "__main__":
    print("symbols %d, so %d models to enumerate"
          % (len(SYMBOLS), 2 ** len(SYMBOLS)))
    print()

    ent, models, seen = model_check(KB, "Deploy", SYMBOLS)
    print("MODEL CHECKING")
    print("   %d of %d models satisfy the KB" % (len(models), seen))
    for m in models:
        print("      " + show(m))
    print("   KB entails Deploy:", ent)
    assert ent and len(models) == 1

    ent2, w = entails_dpll(KB, "Deploy")
    print()
    print("DPLL")
    print("   %d calls, %d unit propagations, %d branches"
          % (w.calls, w.units, w.branches))
    print("   KB entails Deploy:", ent2)
    assert ent2 and ent2 == ent
    assert w.branches == 0, "the facts cascade; no branching needed"

    ent3, steps = resolution(KB, "Deploy")
    print()
    print("RESOLUTION")
    print("   empty clause derived:", ent3, " in %d steps" % steps)
    assert ent3 == ent
    print("   ... which is more work than enumerating all %d models."
          % seen)
    assert steps > seen

    # --- drop the sign-off: the answer becomes no ------------------
    print()
    print("=== drop F3 (SignedOff) ===")
    KB2 = [c for c in KB if c != frozenset({"SignedOff"})]
    ent4, models4, _ = model_check(KB2, "Deploy", SYMBOLS)
    print("   %d models satisfy the KB; entails Deploy: %s"
          % (len(models4), ent4))
    counter = [m for m in models4 if not m["Deploy"]]
    print("   countermodel:")
    print("      " + show(counter[0]))
    assert not ent4 and counter
    ok4, _ = resolution(KB2, "Deploy")
    assert not ok4
    print("   resolution correctly fails to derive the empty clause.")
    print("   The countermodel is the useful output: everything you")
    print("   told me holds, and the release still must not go out.")

    # --- drop the blocker fact: the surprise ----------------------
    print()
    print("=== drop F4 (not Blocker) ===")
    KB3 = [c for c in KB if c != frozenset({"-Blocker"})]
    ent5, models5, _ = model_check(KB3, "Deploy", SYMBOLS)
    print("   %d models satisfy the KB; entails Deploy: %s"
          % (len(models5), ent5))
    assert ent5 and len(models5) == 1
    print("   Blocker is %d in the only model: the KB entails that"
          % models5[0]["Blocker"])
    print("   there is no blocker, without anyone saying so.")
    assert not models5[0]["Blocker"]

    # --- and the fix: put the condition in the rule ---------------
    print()
    print("=== the fix: R2 requires no blocker ===")
    KB4 = [c for c in KB3
           if c != frozenset({"-Staged", "-SignedOff", "Deploy"})]
    KB4.append(frozenset({"-Staged", "-SignedOff", "Blocker", "Deploy"}))
    ent6, models6, _ = model_check(KB4, "Deploy", SYMBOLS)
    print("   %d models satisfy the KB; entails Deploy: %s"
          % (len(models6), ent6))
    assert not ent6
    print("   Now Deploy does NOT follow without knowing about")
    print("   blockers -- which is what the team meant all along.")
