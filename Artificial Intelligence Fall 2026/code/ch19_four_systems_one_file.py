# Chapter 19 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 19 lab -- the four classic systems, reconstructed.

Nothing here is new. GPS is Chapter 9's means-ends analysis, Eliza is
Chapter 18's pattern substitution, Student is Chapter 12's unification
restricted to strings, and Macsyma is Chapter 13's forward chaining
over expression trees.
"""

import re

# ===================================================================
# 1. GPS (1961): means-ends analysis          -- as in Chapter 9
# ===================================================================
OPERATORS = [
    # (name, preconditions, add, delete)
    ("code", set(), {"coded"}, set()),
    ("review", {"coded"}, {"reviewed"}, set()),
    ("test", {"reviewed", "staging free"},
     {"tested", "staging busy"}, {"staging free"}),
    ("deploy", {"tested"}, {"deployed"}, set()),
    ("release staging", {"staging busy"},
     {"staging free"}, {"staging busy"}),
]


def gps(state, goals, depth=0, trace=None):
    """Achieve each goal by reducing differences. Chapter 9's engine."""
    if trace is None:
        trace = []
    if depth > 20:
        return None, trace
    for g in goals:
        if g in state:
            continue
        achieved = False
        for (name, pre, add, dele) in OPERATORS:
            if g not in add:
                continue
            got, trace = gps(state, sorted(pre), depth + 1, trace)
            if got is None:
                continue
            state = got
            if not pre <= state:
                continue
            state = (state - dele) | add
            trace.append(name)
            achieved = True
            break
        if not achieved:
            return None, trace
    return state, trace


# ===================================================================
# 2. Eliza (1966): pattern substitution       -- as in Chapter 18
# ===================================================================
REFLECT = {"my": "your", "i": "you", "me": "you", "am": "are"}
ELIZA_RULES = [
    (r"^i am (.*)$", "How long have you been {0}?"),
    (r"^my (\w+) keeps (.*)$", "Tell me more about your {0}."),
    (r"^because (.*)$", "Is that the real reason?"),
]


def eliza(line):
    low = line.lower().strip().rstrip(".!?")
    for pat, tmpl in ELIZA_RULES:
        m = re.match(pat, low)
        if m:
            groups = [" ".join(REFLECT.get(w, w) for w in g.split())
                      for g in m.groups()]
            return tmpl.format(*groups)
    return "Can you tell me more about that?"


# ===================================================================
# 3. Student (1964): English -> equations     -- Chapter 12's matching
# ===================================================================
STUDENT_PATTERNS = [
    # "the sprint has twice as many bugs as stories"
    (r"^the \w+ has twice as many (\w+) as (\w+)$",
     lambda m: (m.group(1), ("times", 2, m.group(2)))),
    # "the number of stories is 12"
    (r"^the number of (\w+) is (\d+)$",
     lambda m: (m.group(1), ("const", int(m.group(2))))),
]


def student(sentences, question):
    """Turn English into equations, then solve them."""
    eqs = {}
    for s in sentences:
        low = s.lower().strip().rstrip(".")
        for pat, build in STUDENT_PATTERNS:
            m = re.match(pat, low)
            if m:
                var, rhs = build(m)
                eqs[var] = rhs
                break
        else:
            return None, eqs, "no pattern matched: %r" % s

    def value(var, seen=None):
        seen = seen or set()
        if var in seen:
            return None                 # circular definition
        rhs = eqs.get(var)
        if rhs is None:
            return None
        if rhs[0] == "const":
            return rhs[1]
        if rhs[0] == "times":
            inner = value(rhs[2], seen | {var})
            return None if inner is None else rhs[1] * inner
        return None

    m = re.match(r"^how many (\w+) does the \w+ have$",
                 question.lower().strip().rstrip("?"))
    if not m:
        return None, eqs, "question not understood"
    return value(m.group(1)), eqs, None


# ===================================================================
# 4. Macsyma (1968): term rewriting          -- Chapter 13's chaining
# ===================================================================
# an expression is a nested tuple: ("+", a, b) or ("*", a, b),
# a number, or a string (a symbol)
def rewrite_once(e):
    """One pass of the identities. Returns (expr, rule name or None)."""
    if not isinstance(e, tuple):
        return e, None
    op, a, b = e
    if op == "+" and b == 0:
        return a, "x + 0 -> x"
    if op == "+" and a == 0:
        return b, "0 + x -> x"
    if op == "*" and b == 1:
        return a, "x * 1 -> x"
    if op == "*" and (a == 0 or b == 0):
        return 0, "x * 0 -> 0"
    if op == "+" and a == b:
        return ("*", 2, a), "x + x -> 2x"
    # (a*x) + (b*x) -> (a+b)*x
    if (op == "+" and isinstance(a, tuple) and isinstance(b, tuple)
            and a[0] == "*" and b[0] == "*" and a[2] == b[2]
            and isinstance(a[1], int) and isinstance(b[1], int)):
        return ("*", a[1] + b[1], a[2]), "(a*x)+(b*x) -> (a+b)*x"
    # recurse
    na, r = rewrite_once(a)
    if r:
        return (op, na, b), r
    nb, r = rewrite_once(b)
    if r:
        return (op, a, nb), r
    return e, None


def simplify(e, limit=25):
    """Forward chaining to a fixed point -- Chapter 13, over trees."""
    steps = []
    for _ in range(limit):
        e2, rule = rewrite_once(e)
        if rule is None:
            return e, steps
        steps.append(rule)
        e = e2
    return e, steps


def show(e):
    if not isinstance(e, tuple):
        return str(e)
    return "(%s %s %s)" % (show(e[1]), e[0], show(e[2]))


if __name__ == "__main__":
    print("=" * 62)
    print("1. GENERAL PROBLEM SOLVER (1961) -- means-ends analysis")
    print("=" * 62)
    final, plan = gps(frozenset({"staging free"}), ["deployed"])
    print("   goal: deployed")
    print("   plan:", " -> ".join(plan))
    assert plan == ["code", "review", "test", "deploy"]
    print("   This is Chapter 9's engine, unchanged.")

    print()
    print("=" * 62)
    print("2. ELIZA (1966) -- pattern substitution")
    print("=" * 62)
    for line in ("I am worried about the release",
                 "my deploy keeps failing",
                 "because the config is wrong",
                 "the quarterly numbers look fine"):
        print("   > %-34s %s" % (line, eliza(line)))
    print("   No representation of the conversation. It cannot fail,")
    print("   which is the most dangerous failure mode of the four.")

    print()
    print("=" * 62)
    print("3. STUDENT (1964) -- English into equations")
    print("=" * 62)
    facts = ["the sprint has twice as many bugs as stories",
             "the number of stories is 12"]
    q = "how many bugs does the sprint have?"
    ans, eqs, err = student(facts, q)
    for f in facts:
        print("   given: %s" % f)
    print("   equations:", eqs)
    print("   question: %s" % q)
    print("   answer:", ans)
    assert ans == 24

    print()
    print("   ... and the brittleness:")
    for para in ("the bug count is double the story count",
                 "there are two bugs for every story"):
        ans2, _, err2 = student([para, facts[1]], q)
        print("   %-42s -> %s" % (para, err2 or ans2))
        assert ans2 is None
    print("   Each paraphrase needs its own pattern, and the")
    print("   paraphrases are unbounded. This is Chapter 13's")
    print("   knowledge acquisition bottleneck, in another subject.")

    print()
    print("=" * 62)
    print("4. MACSYMA (1968) -- term rewriting")
    print("=" * 62)
    expr = ("+", ("+", ("*", 3, "e"), ("*", 2, "e")), 0)
    print("   before:", show(expr))
    out, steps = simplify(expr)
    print("   after: ", show(out))
    for s in steps:
        print("      applied  %s" % s)
    assert out == ("*", 5, "e")
    print("   Forward chaining to a fixed point, over trees instead")
    print("   of over facts. Chapter 13's engine, same conflict-")
    print("   resolution question: when two rules apply, which?")

    print()
    print("=" * 62)
    print("THE ANALYSIS")
    print("=" * 62)
    ANALYSIS = [
        ("GPS", "reach a goal", "operators", "means-ends",
         "hand-written", "incomplete; goal order"),
        ("Eliza", "converse", "patterns", "match+substitute",
         "hand-written", "NONE DETECTABLE"),
        ("Student", "solve a problem", "patterns+algebra", "match, solve",
         "hand-written", "brittle to paraphrase"),
        ("Macsyma", "simplify", "expression trees", "rewrite",
         "hand-written", "loops; wrong fixed point"),
    ]
    print("%-9s %-16s %-17s %-17s %-13s %s"
          % ("SYSTEM", "TASK", "REPRESENTATION", "INFERENCE",
             "KNOWLEDGE", "FAILURE MODE"))
    print("-" * 100)
    for row in ANALYSIS:
        print("%-9s %-16s %-17s %-17s %-13s %s" % row)

    sources = {r[4] for r in ANALYSIS}
    assert sources == {"hand-written"}
    print()
    print("Every one of the four has the same knowledge source, and")
    print("that single fact explains all four biographies: each")
    print("worked, each needed a person to extend it, and each ran")
    print("out of people before it ran out of domain.")
