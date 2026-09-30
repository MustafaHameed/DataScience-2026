# Chapter 16 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 16 lab -- rating project risk by Mamdani inference.

Two inputs, nine rules, one number out. The interesting outputs are
the smoothness table and the range check: a centroid defuzzifier
cannot reach the ends of its own scale.
"""

import numpy as np


def tri(x, a, b, c):
    """Triangular membership: 0 at a, 1 at b, 0 at c."""
    x = np.asarray(x, dtype=float)
    return np.clip(np.minimum((x - a) / (b - a), (c - x) / (c - b)), 0, 1)


# Fuzzy partitions: peaks evenly spaced, memberships summing to one,
# so every input is described by at most two words.
EFFORT = {                                   # person-days, 0..100
    "low":    lambda x: tri(x, -50, 0, 50),
    "medium": lambda x: tri(x, 0, 50, 100),
    "high":   lambda x: tri(x, 50, 100, 150),
}
UNCERTAINTY = {                              # the team's 0..10 scale
    "low":    lambda x: tri(x, -5, 0, 5),
    "medium": lambda x: tri(x, 0, 5, 10),
    "high":   lambda x: tri(x, 5, 10, 15),
}
RISK = {                                     # score, 0..100
    "low":    lambda x: tri(x, -50, 0, 50),
    "medium": lambda x: tri(x, 0, 50, 100),
    "high":   lambda x: tri(x, 50, 100, 150),
}

# (effort word, uncertainty word) -> risk word
RULES = [
    ("low", "low", "low"), ("low", "medium", "low"),
    ("low", "high", "medium"),
    ("medium", "low", "low"), ("medium", "medium", "medium"),
    ("medium", "high", "high"),
    ("high", "low", "medium"), ("high", "medium", "high"),
    ("high", "high", "high"),
]

GRID = np.linspace(0, 100, 1001)


def mamdani(effort, uncertainty, trace=False):
    # 1. fuzzify
    me = {k: float(f(effort)) for k, f in EFFORT.items()}
    mu = {k: float(f(uncertainty)) for k, f in UNCERTAINTY.items()}
    if trace:
        print("   effort %g      -> %s" % (effort,
              {k: round(v, 3) for k, v in me.items() if v > 0}))
        print("   uncertainty %g -> %s" % (uncertainty,
              {k: round(v, 3) for k, v in mu.items() if v > 0}))

    # 2. firing strength = min of the conditions
    # 3. clip the consequent at that strength
    # 4. aggregate by max
    aggregate = np.zeros_like(GRID)
    fired = []
    for (e, u, r) in RULES:
        w = min(me[e], mu[u])
        if w > 1e-9:
            fired.append((e, u, r, w))
            aggregate = np.maximum(aggregate, np.minimum(w, RISK[r](GRID)))
    if trace:
        for (e, u, r, w) in fired:
            print("   %-6s x %-6s -> %-6s   strength %.3f"
                  % (e, u, r, w))

    # 5. defuzzify by centroid
    area = aggregate.sum()
    crisp = float((GRID * aggregate).sum() / area) if area else 0.0
    return crisp, fired


def threshold_rule(effort, uncertainty):
    """The same knowledge as a crisp rule, for contrast."""
    return 90 if (effort > 40 and uncertainty > 7) else 10


PROJECTS = {
    "Atlas":  (20, 2),
    "Beacon": (45, 4),
    "Cobalt": (85, 9),
    "Delta":  (45, 9),
    "Echo":   (90, 2),
}

if __name__ == "__main__":
    print("project   effort  uncert    risk   rules fired")
    print("-" * 50)
    for name, (e, u) in PROJECTS.items():
        risk, fired = mamdani(e, u)
        print("%-9s %6d %7d %7.2f %8d" % (name, e, u, risk, len(fired)))
        assert len(fired) == 4, "a partition should fire 2x2 rules"

    print()
    print("=== Beacon in full (effort 45, uncertainty 4) ===")
    risk, fired = mamdani(45, 4, trace=True)
    print("   crisp risk = %.2f" % risk)
    assert abs(risk - 49.04) < 0.01

    # the memberships are exactly the hand calculation
    assert abs(float(EFFORT["low"](45)) - 0.1) < 1e-9
    assert abs(float(EFFORT["medium"](45)) - 0.9) < 1e-9
    assert abs(float(UNCERTAINTY["low"](4)) - 0.2) < 1e-9
    assert abs(float(UNCERTAINTY["medium"](4)) - 0.8) < 1e-9

    # --- smoothness ------------------------------------------------
    print()
    print("=== effort fixed at 45, uncertainty swept ===")
    print("   unc   fuzzy    threshold rule")
    print("   " + "-" * 34)
    curve = []
    for u in range(11):
        r, _ = mamdani(45, u)
        curve.append(r)
        print("   %3d %8.2f %12d" % (u, r, threshold_rule(45, u)))

    # the fuzzy output is monotone and moves in small steps
    steps = [curve[i + 1] - curve[i] for i in range(len(curve) - 1)]
    assert all(s > 0 for s in steps), "should be monotone increasing"
    print()
    print("   fuzzy: largest single step %.2f points" % max(steps))
    print("   threshold rule: one step of 80 points, at u = 7 -> 8")
    assert max(steps) < 20

    # --- the range limitation --------------------------------------
    print()
    print("=== what range can the output actually reach? ===")
    lo, hi = 1e9, -1e9
    for e in range(0, 101, 5):
        for u in range(0, 11):
            r, _ = mamdani(e, u)
            lo, hi = min(lo, r), max(hi, r)
    print("   across the whole input space, risk lies in [%.2f, %.2f]"
          % (lo, hi))
    print("   a minimum-effort, minimum-uncertainty project scores")
    print("   %.2f, not 0. Centroid defuzzification cannot reach the"
          % mamdani(0, 0)[0])
    print("   ends of its own scale, so never label this 'percent'.")
    assert lo > 5 and hi < 95

    # --- and the confusion this chapter exists to prevent ----------
    print()
    print("=== membership is not probability ===")
    print("   mu_high(effort=70) = %.3f" % float(EFFORT["high"](70)))
    print("   This does NOT mean a 40% chance the project is high")
    print("   effort. It is 70 person-days -- we are certain of that.")
    print("   The number describes the word 'high', not our ignorance,")
    print("   and no further measurement will change it.")
    total = sum(float(f(70)) for f in EFFORT.values())
    print("   memberships at 70 sum to %.3f (probabilities would sum"
          % total)
    print("   to 1 only by coincidence; here it is by construction)")
