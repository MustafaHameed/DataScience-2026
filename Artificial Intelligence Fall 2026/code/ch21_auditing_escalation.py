# Chapter 21 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 21 lab -- auditing a rule base nobody wrote badly.

The rules mention no team. The escalation rate differs by 38
percentage points between two teams. Everything here is measurement:
the disparity is invisible in the rule base and obvious in the
outcomes.
"""

import random

# ---- Chapter 13's rule base, unchanged ----------------------------
RULES = [
    ("R1",  [("severity", "critical")], ("priority", "P1")),
    ("R2",  [("severity", "major"), ("tier", "platinum")],
     ("priority", "P1")),
    ("R3",  [("severity", "major"), ("tier", "standard")],
     ("priority", "P2")),
    ("R4",  [("severity", "minor")], ("priority", "P3")),
    ("R5",  [("priority", "P1")], ("needs_owner", "yes")),
    ("R6",  [("priority", "P2"), ("workaround", "no")],
     ("needs_owner", "yes")),
    ("R7",  [("priority", "P2"), ("workaround", "yes")],
     ("needs_owner", "no")),
    ("R8",  [("priority", "P3")], ("needs_owner", "no")),
    ("R9",  [("needs_owner", "yes"), ("owner_available", "no")],
     ("escalate", "yes")),
    ("R11", [("needs_owner", "no")], ("escalate", "no")),
]


def forward_chain(facts, rules=RULES):
    wm, fired = dict(facts), set()
    while True:
        applicable = [r for r in rules
                      if r[0] not in fired
                      and all(wm.get(a) == v for (a, v) in r[1])
                      and r[2][0] not in wm]
        if not applicable:
            return wm, fired
        applicable.sort(key=lambda r: (-len(r[1]), rules.index(r)))
        r = applicable[0]
        wm[r[2][0]] = r[2][1]
        fired.add(r[0])


def explain(goal, facts, rules=RULES, seen=None):
    """Backward chaining, for the chain an auditor reads."""
    seen = seen or set()
    if facts.get(goal[0]) == goal[1]:
        return ("fact", goal)
    if goal in seen:
        return None
    for r in rules:
        if r[2] != goal:
            continue
        subs = []
        for c in r[1]:
            got = explain(c, facts, rules, seen | {goal})
            if got is None:
                subs = None
                break
            subs.append(got)
        if subs is not None:
            return ("rule", r[0], goal, subs)
    return None


def render(tree, indent=0):
    pad = "  " * indent
    if tree[0] == "fact":
        return ["%s%s = %s  (given)" % (pad, tree[1][0], tree[1][1])]
    _, name, goal, subs = tree
    out = ["%s%s = %s  (by %s, because)" % (pad, goal[0], goal[1], name)]
    for s in subs:
        out += render(s, indent + 1)
    return out


# ---- two hundred tickets from two teams ---------------------------
def population(n=200, seed=0):
    """The ONLY difference between the teams is how often a
    workaround is documented. Nothing else, and no rule sees the
    team name."""
    rng = random.Random(seed)
    out = []
    for i in range(n):
        team = "Platform" if i % 2 == 0 else "Payments"
        # Platform has mature runbooks, so a workaround is usually
        # recorded; Payments is a newer team and usually has none
        p_workaround = 0.80 if team == "Platform" else 0.15
        out.append({
            "id": "T-%04d" % (4000 + i),
            "team": team,                       # for the AUDIT only
            "severity": rng.choice(["critical", "major", "major",
                                    "minor"]),
            "tier": rng.choice(["platinum", "standard", "standard"]),
            "workaround": "yes" if rng.random() < p_workaround else "no",
            "owner_available": rng.choice(["yes", "no", "no"]),
        })
    return out


def decide(ticket, rules=RULES):
    facts = {k: v for k, v in ticket.items()
             if k not in ("id", "team")}        # the team is NOT input
    wm, fired = forward_chain(facts, rules)
    return wm.get("escalate", "no"), fired, facts


def rates(tickets, rules=RULES):
    counts = {}
    for t in tickets:
        esc, _, _ = decide(t, rules)
        a, b = counts.get(t["team"], (0, 0))
        counts[t["team"]] = (a + (esc == "yes"), b + 1)
    return {k: (a, b, 100.0 * a / b) for k, (a, b) in counts.items()}


if __name__ == "__main__":
    tickets = population()
    print("no rule mentions the reporting team. Rule attributes used:")
    print("   ", sorted({a for r in RULES for (a, _) in r[1]}))
    print()

    r = rates(tickets)
    print("=== step 1: measure the outcomes ===")
    print("   team        tickets  escalated   rate")
    print("   " + "-" * 40)
    for team in sorted(r):
        esc, tot, pct = r[team]
        print("   %-10s %8d %10d %6.0f%%" % (team, tot, esc, pct))
    gap = abs(r["Payments"][2] - r["Platform"][2])
    print("   gap: %.0f percentage points" % gap)
    assert gap > 15

    # --- step 2: which rule carries the gap? -----------------------
    overall = sum(x[0] for x in r.values()) / sum(x[1] for x in r.values())
    print("   overall escalation rate: %.0f%%" % (100 * overall))

    # --- step 2: which rule carries the gap? -----------------------
    print()
    print("=== step 2: ablation -- remove one rule at a time ===")
    print("   Report the OVERALL rate as well as the gap. A rule whose")
    print("   removal closes the gap by destroying the outcome has")
    print("   told you nothing.")
    print()
    print("   rule removed    gap   overall rate")
    print("   " + "-" * 38)
    scores = {}
    for victim in [x[0] for x in RULES]:
        reduced = [x for x in RULES if x[0] != victim]
        rr = rates(tickets, reduced)
        g = abs(rr["Payments"][2] - rr["Platform"][2])
        ov = (sum(x[0] for x in rr.values())
              / sum(x[1] for x in rr.values()))
        scores[victim] = (g, ov)
        flag = ""
        if ov < 0.5 * overall:
            flag = "  <- outcome collapsed; tells us nothing"
        print("   %-14s %5.0f %11.0f%%%s" % (victim, g, 100 * ov, flag))

    # R9 is the only rule concluding escalate=yes, so removing it
    # sets every rate to zero and the gap to zero. That is the trap.
    assert scores["R9"][0] == 0 and scores["R9"][1] == 0

    # the honest question: among rules that leave the outcome intact,
    # which one carries the gap?
    intact = {k: v for k, v in scores.items()
              if v[1] >= 0.5 * overall}
    worst = min(intact, key=lambda k: intact[k][0])
    print()
    print("   Among rules that leave the outcome intact, removing %s"
          % worst)
    print("   closes the gap most: %.0f -> %.0f points."
          % (gap, intact[worst][0]))
    assert worst in ("R3", "R6")

    # --- step 3: read the chain for an affected ticket -------------
    print()
    print("=== step 3: read the explanation ===")
    for t in tickets:
        esc, fired, facts = decide(t)
        if t["team"] == "Platform" and esc == "no" and "R7" in fired:
            print("   ticket %s (%s), not escalated:"
                  % (t["id"], t["team"]))
            tree = explain(("escalate", "no"), facts)
            for line in render(tree):
                print("      " + line)
            break
    print()
    print("   R7 is a good rule. 'workaround' is recorded for 80% of")
    print("   Platform tickets and 15% of Payments tickets, because")
    print("   Platform has runbooks. A difference in DOCUMENTATION is")
    print("   being read as a difference in OPERATIONAL NEED.")

    # --- the competence boundary -----------------------------------
    print()
    print("=== what the rule base does NOT cover ===")
    import itertools
    attrs = {"severity": ["critical", "major", "minor"],
             "tier": ["platinum", "standard"],
             "workaround": ["yes", "no"],
             "owner_available": ["yes", "no"]}
    uncovered = []
    for combo in itertools.product(*attrs.values()):
        facts = dict(zip(attrs.keys(), combo))
        wm, _ = forward_chain(facts)
        if "escalate" not in wm:
            uncovered.append(facts)
    print("   %d of %d possible situations reach no escalate decision"
          % (len(uncovered), 3 * 2 * 2 * 2))
    for u in uncovered[:3]:
        print("      ", u)
    print("   A rule base can be asked this. Enumerate the inputs,")
    print("   find the gaps, and SAY SO -- that is a boundary the")
    print("   system can describe, and it takes an afternoon.")

    # --- the computational limit -----------------------------------
    print()
    print("=== and the limit no engineering removes ===")
    print("   b=4:  depth   nodes          at 1e6 nodes/sec")
    print("   " + "-" * 46)
    for d in (10, 15, 20, 25):
        n = 4 ** d
        secs = n / 1e6
        if secs < 90:
            t = "%.1f seconds" % secs
        elif secs < 5400:
            t = "%.1f minutes" % (secs / 60)
        elif secs < 86400 * 2:
            t = "%.1f hours" % (secs / 3600)
        else:
            t = "%.0f days" % (secs / 86400)
        print("   %11d %14d   %s" % (d, n, t))
    print()
    print("   A 100x faster machine buys three or four levels of")
    print("   depth and then stops. Structure is the only real")
    print("   answer, which is what Parts II to V were about.")
