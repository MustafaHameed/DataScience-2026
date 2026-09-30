# Chapter 13 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 13 lab -- a rule-based escalation adviser.

Twenty rules an operations lead could read and correct. The engine
fires until nothing new follows; the explanation tree is what makes
the result auditable, and it is a by-product of backward chaining
rather than something added afterwards.
"""

# (name, [conditions], conclusion); each is an (attribute, value) pair
RULES = [
    # R0 is a deliberately general fallback, and it is listed FIRST.
    # Plain rule-order resolution therefore reaches it before the
    # narrower R2 and R3, and gets the wrong answer. Specificity does
    # not, which is the point of the comparison below.
    ("R0",  [("severity", "major")], ("priority", "P1")),
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
    ("R10", [("priority", "P1"), ("sla_risk", "high")],
     ("escalate", "yes")),
    ("R11", [("needs_owner", "no")], ("escalate", "no")),
    ("R12", [("sla_hours_left", "under2")], ("sla_risk", "high")),
    ("R13", [("sla_hours_left", "over2")], ("sla_risk", "low")),
    ("R14", [("escalate", "yes"), ("wake_budget", "spent")],
     ("action", "queue_for_morning")),
    ("R15", [("escalate", "yes"), ("wake_budget", "available")],
     ("action", "wake_on_call")),
    ("R16", [("escalate", "no")], ("action", "leave_in_queue")),
    ("R17", [("severity", "critical"), ("tier", "platinum")],
     ("notify_account_team", "yes")),
    ("R18", [("service", "payments"), ("priority", "P1")],
     ("notify_finance", "yes")),
    ("R19", [("escalate", "yes"), ("service", "payments")],
     ("runbook", "payments_incident")),
    ("R20", [("action", "wake_on_call"), ("tier", "platinum")],
     ("page_severity", "urgent")),
]

TICKETS = {
    "T-4471": {"severity": "critical", "tier": "platinum",
               "service": "payments", "workaround": "no",
               "sla_hours_left": "under2", "owner_available": "no",
               "wake_budget": "available"},
    "T-4472": {"severity": "major", "tier": "standard",
               "service": "search", "workaround": "yes",
               "sla_hours_left": "over2", "owner_available": "no",
               "wake_budget": "available"},
    "T-4473": {"severity": "major", "tier": "platinum",
               "service": "auth", "workaround": "no",
               "sla_hours_left": "over2", "owner_available": "no",
               "wake_budget": "spent"},
}


class Trace:
    def __init__(self):
        self.fired = []
        self.tests = 0
        self.cycles = 0


def matches(rule, wm, t):
    for (attr, val) in rule[1]:
        t.tests += 1
        if wm.get(attr) != val:
            return False
    return True


def forward_chain(facts, resolve="specificity", rules=None):
    """Match, resolve, act -- until nothing new follows.

    The 'fired' set is REFRACTION: without it the engine re-fires the
    same rule on the same facts forever.

    'r[2][0] not in wm' means each attribute is DECIDED ONCE, by the
    rule conflict resolution picks. Without it a less specific rule
    that fires later would quietly overwrite a more specific rule's
    conclusion, and conflict resolution would be pointless.
    """
    rules = RULES if rules is None else rules
    wm, t, fired = dict(facts), Trace(), set()
    while True:
        t.cycles += 1
        conflict = [r for r in rules
                    if r[0] not in fired
                    and matches(r, wm, t)
                    and r[2][0] not in wm]
        if not conflict:
            return wm, t
        if resolve == "specificity":
            # more conditions first; ties broken by rule order
            conflict.sort(key=lambda r: (-len(r[1]), rules.index(r)))
        else:
            conflict.sort(key=lambda r: rules.index(r))
        r = conflict[0]
        wm[r[2][0]] = r[2][1]
        fired.add(r[0])
        t.fired.append(r[0])


def backward_chain(goal, facts, seen=None):
    """Prove an (attribute, value) goal. Returns a proof tree."""
    if seen is None:
        seen = set()
    if facts.get(goal[0]) == goal[1]:
        return ("fact", goal)
    if goal in seen:
        return None                 # a cycle in the rule base
    seen = seen | {goal}
    for r in RULES:
        if r[2] != goal:
            continue
        subs = []
        for c in r[1]:
            got = backward_chain(c, facts, seen)
            if got is None:
                subs = None
                break
            subs.append(got)
        if subs is not None:
            return ("rule", r[0], goal, subs)
    return None


def explain(tree, indent=0):
    pad = "  " * indent
    if tree[0] == "fact":
        return ["%s%s = %s   (given in the ticket)"
                % (pad, tree[1][0], tree[1][1])]
    _, name, goal, subs = tree
    out = ["%s%s = %s   (by %s, because)" % (pad, goal[0], goal[1], name)]
    for s in subs:
        out += explain(s, indent + 1)
    return out


def goals_visited(goal, facts, seen=None, count=None):
    """How much work backward chaining does for ONE question."""
    if count is None:
        count = [0]
    count[0] += 1
    if seen is None:
        seen = set()
    if facts.get(goal[0]) == goal[1]:
        return True, count[0]
    if goal in seen:
        return False, count[0]
    seen = seen | {goal}
    for r in RULES:
        if r[2] != goal:
            continue
        if all(goals_visited(c, facts, seen, count)[0] for c in r[1]):
            return True, count[0]
    return False, count[0]


if __name__ == "__main__":
    print("rule base: %d rules" % len(RULES))
    print()
    print("ticket    action              rules  tests  cycles")
    print("-" * 52)
    results = {}
    for tid, facts in TICKETS.items():
        wm, t = forward_chain(facts)
        results[tid] = (wm, t)
        print("%-9s %-19s %5d %6d %7d"
              % (tid, wm.get("action"), len(t.fired), t.tests, t.cycles))

    # the three actions differ, and only one fact separates the last two
    assert results["T-4471"][0]["action"] == "wake_on_call"
    assert results["T-4472"][0]["action"] == "leave_in_queue"
    assert results["T-4473"][0]["action"] == "queue_for_morning"
    print()
    print("T-4471 firing order:", " ".join(results["T-4471"][1].fired))
    print("T-4473 firing order:", " ".join(results["T-4473"][1].fired))
    print()
    print("Both are escalated; they differ only in wake_budget, and")
    print("that is a fact about the ENGINEER, not about the ticket.")
    assert results["T-4471"][0]["escalate"] == "yes"
    assert results["T-4473"][0]["escalate"] == "yes"

    # --- specificity is doing real work --------------------------
    print()
    print("=== specificity against plain rule order ===")
    print("ticket    specificity         rule order")
    print("-" * 46)
    for tid, facts in TICKETS.items():
        a, _ = forward_chain(facts, "specificity")
        b, _ = forward_chain(facts, "order")
        flag = "  <-- DIFFER" if a["action"] != b["action"] else ""
        print("%-9s %-19s %-19s%s"
              % (tid, a["action"], b["action"], flag))

    spec, _ = forward_chain(TICKETS["T-4472"], "specificity")
    order, _ = forward_chain(TICKETS["T-4472"], "order")
    assert spec["priority"] == "P2" and spec["action"] == "leave_in_queue"
    assert order["priority"] == "P1" and order["action"] == "wake_on_call"
    print()
    print("T-4472 is major severity on a STANDARD tier customer.")
    print("R0 ('major -> P1') is general and listed first, so plain")
    print("rule order reaches it and wakes an engineer at 3 a.m. for")
    print("a standard-tier ticket with a workaround. Specificity")
    print("prefers the two-condition R3, gets P2, and leaves it in")
    print("the queue. That is a real engineer's night, decided by the")
    print("conflict-resolution strategy alone.")

    # --- the explanation ------------------------------------------
    print()
    print("=== why is T-4471 escalated? ===")
    tree = backward_chain(("escalate", "yes"), TICKETS["T-4471"])
    assert tree is not None
    for line in explain(tree):
        print("   " + line)

    # --- the two directions DISAGREE, and that is a real bug -----
    print()
    print("=== and is T-4472 escalated? ===")
    fwd_wm, _ = forward_chain(TICKETS["T-4472"])
    up = backward_chain(("escalate", "yes"), TICKETS["T-4472"])
    print("   forward chaining says escalate =", fwd_wm["escalate"])
    print("   backward chaining proves escalate = yes:", up is not None)
    assert fwd_wm["escalate"] == "no" and up is not None
    print()
    print("   They disagree. Backward chaining's proof:")
    for line in explain(up):
        print("      " + line)
    print()
    print("   Backward chaining has no conflict resolution: it asks")
    print("   'is there ANY rule concluding this?' and R0 is one.")
    print("   Forward chaining never used R0 because specificity")
    print("   preferred R3 -- so specificity was MASKING a bad rule,")
    print("   not fixing it. R0 says every major ticket is P1, which")
    print("   is simply false, and no resolution strategy makes a")
    print("   false rule true.")

    # the fix is to delete the over-general rule, not to tune the
    # conflict-resolution strategy
    print()
    print("=== delete R0 and ask both directions again ===")
    fixed = [r for r in RULES if r[0] != "R0"]
    fwd2, _ = forward_chain(TICKETS["T-4472"], rules=fixed)
    globals()["RULES"], saved = fixed, RULES
    up2 = backward_chain(("escalate", "yes"), TICKETS["T-4472"])
    down2 = backward_chain(("escalate", "no"), TICKETS["T-4472"])
    globals()["RULES"] = saved
    print("   forward says escalate =", fwd2["escalate"])
    print("   backward proves escalate = yes:", up2 is not None)
    assert fwd2["escalate"] == "no" and up2 is None and down2 is not None
    print("   They now agree. The explanation:")
    for line in explain(down2):
        print("      " + line)

    # --- what each direction costs --------------------------------
    fwd = len(results["T-4471"][1].fired)
    _, bwd = goals_visited(("escalate", "yes"), TICKETS["T-4471"])
    print()
    print("forward chaining : %d rules fired to derive everything" % fwd)
    print("backward chaining: %d goals visited to answer one question"
          % bwd)
    assert bwd < fwd
    print()
    print("Forward chaining also concluded notify_finance, runbook and")
    print("page_severity -- useful for a dashboard, wasted for a")
    print("single question.")

    # --- adding a rule changes nothing else -----------------------
    print()
    print("=== the operations lead changes the policy ===")
    extended = RULES + [("R21", [("escalate", "yes"),
                                 ("wake_budget", "spent"),
                                 ("tier", "platinum")],
                         ("action", "wake_on_call"))]
    wm_new, t_new = forward_chain(TICKETS["T-4473"], rules=extended)
    print("   'platinum may wake the engineer past the budget'")
    print("   T-4473 action was queue_for_morning, now:",
          wm_new["action"])
    print("   firing order:", " ".join(t_new.fired))
    assert wm_new["action"] == "wake_on_call"
    print("   One rule added. No existing rule was touched, and")
    print("   specificity put it ahead of R14 automatically.")
