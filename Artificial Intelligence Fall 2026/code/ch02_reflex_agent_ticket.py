# Chapter 2 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 2 lab -- a reflex agent and a model-based agent.

Two agent programs, the same percepts, the same performance measure.
The simple reflex agent cannot represent how often it has already woken
the on-call engineer, because no percept reports it. That single
missing piece of state is the whole difference.
"""

import random
from collections import namedtuple

random.seed(0)

# PEAS, as code. A percept is one ticket as the agent sees it.
Ticket = namedtuple("Ticket", "tid hour severity hours_to_breach")

WAKE_LIMIT = 3          # more than this and Ayesha is useless tomorrow
COST = {"major": 10, "minor": 3}
UNNECESSARY = 1
SLEEP = 4


def night(n=40):
    """One night's percept sequence: n tickets, 21:00 to 07:00."""
    out = []
    for tid in range(1, n + 1):
        out.append(Ticket(
            tid=tid,
            hour=21 + (tid * 10) // n,
            severity=random.choice(["major", "minor", "minor"]),
            hours_to_breach=random.choice([1, 2, 4, 8, 12]),
        ))
    return out


def simple_reflex(t, state):
    """Condition-action rules on the current percept only."""
    return t.hours_to_breach <= 4 or t.severity == "major"


def model_based(t, state):
    """Same rules, plus a model of what has already been spent.

    state["wakes"] is not in any percept: the agent maintains it.
    """
    if state["wakes"] >= WAKE_LIMIT:
        # budget gone: escalate only what will certainly breach badly
        return t.severity == "major" and t.hours_to_breach <= 2
    if t.hours_to_breach > 10:
        return False                    # cannot breach tonight; defer
    return t.hours_to_breach <= 4 or t.severity == "major"


def run(agent, tickets):
    """One night. The environment, not the agent, decides what an
    escalation achieves: past WAKE_LIMIT the engineer is too tired to
    fix anything, but the sleep is lost all the same.
    """
    state = {"wakes": 0}
    breach = sleep = waste = woke = 0
    for t in tickets:
        act = agent(t, state)
        fixed = False
        if act:
            state["wakes"] += 1
            woke += 1
            sleep += SLEEP
            if state["wakes"] <= WAKE_LIMIT:
                fixed = True            # still sharp enough to help
            if t.hours_to_breach > 10:
                waste += UNNECESSARY    # would not have breached tonight
        if not fixed and t.hours_to_breach <= 10:
            breach += COST[t.severity]
    return breach + sleep + waste, breach, sleep + waste, woke


if __name__ == "__main__":
    tickets = night()
    print("tickets:", len(tickets),
          " major:", sum(1 for t in tickets if t.severity == "major"),
          " wake limit:", WAKE_LIMIT)
    print()
    print("agent             total   breaches   on-call   escalations")
    print("-" * 58)
    rows = {}
    for name, agent in (("simple reflex", simple_reflex),
                        ("model-based  ", model_based)):
        rows[name.strip()] = run(agent, tickets)
        tot, br, oc, woke = rows[name.strip()]
        print("%-16s %6d %10d %9d %13d" % (name, tot, br, oc, woke))

    t1, b1, o1, w1 = rows["simple reflex"]
    t2, b2, o2, w2 = rows["model-based"]
    assert t2 < t1, "the model-based agent should cost less"
    assert w2 < w1, "and should wake the engineer less often"
    print()
    print("total cost  %d -> %d  (%.0f%% lower)" %
          (t1, t2, 100 * (t1 - t2) / t1))
    print("escalations %d -> %d  (%.1fx fewer)" % (w1, w2, w1 / w2))
    print()
    print("The model-based agent saw no extra percepts. It kept one")
    print("number -- how often it had already woken her -- that no")
    print("percept reports and the simple agent cannot represent.")
