# Chapter 20 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 20 lab -- a tool-using agent with no model in it.

The architecture of a modern agent: a goal, a set of tools, a loop
that decides what to call next, a step budget and a termination test.
The decision procedure here is Chapter 13's rule engine rather than a
language model, so every run is reproducible and every step has a
reason. No network, no API, no randomness.
"""

# ---- the world the tools read --------------------------------------
TASKS = {
    "auth-refresh":  {"points": 5, "depends": ["schema-change"]},
    "schema-change": {"points": 3, "depends": []},
    "ui-polish":     {"points": 2, "depends": []},
}
SPRINT_CAPACITY = 8


# ---- three tools. Each is cheap, local and honest about failure ----
def tool_estimate(task):
    t = TASKS.get(task)
    return t["points"] if t else None


def tool_dependencies(task):
    t = TASKS.get(task)
    return list(t["depends"]) if t else None


def tool_capacity(_):
    return SPRINT_CAPACITY


TOOLS = {
    "estimate": tool_estimate,
    "dependencies": tool_dependencies,
    "capacity": tool_capacity,
}


# ---- the decision procedure: Chapter 13's rule engine ---------------
# (name, [conditions on what we know], action)
# an action is ("call", tool, argument) or ("answer",)
POLICY = [
    ("need-deps",
     lambda k: "deps" not in k,
     lambda k: ("call", "dependencies", k["goal"])),
    # a tool can fail. A policy with no rule for that is the commonest
    # defect in deployed agents -- see the second experiment below.
    ("unknown-task",
     lambda k: k.get("deps", []) is None,
     lambda k: ("give-up",)),
    ("need-dep-estimates",
     lambda k: k.get("deps") and any(d not in k["points"]
                                     for d in k["deps"]),
     lambda k: ("call", "estimate",
                next(d for d in k["deps"] if d not in k["points"]))),
    ("need-own-estimate",
     lambda k: "deps" in k and k["goal"] not in k["points"],
     lambda k: ("call", "estimate", k["goal"])),
    ("need-capacity",
     lambda k: "capacity" not in k,
     lambda k: ("call", "capacity", None)),
    ("have-everything",
     lambda k: True,
     lambda k: ("answer",)),
]

BUDGET = 8


def agent(goal, policy=None, trace=True):
    """The loop of Figure 20.1. Keeps state, holds a goal, selects
    actions -- Chapter 2's model-based, goal-based agent."""
    policy = POLICY if policy is None else policy
    known = {"goal": goal, "points": {}}
    steps = []
    for step in range(1, BUDGET + 1):
        # MATCH and RESOLVE: first rule whose conditions hold
        action, name = None, None
        for (nm, cond, act) in policy:
            if cond(known):
                action, name = act(known), nm
                break
        if action is None:
            return None, steps, "no rule applies"
        if action[0] == "answer":
            steps.append((step, name, "answer", None))
            break
        if action[0] == "give-up":
            steps.append((step, name, "give up", None))
            return None, steps, "unknown task: %s" % goal
        _, tool, arg = action
        result = TOOLS[tool](arg)
        steps.append((step, name, "%s(%s)" % (tool, arg), result))
        # ACT: fold the observation into what we know
        if tool == "dependencies":
            known["deps"] = result
        elif tool == "estimate":
            known["points"][arg] = result
        elif tool == "capacity":
            known["capacity"] = result
    else:
        return None, steps, "budget exhausted"

    total = sum(known["points"].values())
    fits = total <= known["capacity"]
    answer = ("%s plus its dependencies is %d points against a "
              "capacity of %d: %s"
              % (goal, total, known["capacity"],
                 "it fits" if fits else "it does NOT fit"))
    return answer, steps, None


if __name__ == "__main__":
    print("goal: can we take 'auth-refresh' into the next sprint?")
    print()
    answer, steps, err = agent("auth-refresh")
    print("step  rule                 action                   result")
    print("-" * 66)
    for (n, rule, action, result) in steps:
        print("%4d  %-20s %-24s %s" % (n, rule, action, result))
    print()
    print("ANSWER:", answer)
    assert err is None
    # 5 points for auth-refresh, 3 for the dependency it pulled in,
    # against a capacity of 8: it fits exactly, with nothing spare
    assert "it fits" in answer

    # every step has a named reason, and the run is reproducible
    a2, s2, _ = agent("auth-refresh")
    assert [x[1] for x in s2] == [x[1] for x in steps]
    print()
    print("Run it again and the trace is identical: the decision")
    print("procedure is a rule base, so the agent is reproducible")
    print("and every step names the rule that caused it.")

    # --- a tool that fails, handled --------------------------------
    print()
    print("=== when a tool returns nothing ===")
    a3, s3, err3 = agent("does-not-exist")
    print("   goal 'does-not-exist' ->", err3)
    print("   rule that caught it:", s3[-1][1])
    assert err3.startswith("unknown task")
    print("   One rule, 'unknown-task', turns a failed tool call into")
    print("   an honest answer rather than a crash.")

    # --- and what happens without it -------------------------------
    print()
    print("=== the same agent with that rule removed ===")
    careless = [r for r in POLICY if r[0] != "unknown-task"]
    try:
        a4, s4, err4 = agent("does-not-exist", policy=careless)
        print("   result:", err4 or a4)
    except TypeError as exc:
        print("   TypeError: %s" % exc)
        print("   The tool returned None, no rule noticed, and the")
        print("   None travelled all the way to the arithmetic at the")
        print("   end. This is the commonest failure in deployed agent")
        print("   systems, and note where it lives: in the POLICY, not")
        print("   in the model at the centre.")

    # the step budget is the other guard
    print()
    print("=== what the step budget is for ===")
    spinner = [("always-ask", lambda k: True,
                lambda k: ("call", "capacity", None))]
    a5, s5, err5 = agent("auth-refresh", policy=spinner)
    print("   a policy that never terminates ->", err5)
    print("   it ran %d steps and was stopped by the budget." % len(s5))
    assert err5 == "budget exhausted" and len(s5) == BUDGET

    # --- what is symbolic and what would be learned ----------------
    print()
    print("=== which part would a language model replace? ===")
    print("   the LOOP            -- no, ordinary engineering")
    print("   the TOOLS           -- no, they read a database")
    print("   the STEP BUDGET     -- no, and it is load-bearing")
    print("   the TERMINATION     -- no")
    print("   the DECISION RULE   -- YES, this is the part")
    print()
    print("   Five of the six components are unchanged. Most of what")
    print("   makes an agent work or fail lives in the five.")
