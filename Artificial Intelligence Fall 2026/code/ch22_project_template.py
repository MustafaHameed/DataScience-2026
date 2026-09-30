# Chapter 22 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 22 -- the project template.

Replace the marked sections. Everything else is the shape every
project in this course should have: a PEAS specification you can
print, a baseline, a measure fixed in advance, a trace, an audit and
a reproducible run.
"""

import json
import os
import random

SEED = 0                      # every random choice, reproducible

# ===================================================================
# 1. PEAS -- fill this in, and print it in the report
# ===================================================================
PEAS = {
    "performance": "sessions placed; zero clashes; load spread",
    "environment": "sessions, rooms, staff, slots; static, discrete",
    "actuators": "assign(session, room, slot, staff); unassign",
    "sensors": "the full session list and all constraints",
}

# what would count as FAILURE -- written down BEFORE building
FAILURE_CONDITION = (
    "the baseline places as many sessions with a load spread within "
    "2 of ours"
)

# ===================================================================
# 2. The domain, as DATA -- never inside the algorithm
# ===================================================================
SESSIONS = ["ai-lab-1", "ai-lab-2", "ds-lab-1", "ml-lab-1"]
SLOTS = ["mon-am", "mon-pm", "tue-am"]
STAFF = ["ayesha", "bilal"]


def legal(assignment, session, slot, who):
    """The constraints. Reading this should tell a domain expert
    exactly what the system believes."""
    for (s, (sl, w)) in assignment.items():
        if sl == slot and w == who:
            return False              # one person, one slot
    return True


# ===================================================================
# 3. The BASELINE -- build this first, always
# ===================================================================
def baseline(sessions):
    """Greedy: first legal option, in the order given."""
    assignment = {}
    for s in sessions:
        for slot in SLOTS:
            for who in STAFF:
                if legal(assignment, s, slot, who):
                    assignment[s] = (slot, who)
                    break
            if s in assignment:
                break
    return assignment


# ===================================================================
# 4. YOUR SYSTEM -- replace this with the real method
# ===================================================================
def solve(sessions, trace):
    """Backtracking with minimum remaining values would go here.

    The template uses plain backtracking so the scaffold runs; the
    point is that it LOGS, and that the log is what the report,
    the demonstration and the audit are all built from.
    """
    assignment = {}

    def bt(remaining):
        if not remaining:
            return True
        # MRV would order 'remaining' here
        s = remaining[0]
        for slot in SLOTS:
            for who in STAFF:
                if legal(assignment, s, slot, who):
                    assignment[s] = (slot, who)
                    trace.append("assign %s -> %s, %s" % (s, slot, who))
                    if bt(remaining[1:]):
                        return True
                    trace.append("undo %s" % s)
                    del assignment[s]
        return False

    bt(list(sessions))
    return assignment


# ===================================================================
# 5. The MEASURE -- fixed in advance, and the same for both
# ===================================================================
def measure(assignment):
    placed = len(assignment)
    load = {w: 0 for w in STAFF}
    for (_, who) in assignment.values():
        load[who] += 1
    spread = max(load.values()) - min(load.values()) if load else 0
    return {"placed": placed, "spread": spread, "load": load}


# ===================================================================
# 6. The AUDIT -- Chapter 21
# ===================================================================
def audit(assignment):
    """Outcomes by group, and the situations not covered."""
    by_course = {}
    for s, (slot, who) in assignment.items():
        course = s.split("-")[0]
        by_course.setdefault(course, []).append(slot)
    unplaced = [s for s in SESSIONS if s not in assignment]
    return {"by_course": by_course, "unplaced": unplaced}


if __name__ == "__main__":
    random.seed(SEED)
    print("=" * 58)
    print("PEAS")
    print("=" * 58)
    for k, v in PEAS.items():
        print("   %-13s %s" % (k + ":", v))
    print()
    print("   fails if:    %s" % FAILURE_CONDITION)

    print()
    print("=" * 58)
    print("BASELINE vs SYSTEM")
    print("=" * 58)
    base = baseline(SESSIONS)
    trace = []
    ours = solve(SESSIONS, trace)

    mb, mo = measure(base), measure(ours)
    print("   measure      baseline   system")
    print("   " + "-" * 34)
    for key in ("placed", "spread"):
        print("   %-12s %8s %8s" % (key, mb[key], mo[key]))
    print("   load         %8s %8s" % (mb["load"], mo["load"]))

    # state the verdict against the CONDITION FIXED IN ADVANCE
    print()
    beat = (mo["placed"] > mb["placed"]
            or (mo["placed"] == mb["placed"]
                and mo["spread"] < mb["spread"] - 2))
    print("   verdict: %s"
          % ("the system beat the baseline"
             if beat else
             "the system did NOT beat the baseline on the measure "
             "fixed in advance -- report this"))

    print()
    print("=" * 58)
    print("TRACE (first 8 steps)")
    print("=" * 58)
    for line in trace[:8]:
        print("   " + line)
    print("   ... %d steps total" % len(trace))

    print()
    print("=" * 58)
    print("AUDIT")
    print("=" * 58)
    a = audit(ours)
    for course, slots in sorted(a["by_course"].items()):
        print("   %-6s %d session(s): %s"
              % (course, len(slots), ", ".join(slots)))
    print("   unplaced:", a["unplaced"] or "none")

    # reproducibility: the artefact the report cites
    out = {"seed": SEED, "peas": PEAS,
           "failure_condition": FAILURE_CONDITION,
           "baseline": mb, "system": mo, "audit": a,
           "trace_length": len(trace)}
    os.makedirs("project_output", exist_ok=True)
    with open(os.path.join("project_output", "metrics.json"), "w") as f:
        json.dump(out, f, indent=2, default=str)
    print()
    print("wrote project_output/metrics.json -- cite this in the")
    print("report, and commit it. A result nobody can reproduce is")
    print("not a result.")

    # the scaffold's own checks
    assert mo["placed"] == len(SESSIONS), "the system should place all"
    assert os.path.exists(os.path.join("project_output",
                                       "metrics.json"))
