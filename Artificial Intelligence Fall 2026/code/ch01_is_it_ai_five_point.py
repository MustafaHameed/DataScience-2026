# Chapter 1 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 1 lab -- Is it AI? A five-point audit.

Score five programs from a project management office against the five
tests of Section 1.3, and classify each one. Nothing here learns
anything; the exercise is to separate programs that decide from
programs that compute.
"""

TESTS = ("goal", "search", "knowledge", "adapts", "explains")

# One row per tool: which of the five tests it passes, and why the
# most interesting answer is what it is.
TOOLS = {
    "Gantt recalculator": {
        "passes": (),
        "note": "given a procedure, not an objective; one answer, "
                "computed directly",
    },
    "Burndown chart generator": {
        "passes": (),
        "note": "arithmetic and a plot; no alternatives considered",
    },
    "Sprint scheduler": {
        "passes": ("goal", "search", "knowledge", "explains"),
        "note": "constraints are declared, not coded; reports which "
                "constraint blocked a story  (Chapter 7)",
    },
    "Escalation adviser": {
        "passes": ("goal", "search", "knowledge", "explains"),
        "note": "rules an operations lead can read and correct; prints "
                "its firing chain  (Chapter 13)",
    },
    "Duplicate-ticket finder": {
        "passes": ("adapts",),
        "note": "weights retuned monthly from past merges; the one "
                "tool here that belongs to Machine Learning",
    },
}


def verdict(score):
    """The reading of a score given in Section 1.3."""
    if score == 0:
        return "ordinary software"
    if score <= 2:
        return "doing something interesting"
    return "an AI system"


def audit():
    width = max(len(name) for name in TOOLS)
    head = "TOOL".ljust(width) + "  " + " ".join(
        t[:4].upper().rjust(5) for t in TESTS) + "   SCORE  VERDICT"
    print(head)
    print("-" * len(head))
    for name, row in TOOLS.items():
        marks = ["yes".rjust(5) if t in row["passes"] else "-".rjust(5)
                 for t in TESTS]
        score = len(row["passes"])
        print(name.ljust(width) + "  " + " ".join(marks) +
              "   %5d  %s" % (score, verdict(score)))
    print()
    for name, row in TOOLS.items():
        print("%s\n    %s" % (name, row["note"]))


def check():
    """The two claims the chapter makes about this table."""
    scores = {n: len(r["passes"]) for n, r in TOOLS.items()}
    # Adaptation is the only test the AI systems here fail ...
    for name in ("Sprint scheduler", "Escalation adviser"):
        assert "adapts" not in TOOLS[name]["passes"], name
        assert scores[name] == 4, name
    # ... and the only test the machine learning tool passes.
    assert TOOLS["Duplicate-ticket finder"]["passes"] == ("adapts",)
    print("\nchecks passed: intelligent behaviour here does not "
          "require learning,")
    print("and the one tool that learns is the one this course "
          "hands on.")


if __name__ == "__main__":
    audit()
    check()
