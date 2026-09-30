# Chapter 18 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 18 lab -- symbolic language processing on ticket subjects.

A CYK chart parser, a pattern-substitution responder in the style of
Eliza, and a bigram model. Nothing here learns; the knowledge is in
the grammar and the patterns.
"""

import re
from collections import defaultdict

# --- the grammar, in Chomsky normal form ---------------------------
RULES = [
    ("S", ("VP",)),                 # imperative: "restart the server"
    ("S", ("NP", "VP")),            # statement:  "the server crashed"
    ("VP", ("V", "NP")),
    ("VP", ("VP", "PP")),           # PP attaches to the verb phrase
    ("NP", ("Det", "N")),
    ("NP", ("NP", "PP")),           # ... or to the noun phrase
    ("PP", ("P", "NP")),
    ("N", ("N", "N")),              # noun compound: "payments gateway"
]
LEXICON = {
    "restart": ["V"], "check": ["V"], "fix": ["V"],
    "server": ["N"], "payments": ["N"], "login": ["N"],
    "error": ["N"], "timeout": ["N"], "gateway": ["N"],
    "the": ["Det"], "a": ["Det"],
    "on": ["P"], "after": ["P"], "in": ["P"],
}


def cyk(words):
    """Fill the triangular chart. Returns (chart, back-pointers)."""
    n = len(words)
    chart = defaultdict(set)
    back = defaultdict(list)

    def close(cell):
        """Apply unit rules such as S -> VP until nothing changes."""
        changed = True
        while changed:
            changed = False
            for (lhs, rhs) in RULES:
                if len(rhs) == 1 and rhs[0] in chart[cell] \
                        and lhs not in chart[cell]:
                    chart[cell].add(lhs)
                    changed = True

    for i, w in enumerate(words):
        for tag in LEXICON.get(w, []):
            chart[(i, i + 1)].add(tag)
        close((i, i + 1))

    for span in range(2, n + 1):
        for i in range(0, n - span + 1):
            j = i + span
            for k in range(i + 1, j):
                for (lhs, rhs) in RULES:
                    if len(rhs) != 2:
                        continue
                    b, c = rhs
                    if b in chart[(i, k)] and c in chart[(k, j)]:
                        chart[(i, j)].add(lhs)
                        back[(i, j, lhs)].append((k, b, c))
            close((i, j))
    return chart, back


def count_parses(chart, back, i, j, sym, memo=None):
    """Count DERIVATIONS, not membership: an ambiguous sentence
    reaches the same cell by two different routes."""
    if memo is None:
        memo = {}
    key = (i, j, sym)
    if key in memo:
        return memo[key]
    total = 0
    for (k, b, c) in back[key]:
        total += (count_parses(chart, back, i, k, b, memo)
                  * count_parses(chart, back, k, j, c, memo))
    for (lhs, rhs) in RULES:                     # unit rules
        if lhs == sym and len(rhs) == 1 and rhs[0] != sym \
                and rhs[0] in chart[(i, j)]:
            total += count_parses(chart, back, i, j, rhs[0], memo)
    if total == 0 and j == i + 1 and sym in chart[(i, j)]:
        total = 1
    memo[key] = total
    return total


def parse(subject):
    words = subject.split()
    chart, back = cyk(words)
    ok = "S" in chart[(0, len(words))]
    n = count_parses(chart, back, 0, len(words), "S") if ok else 0
    return ok, n, chart


# --- Eliza: pattern substitution, and nothing else ------------------
REFLECT = {"my": "your", "i": "you", "me": "you", "am": "are",
           "mine": "yours", "myself": "yourself"}

PATTERNS = [
    (r"^i am (.*)$", "How long have you been {0}?"),
    (r"^my (\w+) keeps (.*)$", "Tell me more about your {0}."),
    (r"^because (.*)$", "Is that the real reason?"),
    (r"^i (?:can't|cannot) (.*)$", "What stops you from {0}?"),
    (r"^(.*) is broken$", "What makes you say {0} is broken?"),
]
FALLBACK = "Can you tell me more about that?"


def reflect(text):
    return " ".join(REFLECT.get(w, w) for w in text.split())


def eliza(line):
    low = line.lower().strip().rstrip(".!?")
    for (pat, template) in PATTERNS:
        m = re.match(pat, low)
        if m:
            return template.format(*[reflect(g) for g in m.groups()])
    return FALLBACK


# --- a bigram model -------------------------------------------------
CORPUS = [
    "restart the server", "restart the gateway", "restart the service",
    "check the payments gateway", "check the login error",
    "check the server", "fix the login error", "fix the timeout error",
    "restart the payments gateway", "check the timeout",
]


def bigrams(corpus):
    counts = defaultdict(lambda: defaultdict(int))
    for line in corpus:
        toks = ["<s>"] + line.split()
        for a, b in zip(toks, toks[1:]):
            counts[a][b] += 1
    return counts


def suggest(counts, word, k=3):
    nxt = counts.get(word, {})
    total = sum(nxt.values())
    ranked = sorted(nxt.items(), key=lambda kv: -kv[1])[:k]
    return [(w, c / total) for w, c in ranked]


SUBJECTS = [
    "restart the server",
    "check the payments gateway",
    "restart the server after the timeout",
    "fix the login error on the gateway",
    "restart the gateway after the payments error",
]

if __name__ == "__main__":
    print("=== parsing ticket subjects ===")
    print("subject                                         parses  count")
    print("-" * 62)
    counts = {}
    for s in SUBJECTS:
        ok, n, _ = parse(s)
        counts[s] = n
        print("%-46s %7s %6d" % (s, ok, n))
    # three of the five are structurally ambiguous
    ambiguous = [s for s in SUBJECTS if counts[s] > 1]
    assert len(ambiguous) == 3 and all(counts[s] == 2 for s in ambiguous)
    print()
    print("%d of %d subjects have more than one parse."
          % (len(ambiguous), len(SUBJECTS)))
    print("Each extra prepositional phrase roughly doubles the count.")

    print()
    print("=== and one that is not in the language ===")
    for bad in ("the the server restart", "server the",
                "restart restart the"):
        ok, n, _ = parse(bad)
        print("   %-32s parses: %s" % (bad, ok))
        assert not ok
    print("   Returning nothing is a useful answer: the system knows")
    print("   that it does not know.")

    print()
    print("=== the chart ===")
    words = SUBJECTS[4].split()
    chart, _ = cyk(words)
    cells = sum(1 for k in chart if chart[k])
    total = len(words) * (len(words) + 1) // 2
    print("   %d words -> %d cells, %d of them non-empty"
          % (len(words), total, cells))

    print()
    print("=== Eliza ===")
    for line in ("I am worried about the release",
                 "my deploy keeps failing",
                 "because the config is wrong",
                 "I can't reproduce it",
                 "the gateway is broken",
                 "no idea what is happening"):
        print("   > %-34s %s" % (line, eliza(line)))
    print()
    print("   Pattern matching and pronoun swapping. No grammar, no")
    print("   representation of the conversation, no model of")
    print("   anything -- and it reads as though it is listening.")
    assert eliza("something unmatched entirely") == FALLBACK

    print()
    print("=== a bigram model ===")
    bg = bigrams(CORPUS)
    for w in ("<s>", "restart", "the", "login"):
        print("   after %-10s -> %s"
              % (w, ", ".join("%s %.2f" % t for t in suggest(bg, w))))

    print()
    print("=== the limitation no corpus size removes ===")
    hard = ("the server that the engineer who wrote the runbook "
            "maintained was restarted")
    toks = hard.split()
    print("   '%s'" % hard)
    print("   'was restarted' is at position %d; its subject 'server'"
          % toks.index("was"))
    print("   is at position %d -- %d words back."
          % (toks.index("server"), toks.index("was") - toks.index("server")))
    print("   No n-gram with a practical n connects them, and raising")
    print("   n makes sparsity exponentially worse. A grammar has no")
    print("   difficulty with it, because the dependency is structural.")
    assert toks.index("was") - toks.index("server") >= 8
