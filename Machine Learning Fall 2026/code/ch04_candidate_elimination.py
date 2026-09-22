# Chapter 4 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
from itertools import product

# EnjoySport. "?" = any value, "0" = no value (the empty constraint)
VALUES = [("Sunny", "Cloudy", "Rainy"), ("Warm", "Cold"), ("Normal", "High"),
          ("Strong", "Weak"), ("Warm", "Cool"), ("Same", "Change")]
D = [(("Sunny", "Warm", "Normal", "Strong", "Warm", "Same"), True),
     (("Sunny", "Warm", "High", "Strong", "Warm", "Same"), True),
     (("Rainy", "Cold", "High", "Strong", "Warm", "Change"), False),
     (("Sunny", "Warm", "High", "Strong", "Cool", "Change"), True)]
show = lambda hs: "  ".join("<" + ", ".join(h) + ">" for h in sorted(hs))


def covers(h, x):             # does h classify x as positive?
    return all(c == "?" or c == v for c, v in zip(h, x))


def more_general(h1, h2):     # h1 >=g h2
    return "0" in h2 or all(a == "?" or a == b for a, b in zip(h1, h2))


def generalize(s, x):         # minimal generalisation of s covering x
    return tuple(v if c == "0" else (c if c == v else "?")
                 for c, v in zip(s, x))


def specializations(g, x):    # minimal specialisations of g excluding x
    return [g[:i] + (v,) + g[i + 1:] for i, c in enumerate(g) if c == "?"
            for v in VALUES[i] if v != x[i]]


# --- Find-S --------------------------------------------------------------
h = ("0",) * 6
for x, positive in D:
    if positive:
        h = generalize(h, x)
print("Find-S:", show([h]))

# --- Candidate-Elimination -----------------------------------------------
S, G = {("0",) * 6}, {("?",) * 6}
for k, (x, positive) in enumerate(D, 1):
    if positive:
        G = {g for g in G if covers(g, x)}
        S = {generalize(s, x) for s in S}
        S = {s for s in S if any(more_general(g, s) for g in G)}
    else:
        S = {s for s in S if not covers(s, x)}
        new = {g for g in G if not covers(g, x)}
        for g in G - new:
            new |= {h for h in specializations(g, x)
                    if any(more_general(h, s) for s in S)}
        G = {g for g in new
             if not any(o != g and more_general(o, g) for o in new)}
    print(f"day {k}  S: {show(S)}\n       G: {show(G)}")

# --- brute force: list every hypothesis, keep the consistent ones ---------
H = list(product(*[v + ("?",) for v in VALUES]))
VS = [h for h in H if all(covers(h, x) == p for x, p in D)]
print(len(VS), "hypotheses in the version space:")
print(show(VS))
day = ("Sunny", "Warm", "Normal", "Weak", "Warm", "Same")
yes = sum(covers(h, day) for h in VS)
print(f"vote on {day}: {yes} yes, {len(VS) - yes} no")
