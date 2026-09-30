# Chapter 10 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 10 lab -- a semantic network with frames and defaults.

Inheritance is one upward walk. The interesting cases are the third
query, where a default is correctly suppressed by a more specific one,
and the fourth, where two equally specific parents disagree and the
only defensible answer is to say so.
"""

# frame -> {slot: value}. 'is-a' and 'instance-of' are the links the
# inheritance walk follows; everything else is ordinary knowledge.
FRAMES = {
    "Artefact": {
        "owned-by": "Delivery",
    },
    "Release": {
        "is-a": ["Artefact"],
        "needs-signoff": "yes",          # default
        "regression-pass": "full",       # default
        "window": "Thursday",            # default
    },
    "Hotfix": {
        "is-a": ["Release"],
        "regression-pass": "smoke",      # overrides Release
        "window": "any",                 # overrides Release
    },
    "Regulated Change": {
        "is-a": ["Release"],
        "needs-signoff": "yes",
        "window": "Thursday",
    },
    "Release 4.7": {
        "instance-of": ["Hotfix"],
        "changes": "Payments",
    },
    "Developer": {},
    "Ayesha": {"instance-of": ["Developer"], "owns": "Payments"},
}

LINKS = ("instance-of", "is-a")


def parents(frame):
    out = []
    for link in LINKS:
        out += FRAMES.get(frame, {}).get(link, [])
    return out


def ancestors(frame):
    """Every class above this one, nearest first. The seen set is what
    stops an is-a cycle from looping forever."""
    out, seen, queue = [], {frame}, list(parents(frame))
    while queue:
        f = queue.pop(0)
        if f in seen:
            continue
        seen.add(f)
        out.append(f)
        queue += parents(f)
    return out


def get(frame, slot):
    """Most specific value wins. Returns (value, where it came from),
    or a conflict if two equally distant sources disagree."""
    if slot in FRAMES.get(frame, {}):
        return FRAMES[frame][slot], frame

    # breadth-first by distance, so 'equally specific' is well defined
    level, seen = list(parents(frame)), {frame}
    distance = 1
    while level:
        hits = []
        for f in level:
            if f in seen:
                continue
            if slot in FRAMES.get(f, {}):
                hits.append((FRAMES[f][slot], f))
        values = {v for v, _ in hits}
        if len(values) > 1:
            return ("CONFLICT", sorted(hits)), distance
        if hits:
            return hits[0][0], hits[0][1]
        seen |= set(level)
        nxt = []
        for f in level:
            nxt += parents(f)
        level = [f for f in nxt if f not in seen]
        distance += 1
    return None, None


def isa_chain(frame):
    return [frame] + ancestors(frame)


if __name__ == "__main__":
    print("inheritance chain for Release 4.7:")
    print("   " + " -> ".join(isa_chain("Release 4.7")))
    assert "Artefact" in ancestors("Release 4.7")
    print()

    print("queries, with the frame each answer came from")
    print("-" * 56)
    for slot in ("regression-pass", "needs-signoff", "window",
                 "changes", "owned-by"):
        val, src = get("Release 4.7", slot)
        print("   %-16s = %-10s  (from %s)" % (slot, val, src))

    # the default was correctly suppressed
    assert get("Release 4.7", "regression-pass") == ("smoke", "Hotfix")
    assert get("Release", "regression-pass") == ("full", "Release")
    # and a default two levels up is still found
    assert get("Release 4.7", "needs-signoff") == ("yes", "Release")
    print()
    print("'smoke' overrode 'full'; 'yes' was still reached two levels up.")

    # --- now add a second, equally specific parent ----------------
    print()
    print("adding 'Release 4.7 is-a Regulated Change' ...")
    FRAMES["Hotfix"]["needs-signoff"] = "no"      # delivery convention
    FRAMES["Release 4.7"]["is-a"] = ["Regulated Change"]

    val, dist = get("Release 4.7", "needs-signoff")
    print("   needs-signoff =", val[0] if isinstance(val, tuple) else val)
    assert isinstance(val, tuple) and val[0] == "CONFLICT"
    for v, src in val[1]:
        print("      %-4s from %s" % (v, src))
    print()
    print("Two parents at the same distance disagree. Depth-first")
    print("inheritance would answer whichever link was entered first,")
    print("which is a property of the file and not of the knowledge.")
    print("Here that would wave a regulated change through with no")
    print("sign-off, so the only defensible answer is to report the")
    print("conflict and its provenance.")

    # non-monotonicity, demonstrated: adding knowledge REMOVED an answer
    print()
    print("Note what just happened: adding a fact removed a")
    print("conclusion. Classical logic cannot do that, which is")
    print("exactly the trade Chapters 11 and 12 make.")
