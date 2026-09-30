# Chapter 14 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 14 lab -- ship now or delay a week.

Fractions throughout, so every printed number can be checked by hand
against the natural-frequency tree of Figure 14.1.
"""

from fractions import Fraction as F

# --- the smoke test -----------------------------------------------
P_DEFECT = F(1, 10)          # one release in ten has a serious defect
SENS = F(9, 10)              # P(alarm | defect)
FPR = F(2, 10)               # P(alarm | no defect)

# --- what each outcome is worth -----------------------------------
UTIL = {
    ("ship", "defect"): -100,    # incident, rollback, reputation
    ("ship", "ok"): 40,          # a week of revenue earned early
    ("delay", "defect"): -5,     # caught it; a small slip
    ("delay", "ok"): -10,        # a week lost for nothing
}
ACTIONS = ("ship", "delay")


def posterior(prior, sens, fpr, alarm=True):
    """Bayes' rule. P(defect | alarm) or P(defect | no alarm)."""
    like_d = sens if alarm else 1 - sens
    like_ok = fpr if alarm else 1 - fpr
    return (like_d * prior) / (like_d * prior + like_ok * (1 - prior))


def eu(action, p):
    return UTIL[(action, "defect")] * p + UTIL[(action, "ok")] * (1 - p)


def best(p):
    a = max(ACTIONS, key=lambda x: eu(x, p))
    return a, eu(a, p)


def vpi(p):
    """Value of perfect information about the defect."""
    informed = (p * max(UTIL[(a, "defect")] for a in ACTIONS)
                + (1 - p) * max(UTIL[(a, "ok")] for a in ACTIONS))
    return informed - best(p)[1]


def flip_point():
    """Solve EU(ship) = EU(delay) for p, exactly.

    EU(ship)  = Ud*p + Uo*(1-p)   = Uo + (Ud-Uo)p
    EU(delay) = Vd*p + Vo*(1-p)   = Vo + (Vd-Vo)p
    """
    ud, uo = UTIL[("ship", "defect")], UTIL[("ship", "ok")]
    vd, vo = UTIL[("delay", "defect")], UTIL[("delay", "ok")]
    return F(vo - uo, (ud - uo) - (vd - vo))


if __name__ == "__main__":
    print("=== the smoke test ===")
    print("P(defect)            =", P_DEFECT)
    print("P(alarm | defect)    =", SENS)
    print("P(alarm | no defect) =", FPR)

    post = posterior(P_DEFECT, SENS, FPR, alarm=True)
    quiet = posterior(P_DEFECT, SENS, FPR, alarm=False)
    print()
    print("P(defect | alarm)    =", post, "=", round(float(post), 4))
    print("P(defect | silence)  =", quiet, "=", round(float(quiet), 4))
    assert post == F(1, 3)

    # the natural-frequency check of Figure 14.1
    n = 1000
    defective = n * P_DEFECT                 # 100
    true_alarms = defective * SENS           # 90
    false_alarms = (n - defective) * FPR     # 180
    print()
    print("in %d releases: %d defective, %d true alarms, %d false"
          % (n, defective, true_alarms, false_alarms))
    print("   %d of %d alarms are real = %s"
          % (true_alarms, true_alarms + false_alarms,
             F(int(true_alarms), int(true_alarms + false_alarms))))
    assert F(int(true_alarms), int(true_alarms + false_alarms)) == post
    print("   a 90%-sensitive test fired, and the release is still")
    print("   TWICE as likely sound as broken.")

    # --- the decision ---------------------------------------------
    print()
    print("=== the decision ===")
    print("situation          P(defect)   EU(ship)  EU(delay)  choose")
    print("-" * 60)
    rows = (("no test yet   ", P_DEFECT),
            ("alarm fired   ", post),
            ("test is silent", quiet))
    for label, p in rows:
        a, _ = best(p)
        print("%s %10s %10.2f %10.2f  %s"
              % (label, round(float(p), 4), float(eu("ship", p)),
                 float(eu("delay", p)), a))

    star = flip_point()
    print()
    print("flip point p* =", star, "=", round(float(star), 4))
    assert star == F(10, 29)
    # just either side of it
    assert best(star - F(1, 1000))[0] == "ship"
    assert best(star + F(1, 1000))[0] == "delay"

    # the headline: belief changed a lot, decision did not change
    assert best(P_DEFECT)[0] == "ship"
    assert best(post)[0] == "ship"
    margin = eu("ship", post) - eu("delay", post)
    print()
    print("the alarm moved belief from %s to %s -- more than triple --"
          % (P_DEFECT, post))
    print("and the decision did not change, because %s < %s."
          % (post, star))
    print("the margin is now only %.2f utility units." % float(margin))
    assert float(margin) < 2

    # --- the value of information ---------------------------------
    print()
    print("=== value of perfect information ===")
    for label, p in rows + (("at the flip point", star),):
        print("   %-18s p=%-8s VPI = %6.2f"
              % (label.strip(), round(float(p), 4), float(vpi(p))))
    assert vpi(post) > 3 * vpi(P_DEFECT)
    print()
    print("Information is worth most where the decision is hardest:")
    print("%.2f before the test, %.2f after an ambiguous alarm."
          % (float(vpi(P_DEFECT)), float(vpi(post))))
    print("Delaying a week costs 10 when the release is sound, so an")
    print("investigation cheaper than that is clearly worth doing.")

    # --- and the utilities matter as much as the probabilities ----
    print()
    print("=== if shipping a defect cost 60 rather than 100 ===")
    UTIL[("ship", "defect")] = -60
    star2 = flip_point()
    print("   flip point moves from %s to %s (%.3f)"
          % (star, star2, float(star2)))
    print("   the alarm's 1/3 is now far from flipping anything:",
          best(post)[0])
    print("   VPI after an alarm falls from %.2f to %.2f"
          % (31.67, float(vpi(post))))
    assert star2 > star
    print()
    print("The utilities decide the answer as much as the")
    print("probabilities do -- and they are the part nobody measures.")
