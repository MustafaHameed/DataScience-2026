# Chapter 17 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Chapter 17 lab -- the perceptron, and where this course stops.

Two datasets and one proof. The first converges in three passes. The
second never converges, and the grid sweep at the end shows that this
is a fact about straight lines rather than about the training run.
"""

# twelve past tasks: (effort days, requirement clarity) -> late?
TASKS = [
    ((2, 9), 0), ((3, 8), 0), ((4, 9), 0), ((3, 7), 0),
    ((5, 8), 0), ((2, 6), 0),
    ((9, 3), 1), ((8, 2), 1), ((10, 4), 1), ((9, 1), 1),
    ((7, 2), 1), ((8, 4), 1),
]

# the exclusive-or arrangement: late when EXACTLY ONE of
# "large effort" and "low clarity" holds
XOR_TASKS = [
    ((2, 9), 0),   # small effort, high clarity -> on time
    ((9, 2), 0),   # large effort, low clarity  -> on time; they planned
    ((9, 9), 1),   # large effort, high clarity -> late; looked easy
    ((2, 2), 1),   # small effort, low clarity  -> late; looked small
]


def predict(x, w, b):
    return 1 if w[0] * x[0] + w[1] * x[1] + b > 0 else 0


def accuracy(data, w, b):
    return sum(predict(x, w, b) == y for x, y in data) / len(data)


def train(data, eta=0.1, epochs=100, trace=False):
    """The perceptron learning rule. Learns only from mistakes."""
    w, b, updates = [0.0, 0.0], 0.0, 0
    for ep in range(1, epochs + 1):
        errors = 0
        for (x, y) in data:
            err = y - predict(x, w, b)
            if err:
                errors += 1
                updates += 1
                w[0] += eta * err * x[0]
                w[1] += eta * err * x[1]
                b += eta * err
        if trace:
            print("   epoch %2d: %d error(s), w = (%.2f, %.2f), b = %.2f"
                  % (ep, errors, w[0], w[1], b))
        if errors == 0:
            return w, b, ep, updates, True
    return w, b, epochs, updates, False


def best_line_accuracy(data, span=20, step=0.5):
    """Brute force over a grid of ALL lines w0*x + w1*y + b = 0.

    This is the difference between 'training failed' and 'no line
    exists'. It is only possible because the problem is tiny.
    """
    best, arg = 0.0, None
    vals = [i * step for i in range(-span * 2, span * 2 + 1)]
    for w0 in vals:
        for w1 in vals:
            for bb in vals:
                a = accuracy(data, [w0, w1], bb)
                if a > best:
                    best, arg = a, (w0, w1, bb)
    return best, arg


if __name__ == "__main__":
    print("=== twelve past tasks ===")
    w, b, epochs, updates, converged = train(TASKS, trace=True)
    print("   converged: %s, after %d epochs and %d weight updates"
          % (converged, epochs, updates))
    print("   learned rule:  %.2f*effort %+.2f*clarity %+.2f > 0"
          % (w[0], w[1], b))
    print("   accuracy on the training data: %.0f%%"
          % (100 * accuracy(TASKS, w, b)))
    assert converged and epochs == 3 and updates == 2
    assert accuracy(TASKS, w, b) == 1.0
    # the clarity weight went NEGATIVE: more clarity argues against late
    assert w[1] < 0
    print()
    print("   The weight on clarity is negative. Nobody said that;")
    print("   it came from one misclassified example in pass two.")
    print()
    print("   But 100% here is accuracy on the data it LEARNED from.")
    print("   Whether it is right about the next task is a different")
    print("   question, and this course cannot answer it.")

    print()
    print("=== the exclusive-or arrangement ===")
    for (x, y) in XOR_TASKS:
        print("   effort %2d, clarity %2d -> %s"
              % (x[0], x[1], "late" if y else "on time"))

    print()
    print("   epochs allowed   converged   weight updates")
    print("   " + "-" * 44)
    for budget in (100, 1000, 10000):
        _, _, _, ups, conv = train(XOR_TASKS, epochs=budget)
        print("   %14d   %9s   %13d" % (budget, conv, ups))
    _, _, _, _, conv = train(XOR_TASKS, epochs=10000)
    assert not conv
    print()
    print("   The update count just grows with the budget: the")
    print("   weights are still oscillating when time runs out.")
    print("   Nothing here says 'impossible' -- it looks exactly")
    print("   like slow convergence.")

    # so settle it by brute force
    print()
    print("=== can ANY straight line do it? ===")
    best, arg = best_line_accuracy(XOR_TASKS)
    print("   best accuracy over a grid of every line: %.0f%%"
          % (100 * best))
    print("   achieved by %.1f*effort %+.1f*clarity %+.1f > 0"
          % arg)
    assert best == 0.75
    print()
    print("   Three of four, and no line does better. This is not a")
    print("   training failure -- it is a fact about the hypothesis")
    print("   class. No learning rate, initialisation or patience")
    print("   fixes it; only a richer class does, and a hidden layer")
    print("   is exactly that.")

    # and the same sweep on the separable data, for contrast
    best2, _ = best_line_accuracy(TASKS)
    assert best2 == 1.0
    print()
    print("   (The same sweep on the twelve tasks finds a line with")
    print("   100% accuracy, which is why the perceptron found one.)")

    print()
    print("=" * 58)
    print("Machine Learning Fall 2026, Chapter 13, takes over here.")
    print("=" * 58)
