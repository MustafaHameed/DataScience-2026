"""Chapter 6 - Simulating the base-rate fallacy.

The hand calculation in the chapter says a detector with 99% recall and a 1%
false-alarm rate produces alerts that are right under 1% of the time when
attacks are 1 in 10,000. Do not take that on trust -- generate a million
sessions and count.
"""
import numpy as np

rng = np.random.default_rng(42)


def simulate(n=1_000_000, base_rate=0.0001, detection=0.99, false_alarm=0.01):
    is_attack = rng.random(n) < base_rate
    p_alert = np.where(is_attack, detection, false_alarm)
    alert = rng.random(n) < p_alert

    n_alerts = alert.sum()
    n_true = (alert & is_attack).sum()
    return {
        "attacks": int(is_attack.sum()),
        "alerts": int(n_alerts),
        "true alerts": int(n_true),
        "precision": float(n_true / n_alerts) if n_alerts else float("nan"),
    }


print("--- as in the worked example ---")
for k, v in simulate().items():
    print(f"  {k:12s} {v}")

print("\n--- raising DETECTION from 99% to 99.9% barely helps ---")
print(f"  precision: {simulate(detection=0.999)['precision']:.4f}")

print("\n--- lowering FALSE ALARMS from 1% to 0.01% transforms it ---")
print(f"  precision: {simulate(false_alarm=0.0001)['precision']:.4f}")

print("""
The engineering priority that falls out of this:
in rare-event detection, effort spent reducing false alarms is worth
vastly more than effort spent raising the detection rate.""")
