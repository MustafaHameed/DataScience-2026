"""Chapter 28 - Measuring fairness, and demonstrating the impossibility result.

Fairness metrics are per-group confusion matrices. This script also shows,
numerically, that a calibrated model cannot achieve demographic parity when
base rates differ between groups.
"""
import numpy as np
import pandas as pd
from sklearn.metrics import confusion_matrix

rng = np.random.default_rng(3)


def fairness_report(y_true, y_pred, group):
    rows = []
    for g in pd.unique(group):
        m = group == g
        tn, fp, fn, tp = confusion_matrix(y_true[m], y_pred[m]).ravel()
        rows.append({
            "group": g,
            "n": int(m.sum()),
            "base_rate": round(float(y_true[m].mean()), 3),
            "flag_rate": round(float(y_pred[m].mean()), 3),   # demographic parity
            "recall": round(tp / (tp + fn), 3),               # equal opportunity
            "precision": round(tp / (tp + fp), 3),
            "fpr": round(fp / (fp + tn), 3),
        })
    df = pd.DataFrame(rows)
    print(df.to_string(index=False))
    print(f"  max recall gap    : {df.recall.max() - df.recall.min():.3f}")
    print(f"  max flag-rate gap : {df.flag_rate.max() - df.flag_rate.min():.3f}")
    return df


# Two groups with genuinely different base rates (20% and 40%).
n = 4000
group = np.array(["A"] * (n // 2) + ["B"] * (n // 2))
base = np.where(group == "A", 0.20, 0.40)
y = (rng.random(n) < base).astype(int)

# A well-calibrated score: P(y=1) is recovered honestly in both groups.
score = np.clip(base + rng.normal(0, 0.18, n), 0.01, 0.99)
score = np.where(y == 1, score + 0.22, score - 0.10)

print("=== A CALIBRATED MODEL AT A SINGLE THRESHOLD ===")
calibrated = fairness_report(y, (score > 0.35).astype(int), group)

print("\n=== FORCING EQUAL FLAG RATES (demographic parity) ===")
pred = np.zeros(n, dtype=int)
for g in ("A", "B"):
    m = group == g
    cut = np.percentile(score[m], 70)          # flag the top 30% of EACH group
    pred[m] = (score[m] > cut).astype(int)
forced = fairness_report(y, pred, group)

print("""
Read the two tables together:

  - The calibrated model has similar RECALL in both groups but different
    FLAG RATES, because group B genuinely contains twice as many cases.
  - Forcing equal flag rates equalises the second column and breaks the
    first: group B's at-risk students are left unflagged to meet the quota.

Neither table is 'the fair one'. With different base rates you must CHOOSE
which criterion the context demands, and say so in writing.""")
