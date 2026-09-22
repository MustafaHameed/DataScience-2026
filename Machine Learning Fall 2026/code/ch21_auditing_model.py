# Chapter 21 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split, cross_val_score

# --- a synthetic hiring dataset with a protected group -------------------
rng = np.random.default_rng(0)
n = 6000
group = rng.integers(0, 2, n)                     # 0 = group A, 1 = group B
# group B had fewer internships and projects to show (historical bias)
record = rng.normal(0.6 - 0.25 * group, 0.15, n)  # CV: internships, projects
aptitude = rng.normal(0.5, 0.15, n)               # coding test: equal
district = group + rng.normal(0, 0.3, n)          # a proxy for the group
qualified = (0.5 * record + 0.5 * aptitude
             + rng.normal(0, 0.05, n) > 0.5).astype(int)
X = np.column_stack([record, aptitude, district])
X_tr, X_te, y_tr, y_te, g_tr, g_te = train_test_split(
    X, qualified, group, test_size=0.5, random_state=0)


def report(name, pred):
    print(name)
    for g, label in ((0, "A"), (1, "B")):
        m = g_te == g
        sel = pred[m].mean()                               # selection rate
        tpr = pred[m][y_te[m] == 1].mean()                 # recall
        fpr = pred[m][y_te[m] == 0].mean()
        prec = y_te[m][pred[m] == 1].mean()
        print(f"  group {label}: base rate {y_te[m].mean():.2f}  "
              f"selected {sel:.2f}  TPR {tpr:.2f}  FPR {fpr:.2f}  "
              f"precision {prec:.2f}")


# --- 1. the model, group column excluded (it was never a feature) --------
model = LogisticRegression().fit(X_tr, y_tr)
p = model.predict_proba(X_te)[:, 1]
report("threshold 0.5 for everyone:", (p > 0.5).astype(int))

# --- 2. removing the group column does not remove the group --------------
acc = cross_val_score(LogisticRegression(), X, group, cv=5).mean()
print(f"the features recover group membership with accuracy {acc:.2f}")

# --- 3. equalise the true positive rate with a group-specific threshold --
t_b = 0.5
tpr_a = (p[(g_te == 0) & (y_te == 1)] > 0.5).mean()
while (p[(g_te == 1) & (y_te == 1)] > t_b).mean() < tpr_a:
    t_b -= 0.01
pred = np.where(g_te == 1, p > t_b, p > 0.5).astype(int)
report(f"threshold 0.5 for A, {t_b:.2f} for B (equal opportunity):", pred)
