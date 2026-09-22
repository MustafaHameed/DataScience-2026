# Chapter 8 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from scipy.stats import chi2
from sklearn.datasets import make_classification
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import (confusion_matrix, precision_score, recall_score,
                             f1_score, roc_auc_score, average_precision_score)

# --- an imbalanced problem: 5% positives ---------------------------------
X, y = make_classification(n_samples=6000, n_features=12, n_informative=5,
                           weights=[0.95], flip_y=0.01, random_state=1)
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.4,
                                          stratify=y, random_state=1)
print("baseline accuracy (always 0):", round(1 - y_te.mean(), 3))
lr = LogisticRegression(max_iter=1000).fit(X_tr, y_tr)
p = lr.predict_proba(X_te)[:, 1]

# --- 1. the four cells and the metrics built from them -------------------
pred = (p > 0.5).astype(int)
tn, fp, fn, tp = confusion_matrix(y_te, pred).ravel()
print(f"TP={tp} FP={fp} FN={fn} TN={tn}")
print(f"accuracy {np.mean(pred == y_te):.3f}  "
      f"precision {precision_score(y_te, pred):.3f}  "
      f"recall {recall_score(y_te, pred):.3f}  F1 {f1_score(y_te, pred):.3f}")
print(f"ROC-AUC {roc_auc_score(y_te, p):.3f}   "
      f"PR-AUC {average_precision_score(y_te, p):.3f}")

# --- 2. the threshold is a choice: sweep it ------------------------------
for t in (0.1, 0.3, 0.5, 0.7):
    pr = (p > t).astype(int)
    print(f"threshold {t}: precision {precision_score(y_te, pr):.2f}  "
          f"recall {recall_score(y_te, pr):.2f}")

# --- 3. a 95% confidence interval for the error rate ---------------------
n, e = len(y_te), np.mean(pred != y_te)
half = 1.96 * np.sqrt(e * (1 - e) / n)
print(f"error {e:.3f}, 95% interval [{e - half:.3f}, {e + half:.3f}]")

# --- 4. is a decision tree really different? McNemar's test --------------
tree = DecisionTreeClassifier(max_depth=5, random_state=0).fit(X_tr, y_tr)
a_right = pred == y_te
b_right = tree.predict(X_te) == y_te
b = np.sum(~a_right & b_right)          # logistic wrong, tree right
c = np.sum(a_right & ~b_right)          # logistic right, tree wrong
stat = (abs(b - c) - 1) ** 2 / (b + c)
print(f"disagreements b={b}, c={c}; McNemar chi2 = {stat:.2f}, "
      f"p = {chi2.sf(stat, 1):.3f}")
