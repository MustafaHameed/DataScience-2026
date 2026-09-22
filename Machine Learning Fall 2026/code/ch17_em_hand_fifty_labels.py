# Chapter 17 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.datasets import make_classification
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.semi_supervised import SelfTrainingClassifier
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

# --- 1. EM for two Gaussian means (sigma = 1 known, equal weights) -------
x = np.array([1.0, 2.0, 3.0, 7.0, 8.0, 9.0])
mu = np.array([2.0, 5.0])                        # initial guesses
for it in range(1, 5):
    dens = np.exp(-0.5 * (x[:, None] - mu[None, :]) ** 2)   # E-step
    r = dens / dens.sum(axis=1, keepdims=True)    # responsibilities
    ll = np.log(dens.mean(axis=1) / np.sqrt(2 * np.pi)).sum()
    mu = (r * x[:, None]).sum(axis=0) / r.sum(axis=0)       # M-step
    print(f"iteration {it}: log-likelihood {ll:8.3f} -> mu = {mu.round(3)}")

# --- 2. two labels, four unlabelled points -------------------------------
labelled = {2: 0, 3: 1}                 # x=3 is class A (0), x=7 class B (1)
mu = np.array([x[2], x[3]])             # labelled points alone: 3 and 7
print("from the two labels alone: mu =", mu)
for it in range(10):
    dens = np.exp(-0.5 * (x[:, None] - mu[None, :]) ** 2)
    r = dens / dens.sum(axis=1, keepdims=True)
    for i, c in labelled.items():       # a labelled point keeps its label
        r[i] = np.eye(2)[c]
    mu = (r * x[:, None]).sum(axis=0) / r.sum(axis=0)
print("EM with the unlabelled points too: mu =", mu.round(3))

# --- 3. self-training with 50 labels out of 1,260 ------------------------
X, y = make_classification(n_samples=1800, n_features=10, n_informative=6,
                           n_redundant=2, n_classes=4, n_clusters_per_class=1,
                           class_sep=2.0, random_state=6)   # 4 root causes
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.3,
                                          random_state=0, stratify=y)
rng = np.random.default_rng(0)
keep = rng.choice(len(y_tr), size=50, replace=False)
y_part = np.full_like(y_tr, -1)                 # -1 means "unlabelled"
y_part[keep] = y_tr[keep]
base = make_pipeline(StandardScaler(), LogisticRegression())
only = base.fit(X_tr[keep], y_tr[keep]).score(X_te, y_te)
self_t = SelfTrainingClassifier(
    make_pipeline(StandardScaler(), LogisticRegression()),
    threshold=0.9).fit(X_tr, y_part)
print(f"50 labels only: {only:.3f}   self-training with the rest "
      f"unlabelled: {self_t.score(X_te, y_te):.3f}   "
      f"all {len(y_tr)} labels: "
      f"{base.fit(X_tr, y_tr).score(X_te, y_te):.3f}")
