# Chapter 14 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.datasets import make_classification
from sklearn.model_selection import cross_val_score
from sklearn.tree import DecisionTreeClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import (BaggingClassifier, RandomForestClassifier,
                              AdaBoostClassifier, GradientBoostingClassifier,
                              VotingClassifier)

# --- 1. one round of AdaBoost, by hand -----------------------------------
w = np.full(10, 0.1)                            # ten examples, equal weight
wrong = np.array([1, 1, 1, 0, 0, 0, 0, 0, 0, 0], dtype=bool)
eps = w[wrong].sum()                            # weighted error = 0.3
alpha = 0.5 * np.log((1 - eps) / eps)           # say of this learner
w = np.where(wrong, w * np.exp(alpha), w * np.exp(-alpha))
w = w / w.sum()
print(f"alpha = {alpha:.3f}; new weights: wrong {w[0]:.4f}, "
      f"right {w[-1]:.4f}; weight on the mistakes = {w[wrong].sum():.2f}")

# --- 2. one tree against five ensembles ----------------------------------
X, y = make_classification(n_samples=3000, n_features=20, n_informative=8,
                           n_redundant=4, flip_y=0.05, random_state=7)
tree = DecisionTreeClassifier
models = {
    "single tree": tree(random_state=0),
    "bagging (100 trees)": BaggingClassifier(tree(), n_estimators=100,
                                             random_state=0),
    "random forest (100)": RandomForestClassifier(n_estimators=100,
                                                  random_state=0),
    "AdaBoost (depth 2)": AdaBoostClassifier(tree(max_depth=2),
                                             n_estimators=200,
                                             random_state=0),
    "gradient boosting": GradientBoostingClassifier(n_estimators=200,
                                                    random_state=0),
    "vote: LR + tree + RF": VotingClassifier([
        ("lr", LogisticRegression(max_iter=1000)),
        ("dt", tree(max_depth=6, random_state=0)),
        ("rf", RandomForestClassifier(n_estimators=100, random_state=0))],
        voting="soft"),
}
for name, m in models.items():
    s = cross_val_score(m, X, y, cv=5)
    print(f"{name:24s} accuracy {s.mean():.3f} +/- {s.std():.3f}")

# --- 3. a free validation estimate, and what the forest looked at --------
rf = RandomForestClassifier(n_estimators=300, oob_score=True,
                            random_state=0).fit(X, y)
print("random forest out-of-bag accuracy:", round(rf.oob_score_, 3))
top = np.argsort(-rf.feature_importances_)[:5]
print("most used features:", top.tolist(),
      rf.feature_importances_[top].round(3).tolist())
