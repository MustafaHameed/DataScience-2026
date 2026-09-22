# Chapter 1 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score

rng = np.random.default_rng(1)
n = 1200
points = rng.choice([1, 2, 3, 5, 8, 13], n)   # story points
depends = rng.integers(0, 4, n)               # tasks it waited on
years = rng.integers(0, 11, n)                # assignee's experience
changes = rng.integers(0, 5, n)               # requirement changes
z = -3.3 + 0.18 * points + 0.9 * depends + 0.7 * changes - 0.3 * years
y = (rng.random(n) < 1 / (1 + np.exp(-z))).astype(int)   # 1 = late
X = np.column_stack([points, depends, years, changes])
X_tr, X_te, y_tr, y_te = train_test_split(
    X, y, test_size=0.25, random_state=0, stratify=y)

# --- 1. a rule written by a person: big tasks run late ------------------
rule = (X_te[:, 0] >= 8).astype(int)          # 8 or 13 points -> late
print("hand-written rule :", round(accuracy_score(y_te, rule), 3))

# --- 2. a rule found by a learning algorithm ----------------------------
tree = DecisionTreeClassifier(max_depth=3, random_state=0).fit(X_tr, y_tr)
acc = accuracy_score(y_te, tree.predict(X_te))
print("learned tree      :", round(acc, 3))

# --- 3. the baseline every result must beat -----------------------------
majority = int(y_tr.mean() > 0.5)             # here: 0, "on time"
acc = accuracy_score(y_te, [majority] * len(y_te))
print("always 'majority' :", round(acc, 3))
