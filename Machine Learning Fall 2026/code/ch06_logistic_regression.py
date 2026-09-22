# Chapter 6 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression

sigmoid = lambda z: 1 / (1 + np.exp(-z))

# --- 1. the worked example: gradient descent by hand ---------------------
x = np.array([1, 2, 3, 4.0])
y = np.array([0, 0, 1, 1.0])
logloss = lambda p: -np.mean(y * np.log(p) + (1 - y) * np.log(1 - p))
w, b = 0.0, 0.0
p = sigmoid(w * x + b)
print("before: loss", round(logloss(p), 3))
w -= 1.0 * np.mean((p - y) * x)          # gradient: mean of (p - y) * x
b -= 1.0 * np.mean(p - y)
print(f"after one step: w = {w}, b = {b}, "
      f"loss {logloss(sigmoid(w * x + b)):.3f}")
for step in range(5000):                 # ... and many more steps
    p = sigmoid(w * x + b)
    w, b = w - np.mean((p - y) * x), b - np.mean(p - y)
print(f"after 5000 steps: w = {w:.2f}, b = {b:.2f}, "
      f"boundary at x = {-b / w:.2f}")

# --- 2. a real classifier, read as odds ratios ---------------------------
rng = np.random.default_rng(3)
n = 500                                  # past projects
names = ["team_size", "changes", "dependencies", "clarity"]
X = np.column_stack([rng.integers(3, 16, n),          # people
                     rng.poisson(5, n),                # change requests
                     rng.integers(0, 7, n),            # external teams
                     rng.uniform(0.3, 1.0, n)])        # clarity, 0-1
z = -1.5 + 0.05 * X[:, 0] + 0.3 * X[:, 1] + 0.4 * X[:, 2] - 3.5 * X[:, 3]
y = (rng.random(n) < 1 / (1 + np.exp(-z))).astype(int)   # 1 = late
X_tr, X_te, y_tr, y_te = train_test_split(
    X, y, test_size=0.25, random_state=0, stratify=y)
clf = make_pipeline(StandardScaler(), LogisticRegression())
clf.fit(X_tr, y_tr)
print("test accuracy:", round(clf.score(X_te, y_te), 3))
coef = clf[-1].coef_[0]
for i in np.argsort(-np.abs(coef)):      # strongest risk factors first
    print(f"  {names[i]:13s} w = {coef[i]:5.2f}  "
          f"odds x {np.exp(coef[i]):.2f} per standard deviation")
print("P(late) for 3 test projects:",
      clf.predict_proba(X_te[:3])[:, 1].round(3))

# --- 3. three classes: softmax -------------------------------------------
m = 300                                  # support tickets
T = np.column_stack([rng.uniform(0, 4, m),            # log10 users hit
                     rng.integers(0, 2, m)])           # 1 = workaround
score = T[:, 0] - 1.2 * T[:, 1] + rng.normal(0, 0.4, m)
level = np.digitize(score, [1.0, 2.5])   # 0 low, 1 medium, 2 high
soft = make_pipeline(StandardScaler(), LogisticRegression())
soft.fit(T, level)
print("P(low, medium, high) for ticket 60:",
      soft.predict_proba(T[[60]]).round(3))
