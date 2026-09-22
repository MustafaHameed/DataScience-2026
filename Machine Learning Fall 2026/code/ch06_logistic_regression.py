# Chapter 6 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.datasets import load_breast_cancer, load_iris
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
data = load_breast_cancer()
X_tr, X_te, y_tr, y_te = train_test_split(
    data.data, data.target, test_size=0.25, random_state=0,
    stratify=data.target)
clf = make_pipeline(StandardScaler(), LogisticRegression(max_iter=1000))
clf.fit(X_tr, y_tr)
print("test accuracy:", round(clf.score(X_te, y_te), 3))
coef = clf[-1].coef_[0]
for i in np.argsort(coef)[:3]:           # most negative = most malignant
    print(f"  {data.feature_names[i]:22s} w = {coef[i]:5.2f}  "
          f"odds x {np.exp(coef[i]):.2f} per standard deviation")
print("P(benign) for 3 test tumours:",
      clf.predict_proba(X_te[:3])[:, 1].round(3))

# --- 3. three classes: softmax -------------------------------------------
iris = load_iris()
soft = make_pipeline(StandardScaler(), LogisticRegression(max_iter=1000))
soft.fit(iris.data, iris.target)
print("softmax probabilities for flower 60:",
      soft.predict_proba(iris.data[[60]]).round(3))
