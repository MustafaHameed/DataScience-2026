# Chapter 12 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.svm import SVC
from sklearn.datasets import make_circles, load_breast_cancer
from sklearn.model_selection import cross_val_score, GridSearchCV
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

# --- 1. the worked example: a maximum-margin line ------------------------
X = np.array([[2, 2], [3, 2], [2, 4], [0, 0], [-1, -1], [1, -2]])
y = np.array([1, 1, 1, -1, -1, -1])
svm = SVC(kernel="linear", C=1e6).fit(X, y)       # huge C = hard margin
w, b = svm.coef_[0], svm.intercept_[0]
print("w =", w.round(3), " b =", round(b, 3))
print("margin width 2/||w|| =", round(2 / np.linalg.norm(w), 3))
print("support vectors:", svm.support_vectors_.tolist())
print("f(1.5, 1) =", round(float(svm.decision_function([[1.5, 1]])[0]), 3))

# --- 2. a problem no line can solve --------------------------------------
Xc, yc = make_circles(n_samples=400, noise=0.08, factor=0.4, random_state=0)
for kernel in ("linear", "poly", "rbf"):
    acc = cross_val_score(SVC(kernel=kernel, degree=2), Xc, yc, cv=5).mean()
    print(f"circles, {kernel:6s} kernel: accuracy {acc:.3f}")

# --- 3. tune C and gamma together, features scaled -----------------------
Xb, yb = load_breast_cancer(return_X_y=True)
pipe = make_pipeline(StandardScaler(), SVC(kernel="rbf"))
grid = GridSearchCV(pipe, {"svc__C": [0.1, 1, 10, 100],
                           "svc__gamma": [0.001, 0.01, 0.1, 1]}, cv=5)
grid.fit(Xb, yb)
print("best:", grid.best_params_, "CV accuracy",
      round(grid.best_score_, 3))
print("support vectors used:", grid.best_estimator_[-1].n_support_.sum(),
      "of", len(yb))
