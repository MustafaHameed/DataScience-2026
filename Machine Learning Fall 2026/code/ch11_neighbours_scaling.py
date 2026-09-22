# Chapter 11 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.model_selection import cross_val_score, GridSearchCV
from sklearn.neighbors import KNeighborsClassifier
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

# --- 1. the worked example -----------------------------------------------
X = np.array([[90, 85], [80, 70], [60, 75], [55, 40], [72, 46], [40, 60]])
y = np.array(["OnTime", "OnTime", "OnTime", "Late", "Late", "Late"])
q = np.array([[65, 60]])
for k in (1, 3, 5):
    plain = KNeighborsClassifier(k).fit(X, y).predict(q)[0]
    weighted = KNeighborsClassifier(k, weights="distance").fit(X, y)
    print(f"k={k}: vote {plain}, distance-weighted {weighted.predict(q)[0]}")

# --- 2. scaling matters ----------------------------------------------------
rng = np.random.default_rng(11)
n = 600                                          # past projects
Xb = np.column_stack([rng.uniform(20, 100, n),   # requirement stability
                      rng.uniform(20, 100, n),   # team experience
                      rng.uniform(50, 2000, n),  # budget, thousand USD
                      rng.uniform(4, 52, n)])    # planned weeks
yb = (Xb[:, 0] * Xb[:, 1] < 2500).astype(int)    # 1 = late
flip = rng.random(n) < 0.05                      # 5% recorded wrongly
yb[flip] = 1 - yb[flip]
raw = cross_val_score(KNeighborsClassifier(5), Xb, yb, cv=10).mean()
scaled = cross_val_score(make_pipeline(StandardScaler(),
                                       KNeighborsClassifier(5)),
                         Xb, yb, cv=10).mean()
print(f"5-NN accuracy: raw {raw:.3f}, standardised {scaled:.3f}")

# --- 3. choose k by cross-validation --------------------------------------
pipe = make_pipeline(StandardScaler(), KNeighborsClassifier())
grid = GridSearchCV(pipe, {"kneighborsclassifier__n_neighbors":
                           [1, 3, 5, 7, 11, 21, 51, 101]}, cv=10)
grid.fit(Xb, yb)
for k, s in zip([1, 3, 5, 7, 11, 21, 51, 101],
                grid.cv_results_["mean_test_score"]):
    print(f"  k={k:3d}: CV accuracy {s:.3f}")

# --- 4. the curse of dimensionality ----------------------------------------
rng = np.random.default_rng(0)
for dim in (2, 10, 100, 1000):
    P, Q = rng.uniform(size=(500, dim)), rng.uniform(size=(50, dim))
    D = np.sqrt(((Q[:, None, :] - P[None, :, :]) ** 2).sum(-1))
    ratio = (D.min(1) / D.max(1)).mean()
    print(f"d={dim:4d}: nearest/farthest = {ratio:.2f}")
