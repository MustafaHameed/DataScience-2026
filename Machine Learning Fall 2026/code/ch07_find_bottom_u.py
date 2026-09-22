# Chapter 7 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import PolynomialFeatures, StandardScaler
from sklearn.linear_model import LinearRegression, Ridge
from sklearn.model_selection import (train_test_split, KFold,
                                     cross_val_score, GridSearchCV)

# --- noisy samples of a sine wave ----------------------------------------
rng = np.random.default_rng(0)
X = rng.uniform(0, 1, (60, 1))
y = np.sin(2 * np.pi * X[:, 0]) + rng.normal(0, 0.25, 60)
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.25,
                                          random_state=0)
cv = KFold(5, shuffle=True, random_state=0)
rmse = "neg_root_mean_squared_error"

# --- 1. complexity: training error against cross-validated error ---------
for deg in (1, 3, 6, 12):
    m = make_pipeline(PolynomialFeatures(deg), StandardScaler(),
                      LinearRegression())
    m.fit(X_tr, y_tr)
    train = np.sqrt(np.mean((m.predict(X_tr) - y_tr) ** 2))
    cvs = -cross_val_score(m, X_tr, y_tr, cv=cv, scoring=rmse)
    print(f"degree {deg:2d}: train RMSE {train:.3f}   "
          f"CV RMSE {cvs.mean():.3f} +/- {cvs.std():.3f}")

# --- 2. keep degree 12, but regularise; choose alpha by CV ---------------
pipe = make_pipeline(PolynomialFeatures(12), StandardScaler(), Ridge())
grid = GridSearchCV(pipe, {"ridge__alpha": [1e-6, 1e-3, 1e-1, 10, 1000]},
                    cv=cv, scoring=rmse).fit(X_tr, y_tr)
for a, s in zip(grid.cv_results_["param_ridge__alpha"],
                grid.cv_results_["mean_test_score"]):
    print(f"alpha {a:>8g}: CV RMSE {-s:.3f}")
print("chosen:", grid.best_params_)

# --- 3. only now, once: the test set -------------------------------------
pred = grid.predict(X_te)
print("test RMSE:", round(float(np.sqrt(np.mean((pred - y_te) ** 2))), 3))
