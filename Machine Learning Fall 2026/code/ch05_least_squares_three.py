# Chapter 5 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.datasets import load_diabetes
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

# --- 1. the worked example, by formula -----------------------------------
x = np.array([1, 2, 3, 4, 5.0])          # story points
y = np.array([3, 5, 6, 9, 12.0])         # hours
w = ((x - x.mean()) * (y - y.mean())).sum() / ((x - x.mean()) ** 2).sum()
b = y.mean() - w * x.mean()
print(f"least squares:    w = {w:.2f}, b = {b:.2f}")

# --- 2. the same line by gradient descent --------------------------------
wg, bg, alpha = 0.0, 0.0, 0.05
for step in range(2000):
    err = (wg * x + bg) - y              # y_hat - y
    wg -= alpha * 2 * (err * x).mean()
    bg -= alpha * 2 * err.mean()
print(f"gradient descent: w = {wg:.2f}, b = {bg:.2f}")

# --- 3. a real dataset, against the mean baseline ------------------------
X, t = load_diabetes(return_X_y=True)
X_tr, X_te, t_tr, t_te = train_test_split(X, t, test_size=0.25,
                                          random_state=0)
scaler = StandardScaler().fit(X_tr)
model = LinearRegression().fit(scaler.transform(X_tr), t_tr)
pred = model.predict(scaler.transform(X_te))
base = np.full_like(t_te, t_tr.mean())   # always predict the mean
for name, p in [("baseline (mean)", base), ("linear regression", pred)]:
    print(f"{name:18s} MAE {mean_absolute_error(t_te, p):5.1f}  "
          f"RMSE {mean_squared_error(t_te, p) ** 0.5:5.1f}  "
          f"R2 {r2_score(t_te, p):5.2f}")
names = load_diabetes().feature_names
for i in np.argsort(-np.abs(model.coef_))[:3]:
    print(f"  weight on {names[i]:4s}: {float(model.coef_[i]):6.1f}")
