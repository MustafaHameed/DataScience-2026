# Chapter 16 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.decomposition import PCA
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import cross_val_score

# --- 1. the worked example: PCA by eigen-decomposition -------------------
X = np.array([[1, 2], [2, 1], [3, 3], [4, 5], [5, 4]], dtype=float)
Xc = X - X.mean(axis=0)                     # centre
S = np.cov(Xc, rowvar=False)                # covariance, divides by n - 1
vals, vecs = np.linalg.eigh(S)              # eigenvalues, ascending
order = np.argsort(vals)[::-1]
vals, vecs = vals[order], vecs[:, order]
print("covariance:\n", S)
print("eigenvalues:", vals, " explained:", (vals / vals.sum()).round(3))
print("PC1 direction:", vecs[:, 0].round(3))
print("scores on PC1:", (Xc @ vecs[:, 0]).round(3))

# --- 2. 40 dashboard metrics: how many components? -----------------------
rng = np.random.default_rng(16)
n, d = 1000, 40                          # projects, weekly metrics
Z = rng.normal(size=(n, 4))              # hidden: size, quality, churn, staff
A = rng.normal(size=(4, d))              # how each metric reflects them
Xm = Z @ A + rng.normal(0, 1.0, (n, d))  # what the dashboard records
ym = (Z[:, 0] - Z[:, 1] + 0.5 * Z[:, 2] + rng.normal(0, 0.5, n) > 0)  # late
pca = PCA().fit(StandardScaler().fit_transform(Xm))
cum = np.cumsum(pca.explained_variance_ratio_)
for target in (0.5, 0.8, 0.9, 0.95):
    print(f"{target:.0%} of variance needs {np.argmax(cum >= target) + 1} "
          f"of {d} components")
for k in (1, 2, 4, 10, 40):
    pipe = make_pipeline(StandardScaler(), PCA(k), LogisticRegression())
    acc = cross_val_score(pipe, Xm, ym, cv=5).mean()
    print(f"logistic regression on {k:2d} components: accuracy {acc:.3f}")

# --- 3. a self-organizing map in twenty lines ----------------------------
rng = np.random.default_rng(16)
kinds = np.repeat([0, 1, 2], 50)         # maintenance, build, integration
centre = np.array([[-1.5, -1.2, -1.5, -0.5],     # small, short, little new
                   [0.8, 0.9, 1.2, 0.2],         # new build
                   [0.8, 0.6, 0.3, 1.0]])        # many integrations
Xi = centre[kinds] + rng.normal(0, 0.45, (150, 4))
Xi = StandardScaler().fit_transform(Xi)
rows, cols = 6, 6
grid = np.array([(r, c) for r in range(rows) for c in range(cols)])
rng = np.random.default_rng(0)
W = rng.normal(0, 0.5, (rows * cols, Xi.shape[1]))
steps = 3000
for t in range(steps):
    x = Xi[rng.integers(len(Xi))]
    bmu = np.argmin(((W - x) ** 2).sum(axis=1))   # best-matching unit
    eta = 0.5 * (1 - t / steps)                   # learning rate decays
    sigma = 3.0 * (1 - t / steps) + 0.5           # neighbourhood shrinks
    d2 = ((grid - grid[bmu]) ** 2).sum(axis=1)    # distance on the map
    h = np.exp(-d2 / (2 * sigma ** 2))            # neighbourhood weight
    W += eta * h[:, None] * (x - W)
# label each unit with the kind of project that most often maps to it
hits = {}
for x, y in zip(Xi, kinds):
    u = np.argmin(((W - x) ** 2).sum(axis=1))
    hits.setdefault(u, []).append(y)
names = "mbi"                       # maintenance, build, integration
for r in range(rows):
    print(" ".join(names[np.bincount(hits[r * cols + c]).argmax()]
                   if r * cols + c in hits else "." for c in range(cols)))
