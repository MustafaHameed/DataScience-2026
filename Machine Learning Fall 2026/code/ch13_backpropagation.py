# Chapter 13 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.neural_network import MLPClassifier
from sklearn.linear_model import LogisticRegression

# --- 1. XOR: no perceptron can learn it; a 2-3-1 network can -------------
X = np.array([[0, 0], [0, 1], [1, 0], [1, 1]], dtype=float)
t = np.array([[0], [1], [1], [0]], dtype=float)
sig = lambda z: 1 / (1 + np.exp(-z))
rng = np.random.default_rng(1)
W1, b1 = rng.normal(0, 1, (2, 3)), np.zeros(3)     # input -> hidden
W2, b2 = rng.normal(0, 1, (3, 1)), np.zeros(1)     # hidden -> output
eta = 1.0
for epoch in range(5001):
    h = sig(X @ W1 + b1)                           # forward pass
    o = sig(h @ W2 + b2)
    d_out = o * (1 - o) * (t - o)                  # output deltas
    d_hid = h * (1 - h) * (d_out @ W2.T)           # hidden deltas
    W2 += eta * h.T @ d_out                        # weight updates
    b2 += eta * d_out.sum(0)
    W1 += eta * X.T @ d_hid
    b1 += eta * d_hid.sum(0)
    if epoch in (0, 500, 5000):
        err = 0.5 * np.sum((t - o) ** 2)
        print(f"epoch {epoch:4d}: error {err:.4f}  "
              f"outputs {o.ravel().round(2)}")

# --- 2. staffing: too few people is late, and so is too many -------------
rng = np.random.default_rng(0)
n = 800
scope = rng.uniform(100, 1000, n)                 # story points
team = rng.uniform(2, 20, n)                      # people
Xs = np.column_stack([scope, team])
ys = (np.abs(team - scope / 60) > 5).astype(int)  # 1 = late
flip = rng.random(n) < 0.05                       # 5% recorded wrongly
ys[flip] = 1 - ys[flip]
X_tr, X_te, y_tr, y_te = train_test_split(Xs, ys, test_size=0.3,
                                          random_state=0, stratify=ys)
sc = StandardScaler().fit(X_tr)
X_tr, X_te = sc.transform(X_tr), sc.transform(X_te)
lin = LogisticRegression().fit(X_tr, y_tr)
print("logistic regression test accuracy:", round(lin.score(X_te, y_te), 3))
for alpha in (0.0001, 1.0):                        # weight decay
    net = MLPClassifier(hidden_layer_sizes=(16,), alpha=alpha,
                        max_iter=3000, random_state=0).fit(X_tr, y_tr)
    print(f"network, weight decay {alpha:g}: {net.n_iter_} epochs, "
          f"test accuracy {net.score(X_te, y_te):.3f}")
