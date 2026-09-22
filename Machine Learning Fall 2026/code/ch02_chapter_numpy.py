# Chapter 2 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np

# --- vectors and the dot product -----------------------------------------
x = np.array([0.9, 0.6, 4.0])
w = np.array([2.0, 3.0, 0.25])
print("score w.x + b :", round(w @ x - 3, 2))            # 1.6
z = np.array([0.5, 0.4, 1.0])
print("distance      :", round(np.linalg.norm(x - z), 2))  # 3.03

# --- a whole dataset at once ---------------------------------------------
X = np.array([[0.9, 0.6, 4], [0.5, 0.4, 1], [0.7, 0.9, 6]])
print("X w + b       :", X @ w - 3)                      # [1.6 -0.55 2.6]

# --- gradient descent on L(w) = (w - 3)^2 --------------------------------
for alpha in (0.1, 0.5, 1.1):
    wt = 0.0
    path = [wt]
    for step in range(4):
        grad = 2 * (wt - 3)
        wt = wt - alpha * grad
        path.append(round(wt, 3))
    print(f"alpha={alpha}: {path}")

# --- Bayes by simulation: how many alarms are real? ----------------------
rng = np.random.default_rng(0)
n = 1_000_000
attack = rng.random(n) < 0.001
alarm = np.where(attack, rng.random(n) < 0.99, rng.random(n) < 0.01)
print("P(attack | alarm), simulated:", attack[alarm].mean())   # about 0.09

# --- entropy -------------------------------------------------------------
def entropy(counts):
    p = np.array(counts) / np.sum(counts)
    p = p[p > 0]
    return -(p * np.log2(p)).sum()
print("H(9 yes, 5 no) :", round(entropy([9, 5]), 3))       # 0.94
print("H(7 yes, 7 no) :", round(entropy([7, 7]), 3))       # 1.0
