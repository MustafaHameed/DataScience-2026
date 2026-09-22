# Chapter 19 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np

# --- the technical-debt MDP: states H, D, F; actions ship (0), refactor (1)
S, A = ["H", "D", "F"], ["ship", "refactor"]
P = np.zeros((2, 3, 3))              # P[a, s, s']
P[0, 0] = [0.8, 0.2, 0.0]            # ship when healthy
P[0, 1] = [0.0, 0.5, 0.5]            # ship when debt-laden
P[0, 2] = [0.0, 0.0, 1.0]            # ship when frozen: stays frozen
P[1, :, 0] = 1.0                     # refactor: back to healthy
R = np.array([[10.0, 5.0, 0.0],      # R[a, s]: ship
              [-2.0, -4.0, -15.0]])  #          refactor
gamma = 0.9

# --- value iteration -------------------------------------------------------
V = np.zeros(3)
for it in range(1, 1000):
    Q = R + gamma * P @ V            # Q[a, s]
    V_new = Q.max(axis=0)
    if it <= 3:
        print(f"iteration {it}: V = {V_new.round(3)}  "
              f"policy = {[A[a] for a in Q.argmax(axis=0)]}")
    if np.abs(V_new - V).max() < 1e-8:
        break
    V = V_new
print(f"converged after {it} iterations: V = {V.round(2)}")
print("optimal policy:", dict(zip(S, [A[a] for a in Q.argmax(axis=0)])))

# --- policy iteration: evaluate exactly, then improve ----------------------
policy = np.zeros(3, dtype=int)                  # start: always ship
while True:
    Pp = P[policy, np.arange(3)]                 # 3 x 3 under the policy
    Rp = R[policy, np.arange(3)]
    V = np.linalg.solve(np.eye(3) - gamma * Pp, Rp)   # V = R + g P V
    print("evaluate", [A[a] for a in policy], "->", V.round(2))
    new = (R + gamma * P @ V).argmax(axis=0)     # improve greedily
    if (new == policy).all():
        break
    policy = new
