# Chapter 19 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np

# --- the maintenance MDP: states H, D, F; actions operate (0), repair (1)
S, A = ["H", "D", "F"], ["operate", "repair"]
P = np.zeros((2, 3, 3))              # P[a, s, s']
P[0, 0] = [0.8, 0.2, 0.0]            # operate when healthy
P[0, 1] = [0.0, 0.5, 0.5]            # operate when degraded
P[0, 2] = [0.0, 0.0, 1.0]            # operate when failed: stays failed
P[1, :, 0] = 1.0                     # repair: back to healthy
R = np.array([[10.0, 5.0, 0.0],      # R[a, s]: operate
              [-2.0, -4.0, -15.0]])  #          repair
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
policy = np.zeros(3, dtype=int)                  # start: always operate
while True:
    Pp = P[policy, np.arange(3)]                 # 3 x 3 under the policy
    Rp = R[policy, np.arange(3)]
    V = np.linalg.solve(np.eye(3) - gamma * Pp, Rp)   # V = R + g P V
    print("evaluate", [A[a] for a in policy], "->", V.round(2))
    new = (R + gamma * P @ V).argmax(axis=0)     # improve greedily
    if (new == policy).all():
        break
    policy = new

# --- a gridworld: 3 x 4, wall at (1,1), exits +1 and -1, slippery moves ---
rows, cols, wall = 3, 4, (1, 1)
exits = {(0, 3): 1.0, (1, 3): -1.0}
moves = {"^": (-1, 0), "v": (1, 0), "<": (0, -1), ">": (0, 1)}
side = {"^": "<>", "v": "<>", "<": "^v", ">": "^v"}


def step(s, a):
    r, c = s[0] + moves[a][0], s[1] + moves[a][1]
    ok = 0 <= r < rows and 0 <= c < cols and (r, c) != wall
    return (r, c) if ok else s


def q(V, s, a):                                  # expected next value
    return (0.8 * V[step(s, a)] + 0.1 * V[step(s, side[a][0])]
            + 0.1 * V[step(s, side[a][1])])


cells = [(r, c) for r in range(rows) for c in range(cols) if (r, c) != wall]
V = {s: 0.0 for s in cells}
for _ in range(200):
    V = {s: exits[s] if s in exits else
         -0.04 + 1.0 * max(q(V, s, a) for a in moves) for s in cells}
for r in range(rows):
    line = []
    for c in range(cols):
        s = (r, c)
        if s == wall:
            line.append(" wall  ")
        else:
            best = "" if s in exits else max(moves, key=lambda a: q(V, s, a))
            line.append(f"{V[s]:6.3f}{best or ' '}")
    print(" ".join(line))
