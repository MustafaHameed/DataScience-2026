# Chapter 18 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np

# --- the worked example: a project, Healthy (0) or Distressed (1) --------
pi = np.array([0.7, 0.3])                        # initial state
A = np.array([[0.8, 0.2],                        # transitions, row = from
              [0.3, 0.7]])
B = np.array([[0.70, 0.25, 0.05],                # reports: on plan, behind,
              [0.10, 0.50, 0.40]])               #          blocked
obs = [0, 1, 2]                                  # on plan, behind, blocked


def forward(obs, pi, A, B):
    alpha = pi * B[:, obs[0]]
    for o in obs[1:]:
        alpha = (alpha @ A) * B[:, o]
    return alpha                                 # sums to P(obs)


def viterbi(obs, pi, A, B):
    delta, back = pi * B[:, obs[0]], []
    for o in obs[1:]:
        cand = delta[:, None] * A                # cand[i, j]: from i to j
        back.append(cand.argmax(axis=0))
        delta = cand.max(axis=0) * B[:, o]
    path = [int(delta.argmax())]
    for b in reversed(back):
        path.insert(0, int(b[path[0]]))
    return path, delta.max()


alpha = forward(obs, pi, A, B)
print("P(on plan, behind, blocked) =", round(alpha.sum(), 6))
print("P(state at t=3 | obs)  =", (alpha / alpha.sum()).round(3))
path, p = viterbi(obs, pi, A, B)
print("most likely states:", ["HD"[s] for s in path], "prob", round(p, 5))

# --- Monte Carlo: estimate the same probability by sampling --------------
rng = np.random.default_rng(0)


def sample(T):
    s = rng.choice(2, p=pi)
    out = []
    for t in range(T):
        out.append(rng.choice(3, p=B[s]))
        s = rng.choice(2, p=A[s])
    return out


for n in (1_000, 10_000, 100_000):
    hits = sum(sample(3) == obs for _ in range(n))
    print(f"Monte Carlo, {n:6d} samples: {hits / n:.4f}")

# --- Baum-Welch (EM): relearn A and B from one long sampled sequence -----
seq = np.array(sample(3000))
A_hat = np.array([[0.5, 0.5], [0.5, 0.5]])       # deliberately vague start
B_hat = np.array([[0.5, 0.3, 0.2], [0.2, 0.3, 0.5]])
p_hat = np.array([0.5, 0.5])
T = len(seq)
for it in range(60):
    a, c = np.zeros((T, 2)), np.zeros(T)         # scaled forward pass
    a[0] = p_hat * B_hat[:, seq[0]]
    c[0] = a[0].sum()
    a[0] /= c[0]
    for t in range(1, T):
        a[t] = (a[t - 1] @ A_hat) * B_hat[:, seq[t]]
        c[t] = a[t].sum()
        a[t] /= c[t]

    b = np.ones((T, 2))                          # scaled backward pass
    for t in range(T - 2, -1, -1):
        b[t] = A_hat @ (B_hat[:, seq[t + 1]] * b[t + 1]) / c[t + 1]

    g = a * b                                    # E-step: P(state_t | obs)
    xi = (a[:-1, :, None] * A_hat[None] *
          (B_hat[:, seq[1:]].T * b[1:])[:, None, :]) / c[1:, None, None]
    A_hat = xi.sum(0) / g[:-1].sum(0)[:, None]   # M-step: expected counts
    B_hat = np.array([g[seq == k].sum(0) for k in range(3)]).T
    B_hat /= g.sum(0)[:, None]
    p_hat = g[0]
print("relearned A:\n", A_hat.round(2))
print("relearned B:\n", B_hat.round(2))
