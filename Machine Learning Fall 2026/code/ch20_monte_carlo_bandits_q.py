# Chapter 20 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np

rng = np.random.default_rng(0)

# --- 1. Monte Carlo inference: estimate an expectation by averaging ------
# the chance that a project stays healthy for 10 weeks (it slips with
# probability 0.2 each week): exact answer 0.8 ** 10 = 0.107
for n in (100, 10_000, 1_000_000):
    runs = rng.random((n, 10)) > 0.2            # True = stayed healthy
    print(f"{n:9,d} simulated runs: P(healthy 10 weeks) ~ "
          f"{runs.all(axis=1).mean():.4f}")


# --- 2. exploration vs exploitation: a 10-armed bandit -------------------
def bandit(strategy, steps=1000, arms=10, runs=300):
    total = np.zeros(steps)
    for _ in range(runs):
        true = rng.normal(0, 1, arms)           # unknown mean payoffs
        Q, N = np.zeros(arms), np.zeros(arms)
        for t in range(1, steps + 1):
            if strategy == "greedy":
                a = int(np.argmax(Q))
            elif strategy == "e-greedy":
                a = (int(rng.integers(arms)) if rng.random() < 0.1
                     else int(np.argmax(Q)))
            else:                               # UCB, c = 2
                ucb = Q + 2 * np.sqrt(np.log(t) / np.maximum(N, 1e-9))
                a = int(np.argmax(np.where(N == 0, np.inf, ucb)))
            r = rng.normal(true[a], 1)
            N[a] += 1
            Q[a] += (r - Q[a]) / N[a]           # incremental average
            total[t - 1] += r
    return total / runs


for s in ("greedy", "e-greedy", "ucb"):
    avg = bandit(s)
    print(f"{s:8s}: mean reward, steps 1-100 {avg[:100].mean():.2f}; "
          f"steps 901-1000 {avg[900:].mean():.2f}")

# --- 3. Q-learning on the technical-debt MDP, without being told P -------
P = {("H", "ship"): ([0.8, 0.2, 0.0], 10),
     ("D", "ship"): ([0.0, 0.5, 0.5], 5),
     ("F", "ship"): ([0.0, 0.0, 1.0], 0),
     ("H", "refactor"): ([1, 0, 0], -2), ("D", "refactor"): ([1, 0, 0], -4),
     ("F", "refactor"): ([1, 0, 0], -15)}
S, A = ["H", "D", "F"], ["ship", "refactor"]


def env_step(s, a):                         # the agent only sees samples
    probs, reward = P[(S[s], A[a])]
    return int(rng.choice(3, p=probs)), reward


Q = np.zeros((3, 2))
alpha, gamma, eps = 0.1, 0.9, 0.1
s = 0
for t in range(200_000):
    a = int(rng.integers(2)) if rng.random() < eps else int(Q[s].argmax())
    s2, r = env_step(s, a)
    Q[s, a] += alpha * (r + gamma * Q[s2].max() - Q[s, a])   # Q-learning
    s = s2
print("learned Q (rows H, D, F; columns ship, refactor):")
print(Q.round(1))
print("learned policy:", dict(zip(S, [A[a] for a in Q.argmax(axis=1)])))
