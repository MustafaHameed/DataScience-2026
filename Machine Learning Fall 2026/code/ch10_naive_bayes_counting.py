# Chapter 10 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
from sklearn.naive_bayes import GaussianNB, MultinomialNB
from sklearn.feature_extraction.text import CountVectorizer

# --- 1. naive Bayes on the sprint data, by counting ----------------------
sprints = [("Changing High High Few", "No"),
           ("Changing High High Many", "No"),
           ("Frozen High High Few", "Yes"),
           ("Vague Medium High Few", "Yes"),
           ("Vague Low Normal Few", "Yes"), ("Vague Low Normal Many", "No"),
           ("Frozen Low Normal Many", "Yes"),
           ("Changing Medium High Few", "No"),
           ("Changing Low Normal Few", "Yes"),
           ("Vague Medium Normal Few", "Yes"),
           ("Changing Medium Normal Many", "Yes"),
           ("Frozen Medium High Many", "Yes"),
           ("Frozen High Normal Few", "Yes"),
           ("Vague Medium High Many", "No")]
X = [s.split() for s, _ in sprints]
y = [c for _, c in sprints]
n_values = [3, 3, 2, 2]                  # values per attribute


def naive_bayes(x, laplace=0):
    scores = {}
    for c in ("Yes", "No"):
        rows = [r for r, lab in zip(X, y) if lab == c]
        p = len(rows) / len(X)                            # prior P(c)
        for j, v in enumerate(x):                         # P(x_j | c)
            count = sum(r[j] == v for r in rows)
            p *= (count + laplace) / (len(rows) + laplace * n_values[j])
        scores[c] = p
    total = sum(scores.values())
    return {c: round(s / total, 3) for c, s in scores.items()}, scores


post, raw = naive_bayes(["Changing", "Low", "High", "Many"])
print("unnormalised:", {c: round(s, 5) for c, s in raw.items()})
print("posterior:   ", post)
frozen = ["Frozen", "High", "High", "Many"]
print("Frozen sprint, no smoothing:", naive_bayes(frozen)[1])
print("Frozen sprint, Laplace     :", naive_bayes(frozen, laplace=1)[0])

# --- 2. Gaussian naive Bayes on the code-churn example -------------------
rng = np.random.default_rng(0)
churn = np.r_[rng.normal(500, 100, 1900),       # modules that pass QA
              rng.normal(900, 150, 100)]         # the 5% that fail
fails = np.r_[np.zeros(1900), np.ones(100)]
gnb = GaussianNB().fit(churn.reshape(-1, 1), fails)
print("learned means:", gnb.theta_.ravel().round(0))
print("P(fail | 750 lines):", gnb.predict_proba([[750]])[0, 1].round(2))
grid = np.arange(600, 1000).reshape(-1, 1)
print("boundary near", int(grid[gnb.predict(grid) == 1][0, 0]), "lines")

# --- 3. a ticket router from twelve tickets ------------------------------
tickets = ["app crashes when saving", "error on the export page",
           "crash after the update", "null pointer error in upload",
           "report page crashes on load", "error message on checkout",
           "add an option to export pdf", "please add a dark mode",
           "option to filter by date", "add support for csv import",
           "add a bulk edit option", "please add weekly summary"]
labels = [1] * 6 + [0] * 6                         # 1 = bug
vec = CountVectorizer()
nb = MultinomialNB(alpha=1.0).fit(vec.fit_transform(tickets), labels)
for t in ["crash after export", "add export option", "crash option crash"]:
    p = nb.predict_proba(vec.transform([t]))[0, 1]
    print(f"{t!r:22s} P(bug) = {p:.3f}")
