# Chapter 10 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import cross_val_score
from sklearn.naive_bayes import GaussianNB, MultinomialNB
from sklearn.feature_extraction.text import CountVectorizer

# --- 1. naive Bayes on PlayTennis, by counting ---------------------------
days = [("Sunny Hot High Weak", "No"), ("Sunny Hot High Strong", "No"),
        ("Overcast Hot High Weak", "Yes"), ("Rain Mild High Weak", "Yes"),
        ("Rain Cool Normal Weak", "Yes"), ("Rain Cool Normal Strong", "No"),
        ("Overcast Cool Normal Strong", "Yes"),
        ("Sunny Mild High Weak", "No"), ("Sunny Cool Normal Weak", "Yes"),
        ("Rain Mild Normal Weak", "Yes"), ("Sunny Mild Normal Strong", "Yes"),
        ("Overcast Mild High Strong", "Yes"),
        ("Overcast Hot Normal Weak", "Yes"), ("Rain Mild High Strong", "No")]
X = [d.split() for d, _ in days]
y = [c for _, c in days]
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


post, raw = naive_bayes(["Sunny", "Cool", "High", "Strong"])
print("unnormalised:", {c: round(s, 5) for c, s in raw.items()})
print("posterior:   ", post)
day = ["Overcast", "Hot", "High", "Strong"]
print("Overcast day, no smoothing:", naive_bayes(day)[1])
print("Overcast day, Laplace     :", naive_bayes(day, laplace=1)[0])

# --- 2. Gaussian naive Bayes on 30 continuous measurements ---------------
Xb, yb = load_breast_cancer(return_X_y=True)
print("GaussianNB 5-fold accuracy:",
      cross_val_score(GaussianNB(), Xb, yb, cv=5).mean().round(3))

# --- 3. a spam filter from twelve messages -------------------------------
msgs = ["win a free prize now", "free money win cash", "claim your free gift",
        "cheap loans win now", "free entry win big",
        "urgent cash prize claim",
        "meeting moved to monday", "please review the report",
        "lunch meeting tomorrow", "report due friday please",
        "project meeting notes", "draft report for review"]
labels = [1] * 6 + [0] * 6                         # 1 = spam
vec = CountVectorizer()
nb = MultinomialNB(alpha=1.0).fit(vec.fit_transform(msgs), labels)
for m in ["free meeting free", "report on the prize", "win win win"]:
    p = nb.predict_proba(vec.transform([m]))[0, 1]
    print(f"{m!r:24s} P(spam) = {p:.3f}")
