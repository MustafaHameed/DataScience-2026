# Chapter 9 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
import pandas as pd
from sklearn.datasets import make_classification
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.tree import DecisionTreeClassifier, export_text

# --- 1. ID3's first choice, computed by hand -----------------------------
rows = ["Sunny Hot High Weak No", "Sunny Hot High Strong No",
        "Overcast Hot High Weak Yes", "Rain Mild High Weak Yes",
        "Rain Cool Normal Weak Yes", "Rain Cool Normal Strong No",
        "Overcast Cool Normal Strong Yes", "Sunny Mild High Weak No",
        "Sunny Cool Normal Weak Yes", "Rain Mild Normal Weak Yes",
        "Sunny Mild Normal Strong Yes", "Overcast Mild High Strong Yes",
        "Overcast Hot Normal Weak Yes", "Rain Mild High Strong No"]
df = pd.DataFrame([r.split() for r in rows],
                  columns=["Outlook", "Temp", "Humidity", "Wind", "Play"])


def entropy(labels):
    p = labels.value_counts(normalize=True)
    return float(-(p * np.log2(p)).sum())


def gain(data, attr):
    rest = sum(len(g) / len(data) * entropy(g.Play)
               for _, g in data.groupby(attr))
    return entropy(data.Play) - rest


for a in ["Outlook", "Temp", "Humidity", "Wind"]:
    print(f"Gain(S, {a:8s}) = {gain(df, a):.3f}")
sunny = df[df.Outlook == "Sunny"]
print("inside Sunny:", {a: round(gain(sunny, a), 3)
                        for a in ["Temp", "Humidity", "Wind"]})

# --- 2. noisy data: a tree grown without limit, then pruned -------------
X, y = make_classification(n_samples=1500, n_features=10, n_informative=4,
                           flip_y=0.15, random_state=4)  # 15% labels flipped
X_tr, X_te, y_tr, y_te = train_test_split(X, y, test_size=0.5,
                                          random_state=0)
full = DecisionTreeClassifier(criterion="entropy", random_state=0)
full.fit(X_tr, y_tr)
print(f"unpruned: {full.tree_.node_count:3d} nodes, "
      f"train {full.score(X_tr, y_tr):.3f}, "
      f"test {full.score(X_te, y_te):.3f}")

# cost-complexity pruning: choose alpha by 5-fold CV on the training data
alphas = full.cost_complexity_pruning_path(X_tr, y_tr).ccp_alphas[:-1]
cv = [cross_val_score(DecisionTreeClassifier(criterion="entropy",
      ccp_alpha=a, random_state=0), X_tr, y_tr, cv=5).mean() for a in alphas]
best = alphas[int(np.argmax(cv))]
pruned = DecisionTreeClassifier(criterion="entropy", ccp_alpha=best,
                                random_state=0).fit(X_tr, y_tr)
print(f"pruned:   {pruned.tree_.node_count:3d} nodes, "
      f"train {pruned.score(X_tr, y_tr):.3f}, "
      f"test {pruned.score(X_te, y_te):.3f}  (alpha = {best:.4f})")
print(export_text(pruned, max_depth=1))
