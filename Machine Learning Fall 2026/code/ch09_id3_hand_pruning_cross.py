# Chapter 9 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
import pandas as pd
from sklearn.datasets import make_classification
from sklearn.model_selection import train_test_split, cross_val_score
from sklearn.tree import DecisionTreeClassifier, export_text

# --- 1. ID3's first choice, computed by hand -----------------------------
rows = ["Changing High High Few No", "Changing High High Many No",
        "Frozen High High Few Yes", "Vague Medium High Few Yes",
        "Vague Low Normal Few Yes", "Vague Low Normal Many No",
        "Frozen Low Normal Many Yes", "Changing Medium High Few No",
        "Changing Low Normal Few Yes", "Vague Medium Normal Few Yes",
        "Changing Medium Normal Many Yes", "Frozen Medium High Many Yes",
        "Frozen High Normal Few Yes", "Vague Medium High Many No"]
df = pd.DataFrame([r.split() for r in rows],
                  columns=["Reqs", "Pressure", "Load", "Deps", "Met"])


def entropy(labels):
    p = labels.value_counts(normalize=True)
    return float(-(p * np.log2(p)).sum())


def gain(data, attr):
    rest = sum(len(g) / len(data) * entropy(g.Met)
               for _, g in data.groupby(attr))
    return entropy(data.Met) - rest


for a in ["Reqs", "Pressure", "Load", "Deps"]:
    print(f"Gain(S, {a:8s}) = {gain(df, a):.3f}")
changing = df[df.Reqs == "Changing"]
print("inside Changing:", {a: round(gain(changing, a), 3)
                           for a in ["Pressure", "Load", "Deps"]})

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
