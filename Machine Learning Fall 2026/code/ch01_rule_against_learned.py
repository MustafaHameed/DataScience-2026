# Chapter 1 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
from sklearn.datasets import load_breast_cancer
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score

data = load_breast_cancer()
X, y = data.data, data.target            # y: 0 = malignant, 1 = benign
X_tr, X_te, y_tr, y_te = train_test_split(
    X, y, test_size=0.25, random_state=0, stratify=y)

# --- 1. a rule written by a person: large tumours are malignant ---------
radius = list(data.feature_names).index("mean radius")
rule = (X_te[:, radius] < 15).astype(int)     # small -> benign (1)
print("hand-written rule :", accuracy_score(y_te, rule))

# --- 2. a rule found by a learning algorithm ----------------------------
tree = DecisionTreeClassifier(max_depth=3, random_state=0).fit(X_tr, y_tr)
print("learned tree      :", accuracy_score(y_te, tree.predict(X_te)))

# --- 3. the baseline every result must beat -----------------------------
majority = int(y_tr.mean() > 0.5)
print("always 'majority' :", accuracy_score(y_te, [majority] * len(y_te)))
