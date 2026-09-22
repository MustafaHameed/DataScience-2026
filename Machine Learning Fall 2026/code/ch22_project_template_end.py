# Chapter 22 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
"""Project skeleton: from raw data to a saved, documented model."""
import csv
import datetime
import json
import os

import joblib
import sklearn
from sklearn.datasets import make_classification
from sklearn.dummy import DummyClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, recall_score
from sklearn.model_selection import (GridSearchCV, StratifiedKFold,
                                     cross_val_score, train_test_split)
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler

SEED = 42                                   # one seed, used everywhere
OUT = "project_output"
os.makedirs(OUT, exist_ok=True)

# --- 1. data, and a split made once and never changed --------------------
# a stand-in for your tracker export: 900 projects, 8 month-one metrics
X, y = make_classification(n_samples=900, n_features=8, n_informative=5,
                           n_redundant=1, weights=[0.75], flip_y=0.02,
                           random_state=5)  # y: 1 = overran, must not miss
X_tr, X_te, y_tr, y_te = train_test_split(
    X, y, test_size=0.25, stratify=y, random_state=SEED)
cv = StratifiedKFold(5, shuffle=True, random_state=SEED)
metric = "recall"                           # chosen in the proposal

# --- 2. baseline, then candidates, all by cross-validation ---------------
candidates = {
    "baseline": DummyClassifier(strategy="most_frequent"),
    "logistic": make_pipeline(StandardScaler(), LogisticRegression()),
    "forest": RandomForestClassifier(n_estimators=300, random_state=SEED),
}
with open(os.path.join(OUT, "experiments.csv"), "w", newline="") as fh:
    log = csv.writer(fh)
    log.writerow(["model", f"cv_{metric}_mean", f"cv_{metric}_std"])
    for name, model in candidates.items():
        s = cross_val_score(model, X_tr, y_tr, cv=cv, scoring=metric)
        log.writerow([name, round(s.mean(), 4), round(s.std(), 4)])
        print(f"{name:9s} CV {metric} {s.mean():.3f} +/- {s.std():.3f}")

# --- 3. tune the chosen family -------------------------------------------
params = {"logisticregression__C": [0.01, 0.1, 1, 10],
          "logisticregression__class_weight": [None, "balanced"]}
grid = GridSearchCV(candidates["logistic"], params, cv=cv,
                    scoring=metric).fit(X_tr, y_tr)
print("chosen settings:", grid.best_params_)

# --- 4. the test set, once -----------------------------------------------
pred = grid.predict(X_te)
test_recall = recall_score(y_te, pred)
print(classification_report(y_te, pred,
                            target_names=["on time", "overran"], digits=3))

# --- 5. save the model and write its model card --------------------------
joblib.dump(grid.best_estimator_, os.path.join(OUT, "model.joblib"))
card = {
    "model": "logistic regression, standardised features",
    "intended use": "rank projects for a project-office review; never to "
                    "judge a team or a manager",
    "data": f"{len(y)} past projects, 8 month-one metrics",
    "split": f"75/25 stratified, seed {SEED}; 5-fold CV for selection",
    "settings": grid.best_params_,
    f"test {metric}": round(float(test_recall), 3),
    "limitations": "one company's projects; re-validate after any change "
                   "to how projects are run",
    "date": datetime.date.today().isoformat(),
    "library": f"scikit-learn {sklearn.__version__}",
}
with open(os.path.join(OUT, "model_card.json"), "w") as fh:
    json.dump(card, fh, indent=2, default=str)
print("saved:", sorted(os.listdir(OUT)))
