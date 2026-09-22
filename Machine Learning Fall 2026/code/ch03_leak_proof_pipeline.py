# Chapter 3 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import numpy as np
import pandas as pd
from sklearn.model_selection import train_test_split
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline, make_pipeline
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import StandardScaler, OneHotEncoder
from sklearn.linear_model import LogisticRegression
from sklearn.dummy import DummyClassifier

# --- 600 past projects, with gaps ---------------------------------------
rng = np.random.default_rng(0)
n = 600
df = pd.DataFrame({
    "clarity": rng.uniform(0.3, 1.0, n),         # requirement clarity
    "coverage": rng.uniform(0.0, 1.0, n),        # test coverage
    "changes": rng.poisson(5, n).astype(float),  # change requests
    "ptype": rng.choice(["Web", "Mobile", "Data"], n),
})
risk = 2.6 - 4 * df.clarity - 2 * df.coverage + 0.1 * df.changes
df["late"] = (risk + rng.normal(0, 0.5, n) > 0).astype(int)
df.loc[rng.random(n) < 0.1, "coverage"] = np.nan        # 10% missing
print("late rate:", df.late.mean().round(3))

X, y = df.drop(columns="late"), df["late"]

# --- 1. split FIRST; the test set is locked away ------------------------
X_tr, X_te, y_tr, y_te = train_test_split(
    X, y, test_size=0.2, stratify=y, random_state=0)

# --- 2. every preparation step lives inside the pipeline ----------------
num = ["clarity", "coverage", "changes"]
prep = ColumnTransformer([
    ("num", make_pipeline(SimpleImputer(strategy="median"),
                          StandardScaler()), num),
    ("cat", OneHotEncoder(handle_unknown="ignore"), ["ptype"]),
])
model = Pipeline([("prep", prep), ("clf", LogisticRegression())])

# --- 3. a baseline first, then the model --------------------------------
base = DummyClassifier(strategy="most_frequent").fit(X_tr, y_tr)
print("baseline accuracy:", round(base.score(X_te, y_te), 3))
model.fit(X_tr, y_tr)
print("model accuracy   :", round(model.score(X_te, y_te), 3))

# --- 4. look inside: the scaler learned from training rows only ---------
scaler = model.named_steps["prep"].named_transformers_["num"][-1]
print("scaler means:", scaler.mean_.round(2))
