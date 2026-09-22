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

# --- a synthetic cohort of 600 students, with gaps ----------------------
rng = np.random.default_rng(0)
n = 600
df = pd.DataFrame({
    "attendance": rng.uniform(0.3, 1.0, n),
    "quiz_avg": rng.uniform(0.0, 1.0, n),
    "logins": rng.poisson(5, n).astype(float),
    "programme": rng.choice(["BSIT", "BSCS", "BSSE"], n),
})
risk = 3.6 - 4 * df.attendance - 2 * df.quiz_avg - 0.1 * df.logins
df["fail"] = (risk + rng.normal(0, 0.5, n) > 0).astype(int)
df.loc[rng.random(n) < 0.1, "quiz_avg"] = np.nan        # 10% missing
print("fail rate:", df.fail.mean().round(3))

X, y = df.drop(columns="fail"), df["fail"]

# --- 1. split FIRST; the test set is locked away ------------------------
X_tr, X_te, y_tr, y_te = train_test_split(
    X, y, test_size=0.2, stratify=y, random_state=0)

# --- 2. every preparation step lives inside the pipeline ----------------
num = ["attendance", "quiz_avg", "logins"]
prep = ColumnTransformer([
    ("num", make_pipeline(SimpleImputer(strategy="median"),
                          StandardScaler()), num),
    ("cat", OneHotEncoder(handle_unknown="ignore"), ["programme"]),
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
