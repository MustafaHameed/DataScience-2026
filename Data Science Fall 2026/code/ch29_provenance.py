"""Chapter 29 - Reproducibility boilerplate.

Put this at the top of every analysis script. It costs five minutes and moves
a result from 'not reproducible' to 'reproducible' (Figure 29.2).
"""
import json
import os
import random
import subprocess
from datetime import datetime, timezone

import numpy as np

SEED = 42
random.seed(SEED)
np.random.seed(SEED)
os.environ["PYTHONHASHSEED"] = str(SEED)
# torch.manual_seed(SEED); torch.use_deterministic_algorithms(True)


def _git(*args, default="unavailable"):
    try:
        return subprocess.check_output(["git", *args],
                                       stderr=subprocess.DEVNULL).decode().strip()
    except Exception:
        return default


def provenance(outfile="run_provenance.json", **params):
    """Record everything needed to regenerate this result."""
    dirty = bool(_git("status", "--porcelain", default=""))
    record = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "git_commit": _git("rev-parse", "HEAD"),
        "git_dirty": dirty,
        "seed": SEED,
        "params": params,
    }
    if dirty:
        print("WARNING: uncommitted changes -- this run is not reproducible")
    with open(outfile, "w") as f:
        json.dump(record, f, indent=2)
    return record


if __name__ == "__main__":
    rec = provenance(model="logreg", threshold=0.35, n_features=12)
    print(json.dumps(rec, indent=2))
