"""Chapter 10 - Building a confounder and watching it fool you.

`attends` has NO effect on `mark` in this simulation. Both are caused by Z.
The naive comparison nevertheless reports a large, highly significant effect.
"""
import numpy as np
import pandas as pd

rng = np.random.default_rng(7)
n = 2000

# Z = conscientiousness. It causes BOTH attendance and marks.
Z = rng.normal(0, 1, n)

attends = (Z + rng.normal(0, 0.6, n)) > 0.3      # caused by Z
mark = 55 + 9 * Z + rng.normal(0, 5, n)          # caused by Z ONLY
#      note: `attends` does not appear here at all -- the true effect is zero

df = pd.DataFrame({"Z": Z, "attends": attends, "mark": mark})

naive = df[df.attends].mark.mean() - df[~df.attends].mark.mean()
print(f"naive 'effect' of attending: {naive:+.2f} marks")

# It is also statistically significant, which is the trap:
from scipy import stats  # noqa: E402
t, p = stats.ttest_ind(df[df.attends].mark, df[~df.attends].mark)
print(f"  t = {t:.2f}, p = {p:.2e}   <- significant, and entirely spurious")

# Adjust for the confounder by comparing within strata of Z.
df["band"] = pd.qcut(df.Z, 5)
adj = (
    df.groupby("band", observed=True)
    .apply(lambda g: g[g.attends].mark.mean() - g[~g.attends].mark.mean(),
           include_groups=False)
    .mean()
)
print(f"after adjusting for Z:       {adj:+.2f} marks   <- the truth, ~0")

print("""
No amount of extra data would have corrected the naive estimate.
Only the adjustment did -- and in a real study you would only know to
adjust if you had thought of Z in the first place. That is why
randomisation, which balances confounders you never thought of, is
worth more than any amount of statistical control.""")
