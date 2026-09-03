"""appB_tests.py -- Advanced Research Methodology, Appendix B.

Python versions of the procedures in the handout, for those who prefer it.
Each block generates its own illustrative data, so the file runs end to end.

Install once:
    pip install numpy pandas scipy statsmodels pingouin scikit-learn

Every number produced here is from simulated data. Nothing in this file is
a research finding.
"""

import numpy as np
import pandas as pd
import pingouin as pg
import statsmodels.formula.api as smf
from statsmodels.stats.multitest import multipletests
from statsmodels.stats.power import TTestIndPower

RNG = np.random.default_rng(20260902)  # every analysis fixes its seed


def two_groups():
    """Two independent groups: Welch, effect size, bootstrap interval."""
    a = RNG.normal(6.9, 2.5, 60)
    b = RNG.normal(6.0, 2.5, 60)

    # Welch by default -- does not assume equal variances.
    print(pg.ttest(a, b, correction=True))
    print("Cohen's d:", pg.compute_effsize(a, b, eftype="cohen"))

    # Distribution-free alternative.
    print(pg.mwu(a, b))

    # Bootstrap interval: no distributional assumption at all.
    diffs = [RNG.choice(a, 60).mean() - RNG.choice(b, 60).mean()
             for _ in range(5000)]
    print("Bootstrap 95% CI:", np.percentile(diffs, [2.5, 97.5]))


def nested():
    """Nested data. Ignoring the structure inflates significance."""
    n_dev, n_art = 34, 4
    rows = []
    dev_eff = RNG.normal(0, 1.4, n_dev)
    art_eff = RNG.normal(0, 0.8, n_art)
    for d in range(n_dev):
        for a in range(n_art):
            assistant = (d + a) % 2
            rows.append(dict(
                dev=d, artefact=a, assistant=assistant,
                found=6 + 0.9 * assistant + dev_eff[d] + art_eff[a]
                      + RNG.normal(0, 1.6)))
    df = pd.DataFrame(rows)

    naive = smf.ols("found ~ assistant", data=df).fit()
    mixed = smf.mixedlm("found ~ assistant", data=df, groups=df["dev"]).fit()

    print("Naive SE:", naive.bse["assistant"])
    print("Mixed SE:", mixed.bse["assistant"])
    print("-- the ratio is how much the naive model overstated precision.")


def counts():
    """Software counts are overdispersed. Poisson understates the error."""
    size = RNG.poisson(40, 300)
    defects = RNG.negative_binomial(1.2, 1.2 / (1.2 + 0.05 * size))
    df = pd.DataFrame({"size": size, "defects": defects})

    pois = smf.poisson("defects ~ size", data=df).fit(disp=0)
    negb = smf.negativebinomial("defects ~ size", data=df).fit(disp=0)

    print("Poisson SE:", pois.bse["size"], " NegBin SE:", negb.bse["size"])
    print("Poisson intervals are too narrow when variance exceeds the mean.")


def power_and_multiplicity():
    """A-priori power for the SESOI, and correction for declared tests."""
    n = TTestIndPower().solve_power(effect_size=0.40, alpha=0.05, power=0.80)
    print(f"Required n per group for d=0.40: {np.ceil(n):.0f}")

    pvals = [0.011, 0.031, 0.042, 0.180, 0.560]
    print("holm:", multipletests(pvals, method="holm")[1].round(4))
    print("BH:  ", multipletests(pvals, method="fdr_bh")[1].round(4))


def leakage_demo():
    """The leakage that matters most in software data: grouped splitting.

    Random row-level splitting lets a model memorise per-group base rates.
    The same model, honestly validated, performs far worse -- and that drop
    is the finding, not a failure (see Chapter 22).
    """
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import cross_val_score, GroupKFold, KFold

    n_groups, per_group = 20, 60
    groups = np.repeat(np.arange(n_groups), per_group)
    base = RNG.uniform(0.05, 0.6, n_groups)[groups]   # group-specific rate
    X = np.column_stack([RNG.normal(size=len(groups)),
                         base + RNG.normal(0, 0.05, len(groups))])
    y = RNG.binomial(1, base)

    clf = RandomForestClassifier(n_estimators=200, random_state=0)
    naive = cross_val_score(clf, X, y, cv=KFold(5, shuffle=True,
                                                random_state=0), scoring="roc_auc")
    honest = cross_val_score(clf, X, y, groups=groups,
                             cv=GroupKFold(5), scoring="roc_auc")

    print(f"Random  k-fold AUC: {naive.mean():.3f}  <- leaks group identity")
    print(f"Grouped k-fold AUC: {honest.mean():.3f}  <- honest")


if __name__ == "__main__":
    for name, fn in [("TWO GROUPS", two_groups),
                     ("NESTED DATA", nested),
                     ("COUNTS", counts),
                     ("POWER AND MULTIPLICITY", power_and_multiplicity),
                     ("LEAKAGE", leakage_demo)]:
        print(f"\n{'=' * 60}\n{name}\n{'=' * 60}")
        fn()
    print("\nAll data simulated. See Appendix B of the handout.")
