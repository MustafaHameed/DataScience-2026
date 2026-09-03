# =====================================================================
#  appB_tests.R -- Advanced Research Methodology, Appendix B
#
#  Runnable versions of the statistical procedures in the handout.
#  Each block is self-contained and generates its own illustrative data,
#  so the file runs end to end on a clean machine.
#
#  Install once:
#    install.packages(c("effsize", "lme4", "lmerTest", "MASS", "pwr",
#                       "TOSTER", "psych", "boot", "PMCMRplus"))
#
#  Every number produced here is from simulated data. Nothing in this
#  file is a research finding.
# =====================================================================

set.seed(20260902)   # every analysis in this course fixes its seed

# ---------------------------------------------------------------------
# 1. Two independent groups
#    Welch's t-test is R's default and should stay that way: it does not
#    assume equal variances and costs almost nothing when they are equal.
# ---------------------------------------------------------------------
d <- data.frame(
  g = rep(c("checklist", "unguided"), each = 60),
  y = c(rnorm(60, mean = 6.9, sd = 2.5), rnorm(60, mean = 6.0, sd = 2.5))
)

boxplot(y ~ g, data = d, main = "Look at the data before testing it")
qqnorm(residuals(lm(y ~ g, data = d))); qqline(residuals(lm(y ~ g, data = d)))

print(t.test(y ~ g, data = d))                       # estimate + interval
print(effsize::cohen.d(y ~ g, data = d))             # effect size + CI
print(effsize::cliff.delta(y ~ g, data = d))         # distribution-free

# ---------------------------------------------------------------------
# 2. Bootstrap interval -- no distributional assumption at all.
#    Prefer this to arguing about normality.
# ---------------------------------------------------------------------
diff_means <- function(data, i) {
  s <- data[i, ]
  mean(s$y[s$g == "checklist"]) - mean(s$y[s$g == "unguided"])
}
b <- boot::boot(d, diff_means, R = 5000)
print(boot::boot.ci(b, type = "perc"))

# ---------------------------------------------------------------------
# 3. Nested data -- the error that most inflates significance.
#    Reviews are nested in reviewers and crossed with artefacts.
# ---------------------------------------------------------------------
n_dev <- 34; n_art <- 4
reviews <- expand.grid(dev = factor(1:n_dev), artefact = factor(1:n_art))
reviews$assistant <- rep(c(0, 1), length.out = nrow(reviews))
dev_effect <- rnorm(n_dev, 0, 1.4)[reviews$dev]
art_effect <- rnorm(n_art, 0, 0.8)[reviews$artefact]
reviews$found <- 6 + 0.9 * reviews$assistant + dev_effect + art_effect +
                 rnorm(nrow(reviews), 0, 1.6)

naive <- lm(found ~ assistant, data = reviews)
mixed <- lmerTest::lmer(found ~ assistant + (1 | dev) + (1 | artefact),
                        data = reviews)

cat("\nNaive SE: ", summary(naive)$coefficients["assistant", "Std. Error"],
    "\nMixed SE: ", summary(mixed)$coefficients["assistant", "Std. Error"],
    "\n-- the ratio is how much the naive model overstated precision.\n")
print(confint(mixed, method = "Wald"))

# ---------------------------------------------------------------------
# 4. Counts -- software counts are overdispersed. Poisson gives standard
#    errors that are far too narrow.
# ---------------------------------------------------------------------
cd <- data.frame(size = rpois(300, 40))
cd$defects <- rnbinom(300, mu = 0.05 * cd$size, size = 1.2)   # overdispersed

pois <- glm(defects ~ size, data = cd, family = poisson)
negb <- MASS::glm.nb(defects ~ size, data = cd)

cat("\nPoisson dispersion (should be ~1):",
    sum(residuals(pois, type = "pearson")^2) / pois$df.residual, "\n")
cat("Poisson SE:", summary(pois)$coefficients["size", "Std. Error"],
    " NegBin SE:", summary(negb)$coefficients["size", "Std. Error"], "\n")

# ---------------------------------------------------------------------
# 5. A-priori power. Do this BEFORE collecting, for the smallest effect
#    of interest -- not for the effect you hope to find.
# ---------------------------------------------------------------------
print(pwr::pwr.t.test(d = 0.40, sig.level = 0.05, power = 0.80,
                      type = "two.sample"))

# ---------------------------------------------------------------------
# 6. Equivalence. "No significant difference" is NOT evidence of no
#    difference. Bounds come from your SESOI, not from convention.
# ---------------------------------------------------------------------
print(TOSTER::tsum_TOST(m1 = 6.1, m2 = 5.9, sd1 = 2.4, sd2 = 2.6,
                        n1 = 60, n2 = 60, eqb = 0.40))

# ---------------------------------------------------------------------
# 7. Multiplicity. Holm for confirmatory comparisons declared in the
#    protocol; Benjamini-Hochberg for exploratory work.
# ---------------------------------------------------------------------
pvals <- c(0.011, 0.031, 0.042, 0.180, 0.560)
print(rbind(raw  = pvals,
            holm = p.adjust(pvals, method = "holm"),
            BH   = p.adjust(pvals, method = "BH")))

# ---------------------------------------------------------------------
# 8. Reliability. Check dimensionality FIRST; report omega, not alpha
#    alone -- alpha rises with item count and assumes tau-equivalence.
# ---------------------------------------------------------------------
lat <- rnorm(200)
items <- sapply(1:6, function(i) lat * runif(1, 0.5, 0.8) + rnorm(200, 0, 0.6))
colnames(items) <- paste0("Q", 1:6)

print(psych::fa.parallel(items, fa = "fa", plot = FALSE)$nfact)
print(psych::omega(items, nfactors = 1, plot = FALSE)$omega.tot)
print(psych::alpha(items)$total$raw_alpha)

# ---------------------------------------------------------------------
# 9. Comparing classifiers across datasets. NOT pairwise t-tests:
#    multiplicity is uncorrected and normality is untenable.
# ---------------------------------------------------------------------
res <- matrix(c(0.71,0.69,0.66,0.64,  0.78,0.77,0.74,0.70,
                0.65,0.66,0.61,0.60,  0.82,0.80,0.79,0.74,
                0.69,0.68,0.65,0.63,  0.74,0.73,0.70,0.68),
              nrow = 6, byrow = TRUE,
              dimnames = list(paste0("dataset", 1:6),
                              c("gbm", "rf", "logreg", "size_only")))
print(friedman.test(res))
print(PMCMRplus::frdAllPairsNemenyiTest(res))
cat("\nMean ranks (what the test actually compared):\n")
print(colMeans(t(apply(-res, 1, rank))))

cat("\nAll data here is simulated. See Appendix B of the handout.\n")
