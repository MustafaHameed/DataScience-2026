# Chapter 15 lab -- extracted from parts/ by sync_labs.py. Edit the chapter, not this file.
import warnings
import numpy as np
from scipy.cluster.hierarchy import linkage, fcluster
from sklearn.datasets import make_blobs
from sklearn.cluster import KMeans, AgglomerativeClustering
from sklearn.metrics import silhouette_score, adjusted_rand_score
from sklearn.preprocessing import StandardScaler

warnings.filterwarnings("ignore", message="KMeans is known to have a memory")

# --- 1. the worked example: k-means from A2 and A4 -----------------------
P = np.array([[1, 1], [1.5, 2], [3, 4], [5, 7], [3.5, 5], [4.5, 5],
              [3.5, 4.5]])
km = KMeans(n_clusters=2, init=P[[1, 3]], n_init=1).fit(P)
print("centroids:", km.cluster_centers_.round(3).tolist(),
      " WCSS:", round(km.inertia_, 3), " iterations:", km.n_iter_)

# --- 2. choosing k: the elbow and the silhouette -------------------------
X, truth = make_blobs(n_samples=600, centers=4, cluster_std=1.0,
                      random_state=3)     # 600 projects of four kinds
for k in range(2, 8):
    m = KMeans(n_clusters=k, n_init=10, random_state=0).fit(X)
    print(f"k={k}: WCSS {m.inertia_:8.1f}   "
          f"silhouette {silhouette_score(X, m.labels_):.3f}")

# --- 3. hierarchical clustering: linkage heights on five numbers ---------
x = np.array([[1.0], [2.0], [4.0], [9.0], [11.0]])
for method in ("single", "complete", "average"):
    Z = linkage(x, method)
    print(f"{method:8s} merge heights: {Z[:, 2].round(2).tolist()}")
print("complete linkage cut into 2 groups:",
      fcluster(linkage(x, "complete"), 2, criterion="maxclust").tolist())

# --- 4. against the truth, and why scale matters -------------------------
km4 = KMeans(4, n_init=10, random_state=0).fit_predict(X)
hac = AgglomerativeClustering(n_clusters=4, linkage="ward").fit(X)
print("adjusted Rand, k-means:", round(adjusted_rand_score(truth, km4), 3),
      " ward HAC:", round(adjusted_rand_score(truth, hac.labels_), 3))
Xs = X.copy()
Xs[:, 1] *= 100                                   # one feature in other units
raw = KMeans(4, n_init=10, random_state=0).fit_predict(Xs)
fix = KMeans(4, n_init=10, random_state=0).fit_predict(
    StandardScaler().fit_transform(Xs))
print("rescaled feature, adjusted Rand: raw",
      round(adjusted_rand_score(truth, raw), 3), " standardised",
      round(adjusted_rand_score(truth, fix), 3))
