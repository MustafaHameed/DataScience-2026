"""Chapter 4 - Distance and gradient descent from first principles.

Verifies the two worked examples in the chapter:
  - Euclidean distance between two student vectors, before and after scaling
  - three steps of gradient descent on L(w) = w^2
"""
import numpy as np

# --- vectors and distance ---------------------------------------------
a = np.array([92, 8, 12])   # engaged student:    attendance, quizzes, posts
b = np.array([45, 3, 0])    # disengaged student

print("dot product :", a @ b)
print("norm of a   :", round(float(np.linalg.norm(a)), 3))
print("distance a-b:", round(float(np.linalg.norm(a - b)), 3))   # ~48.8

# How much did each feature contribute? Attendance dominates purely because
# of its range -- which is why scaling is mandatory before any distance method.
contrib = (a - b) ** 2
print("share of squared distance per feature:",
      (contrib / contrib.sum()).round(3))                        # ~[0.93, 0.01, 0.06]

# --- after min-max scaling --------------------------------------------
lo, hi = np.array([0, 0, 0]), np.array([100, 10, 25])
a_s, b_s = (a - lo) / (hi - lo), (b - lo) / (hi - lo)
contrib_s = (a_s - b_s) ** 2
print("\nscaled distance:", round(float(np.linalg.norm(a_s - b_s)), 3))   # ~0.837
print("share per feature now:", (contrib_s / contrib_s.sum()).round(3))   # ~equal

# --- gradient descent on L(w) = w^2 -----------------------------------
print("\ngradient descent, alpha = 0.3")
w, alpha = 3.0, 0.3
for step in range(6):
    grad = 2 * w                       # dL/dw
    w = w - alpha * grad
    print(f"  step {step + 1}: w={w:8.5f}  L={w ** 2:9.6f}")

# Too large a learning rate diverges -- the right-hand panel of Figure 4.4.
print("\nsame problem, alpha = 1.2 (diverges)")
w = 3.0
for step in range(4):
    w = w - 1.2 * (2 * w)
    print(f"  step {step + 1}: w={w:9.4f}  L={w ** 2:11.4f}")
