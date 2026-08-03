"""Chapter 17 - A two-layer neural network with no framework.

Forward pass, loss, backpropagation and weight update, written out in full
so the four steps of the training cycle are visible.
"""
import numpy as np

# --- one neuron: verify the worked example ----------------------------
x = np.array([0.8, 0.3, 0.5])
w = np.array([1.2, -0.5, 0.9])
b = -0.4
z = w @ x + b
print(f"one neuron: z = {z:.2f}, relu(z) = {max(0.0, z):.2f}")   # 0.86, 0.86

x2 = np.array([0.2, 0.9, 0.1])
z2 = w @ x2 + b
print(f"other input: z = {z2:.2f}, relu(z) = {max(0.0, z2):.2f}\n")  # -0.52, 0

# --- a tiny two-layer network trained by hand -------------------------
rng = np.random.default_rng(0)
X = rng.normal(size=(500, 3))
y = ((X[:, 0] + X[:, 2] - X[:, 1]) > 0).astype(float).reshape(-1, 1)

W1, b1 = rng.normal(size=(3, 8)) * 0.5, np.zeros((1, 8))
W2, b2 = rng.normal(size=(8, 1)) * 0.5, np.zeros((1, 1))


def train(W1, b1, W2, b2, lr=0.1, epochs=400, label=""):
    for epoch in range(epochs):
        # 1. FORWARD
        h = np.maximum(0, X @ W1 + b1)                  # ReLU hidden layer
        p = 1 / (1 + np.exp(-(h @ W2 + b2)))            # sigmoid output

        # 2. LOSS
        loss = -np.mean(y * np.log(p + 1e-9) + (1 - y) * np.log(1 - p + 1e-9))

        # 3. BACKPROPAGATION
        dz2 = (p - y) / len(X)
        dW2, db2 = h.T @ dz2, dz2.sum(0, keepdims=True)
        dh = dz2 @ W2.T
        dz1 = dh * (h > 0)                              # derivative of ReLU
        dW1, db1 = X.T @ dz1, dz1.sum(0, keepdims=True)

        # 4. UPDATE
        W1 -= lr * dW1; b1 -= lr * db1
        W2 -= lr * dW2; b2 -= lr * db2

        if epoch % 100 == 0:
            print(f"  {label}epoch {epoch:3d}  loss {loss:.4f}")
    acc = ((p > 0.5) == y).mean()
    print(f"  {label}final training accuracy: {acc:.3f}\n")


print("learning rate 0.1 (converges)")
train(W1.copy(), b1.copy(), W2.copy(), b2.copy(), lr=0.1)

print("learning rate 5.0 (watch the loss bounce)")
train(W1.copy(), b1.copy(), W2.copy(), b2.copy(), lr=5.0, label="  ")
