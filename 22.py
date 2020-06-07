import numpy as np

N = 4
M = 5

V = np.random.randint(-10, 10, (N, M))
print("Матрица:\r\n{}\n".format(V))

col = [i % 2 for i in np.sum(V, axis=1)]
V = np.insert(V, M, col, axis=1)

print("Новая матрица:\r\n{}\n".format(V))