import numpy as np

N = 4
M = 5
K = np.random.randint(1, 3)

V = np.random.randint(-10, 10, (N, M))
print("Матрица:\r\n{}\n".format(V))

M_n = np.sum(V, axis=0) * (-1)
V = np.vstack((V, M_n))

print("Новая матрица:\r\n{}\n".format(V))