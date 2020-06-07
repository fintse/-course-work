import numpy as np

N = 4
M = 5
K = np.random.randint(1, 3)

V = np.random.randint(-10, 10, (N, M))
print("Матрица:\r\n{}\n".format(V))
print("K = " + str(K))
V = np.delete(V, (K-1), axis=1)
print("Новая матрица:\r\n{}\n".format(V))