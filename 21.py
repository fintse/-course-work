import numpy as np

N = 4
M = 4

V = np.random.randint(-10,10, (N, M))
print("Матрица:\r\n{}\n".format(V))

G = (V + V.T)/2
print("Новая матрица:\r\n{}\n".format(G))