import numpy as np

N = 4
M = 5

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

Max = V.max(axis=0)
V=V/Max

print("Новая матрица:\r\n{}\n".format(V))