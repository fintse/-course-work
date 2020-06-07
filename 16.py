import numpy as np

N = 4
M = 5
L = np.random.randint(1, 3)

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

print("L = " + str(L))
V = np.delete(V, (L-1), axis=0)

print("Новая матрица:\r\n{}\n".format(V))