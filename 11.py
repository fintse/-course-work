import numpy as np

N = 4
M = 5
L = np.random.randint(1, 3)

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

L_arr = np.array(V[L-1, :])
print("L страка: \r\n{}\n".format(L_arr))
V = V + L_arr

print("Новая матрица:\r\n{}\n".format(V))