import numpy as np

N = 4
M = 5
K = np.random.randint(1, 3)

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

K_arr = np.array(V[:, K-1])
K_arr = K_arr[: , np.newaxis]
print("K-ый столбец: \r\n{}\n".format(K_arr))
V = V * K_arr

print("Новая матрица:\r\n{}\n".format(V))