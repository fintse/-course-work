import numpy as np

N = 4
M = 5
L = np.random.randint(0, 3)
K = np.random.randint(0, 3)

B = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(B))

n = 0
for i in range(0, L):
    for j in range(0, M):
        if B[i,j] == 0:
            n += 1
print("Количество нулевых элементов в верхних L солбцах: {}\n".format(n))
m = 0
for i in range(0, K):
    for j in range(0, N):
        if B[j,i] == 0:
            m += 1
print("Количество нулевых элементов в левых L солбцах: {}\n".format(m))