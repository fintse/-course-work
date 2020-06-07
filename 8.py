import numpy as np

N = 4
M = 5

B = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(B))

X = []
for i in range(0, N):
    K = 0
    for j in range(0, M):
        if B[i,j] < 0:
            K += 1
    X.append(K)
Y = []
for i in range(0, M):
    K = 0
    for j in range(0, N):
        if B[j,i] < 0:
            K += 1
    Y.append(K)
X = np.array(X)[: , np.newaxis]
A = np.hstack((B, X))
Y = np.hstack((Y, [0.]))
A = np.vstack((B, Y))

print("Новая матрица:\r\n{}\n".format(B))