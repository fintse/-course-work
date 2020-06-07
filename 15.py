import numpy as np

N = 4
M = 5
H = np.random.randint(1, 4)

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

a = []
b = []
for i in range(M):
    if H in V[:, i]:
        a.append(i+1)
    else:
        b.append(i+1)
print("Столбцы, которые имеют хотя бы одно число H - {}\n".format(a))
print("Столбцы, которые не имеют это число - {}\n".format(b))