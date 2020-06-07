import numpy as np

N = 4
M = 5
L = np.random.randint(1, 3)

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

a = np.random.randint(0, 10, M)
print("Допонлительная строка: " + str(a))
print("\n L = " + str(L))

row = np.random.randint(low=-9, high=10, size=M)
print("Строка для вставки: {}".format(row))
V = np.insert(V, L, row, axis=0)

print("Новая матрица:\r\n{}\n".format(V))