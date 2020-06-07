import numpy as np

N = 4
M = 5
L = np.random.randint(1, 3)
K = np.random.randint(1, 3)

V = np.random.randint(-10, 0, (N, M))
print("Матрица:\r\n{}\n".format(V))
x =V[0:L ,0:K]
x_sum = x.sum()
print("Вверхняя левая часть: сумма равна = " + str(x_sum) + "\n" + str(x))
y = V[L: ,0 :K]
y_sum = y.sum()
print("\nНижняя левая часть: сумма равна = " + str(y_sum) + "\n" + str(y))
z = V[0:L ,K:]
z_sum = z.sum()
print("\nВверхняя правая часть: сумма равна = " + str(z_sum) + "\n" + str(z))
a = V[L: ,K:]
a_sum = a.sum()
print("\nНижняя правая часть: сумма равна = " + str(a_sum) + "\n" + str(a))