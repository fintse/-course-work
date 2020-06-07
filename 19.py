import numpy as np

N = 4
M = 5

V = np.random.randint(low=-9, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

a = np.diagonal(V, 1)
a_sum = a.sum()
print("Элементы которые выше главной диагонали: \n" + str(a) + "\nИх сумма = " + str(a_sum))
b = np.diagonal(V, -1)
b_sum = b.sum()
print("Элементы которые ниже главной диагонали: \n" + str(b) + "\nИх сумма = " + str(a_sum))