import numpy as np

N = 4
M = 5

V = np.random.randint(low=-10, high=10, size=(N, M))
print("Матрица:\r\n{}\n".format(V))

a = b = np.fliplr(V).diagonal(1)
a_prod = a.prod()
print("Элементы которые выше побочной диагонали: \n" + str(a) + "\nИх сумма = " + str(a_prod))
b = np.fliplr(V).diagonal(-1)
b_prod = b.prod()
print("Элементы которые ниже побочной диагонали: \n" + str(b) + "\nИх сумма = " + str(b_prod))