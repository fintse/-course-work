import numpy as np

N = 4
M = 5

V = np.random.randint(-10, 10, (N, M))
print("Матрица:\r\n{}\n".format(V))

iu = np.triu_indices(N, 1)
a = V[iu]
a = np.sum(np.array(a))
print("\nCумма элементов выше главной диагонали = " + str(a))
b = np.fliplr(V)[iu]
b = np.prod(np.array(b))
print("\nПроизведение элементов выше побочной диагонали = " + str(b))