import numpy as np
print("Введите число строк:")
n = int(input())
print("Введите число столбцов:")
m = int(input())
arr = np.random.randint(1,20,(n,m))
arrsu = arr.sum()
print(arrsu)
arrsum = arr.sum(axis=0)
c = arrsu/arrsum
arr1 =  np.vstack((arr, c))
print(arr1)