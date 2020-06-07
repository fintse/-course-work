import random

M = random.randint(1,10)

a = [random.randint(1,10) for i in range(M)]
n = len(a)
print("данный: ", a)

for i in range(0, n - 1):
    c = i
    for k in range(i + 1, n):
        if a[k] < a[c]:
            c = k
        if i != c:
            l = a[i]
            ll = a[c]
            a[i] = ll
            a[c] = l

print("сорт", a)