print(list(range(2,101)))


for i in range(100,1000):
    sum=0
    for l in str(i):
        b=int(l)*int(l)*int(l)
        sum=sum+b
    if i == sum:
        print(i)
