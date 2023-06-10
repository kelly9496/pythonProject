#1-100的奇数求和

sum=0
a=1
while True:
    if a%2==0:
        a=a+1
        continue
    sum=sum+a
    a=a+1
    if a==100:
        break
print(sum)