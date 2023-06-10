height = input("身高：")
height = float(height) / 100
print("恭喜您！您的身高约为："+str(height)+"米")
# print("身高是否高于180:"+str(height>1.8))
if height>=1.8:
    if height>1.8:
        print("恭喜您！您的身高大于1.8米！")
    else:
        print("恭喜您！您的身高等于1.8米！")
elif height==1.8:
    print("2恭喜您！您的身高等于1.8米！")

else:
    print("很遗憾！您的身高小于1.8米！")