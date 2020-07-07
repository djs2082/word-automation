n=input("enter no:").split(",")
a=int(n[0])
b=int(n[1])
c=int(n[2])

if(a>b and a>c):
    print(f"{a} is greatest")

elif (b>a and b>c):
    print(f"{b} is greatest")

else:
    print(f"{c} is greatest")