n=int(input("give the number")) 
ismerkai=True
if n/10<1 and n<10:
    print("yes")
    quit()
else:    
    while n>=10:
        j=int(n%10)
        l=int(n/10)
        i=int(l%10)
        k=i-j
        if k<0:
            k=-k
        if k!=1:
            ismerkai=False
            break
        n=int(n/10)
if ismerkai:
    print("yes")
else: 
    print("no") 