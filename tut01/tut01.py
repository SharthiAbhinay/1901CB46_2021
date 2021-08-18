
l=[12, 14, 56, 78, 98, 54, 678, 134, 789, 0, 7, 5, 123, 45, 76345,987654321]
mer=0
nmer=0
for i in l:
    n=i
    ismerkai=True
    if n/10<1 and n<10:
        mer=mer+1
        print("yes -",i,"is a Meraki value")
        continue
    else:    
        while n>=10:
            j=int(n%10)
            l=int(n/10)
            p=int(l%10)
            k=p-j
            if k<0:
                k=-k
            if k!=1:
                ismerkai=False
                break
            n=int(n/10)
    if ismerkai:
        mer=mer+1
        print("Yes -",i,"is a Meraki value")
    else: 
        nmer=nmer+1
        print("No -",i,"is not a Meraki value") 
print("The input list contains",mer,"Meraki values and",nmer,"Non-meraki values")
