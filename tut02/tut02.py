def get_memory_score( a):
    check=[]
    score=0
    p=False
    for i in a:
        for j in check:
            if(j==i):
               p=True
        if(p):
            score=score+1
            p=False
        else:
            check.append(i)           
    return score




input_nums=["3","4","1","6","3","3","9","0","0","0"]
invalid=[]
res=True
for i in input_nums:
    if(i.isdigit()) :
        continue
    else :
       res=False 
       invalid.append(i) 

if(res==True):
    print("Score: ",get_memory_score(input_nums))
else :
    print("Please enter a valid input list") 
    print("Invalid inputs detected",invalid)   