'''
Created on 19-06-2019
@author Neeru Rani
'''

# description -> method to define bubble sort
def bubble_sort(val):
    l=len(val)
    for i in range(l-1,0,-1):
        for j in range(i):
            if val[j] > val[j+1]:
                temp = val[j]
                val[j] = val[j+1]
                val[j+1] = temp
    print(val)

num = [0,3,8,5,2,-2]
bubble_sort(num)
