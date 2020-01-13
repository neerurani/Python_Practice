'''
Created on 10-06-2019
@author Neeru Rani
'''
''' Program to print 
   *
  * *
 * * *
* * * *
pattern '''

def piramidPattern(noOfRow,char):
    n=int(noOfRow)
    for i in range(0,n):
        for j in range(0,n-i-1):
            print(end=" ")
        for k in range(0,i+1):
            print(char,end=" ")
        print() # for next line

nor=input("Enter the no. of rows: ")
#piramidPattern(nor,"*")

''' Program to print 
*
**
***
****
*****
   pattern  '''
def ladderPattern(noOfRow,char):
    n=int(noOfRow)
    for i in range(0,n+1):
        for j in range(0,i+1):
            print(char,end=" ")
        print()

#ladderPattern(nor,"#")

'''  Program to print
1
12
123
1234
12345
   pattern  '''
def digitPattern1(noOfRow):
    n=int(noOfRow)
    for i in range(1,n+1):
        for j in range(1,i+1):
            print(j,end=" ")
        print()

digitPattern1(5)