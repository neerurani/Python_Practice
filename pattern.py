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

def patternPiramid(noOfRow,char):
    n=int(noOfRow)
    for i in range(0,n):
        for j in range(0,n-i-1):
            print(end=" ")
        for k in range(0,i+1):
            print(char,end=" ")
        print() # for next line

nor=input("Enter the no. of rows: ")
patternPiramid(nor,"*")