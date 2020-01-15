'''
Created on 22-06-2019
@author Neeru Rani
'''

import openpyxl

#................... Method to Open an excel file....................
def penExcel(filePath):
    wb_obj = openpyxl.load_workbook(filePath)
    sheet_obj = wb_obj.active
    return sheet_obj,wb_obj

#................... Method to print all data of excel file....................
def readExcel(fpath):
    st_obj,w_obj=openExcel(fpath)
    m_row = st_obj.max_row
    m_col = st_obj.max_column
    for i in range(1, m_row + 1):
        for j in range(1, m_col + 1):
            cell_obj = st_obj.cell(row=i, column=j)
            print(cell_obj.value, end=" ")
        print()

#path = "C:\\neeru\\Python_Practice\\ItemDetails.xlsx"
#ReadExcel(path)

#................... Method to insert data in excel file....................
def insertInExcel(fpath):
    st_obj, w_obj = openExcel(fpath)
    m_row = st_obj.max_row
    m_col = st_obj.max_column
    for j in range(1, m_col + 1):
        cell_obj = st_obj.cell(row=m_row+1, column=j)
        c1=st_obj.cell(row=1, column=j)
        cell_obj.value=input("Enter " + c1.value + " value:- ")
    w_obj.save(fpath)

'''path = "C:\\neeru\\Python_Practice\\ItemDetails.xlsx"
nor = int(input("Enter no. of row that u want to add: "))
for i in range(0,nor):
    insertInExcel(path)
readExcel(path)'''

#................... Method to delete data from excel file....................
def deleteFromExcel():
    