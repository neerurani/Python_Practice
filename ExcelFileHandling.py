'''
Created on 22-06-2019
@author Neeru Rani
'''

import openpyxl
import pandas as pd

#................... Method to Open an excel file....................
def openExcel(filePath):
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

#................... Method to update data in excel file....................
def updateExcel(uniqueColVal,colNameToUpdate,newVal,fpath):
    #uniqueColVal  -->  Unique value to match (of which row's column's value want to change)
    #colNameToUpdate --> column name (of which particular column's value want to change)
    st_obj, w_obj = openExcel(fpath)
    for i in range(1, st_obj.max_row+1):
        for j in range(1, st_obj.max_column + 1):
            cell_obj = st_obj.cell(row=i, column=j)
            if cell_obj.value == uniqueColVal:
                for col in range(1, st_obj.max_column + 1):
                    c1 = st_obj.cell(row=1, column=col)
                    if (colNameToUpdate == c1.value):
                        changeing_pos = st_obj.cell(row=i, column=col)
                        changeing_pos.value=newVal
    w_obj.save(fpath)

path = "C:\\neeru\\Python_Practice\\ItemDetails.xlsx"
updateExcel("ff","DateOfMfd",100,path)
readExcel(path)

#................... Method to delete data from excel file....................
def deleteFromExcel():
    # Instantiating a Workbook object by excel file path
    wb_obj = self.Workbook(self.dataDir + 'Book1.xls')
    # Accessing the first worksheet in the Excel file
    worksheet = workbook.getWorksheets().get(0)

    # Deleting 3rd row from the worksheet
    worksheet.getCells().deleteRows(2, 1, True)

    # Saving the modified Excel file in default (that is Excel 2003) format
    workbook.save(self.dataDir + "Delete Row.xls")

    print
    "Delete Row Successfully."

'''path = "C:\\neeru\\Python_Practice\\ItemDetails.xlsx"
st_obj, w_obj = openExcel(path)
#range(a6).EntireColumn.Delete
#readExcel(path)
for sheet in w_obj.sheets():
    for row in range(sheet.nrows):
        for column in range(sheet.ncols):
            print ("row::::: ", row)
            print ("column:: ", column)
            print ("value::: ", sheet.cell(row,column).value)
            '''



'''st_obj.move_range("D4:B7", rows=-1, cols=2)
wb_obj.save(path)'''


'''data = pd.read_csv("nba.csv", index_col="Name")

# dropping passed values 
data.drop(["Avery Bradley", "John Holland", "R.J. Hunter",
           "R.J. Hunter"], inplace=True)

# display 
data '''