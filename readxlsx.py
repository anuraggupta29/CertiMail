# Reading an excel file using Python
import xlrd

dataArr =[]

# Give the location of the file
loc = ("robotics.xlsx")

# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# get no of rows
rows = sheet.nrows

#store name,event,email in array
for i in range(rows):
    dataArr.append((sheet.cell_value(i, 1), sheet.cell_value(i, 6), sheet.cell_value(i, 5)))

print(dataArr)
