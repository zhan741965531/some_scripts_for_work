import xlrd
import xlwt
import os

workbook1 = xlrd.open_workbook("a.xls")
workbook2 = xlrd.open_workbook("b.xls")


sheet1 = workbook1.sheet_by_index(0)
sheet2 = workbook2.sheet_by_index(0)

nrows1 = sheet1.nrows
nrows2 = sheet2.nrows

list1 = []
list2 = []

for x in range(nrows1):
    list1.append(sheet1.cell_value(x,1))
for y in range(nrows2):
    list2.append(sheet2.cell_value(y,0))

writebook = xlwt.Workbook(encoding="utf-8")
table = writebook.add_sheet("test")
count = 0
for index1,item1 in enumerate(list1):
    for index2,item2 in enumerate(list2):
        try:
            if item1 == item2:
                print("第",count,"行")
                print(index1,":",item1,"==>",index2,":",item2)
                table.write(count,0,sheet1.cell_value(index1,0))
                table.write(count,1,sheet1.cell_value(index1,1))
                table.write(count,2,sheet1.cell_value(index1,2))
                table.write(count,3,sheet1.cell_value(index1,3))
                table.write(count,4,sheet1.cell_value(index1,4))
                table.write(count,5,sheet1.cell_value(index1,5))
                table.write(count,6,sheet1.cell_value(index1,6))
                table.write(count,7,sheet1.cell_value(index1,7))
                table.write(count,8,sheet1.cell_value(index1,8))
                table.write(count,9,sheet1.cell_value(index1,9))
                table.write(count,10,sheet2.cell_value(index2,0))
                table.write(count,11,sheet2.cell_value(index2,1))
                table.write(count,12,sheet2.cell_value(index2,2))
                table.write(count,13,sheet2.cell_value(index2,3))
                table.write(count,14,sheet2.cell_value(index2,4))
                table.write(count,15,sheet2.cell_value(index2,5))
                table.write(count,16,sheet2.cell_value(index2,6))
                count = count + 1
                writebook.save("results.xls")
        finally:
            continue



  


        

