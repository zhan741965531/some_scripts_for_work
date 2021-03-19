import xlrd
import os
import glob
from datetime import datetime

now = datetime.now()
count1 = 0

def timenumber(index1):
    if index1 // 6 > 1 :
        index1 = index1 - 6*(index1 // 6)
    else:
        pass
    return index1


x = glob.glob('*.xls*')
for file in x:
    path2 = os.path.abspath(file)
    data = xlrd.open_workbook(path2)
    names = data.sheet_names()

    for name in names:
        table = data.sheet_by_name(name)
        nrows = table.nrows
        list1 = []
        list2 = []
       
    for i in range(nrows):
        list1.append(str(table.cell_value(i,0)))
        list2.append(str(table.cell_value(i,1)))
        list3 = set(list1)
    
    with open('scans.txt','a+',encoding='utf-8') as f1:
        for index1,item1 in enumerate(list3):
            date = "|"+ str(now.year)+"-"+str(now.month)+"-"+str(now.day+1)+" 0"+str(timenumber(index1))+":00:00"
            f1.write(item1 + date + "\n")
            item1 = str(item1)
            item1 = item1.replace(' ','')
            item1 = item1.replace('\\n','-')
            item1 = item1.replace('/','-')
            item1 = item1.replace('（','')
            item1 = item1.replace('）','')
            fname = item1 + '.txt'
            count = 0
            for index2,item2 in enumerate(list1):
                if item1 == item2:
                  with open(fname,'a+') as f2:
                    print(item2,"===>",index2,str(table.cell_value(index2,1)))
                    f2.write(str(table.cell_value(index2,1))+"\n")
                    count = count + 1
            with open("log.txt","a+") as f3:
                f3.write(str(item1)+"资产数量："+str(count)+"\n")
            print(item1,"资产数量：",count)

   