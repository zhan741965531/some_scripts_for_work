import xlrd
import xlwt

asset_xls = xlwt.Workbook()
table2 = asset_xls.add_sheet("Sheet1",cell_overwrite_ok=True)
old_data = xlrd.open_workbook("test.xlsx")
table = old_data.sheet_by_name("Sheet2")

asset_name  = table.col_values(0)
yewuxitong_name = table.col_values(1)
ip = table.col_values(2)
asset_type = table.col_values(3)
server_fuzeren = table.col_values(4)
system_fuzeren = table.col_values(5)
sql_fuzeren = table.col_values(6)

list_all = []
for x in range(len(yewuxitong_name)):
    list_info = []
    list_info.append(asset_name[x])
    list_info.append(yewuxitong_name[x])
    list_info.append(ip[x])
    list_info.append(asset_type[x])
    list_info.append(server_fuzeren[x])
    list_info.append(system_fuzeren[x])
    list_info.append(sql_fuzeren[x])
    list_all.append(list_info)

list_ip = []
for index1 in range(len(list_all)):
    list_ip.append(list_all[index1][2])
list_ip = list(set(list_ip))

def print_inf0(data_type,data_info):
    flags = 0
    if "主机" in str(data_type): #I资产类型
        flags = 1
        print(data_info)
        print("--------------------------------------------------------------------------------------------------------------------")
        with open("log.txt","a+",encoding="utf-8") as f:
            f.write(str(data_info))
            f.write("\n")
    elif "应用" in str(data_type) or "中间件" in str(data_type):
        flags = 2
        print(data_info)
        print("---------------------------------------------------------------------------------------------------------------------")
        with open("log.txt","a+",encoding="utf-8") as f:
            f.write(str(data_info))
            f.write("\n")
    elif "数据库" in str(data_type):
        flags = 3
        print(data_info)
        print("---------------------------------------------------------------------------------------------------------------------")
        with open("log.txt","a+",encoding="utf-8") as f:
            f.write(str(data_info))
            f.write("\n")
    else:
        flags = 4
        print(data_info)
        with open("log.txt","a+",encoding="utf-8") as f:
            f.write(str(data_info))
            f.write("\n")
    return flags,index2

def zhuji(index_set):#4==>应用，5==>主机，6==>数据库
    for index0 in index_set:
        list_update = []
        if list_all[index0][5] != '':
            info = list_all[index0][2] + "-主机脚本生成  " +list_all[index0][1] +"  " + list_all[index0][2] + " 2 /主机/Linux  " + list_all[index0][5]
            print(info)
            list_update.append(list_all[index0][2]+"-主机--20210205")
            list_update.append(list_all[index0][1])
            list_update.append(list_all[index0][2])
            list_update.append("2 /主机/Linux")
            list_update.append(list_all[index0][5])
            with open("test.txt","a+",encoding="utf-8") as f1:
                f1.write(info)
                f1.write("\n")
            print("创建主机")
    return list_update

def yingyong(index_set):#4==>应用，5==>主机，6==>数据库
    for index0 in index_set:
        list_update = []
        if list_all[index0][4] != '':
            info = list_all[index0][2] + "-应用-脚本生成  " +list_all[index0][1] +"  " + list_all[index0][2] + " 63 /中间件/weblogic " + list_all[index0][4]
            print(info)
            list_update.append(list_all[index0][2]+"-应用--20210205")
            list_update.append(list_all[index0][1])
            list_update.append(list_all[index0][2])
            list_update.append("61 /中间件")
            list_update.append(list_all[index0][4])
            with open("test.txt","a+",encoding="utf-8") as f1:
                f1.write(info)
                f1.write("\n")
    print("创建应用")
    return list_update

def shujuku(index_set):   #4==>应用，5==>主机，6==>数据库
    for index0 in index_set:
        list_update = []
        if list_all[index0][6] != '':
            info = list_all[index0][2] + "-数据库-脚本生成  " +list_all[index0][1] +"  " + list_all[index0][2] + "  53 /数据库/mysql " + list_all[index0][6]
            print(info)
            list_update.append(list_all[index0][2]+"-数据库--20210205")
            list_update.append(list_all[index0][1])
            list_update.append(list_all[index0][2])
            list_update.append("50 /数据库")
            list_update.append(list_all[index0][6])
            with open("test.txt","a+",encoding="utf-8") as f1:
                f1.write(info)
                f1.write("\n")
    print("创建数据库")
    return list_update

def judge(*flags):
    flags = flags
    list_all_update = []
    for data in flags:
        flag_Set = []
        index_set = []
        for index,data2 in enumerate(data):
            flag = str(data2[0])
            flag_Set.append(flag)
            index_set.append(data2[1])
        flag_Set = str(flag_Set)
        print(flag_Set)
        print(index_set)
        for lunci in range(3):
            if lunci == 0:
                if "1" not in flag_Set:#主机
                    list_all_update.append(zhuji(index_set))
            if lunci == 1:
                if "2" not in flag_Set:#应用
                    list_all_update.append(yingyong(index_set))
            if lunci == 2:
                if "3" not in flag_Set:#数据库
                    list_all_update.append(shujuku(index_set))
            if lunci == 3:
                list_all_update.append(zhuji(index_set))
                list_all_update.append(yingyong(index_set))
                list_all_update.append(shujuku(index_set))
            else:
                pass
    return list_all_update

sum = 0
for ip_index,ip in enumerate(list_ip):
    print("==================================================================================================================")
    flags = []
    for index2 in range(len(list_all)):
        if ip  == list_all[index2][2]: #ip
            flags.append(print_inf0(list_all[index2][3],list_all[index2]))
            print(flags)
        else:
            pass
    list_all_update = list(judge(list(flags)))
    print(list_all_update)
    print("====",ip_index,"===")
    for data in list_all_update:
        if data != []:
            table2.write(sum,0,data[0])
            table2.write(sum,1,data[1])
            table2.write(sum,2,data[2])
            table2.write(sum,3,data[3])
            table2.write(sum,4,data[4])
            print(data[0])
            count = 1
            ip_index
            print(count)
            sum = sum + count
            print(sum)
        else:
            pass
        
asset_xls.save("资产生成-20210205-初始-v2.xls")