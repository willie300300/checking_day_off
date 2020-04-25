import openpyxl

#打開EXCEL
wb = openpyxl.load_workbook('109年04月(10904212122).xlsx')

#打開工作表
sheet = wb['工作表3']
data = sheet.values

#儲存超勤核對資料及人員姓名
names = []
everyone = {}
for i in data:
    if i[2] != None :
        everyone[i[2]] = i
        names.append(i[2])

# print(everyone['李汶霖'])
# print(everyone['李汶霖'][7])            

#打開工作表
sheet_2 = wb['工作表2']
data_2 = sheet_2.values

#儲存輪休表資料
everyone_2 = {}
for i in data_2:
    everyone_2[i[1]] = i
# print(everyone_2['李汶霖'])
# print(everyone_2['李汶霖'][3])            


#檢查兩個工作表是否對得起來
def checking(name):
    print('正在檢查',name)
    x = 0
    while x < 30:
        x = x + 1
        count_1 = x + 5
        count_2 = x + 1
        #檢查上班日(空格)
        if everyone[name][count_1] == 24:
            if everyone_2[name][count_2] != None:
                print(x, '號')
                print('錯誤') 
        if everyone_2[name][count_2] != None:
            if everyone[name][count_1] == 24:
                print(x, '號')
                print('錯誤')
        #檢查休假整天(▲)，要把整天的都改成「▲」三角形
        if everyone[name][count_1] == 0:
            if everyone_2[name][count_2] != '▲':
                print(x, '號')
                print('錯誤') 
        if everyone_2[name][count_2] != '▲':
            if everyone[name][count_1] == 0:
                print(x, '號')
                print('錯誤')
        #檢查外宿(○)
        if everyone[name][count_1] == 10:
            if everyone_2[name][count_2] != '○':
                print(x, '號')
                print('錯誤') 
        if everyone_2[name][count_2] != '○':
            if everyone[name][count_1] == 10:
                print(x, '號')
                print('錯誤')
        #檢查日休(●)
        if everyone[name][count_1] == 14:
            if everyone_2[name][count_2] != '●':
                print(x, '號')
                print('錯誤') 
        if everyone_2[name][count_2] != '●':
            if everyone[name][count_1] == 14:
                print(x, '號')
                print('錯誤')                

#一個人一個人核對
for i in names:
    checking(i)

#關閉EXCEL
wb.close()