import openpyxl

# def get_data_openpyxl(file_name,sheet):
#     #打開excel文檔
#     wb = openpyxl.load_workbook(file_name)
#     #訪問sheet頁
#     sheet = wb[sheet]
#     # 獲取excel文檔所有的數據,返回的是一個generator對象
#     data = sheet.values
#     print(data)
#     #叠代輸出所有數據
#     for i in data:
#         print(i)
#     wb.close()
# get_data_openpyxl('109年04月.xlsx','工作表2')



wb = openpyxl.load_workbook('109年04月(10904212122).xlsx')
sheet = wb['工作表3']
data = sheet.values
# print(data)
names = []
everyone = {}
for i in data:
    # print(i)
    if i[2] != None :
        everyone[i[2]] = i
        names.append(i[2])

print(everyone['李汶霖'])
print(everyone['李汶霖'][7])            


sheet_2 = wb['工作表2']
data_2 = sheet_2.values
# print(data)
everyone_2 = {}
for i in data_2:
    # print(i)
    everyone_2[i[1]] = i
print(everyone_2['李汶霖'])
print(everyone_2['李汶霖'][3])            


x = 0
while x < 30:
    x = x + 1
    print(x, '號')
    count_1 = x + 5
    count_2 = x + 1
    if everyone['李汶霖'][count_1] == 24:
        if everyone_2['李汶霖'][count_2] == None:
            print('正確')
        else:
            print('錯誤')     
    else:
        print('無檢查')

print("===================================")
x = 0
while x < 30:
    x = x + 1
    count_1 = x + 5
    count_2 = x + 1
    if everyone['李汶霖'][count_1] == 24:
        if everyone_2['李汶霖'][count_2] != None:
            print(x, '號')
            print('錯誤')     

print('===============================')
def checking(name):
    print('正在檢查',name)
    x = 0
    while x < 30:
        x = x + 1
        count_1 = x + 5
        count_2 = x + 1
        if everyone[name][count_1] == 24:
            if everyone_2[name][count_2] != None:
                print(x, '號')
                print('錯誤') 
        if everyone_2[name][count_2] != None:
            if everyone[name][count_1] == 24:
                print(x, '號')
                print('錯誤')
# checking('李汶霖')

for i in names:
    checking(i)





# x = 0
# while x < 30:
#     x = x + 1
#     print(x, '號')
#     count_1 = x + 5
#     count_2 = x + 1
#     if everyone['李汶霖'][count_1] == 24:
#         if everyone_2['李汶霖'][count_2] == None:
#             print('正確')
#         else:
#             print('錯誤')     
#     else:
#         print('無檢查')

# everyone = {}

# for i in data:
#     # print(i)
#     everyone[i[1]] = i

# print(everyone['李汶霖'])
# print(everyone['李汶霖'][0])
# print(everyone['李汶霖'][1])
# print(everyone['李汶霖'][2])



wb.close()