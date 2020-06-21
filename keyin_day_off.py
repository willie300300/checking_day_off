from selenium import webdriver
import openpyxl
import time
from selenium.webdriver.common.keys import Keys  
import os


class Combination(object):
    def __init__(self):
        self.open_chrome()
        self.open_web()
        # self.drop_down_menu("/html/body/center/form/table/tbody/tr/td/table/tbody/tr[5]/td/a[1]")
        self.open_excel()
        # pass

    def open_chrome(self):
        #開啟瀏覽器
        self.browser = webdriver.Chrome()

    def open_web(self):
        #開啟網頁
        self.browser.get('file:///C:/Users/willi/Desktop/%E6%B6%88%E9%98%B2%E5%B1%80e%E5%8C%96%E7%B3%BB%E7%B5%B1.html')
        #可以等待網頁開好
        # print("sleep 5 seconds.")
        time.sleep(5)
        self.browser.find_element_by_xpath("/html/body/center/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[4]/td[2]/select/option[25]").send_keys('15')


    def drop_down_menu(self,xpath):
        #選擇選單，直接貼上要選的東西的XPATH
        #clear 清除元素的内容  
        # send_keys 模拟按键输入
        # click 点击元素
        # submit 提交表单
        self.browser.find_element_by_xpath(xpath).click()

    def open_excel(self):
        #打開EXCEL
        self.wb = openpyxl.load_workbook('109年05月(以此版本為主)(1090521-1308).xlsx')

        #打開工作表(輪休表)
        self.sheet_2 = self.wb['工作表2']
        self.data_2 = self.sheet_2.values

        #儲存輪休表資料及人員姓名
        self.names = []
        self.everyone_2 = {}
        for i in self.data_2:
            if i[1] != None :
                self.everyone_2[i[1]] = i
                self.names.append(i[1])

        # print(self.names)
        # print(self.everyone_2)

    def keyin_data(self):   
        for name in self.names:
            print('正在檢查', name)
            x = 0
            while x < 31:
                x = x + 1
                count = x + 1

                #登打上班日(空格)
                if self.everyone_2[name][count] == None:
                    pass

                #登打休假整天(▲)，要把整天的都改成「▲」三角形
                if self.everyone_2[name][count] == '▲':
                    pass
                    self.drop_down_menu(xpath)

                #登打外宿(○)
                if self.everyone_2[name][count] == '○':
                    pass
                    self.drop_down_menu(xpath)

                #登打日休(●)
                if self.everyone_2[name][count] == '●':
                    pass
                    self.drop_down_menu(xpath)                






s = Combination()
s.open_excel()


# s.browser.find_element_by_xpath("/html/body/form/div/div[2]/div[3]/div/div/div[1]/table/tbody/tr[2]/td[2]/div/input[1]").send_keys("Qq8888....")



# print(s.names1
# print(s.everyone_2)

# a = input("是否已經登入到差勤管理系統(Y/N)")
# if a == "Y":
#     print("Y")
#     # s.browser.find_element_by_xpath('/html/body/center/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/select/option[11]').click()    

#     #获取当前窗口句柄
#     now_handle = s.browser.current_window_handle
#     print(now_handle)

# b = input("是否已經登入到差勤管理系統2(Y/N)")
# if b == "Y":
#     print("Y")    
#     all_handles = s.browser.window_handles 
#     print(all_handles)
#     print(all_handles[0])

# c = input("是否已經登入到差勤管理系統3(Y/N)")
# if c == "Y":    
#     all_handles = s.browser.window_handles 
#     print(all_handles[0])
#     s.browser.switch_to_window(all_handles[0])
#     time.sleep(1)
#     print("看有沒有到這裡")
#     s.browser.find_element_by_xpath('/html/body/center/form/table/tbody/tr/td/table/tbody/tr[3]/td/table/tbody/tr[2]/td[2]/select').send_keys("13")
    
#     print("看有沒有到這裡2")




# # browser = webdriver.Chrome()
# # browser.get('https://www.google.com')

# wb.close()