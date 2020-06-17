from selenium import webdriver
import openpyxl
import time


class Combination(object):
    def __init__(self):
        self.open_chrome()
        self.open_web()
        self.drop_down_menu("/html/body/main/div/section/div/div[1]/article/div[1]/form/div[1]/select[1]/option[118]")
        self.open_excel()
        # pass

    def open_chrome(self):
        #開啟瀏覽器
        self.browser = webdriver.Chrome()

    def open_web(self):
        #開啟網頁
        self.browser.get('https://toolbxs.com/zh-TW/calculator/date')
        #可以等待網頁開好
        # print("sleep 5 seconds.")
        # time.sleep(5)

    def drop_down_menu(self,xpath):
        #選擇選單，直接貼上要選的東西的XPATH
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
print(s.names)
print(s.everyone_2)








# # browser = webdriver.Chrome()
# # browser.get('https://www.google.com')

# wb.close()