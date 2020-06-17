from selenium import webdriver





class Combination(object):
    def __init__(self):
        self.open_chrome()
        self.open_web()

    def open_chrome(self):
        #開啟瀏覽器
        self.browser = webdriver.Chrome()

    def open_web(self):
        #開啟網頁
        self.browser.get('https://toolbxs.com/zh-TW/calculator/date')


s = Combination()
# browser = webdriver.Chrome()
# browser.get('https://www.google.com')