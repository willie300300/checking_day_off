from selenium import webdriver

import time
import datetime
from selenium import webdriver
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.keys import Keys
from contextlib import contextmanager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import \
    staleness_of
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException, ElementNotInteractableException, \
    ElementClickInterceptedException
from selenium.webdriver.chrome.options import Options
from helper_class import wait_for_page_load
from pprint import pprint
import json




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