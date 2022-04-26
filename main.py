from selenium import webdriver

class Main:

    def __init__(self):
        wb = webdriver.Chrome()
        wb.get('https://www.baidu.com')