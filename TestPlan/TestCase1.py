from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Lib.commonUtils import *
from Lib.App_CommonUtils import *
import time
import xlrd
import sys

class Testcase1(App_Common_utils):
    def __init__(self,brName,username,password,url):
        super(Testcase1,self).__init__(brName,username,password,url)

#***** initializing Testcase1
obj = Testcase1("ie","admin","manager","http://127.0.0.1/login.do")
#***** Login
obj.Login_App("Login","edt_User Name","edt_Password","btn_LoginBtn")
#***** create new task
obj.Create_NewTask()
#***** logout
obj.Logout()
#***** Close and Release the browser
obj.browser.quit()


