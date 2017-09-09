from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Lib.commonUtils import *
import time
import xlrd
import sys

class Testcase1(UIdriver):
    def __init__(self,brName,username,password,url):
        super(Testcase1,self).__init__(brName,username,password,url)

obj = Testcase1("chrome","admin","manager","http://127.0.0.1/login.do")

#***** get driver

browser = obj.Get_Browser()
#***** Launch Browser
obj.Launch_Application(browser)

#***** Login
if obj.App_Sync(browser,"Login","edt_User Name"):
    #browser.get_screenshot_as_file("launch.png")
    obj.SetFieldValue(browser,"Login","edt_User Name",obj.username)
    obj.SetFieldValue(browser, "Login", "edt_Password", obj.password)
    browser.find_element_by_xpath("//a[@id='loginButton']").click()
else:
    print("Login Fields not displayed")
    sys.exit()

if obj.App_Sync(browser,"Home","lnk_Logout"):
    print("Login Successful")
else:
    print("Login Failed hence quitting the execution")
    sys.exit()
#***** Click on Create New Task
elem = obj.Get_UIObject(browser,"xpath","//a[text()='Create new tasks']")
elem.click()
#time.sleep(5)
obj.Switch_window(browser)
# ***** Select Customer
obj.SetFieldValue(browser,"CreateNewTask","lst_Customer","-- new customer --")
time.sleep(5)
#browser.switch_to.window(main_window_handle)
obj.Switch_window(browser)
elem = obj.Get_UIObject(browser,"xpath","//a[text()='Logout']")
elem.click()
browser.quit()



# if elem.is_displayed():
#     browser.maximize_window()
#     elem.send_keys("admin")
#     browser.find_element_by_xpath("//input[@name='pwd']").send_keys("manager")
#     browser.find_element_by_xpath("//a[@id='loginButton']").click()
#
# #***** Sync
# for Lpc in range(1,60,1):
#     try:
#         elem = browser.find_element_by_xpath("//a[text()='Logout']")
#         if elem is not None:
#             break
#     except:
#         print("element Not Found")
# #***** Click on Create New Task
#
# elem = browser.find_element_by_xpath("//a[text()='Create new tasks']")
# if elem.is_displayed():
#     browser.find_element_by_xpath("//a[text()='Create new tasks']").click()
#     time.sleep(5)
#     main_window_handle = browser.current_window_handle
#     child_window_handle = None
#     while child_window_handle is None:
#         for handle in browser.window_handles:
#             if handle !=main_window_handle:
#                 child_window_handle=handle
#                 break
#     browser.switch_to.window(child_window_handle)
#
# #***** Select Customer
# selectoption = Select(browser.find_element_by_xpath("//select[@name='customerId']"))
# selectoption.select_by_visible_text('-- new customer --')
# #***** Customer name
# browser.switch_to.window(main_window_handle)
# browser.find_element_by_xpath("//a[text()='Logout']").click()

