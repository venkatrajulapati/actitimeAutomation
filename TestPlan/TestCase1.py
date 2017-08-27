from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import time
import xlrd

#***** get driver
browser = webdriver.Chrome()
#***** Launch Browser
browser.get("http://127.0.0.1/login.do")

#***** Sync
for Lpc in range(1,60,1):
    try:
        elem = browser.find_element_by_xpath("//input[@name='username']")
        if elem is not None:
            break
    except:
        print("element Not Found")


if elem.is_displayed():
    browser.maximize_window()
    elem.send_keys("admin")
    browser.find_element_by_xpath("//input[@name='pwd']").send_keys("manager")
    browser.find_element_by_xpath("//a[@id='loginButton']").click()

#***** Sync
for Lpc in range(1,60,1):
    try:
        elem = browser.find_element_by_xpath("//a[text()='Logout']")
        if elem is not None:
            break
    except:
        print("element Not Found")
#***** Click on Create New Task

elem = browser.find_element_by_xpath("//a[text()='Create new tasks']")
if elem.is_displayed():
    browser.find_element_by_xpath("//a[text()='Create new tasks']").click()
    time.sleep(5)
    main_window_handle = browser.current_window_handle
    child_window_handle = None
    while child_window_handle is None:
        for handle in browser.window_handles:
            if handle !=main_window_handle:
                child_window_handle=handle
                break
    browser.switch_to.window(child_window_handle)

#***** Select Customer
selectoption = Select(browser.find_element_by_xpath("//select[@name='customerId']"))
selectoption.select_by_visible_text('-- new customer --')
#***** Customer name
browser.switch_to.window(main_window_handle)
browser.find_element_by_xpath("//a[text()='Logout']").click()
