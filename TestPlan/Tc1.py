from  selenium import webdriver
from Lib.commonUtils import *



driver = Get_Browser('chrome')

openApplication(driver,"https://www.google.co.in/")

elem = Get_UIObject(driver,'id','lst-ib')
performAction(elem,"python")

elem.send_keys("python")
elem.send_keys(Keys.ENTER)

