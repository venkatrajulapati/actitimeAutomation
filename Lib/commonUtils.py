from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import time

def Get_Browser(BrName):
    BrName = str(BrName)
    if BrName.lower() == 'chrome':
        browser = webdriver.Chrome()
    elif BrName.lower() == 'ie':
        browser = webdriver.Ie
    else:
        browser = webdriver.Firefox
        browser.find_element_by_id()
        browser.find_element_by_name()
        browser.find_element_by_css_selector()
        browser.find_element_by_link_text()
        browser.find_element_by_xpath()
    browser.maximize_window()
    return browser


def openApplication(driver,Appurl):
    driver.get(Appurl)




def Get_UIObject(driver,Locator,Locatorvalue):
    Locator = str(Locator)
    Element = ""
    if Locator.lower() == 'id':
        Element = driver.find_element_by_id(Locatorvalue)
    elif Locator.lower() == 'name':
        Element = driver.find_element_by_name(Locatorvalue)
    elif Locator.lower()== 'css':
        Element = driver.find_element_by_css_selector(Locatorvalue)
    elif Locator.lower()=='Linktext':
        Element=driver.find_element_by_link_text(Locatorvalue)
    elif Locator.lower()=='xpath':
        Element = driver.find_element_by_xpath(Locatorvalue)

    return  Element
def performAction(driver,objectClass,obj,valuetoSet):
    if str(objectClass).lower()== "inputfield":
        obj.send_keys(valuetoSet)

    elif str(objectClass).lower() == "webelement":













# driver = webdriver.Chrome()
# driver.get()
# ele =driver.find_element_by_xpath(".//td[contains(text(),'Access')]/parent::tr/td[3]/select")
# select = Select(ele)
# select.select_by_value("disabled")

#
# driver.maximize_window()
#
# driver.get("https://www.google.co.in")
#
# time.sleep(1)
#
# elem = driver.find_element_by_id("lst-ib")
#
# elem.send_keys("Python")
# elem.send_keys(Keys.ENTER)
