from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
import time
import xlrd

class UIdriver:

    # Common utility functions for python selenium
    # Get the Browser driver

    def __init__(self,brname,UN,Pwd,url):
        self.BrName = brname
        self.username = UN
        self.password = Pwd
        self.url = url

    def Get_Browser(self):

        if str(self.BrName).lower() == 'chrome':
            self.browser = webdriver.Chrome()
        elif self.BrName.lower() == 'ie':
            self.browser = webdriver.Ie
        else:
            self.browser = webdriver.Firefox
        self.browser.maximize_window()
        return self.browser

    # To launch Application

    def Launch_Application(self,driver):

        driver.get(self.url)

    def App_Sync(self,driver,strPageName,strObjectName):

        objArr = []
        objArr = self.Get_Object_ObjectRepository(strPageName, strObjectName)
        objType = objArr.__getitem__(0)
        Locator = objArr.__getitem__(1)
        Locatorval = objArr.__getitem__(2)

        for Lpc in range(1, 60, 1):
            try:
                elem = self.Get_UIObject(driver,Locator,Locatorval)
                if elem is not None:
                    print("element Found")
                    break
            except:
                print("Please wait element Not Found")
        return True

    def Get_UIObject(self,driver,Locator,Locatorvalue):
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
        return Element

    # def ClickObject(self,driver,pagename,objName):
    #     objType = str(objType)
    #     if objType.lower() == "editfield":
    #         elem = self.Get_UIObject(driver,Locator,Locatorval)
    #         if elem.is_displayed:
    #             elem.send_keys(fval)
    #             print(objName + "value is entered as " + fval)
    #         else:
    #             print(objName + " element not displayed")
    #     elif objType.lower() == "dropdown":
    #         elem = Get_UIObject(driver, Locator, Locatorval)
    #         if elem.is_displayed:
    #             selectoption = Select(elem)
    #             selectoption.select_by_visible_text(fval)
    #             print(objName + "value is selected as " + fval)
    #         else:
    #             print(objName + "element not displayed")
    #     elif objType.lower() == "chkbox":
    #         elem = Get_UIObject(driver, Locator, Locatorval)
    #         if elem.is_displayed:
    #             elem.click()
    #             print(objName + "check box is selected")
    #         else:
    #             print(objName + "element not displayed")
    #

    def Switch_window(self,driver):
        main_window_handle = driver.current_window_handle
        child_window_handle = None
        while child_window_handle is None:
             for handle in driver.window_handles:
                 if handle !=main_window_handle:
                     child_window_handle=handle
                     break
             driver.switch_to.window(child_window_handle)
        return driver

    def GetxlColumnNumber(self,oSheet,strColName):
        noofcols = oSheet.ncols
        for c in range(0,noofcols):
            if(oSheet.cell(0,c).value == strColName):
                return c
                break

    def GetxlRowNumber(self,oSheet,strColName,strColumnvalue):
        reqcol = self.GetxlColumnNumber(oSheet, strColName)
        noofrows = oSheet.nrows
        for r in range(0,noofrows):
            if(oSheet.cell(r,reqcol).value == strColumnvalue):
                return r
                break

    def GetxlRowNumberbytwocolvals(self,oSheet,strColName1,strColumnvalue1,strColName2,strColumnvalue12):
        reqcol1 = self.GetxlColumnNumber(oSheet, strColName1)
        reqcol2 = self.GetxlColumnNumber(oSheet, strColName2)
        noofrows = oSheet.nrows
        for r in range(0,noofrows):
            if(oSheet.cell(r,reqcol1).value == strColumnvalue1 and oSheet.cell(r,reqcol2).value == strColumnvalue12):
                return r
                break

    def Get_Object_ObjectRepository(self,strPageName,strObjectName):

        try:
            oWB = xlrd.open_workbook("E:\\actitimeAutomation\\TestPlan\\Actitime objects.xls")
            oSheet = oWB.sheet_by_name("ObjectRepository")
            c1 = self.GetxlColumnNumber(oSheet, "PageName")
            c2 = self.GetxlColumnNumber(oSheet, "ObjectName")
            #r1 = GetxlRowNumber(oSheet, "PageName", strPageName)
            #r2 = GetxlRowNumber(oSheet, "ObjectName", strObjectName)
            r = self.GetxlRowNumberbytwocolvals(oSheet, "PageName", strPageName, "ObjectName", strObjectName)
            st = []
            #if r1==r2:
            Objtypecolnum = self.GetxlColumnNumber(oSheet, "ObjectType")
            ObjLocatorcolnum = self.GetxlColumnNumber(oSheet, "Locator")
            ObjLocatorvalcolnum= self.GetxlColumnNumber(oSheet, "LocatorValue")
            strObjectType = oSheet.cell(r,Objtypecolnum).value
            strLocator = oSheet.cell(r, ObjLocatorcolnum).value
            strLocatorval = oSheet.cell(r, ObjLocatorvalcolnum).value
            st.append(strObjectType)
            st.append(strLocator)
            st.append(strLocatorval)
            #print(st)
            return st
        except:
            print("Failed to load Object repository please check the path")

        #else:
         #   print("object not found please check the PageName and Object Name you are looking for")

    def SetFieldValue(self,driver,strPageName,strObjectName,fval):
        objArr =[]
        objArr= self.Get_Object_ObjectRepository(strPageName, strObjectName)
        objType = objArr.__getitem__(0)
        Locator = objArr.__getitem__(1)
        Locatorval = objArr.__getitem__(2)

        objType = str(objType)
        if objType.lower() == "editfield":
            elem = self.Get_UIObject(driver,Locator,Locatorval)
            if elem.is_displayed:
                elem.send_keys(fval)
                print(strObjectName + "value is entered as " + fval)
            else:
                print(strObjectName+ " element not displayed")
        elif objType.lower() == "dropdown":
            elem = self.Get_UIObject(driver, Locator, Locatorval)
            if elem.is_displayed:
                selectoption = Select(elem)
                selectoption.select_by_visible_text(fval)
                print(strObjectName + "value is selected as " + fval)
            else:
                print(strObjectName + "element not displayed")
        elif objType.lower() == "chkbox":
            elem = self.Get_UIObject(driver, Locator, Locatorval)
            if elem.is_displayed:
                elem.click()
                print(strObjectName + "check box is selected")
            else:
                print(strObjectName + "element not displayed")
