from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Lib.commonUtils import *
import time
import xlrd
import sys

class App_Common_utils(UIdriver):

    def __init__(self):
        super(App_Common_utils,self).__init__()

    def Login_App(self,odata,TCID,BSID):
        driver = self.Get_Browser()
        self.Launch_Application(driver)
        if self.App_Sync(driver,"Login","edt_User Name"):
            self.SetFieldValue(driver,"Login","edt_User Name",self.username)
            self.SetFieldValue(driver,"Login","edt_Password",self.password)
            self.ClickObject(driver,"Login","btn_LoginBtn")
            print("Application Login Successful")
        else:
            print("Failed to load Login page")
            sys.exit()

    def Create_NewTask(self,odata,TCID,BSID):
        #odata=str(odata).replace("'","")
        owb = xlrd.open_workbook("E:\\actitimeAutomation\\TestData\\" + odata + ".xls")
        oDataset = owb.sheet_by_index(0)
        reqrow = self.GetxlRowNumberbytwocolvals(oDataset, "TC_ID", TCID, "BS_ID", BSID)
        print(oDataset.nrows)
        if self.App_Sync(self.browser, "Home", "lnk_CreateNewtasK"):
            self.ClickObject(self.browser,"Home","lnk_CreateNewtasK")
            self.Switch_window(self.browser)
            if self.App_Sync(self.browser, "CreateNewTask", "lst_Customer"):
                customercol = self.GetxlColumnNumber(oDataset,"Customer")
                it = oDataset.cell(reqrow,customercol).value
                self.SetFieldValue(self.browser, "CreateNewTask", "lst_Customer", it)
                time.sleep(4)
                self.Switch_window(self.browser)
            else:
                print("Failed to Load Create New Task Popup window")
        else:
            print("Failed to load welcome page")

    def Logout(self,odata,TCID,BSID):
        try:

            self.ClickObject(self.browser,"Home","lnk_Logout")

            if self.App_Sync(self.browser,"Login","edt_User Name"):
                print("Application Successfully Logged out")
            else:
                print("Logout failed")

        except:
            print('Failed to click logout')
