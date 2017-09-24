from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Lib.commonUtils import *
import time
import xlrd
import sys

class App_Common_utils(UIdriver):

    def __init__(self,brName,username,password,url):
        super(App_Common_utils,self).__init__(brName,username,password,url)

    def Login_App(self,Pagename,Userobj,Passwordobj,Loginbtnobj):
        driver = self.Get_Browser()
        self.Launch_Application(driver)
        if self.App_Sync(driver,Pagename,Userobj):
            self.SetFieldValue(driver,Pagename,Userobj,self.username)
            self.SetFieldValue(driver,Pagename,Passwordobj,self.password)
            self.ClickObject(driver,Pagename,Loginbtnobj)
            print("Application Login Successful")
        else:
            print("Failed to load Login page")
            sys.exit()

    def Create_NewTask(self):

        if self.App_Sync(self.browser, "Home", "lnk_CreateNewtasK"):
            self.ClickObject(self.browser,"Home","lnk_CreateNewtasK")
            self.Switch_window(self.browser)
            if self.App_Sync(self.browser, "CreateNewTask", "lst_Customer"):
                self.SetFieldValue(self.browser, "CreateNewTask", "lst_Customer", "-- new customer --")
                self.Switch_window(self.browser)
            else:
                print("Failed to Load Create New Task Popup window")
        else:
            print("Failed to load welcome page")

    def Logout(self):
        try:
            self.ClickObject(self.browser,"Home","lnk_Logout")
            if self.App_Sync(self.browser,"Login","edt_User Name"):
                print("Application Successfully Logged out")
            else:
                print("Logout failed")
        except:
            print('Failed to click logout')





