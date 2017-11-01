from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import urllib
import time
import xlrd
import os
from datetime import datetime
import sys


class UIdriver(object):

    # Common utility functions for python selenium
    # Get the Browser driver

    def __init__(self):

        oWB = xlrd.open_workbook("E:\\actitimeAutomation\\TestData\\MasterData.xls")
        oSheet = oWB.sheet_by_name("MasterData")
        URL = self.GetxlColumnNumber(oSheet, "URL")
        BR = self.GetxlColumnNumber(oSheet, "Browser")
        UN = self.GetxlColumnNumber(oSheet, "UserName")
        PWD = self.GetxlColumnNumber(oSheet, "Password")
        Rootpath = self.GetxlColumnNumber(oSheet, "Rootpath")
        # self.BrName = brname
        # self.username = UN
        # self.password = Pwd
        # self.url = url
        self.BrName = str(oSheet.cell(1,BR).value)
        self.username = str(oSheet.cell(1,UN).value)
        self.password = str(oSheet.cell(1,PWD).value)
        self.url = str(oSheet.cell(1,URL).value)
        self.Rootpath = str(oSheet.cell(1,Rootpath).value)
        #self.gbl_intScreenCount = 0
        self.Reportfile = ""
        self.screenshotfolder = ""
    def Get_Browser(self):

        if str(self.BrName).lower() == 'chrome':
            self.browser = webdriver.Chrome()
        elif self.BrName.lower() == 'ie':
            capabilities = DesiredCapabilities.INTERNETEXPLORER


            capabilities["ignoreProtectedModeSettings"] = True


            capabilities["ignoreZoomSetting"] = True
            urllib.getproxies = lambda: {}


            self.browser = webdriver.Ie(capabilities=capabilities,
                                   executable_path="C:\\Users\\RAJULAPATI\\AppData\\Local\\Programs\\Python\\Python35-32\\Scripts\\IEDriverServer.exe")


        else:
            self.browser = webdriver.Firefox()
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
            print("Waiting for the Object " + strPageName + "." + strObjectName)
            try:
                elem = self.Get_UIObject(driver,Locator,Locatorval)
                if elem is not None:
                    print(strPageName + "." + strObjectName + " object Found")
                    break
            except:
                print("Please wait element Not Found")
        return True

    def Get_UIObject(self,driver,Locator,Locatorvalue):
        Locator = str(Locator)
        Element = ""
        try:
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
        except:
            print("unable to find the Locator : "+ Locatorvalue)

    def ClickObject(self,driver,pagename,objName):
        objArr = []
        objArr = self.Get_Object_ObjectRepository(pagename,objName)
        objType = objArr.__getitem__(0)
        Locator = objArr.__getitem__(1)
        Locatorval = objArr.__getitem__(2)
        elem = self.Get_UIObject(driver,Locator,Locatorval)

        if elem is not None:
            elem.click()
            print("Clicked on the Object : " + pagename + "." + objName)
        else:
            print("element not found please check the object description")


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
            #c1 = self.GetxlColumnNumber(oSheet, "PageName")
            #c2 = self.GetxlColumnNumber(oSheet, "ObjectName")
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

    def Create_HTML_Report(self,TCName):

        g_tStart_Time = datetime.now()

        # Name of Report-folders and Report-File-Name for this Run
        arrStartTime = str(g_tStart_Time).split(" ")
        strname1 = arrStartTime[0]
        strname1 = strname1.replace("-", "")
        print(strname1)
        strname2 = arrStartTime[1]
        strname2 = strname2.replace(":", "")
        strname2 = strname2.split(".")
        strname2 = strname2[0]
        strname = strname1 + "_" + strname2
        print(strname)
        strEnvironment = ""
        Rp = self.Rootpath
        if not os.path.exists(Rp + "Results"):
            os.mkdir(Rp + "Results")

        #TCName = "Dummy"
        ReportFolder = Rp + "Results\\" + TCName + "_" + strname

        if not os.path.exists(ReportFolder):
            os.mkdir(ReportFolder)

        Reportfile = Rp + "Results\\" + TCName + "_" + strname + ".html"
        screenshotfolder = Rp + "Results\\" + TCName + "_" + strname

        self.Reportfile = Reportfile
        self.screenshotfolder = screenshotfolder
        if not os.path.exists(screenshotfolder):
            os.mkdir(screenshotfolder)

        resfile = open(Reportfile, "a")
        # Write header
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        Test_Automation_Test_Report_Logo = Rp + "Logo.png"
        dttime = datetime.now()
        dttime = str(dttime)
        # Write Report - Header
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write(
            "<TR COLS=2><TD BGCOLOR=WHITE WIDTH=6%><IMG SRC='" + Test_Automation_Test_Report_Logo + "'></TD><TD WIDTH=100% BGCOLOR=WHITE><FONT FACE=VERDANA COLOR=NAVY SIZE=4><B>&nbspactiTime Test Automation Results - [" + dttime + "] </B></FONT></TD></TR></TABLE>")
        resfile.write("<TABLE BORDER=1 BGCOLOR=BLACK CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write("</TABLE></BODY></HTML>")

        # Write Report - Test-Set Name OR Test-Script Name
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write("<TR COLS=1>" \
                      "<TD ALIGN=LEFT BGCOLOR=#66699><FONT FACE=VERDANA COLOR=WHITE SIZE=3><B>" + TCName + "</BR>" + "</B></FONT></TD>" \
                                                                                                                     "</TR>")
        resfile.write("</TABLE></BODY></HTML>")

        # Write Report - Column Headers
        resfile.write("<HTML><BODY><TABLE BORDER=1 CELLPADDING=3 CELLSPACING=1 WIDTH=100%>")
        resfile.write("<TR COLS=4>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=20%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Test Step</B></FONT></TD>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Expected Result</B></FONT></TD>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99 WIDTH=30%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Actual Result</B></FONT></TD>" \
                      "<TH ALIGN=MIDDLE BGCOLOR=#FFCC99   WIDTH=7%><FONT FACE=VERDANA COLOR=BLACK SIZE=2><B>Step-Result</B></FONT></TD>" \
                      "</TR>")
        return resfile
        #resfile.close()



    def fn_HtmlReport_TestStep(self,strRepfilepath,strScreenshotfolder,gbl_intScreenCount,strDesc, strExpected, strActual, strResult):

        #***** Set Result parameters
        if str(strResult).upper() == "PASS":
            strResultColor = "GREEN"
            strResultSign = "P"
            blnCaptureImsge = True
        elif str(strResult).upper() == "FAIL":
            strResultColor = "RED"
            strResultSign = "O"
            blnCaptureImsge = True
        else:
            blnCaptureImsge = False
            strResultColor = "GREEN"
            strResultSign = "P"
            strActualHREF = strActual
        #Set Image Path and capture image
        if (blnCaptureImsge == True):
            #gbl_intScreenCount = gbl_intScreenCount + 1
            #Capture Image
            strImagePath = strScreenshotfolder + "\\Screen_000" + str(gbl_intScreenCount) + ".png"
            self.browser.get_screenshot_as_file(strImagePath)
            strActualHREF = "<A HREF='" + strImagePath + "'>" + strActual + "</A>"

        elif blnCaptureImsge == "False":
            strActualHREF = "<A>" + strActual + "</A>"
        #Update HTML Report
        if not strExpected is None:
            strRepfilepath.write("<TR COLS=4>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=2>" + strDesc + "</FONT></TD>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=VERDANA SIZE=2>" + strExpected + "</FONT></TD>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=30%><FONT FACE=WINGDINGS SIZE=4>2</FONT><FONT FACE=VERDANA SIZE=2>" + strActualHREF + "</FONT></TD>"\
            "<TD ALIGN=MIDDLE BGCOLOR=#EEEEEE WIDTH=7%><FONT FACE='WINGDINGS 2' SIZE=5 COLOR=" + strResultColor + ">" + strResultSign + "</FONT><FONT FACE=VERDANA SIZE=2 COLOR=" + strResultColor + "><B>" + strResult + "</B></FONT></TD>"\
            "</TR>")
        if strExpected is None:
            strRepfilepath.write("<TR COLS=4>"\
            "<TD BGCOLOR=#EEEEEE WIDTH=20%><FONT FACE=VERDANA SIZE=5 COLOR=GREEN>" + strDesc + "</FONT></TD>"\
            "</TR>")




