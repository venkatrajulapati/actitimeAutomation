from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from Lib.commonUtils import *
from Lib.App_CommonUtils import *
import time
import xlrd
import sys

######################################################### Test Suite Driver #######################################################################################

class TestSuiteDriver(App_Common_utils):
    def __init__(self):
        super(TestSuiteDriver,self).__init__()

###################################################################################################################################################################

# Connect to the Test test case repository

oWb = xlrd.open_workbook("E:\\actitimeAutomation\\TestSuite.xls")
oTestsuite = oWb.sheet_by_name("TestSuite")
oBusinessFlow = oWb.sheet_by_name("BusinessFlow")
nooftcs = oTestsuite.nrows
oui = UIdriver()
obj=TestSuiteDriver()

# Run the Test set

for i in range(1,nooftcs):
    TCID = oTestsuite.cell(i,0).value
    BSID = oTestsuite.cell(i,1).value
    TCName = oTestsuite.cell(i,2).value

    tobeexecute = oTestsuite.cell(i,3).value
    if str(tobeexecute).lower() == "y":
        #exec ("obj = " + TCName + "()")
        print("Running the Test case : " + TCName)
        reqrow = oui.GetxlRowNumberbytwocolvals(oBusinessFlow,"TC_ID",TCID,"BS_ID",BSID)
        noofsteps = oBusinessFlow.ncols
        for j in range(3,noofsteps):
            Keyword = oBusinessFlow.cell(reqrow,j).value
            if not Keyword == "end":
                print("running Keyword : " + Keyword)
                Keyword = Keyword + "(" + chr(34)+TCName+chr(34) + "," + chr(34) + TCID + chr(34) + "," + chr(34) + BSID + chr(34) +" )"
                #keyword = "%s%s%s%s" %(Keyword ,"(",oDataset,")")
                exec ("obj." + Keyword)
            elif Keyword == "end":
                print("end of the test")
                obj.browser.quit()
                break
