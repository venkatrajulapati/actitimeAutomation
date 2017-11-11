from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import urllib.request
from Lib.commonUtils import *
from Lib.App_CommonUtils import *
import time
import xlrd
import sys
import xlwt
from xlwt import *
from xlutils.copy import copy


######################################################### Test Suite Driver #######################################################################################

class TestSuiteDriver(App_Common_utils):
    def __init__(self):
        super(TestSuiteDriver,self).__init__()

###################################################################################################################################################################

# Connect to the Test test case repository

oRb = xlrd.open_workbook("E:\\actitimeAutomation\\TestPlan\\TestSuite.xls")
oWb = copy(oRb)
#Read only copy of Test Suite
oRTestsuite = oRb.sheet_by_name("TestSuite")
#Editable copy of Test suite
oWTestsuite = oWb.get_sheet("TestSuite")
#Business Flow sheet
oBusinessFlow = oRb.sheet_by_name("BusinessFlow")
nooftcs = oRTestsuite.nrows
oui = UIdriver()
obj=TestSuiteDriver()
# font0 = Font()
# font0.colour_index. "#0000FF"
# style = XFStyle()
# style.font = font0

# Run the Test set

for i in range(1,nooftcs):
    TCID = oRTestsuite.cell(i,0).value
    TDID = oRTestsuite.cell(i,1).value
    TCName = oRTestsuite.cell(i,2).value

    tobeexecute = oRTestsuite.cell(i,3).value
    if str(tobeexecute).lower() == "y":
        print("Running the Test case : " + TCName)
        #Create HTML Report
        f = obj.Create_HTML_Report(TCName)
        reppath = obj.Reportfile
        screenshotpath = obj.screenshotfolder
        reqrow = oui.GetxlRowNumberbytwocolvals(oBusinessFlow,"TC_ID",TCID,"TD_ID",TDID)
        noofsteps = oBusinessFlow.ncols
        iPasscount = 0
        stepcount = 0
        #Get the Data sheet name
        DatasheetName = oBusinessFlow.cell(reqrow,2).value
        if DatasheetName=="":
            DatasheetName = oRTestsuite.cell(reqrow,2).value
        screenshotcount=0
        #Execute the Business flow Keywords and update the HTML Report
        for j in range(3,noofsteps):
            Keyword = oBusinessFlow.cell(reqrow,j).value
            temp = Keyword
            if not Keyword == "end":
                print("running Keyword : " + Keyword)
                Keyword = Keyword + "(" + chr(34)+DatasheetName+chr(34) + "," + chr(34) + TCID + chr(34) + "," + chr(34) + TDID + chr(34) +" )"
                #keyword = "%s%s%s%s" %(Keyword ,"(",oDataset,")")
                if eval ("obj." + Keyword):
                    print(temp + " Keyword Passed")
                    screenshotcount = screenshotcount + 1
                    iPasscount = iPasscount + 1
                    stepcount = stepcount + 1
                    obj.fn_HtmlReport_TestStep(f,screenshotpath,screenshotcount,"running Step : " + str(temp),str(temp) + " Should be Passed",str(temp) + " is Passed","PASS")
                else:
                    print(temp+" Keyword Failed hence Quitting the current test execution")
                    obj.fn_HtmlReport_TestStep(f, screenshotpath, screenshotcount, "running Step : " + str(temp),str(temp) + " Should be Passed", str(temp) + " is Failed", "FAIL")
                    stepcount = stepcount + 1
                    obj.browser.quit()
                    break
            elif Keyword == "end":
                print("end of the test")
                obj.browser.quit()
                f.close()
                break
    # Update The Testcase wise status and Report path in the test suite
    if iPasscount==stepcount:
        oWTestsuite.write(i,4,"Pass")
        oWTestsuite.write(i, 5, xlwt.Formula('HYPERLINK("%s";"Clickto view report")' % reppath))
    else:
        oWTestsuite.write(i,4,"Fail")
        oWTestsuite.write(i, 5, xlwt.Formula('HYPERLINK("%s";"Clickto view report")' % reppath))

    oWb.save(os.path.join(obj.Rootpath, "TestPlan\\TestSuite.xls"))
