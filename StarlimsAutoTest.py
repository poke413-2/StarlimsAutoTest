##########################################################################
# Author:   Ryan Hankins                                                 #
#                                                                        #
# Purpose:  Automate the creation/completion of a Starlims test in V11   #
#                                                                        #
# Usage:    Open command prompt in script location                       #
#           Execute 'python StarlimsAutoTest.py'                         #
#           Follow prompt                                                #
#                                                                        #    
# Pre-req:  1920 x 1080 screen resolution                                #
#           Starlims open on primary monitor                             #
#           Python installed                                             #
#           Pyautogui library installed                                  #
#           Pyodbc library installed                                     #
##########################################################################


#Import libraries
import pyautogui
import time
import pyodbc
import PIL.ImageGrab
from tkinter import Tk
import re


############# HELPER FUNCTIONS ##############
def Retest(testCode):
    
    testName = GetTestName(testCode)
    

    
    #OK
    #pyautogui.click(PendingTestsCoord["Alerts_btnDone"])
    #time.sleep(3)
    
    #Add sample to run
    pyautogui.click(PendingTestsCoord["tabRun"])
    time.sleep(2)
    pyautogui.click(PendingTestsCoord["Run_btnSelectSample"])
    time.sleep(2)
    pyautogui.click(PendingTestsCoord["Run_btnAddSampleToRun"])
    time.sleep(2)
    
    #Go to results tab
    pyautogui.click(PendingTestsCoord["tabResults"])
    time.sleep(2)
    
    #Get results for the retest
    combo = PromptForAnalyteResults(testCode)
    
    #Add results to sample
    EnterResult(testCode, combo)
    
    print("\n-- Retest [" + testCode + ": " + testName + "] is complete.")
    
    return

def ReflexTest(reflexTestCode):

    reflexTestName = GetTestName(reflexTestCode)
       
    #Select the Run Template
    pyautogui.click(PendingTestsCoord["drpRunTemplate"])
    time.sleep(1)
    pyautogui.typewrite(reflexTestName)
    pyautogui.press('down')
    pyautogui.press('up')
    time.sleep(2)
    pyautogui.rightClick(PendingTestsCoord["drpRunTemplate"])
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["drpRunTemplate_Copy"])
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(4)
    
    #Store run template for parsing
    runTemplate = Tk().clipboard_get()
    result = re.search("/ Default\)", runTemplate)
    firstBracket = runTemplate.rfind("(", 0)
    equipment = runTemplate[firstBracket+1:result.start()]

    #Select run tab in case runs already exist as it defaults to result tab
    pyautogui.click(PendingTestsCoord["tabRun"])
    time.sleep(2)
   
    #If a drafted run doesnt exist, create one and select it
    runno = CheckExistingRun(reflexTestCode, equipment)
    if runno == "0":
        runno = CheckExistingRun(reflexTestCode, equipment)
    
    #Add sample to run (assuming the sample is at top of list)
    pyautogui.click(PendingTestsCoord["Run_btnSelectSample"])
    time.sleep(2)
    pyautogui.click(PendingTestsCoord["Run_btnAddSampleToRun"])
    time.sleep(2)

    #Create run
    pyautogui.click(PendingTestsCoord["btnCreateRun"])
    time.sleep(4)

    #Get results for the retest
    reflexCombo = PromptForAnalyteResults(reflexTestCode)

    #Add results to sample
    EnterResult(reflexTestCode, reflexCombo)
    
    #Is the OOS Alert button visible?
    pyautogui.moveTo(PendingTestsCoord["btnOOSAlert"])
    time.sleep(1)
    step = 1
    
    #Process OOS until Alert button is not visible
    rgb = PIL.ImageGrab.grab().load()[1860,222]
    while str(rgb) == "(255, 226, 144)":
        OOS(reflexTestCode, sampleNo)
        pyautogui.moveTo(PendingTestsCoord["btnOOSAlert"])
        time.sleep(1)
        step = step + 1
        rgb = PIL.ImageGrab.grab().load()[1860,222]
    
    #Go back to original test, as we were inside a reflex test
    #newest run should automatically be selected
    pyautogui.click(PendingTestsCoord["drpRunTemplate"])
    time.sleep(1)
    pyautogui.typewrite(reflexTestName)
    pyautogui.press('down')
    pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(4)
    
    SelectRun(runno)
    
    #Finish run
    pyautogui.click(PendingTestsCoord["btnFinishRun"])
    time.sleep(3)

    print("\n-- Reflex test [" + reflexTestCode + ": " + reflexTestName + "] is complete.")

    return

def OOS(testCode, ordno):
    
    #Connect to DB
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=LABS-Starlims20;'
                          'Database=MD_DEV_v11DATA;'
                          'uid=limssuper;pwd=limssuper')

    cursor = conn.cursor()
    cursor.execute('select TRIGGERACTION, SPECREPORT from MD_DEV_v11DATA.dbo.EVENTACTIONS where TESTCODE = ' + testCode + 'and ordno = \'' + ordno + '\' and DISPSTATUS = \'Pending\'')
    
    triggers = []
    
    #Get all the triggers for this test
    for row in cursor:
        triggers.append([row[0],row[1]])
    
    conn.close()

    arrReflex = []
    offset = 445
    
    #Execute the triggerss
    
    #Execute all the triggers
    for y in triggers:
        pyautogui.click(PendingTestsCoord["btnOOSAlert"])
        time.sleep(2)
        pyautogui.click(PendingTestsCoord["Alerts_btnExecute"])
        time.sleep(2)
        pyautogui.click(PendingTestsCoord["Alerts_btnDone"])
        time.sleep(2)
    
    for y in triggers:
        if y[0].lower() == "retest":
            Retest(testCode)
        elif y[0].lower() == "reflex test":
            #arrReflex.append(y[1])
            #pyautogui.click(590,offset)
            #time.sleep(1)
            #pyautogui.click(PendingTestsCoord["Alerts_btnExecute"])
            #time.sleep(2)
            ReflexTest(y[1])
        else:
            time.sleep(.5)
        #offset = offset + 20
    
    #If we're on a reflex step the Alerts windows remains open, 
    #close it before looping through reflex tests
    # if len(arrReflex) > 0:
        # pyautogui.click(PendingTestsCoord["Alerts_btnDone"])
        # time.sleep(3)
        # for reflexTestCode in arrReflex:
            # ReflexTest(reflexTestCode)
    
    return

def GetTestName(testCode):
    
    #Connect to DB
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=LABS-Starlims20;'
                          'Database=MD_DEV_v11DATA;'
                          'uid=limssuper;pwd=limssuper')

    cursor = conn.cursor()
    cursor.execute('select top 1 TESTNO from MD_DEV_v11Data.dbo.TESTS where testcode = ' + testCode)
    cursor = cursor.fetchall()
      
    testName = cursor[0][0]  
        
    conn.close()
    
    return testName

def GetPanelName(testCode):
    
    #Connect to DB
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=LABS-Starlims20;'
                          'Database=MD_DEV_v11DATA;'
                          'uid=limssuper;pwd=limssuper')

    cursor = conn.cursor()
    cursor.execute('select distinct testgroupname from MD_DEV_v11Data.dbo.CONDITIONTESTTRIGGERS where testcode = ' + testCode)
    cursor = cursor.fetchall()
    
    panelName = "Invalid testcode"
    
    #Was this testcode found? Return test name
    if len(cursor) == 1:
        for row in cursor:
            panelName = row[0]
        print("\n" + panelName + " selected.")
    elif len(cursor) > 1:
        print ("\nMultiple panels include this test:")
        i = 0
        for row in cursor:
            i = i + 1
            print ("   " + str(i) + ". " + row[0])
               
        sel = 0
        while sel <= 0 or sel > i:
            sel = int(input("Make selection (1, 2, etc): "))
            if sel <=0 or sel > i:
                print("\nInvalid selection.\n")
        
        panelName = cursor[int(sel)-1][0]
        
    conn.close()
    
    return panelName

def BackToDash(panelName, delay):

    #Start from the dashboard
    pyautogui.click(39,123)
    time.sleep(2)

    #Clear text from search bar
    pyautogui.click(1895,178)
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])

    #Refresh the panel
    pyautogui.click(1872,154)
    time.sleep(2)

    #Search for the panel
    pyautogui.click(1895,178)
    pyautogui.typewrite(panelName)
    time.sleep(1)
    pyautogui.click(1743,222)
    time.sleep(1)

    #Clear the currently open application popup if there
    pyautogui.click(929,568)
    time.sleep(delay)

def BackToDashMod(panelName, delay):

    #Start from the dashboard
    pyautogui.click(39,123)
    time.sleep(1)

    #Clear text from search bar
    pyautogui.click(1895,178)
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    pyautogui.typewrite(['backspace'])
    
    #Refresh the panel
    pyautogui.click(1872,154)
    time.sleep(2)
    
    #Search for the panel
    pyautogui.click(1895,178)
    pyautogui.typewrite(panelName)
    time.sleep(1)
    pyautogui.click(1754,274)
    time.sleep(1)

    #Clear the currently open application popup if there
    pyautogui.click(929,568)
    time.sleep(delay)    

def PromptForAnalyteResults(testCode):

    #Connect to DB
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=LABS-Starlims20;'
                          'Database=MD_DEV_v11DATA;'
                          'uid=limssuper;pwd=limssuper')

    cursor = conn.cursor()
    cursor.execute('select distinct analyte from MD_DEV_v11DATA.dbo.POSSIBLERESULTS where testcode = ' + testCode)
      
    arrAnalytes = []
    combo = []
    
    #Create list of analytes for this test, strip whitespace
    for row in cursor:
        arrAnalytes.append(row[0].strip()) 
    
    #Loop through each analyte to prompt user for result selection
    for analyte in arrAnalytes:
        cursor = conn.cursor()
        cursor.execute("select LOOKUPEXPRESSION FROM ANALYTES an" +
                          " join SPECSCHEMAS ss" +
                          " on an.SCHEMANAME = ss.SCHEMANAME" +
                          " join SPECSCHEMACALCS ssc" +
                          " on ss.SPECSCHEMACODE = ssc.SPECSCHEMACODE" +
                          " join METADATAFIELDS md" +
                          " on ssc.CALCULATION = md.METADATAMETHODCODE" +
                          " where an.TESTCODE = " + testCode +
                          " and ssc.TYPEOFITEM = \'Prompt\'" + 
                          " and LOOKUPEXPRESSION is not NULL" +
                          " and an.analyte = \'" + analyte + "\'")   
        cursor = cursor.fetchall()
       
        if len(cursor) > 1:
            #If more than 2 prompts found, do manual result entry
            print("\nMore than 2 prompts found for " + analyte + ", using manual result entry.")
            sel = 0
            result = "Manual"
        elif len(cursor) == 1:
            #1 prompt was found, try to parse the string for the list
            #Parse the schema prompt for the results
            sql = str(cursor[0])
            result = re.search("arrResult(s)?\s*:=\s*{\s*", sql)
            
            if result is None:
                #If we could not parse the result string for the list
                print("\nUnable to parse the prompt string, using manual result entry.")
                sel = 0
                result = "Manual"
            else: 
                #String was able to be parsed for the result list, so use that
                print("\nChoose test result for analyte " + analyte + ": ")
                firstBracket = sql.find("{", result.start())
                lastBracket = sql.find("}", result.start())

                txt = sql[firstBracket+1:lastBracket]
                spl = txt.split(",")
                cursor = [i.replace('"', '') for i in spl] 

                i = 0
                for row in cursor:
                    i = i + 1
                    print ("   " + str(i) + ". " + row)

                sel = 0
                while sel <= 0 or sel > i:
                    sel = int(input("Make selection (1, 2, etc): "))
                    if sel <=0 or sel > i:
                        print("\nInvalid selection.\n")

                result = cursor[int(sel)-1]                      
        else:
            print("\nChoose test result for analyte " + analyte + ": ")
            
            #If no prompt use Snomed codes
            cursor = conn.cursor()
            cursor.execute('select result from MD_DEV_v11DATA.dbo.POSSIBLERESULTS where testcode = ' + testCode + " and analyte = \'" + analyte + "\'")
            cursor = cursor.fetchall()   

            i = 0
            for row in cursor:
                i = i + 1
                print ("   " + str(i) + ". " + row[0])
                
            sel = 0
            while sel <= 0 or sel > i:
                sel = int(input("Make selection (1, 2, etc): "))
                if sel <=0 or sel > i:
                    print("\nInvalid selection.\n")
                
            result = cursor[int(sel)-1][0]   
        
        sameFirstLetter = False
        
        if result != "Manual":
            #Does this analyte have results that have the same first letter? if so, we handle that differently
            for outer, res in enumerate(cursor):
                resultLetter = res[0]      
                for inner, first in enumerate(cursor):
                    if outer != inner: #dont compare the same result option
                        firstLetter = first[0]
                        if resultLetter[:1] == firstLetter[:1]:
                            sameFirstLetter = True
                            break
                else:
                    continue
                break
        
        #print(analyte + " " + result)
        combo.append([analyte, result, int(sel)-1, sameFirstLetter])
        
    conn.close()

    return combo

def EnterResult(testCode, combo):
    
    offset = 483
        
    for analyte in combo:
    
        #Is this a manual entry?
        if analyte[1] == "Manual":
            input("-- Manual result entry enabled, press Enter to continue (after entering result and closing the result window)...")
        else:
            #Add results to sample
            pyautogui.doubleClick(825,offset)
            time.sleep(1)
            pyautogui.doubleClick(825,offset)
            time.sleep(1)
            pyautogui.moveTo(1216,373) #random spot on screen
            time.sleep(1)
            
            #If the same first letter result doesnt exist, then just type the result
            if analyte[3] == False:
                #Type out the result
                firstLetter = analyte[1]
                pyautogui.typewrite(firstLetter[:1])
                time.sleep(1) 
            #The same letter result exists, so we can't type the result selection
            else:
                #Cursor down until we're at the chosen result
                for x in range(analyte[2]):
                    pyautogui.press('down')
                    time.sleep(1)
            
            #Submit result selection    
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(2)
            
            #Go down to next analyte
            offset = offset + 20
    
    return

def CheckExistingRun(testCode, eqid):

    #Connect to DB
    conn = pyodbc.connect('Driver={SQL Server};'
                          'Server=LABS-Starlims20;'
                          'Database=MD_DEV_v11DATA;'
                          'uid=limssuper;pwd=limssuper')

    cursor = conn.cursor()
    cursor.execute('select top 1 RUNNO from Runs where testlst like \'%' + testCode + '%\' and RUNSTATUS = \'Draft\' and USRNAM = + \'' + userName + '\' and EQID = \'' + eqid +  '\'')
    cursor = cursor.fetchall()
    
    runno = "0"
    
    if len(cursor) > 0:
        runno = str(cursor[0][0])
        SelectRun(runno)   
    else:
        CreateNewRun()
    
    conn.close()
    
    return runno

def SelectRun(runno):

    #Clear the runno filter and enter the param
    pyautogui.click(PendingTestsCoord["RunnoColHeader"])
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["RunnoColHeader"])
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.typewrite(str(runno))
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["btnSelectFirstRun"])
    time.sleep(2)

    return

def FilterExistingSample(sampleNo):

    #Filter sample list for our sample
    pyautogui.click(PendingTestsCoord["Run_txtSampleNo"])
    time.sleep(1)
    pyautogui.typewrite(str(sampleNo))
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
    return

def CreateNewRun():

    #Clear current run selection
    pyautogui.click(PendingTestsCoord["RunnoColHeader"])
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["RunnoColHeader"])
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('backspace')
    pyautogui.press('enter')
    time.sleep(2)
    
    #Create a new run for the currently selected run template
    pyautogui.click(PendingTestsCoord["btnSelectFirstRun"])
    time.sleep(2)
    pyautogui.click(PendingTestsCoord["btnAddRun"])
    time.sleep(2)
    pyautogui.click(PendingTestsCoord["SelectRunTemplate_btnOk"])
    time.sleep(4)
    pyautogui.click(PendingTestsCoord["Run_btnSelectSample"])
    time.sleep(2)
    
    return


############# WORK FLOW FUNCTIONS ###########
def ReportDeliveryQueue(password):

    #Start from the dashboard
    BackToDashMod('Report Delivery Queue', 3)

    #Print Report
    pyautogui.click(DeliveryQueueCoord["btnSelectReport"])
    time.sleep(2)
    pyautogui.click(DeliveryQueueCoord["btnStartDelivery"])
    time.sleep(2)
    pyautogui.typewrite(password)
    time.sleep(2)
    pyautogui.click(DeliveryQueueCoord["ReportDelivery_btnOk"])
    time.sleep(2)

    print('5. Report ordered.')

def ReleaseByPanel(password):

    #Start from the dashboard
    BackToDash('Release by Panel', 6)

    #Release the record (assuming it is at top of list)
    pyautogui.click(ReleaseByPanelCoord["btnSelectSample"])
    time.sleep(2)
    pyautogui.click(ReleaseByPanelCoord["btnReleaseCurrent"])
    time.sleep(2)
    pyautogui.typewrite(password)
    time.sleep(2)
    pyautogui.click(ReleaseByPanelCoord["ReleaseOrder_btnOk"])
    time.sleep(2)

    print('4. Record released.')

def MyTeamsPendingTests(testCode, testName, combo):

    #Start from the dashboard
    BackToDash('My Teams Pending Tests', 5)
    
    #Store sample no for later
    global sampleNo
    sampleNo = Tk().clipboard_get()
    
    #Select the Run Template
    pyautogui.click(PendingTestsCoord["drpRunTemplate"])
    time.sleep(1)
    pyautogui.typewrite(testName)
    pyautogui.press('down')
    pyautogui.press('up')
    time.sleep(2)
    pyautogui.rightClick(PendingTestsCoord["drpRunTemplate"])
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["drpRunTemplate_Copy"])
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(4)
    
    #Store run template for parsing
    runTemplate = Tk().clipboard_get()
    result = re.search("/ Default\)", runTemplate)
    firstBracket = runTemplate.rfind("(", 0)
    equipment = runTemplate[firstBracket+1:result.start()]
    
    #Select run tab in case runs already exist as it defaults to result tab
    pyautogui.click(PendingTestsCoord["tabRun"])
    time.sleep(2)
    
    #Filter samples while the container id still exists in clipboard
    FilterExistingSample(sampleNo)
    
    #Active the filter row for runs and clear the error
    pyautogui.rightClick(PendingTestsCoord["Corner"])
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["FilterRow"])
    time.sleep(1)
    pyautogui.click(PendingTestsCoord["RunnoColHeader"])
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    
    #If a drafted run doesnt exist, create one and select it
    runno = CheckExistingRun(testCode, equipment)
    if runno == "0":
        runno = CheckExistingRun(testCode, equipment)

    #Add sample to run (assuming the sample is at top of list)
    pyautogui.click(PendingTestsCoord["Run_btnSelectSample"])
    time.sleep(2)
    pyautogui.click(PendingTestsCoord["Run_btnAddSampleToRun"])
    time.sleep(2)

    #Create run
    pyautogui.click(PendingTestsCoord["btnCreateRun"])
    time.sleep(4)

    #Add results to sample
    EnterResult(testCode, combo)
    
    #Is the OOS Alert button visible?
    pyautogui.moveTo(PendingTestsCoord["btnOOSAlert"])
    time.sleep(1)
    step = 1
    
    #Process OOS until Alert button is not visible
    rgb = PIL.ImageGrab.grab().load()[1860,222]
    while str(rgb) == "(255, 226, 144)":
        OOS(testCode, sampleNo)
        pyautogui.moveTo(PendingTestsCoord["btnOOSAlert"])
        time.sleep(1)
        step = step + 1
        rgb = PIL.ImageGrab.grab().load()[1860,222]
    
    #Go back to original test, as we were inside a reflex test
    #newest run should automatically be selected
    pyautogui.click(PendingTestsCoord["drpRunTemplate"])
    time.sleep(1)
    pyautogui.typewrite(testName)
    pyautogui.press('down')
    pyautogui.press('up')
    pyautogui.press('enter')
    time.sleep(4)
    
    SelectRun(runno)
    
    #Finish run
    pyautogui.click(PendingTestsCoord["btnFinishRun"])
    time.sleep(3)

    print('3. Run created and finalized.')

def ReceiveByTeam():

    #Start from the dashboard
    BackToDash('Receive By Team', 4)

    #Select Samples
    pyautogui.click(ReceiveByTeamCoord["btnSelectSamples"])
    time.sleep(2)
    pyautogui.click(ReceiveByTeamCoord["SelectSamples_drpContainerID"])
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1)
    pyautogui.click(ReceiveByTeamCoord["SelectSamples_drpLocation"])
    time.sleep(1)
    pyautogui.click(ReceiveByTeamCoord["SelectSamples_drpLocation_FirstSelection"])
    time.sleep(1)
    pyautogui.click(ReceiveByTeamCoord["SelectSamples_btnOk"])
    time.sleep(2)
    pyautogui.click(ReceiveByTeamCoord["SelectSamples_btnFinished"])
    time.sleep(2)

    print('2. Receive By Team complete.')

def ClinicalSampleLogin(testName, panelName):

    #Start from the dashboard
    BackToDash('Clinical Sample Login', 10)

    #Add test
    pyautogui.click(ClinLoginCoord["btnAdd"])
    time.sleep(3)

    #Enter panel name
    pyautogui.click(ClinLoginCoord["Add_txtSearchPanels"])
    time.sleep(1)
    pyautogui.typewrite(panelName)
    time.sleep(4)
    
    #If the test is currently checked, uncheck it
    rgb = PIL.ImageGrab.grab().load()[778,383]
    if str(rgb) == "(0, 0, 0)":
        pyautogui.click(ClinLoginCoord["Add_chkboxPanel"])
        time.sleep(1)
    #If test is unchecked, then multiple tests exist in this panel, double click to clear them
    else:
        pyautogui.click(ClinLoginCoord["Add_chkboxPanel"])
        time.sleep(1)
        pyautogui.click(ClinLoginCoord["Add_chkboxPanel"])
        time.sleep(1)
    
    #Enter test name
    pyautogui.click(ClinLoginCoord["Add_txtSearchTests"])
    time.sleep(1)
    pyautogui.typewrite(testName)
    time.sleep(2)
    
    #If the test is not checked, then other tests exist and it wont automatically be selected, so select it
    rgb = PIL.ImageGrab.grab().load()[778,383]
    if str(rgb) != "(0, 0, 0)":   
        pyautogui.click(ClinLoginCoord["Add_chkboxTest"])
        time.sleep(2)
    
    #OK
    pyautogui.click(ClinLoginCoord["Add_btnOk"])
    time.sleep(3)
    pyautogui.click(ClinLoginCoord["PanelAndMetadata_btnOk"])
    time.sleep(11)

    #Enter submitter info
    pyautogui.click(ClinLoginCoord["AccessionInfo_SubmitterElipse"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["SelectSubmitter_drpSubmitter"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["SelectSubmitter_drpSubmitter_FirstSelection"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["SelectSubmitter_chkboxMyers"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["SelectSubmitter_btnOk"])
    time.sleep(2)

    #Enter lastname
    pyautogui.click(ClinLoginCoord["AccessionInfo_txtLastName"])
    time.sleep(1)
    pyautogui.typewrite('hankins')

    #Enter firstname
    pyautogui.click(ClinLoginCoord["AccessionInfo_txtFirstName"])
    time.sleep(1)
    pyautogui.typewrite('ryan')
    time.sleep(1)
    pyautogui.typewrite(['tab'])
    time.sleep(4)

    #Select patient info
    pyautogui.click(ClinLoginCoord["SelectPatient_btnOk"])
    time.sleep(2)

    #Select Authorized by
    pyautogui.click(ClinLoginCoord["AccessionInfo_drpAuthBy"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["AccessionInfo_drpAuthBy_SecondSelection"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["btnCommit"])
    time.sleep(2)

    #this is needed for some reason, not sure why
    pyautogui.click(ClinLoginCoord["btnCommit"])
    time.sleep(2)
    pyautogui.click(ClinLoginCoord["Commit_btnCancel"])
    time.sleep(2)

    #Get container ID for receiving
    pyautogui.click(ClinLoginCoord["tabContainers"])
    time.sleep(3)
    pyautogui.click(ClinLoginCoord["Containers_ContainerID_FirstSelection"])
    pyautogui.rightClick(ClinLoginCoord["Containers_ContainerID_FirstSelection"])
    time.sleep(1)
    pyautogui.click(ClinLoginCoord["Containers_CopyToClipboard"])
    time.sleep(1)
    pyautogui.moveTo(ClinLoginCoord["Containers_CopyToClipboard_AllRows"])
    time.sleep(.5)
    pyautogui.click(ClinLoginCoord["Containers_CopyToClipboard_SelectedCell"])
    time.sleep(2)

    #Commit test
    pyautogui.click(ClinLoginCoord["btnCommit"])
    time.sleep(2)
    pyautogui.click(ClinLoginCoord["Commit_btnOk"])
    time.sleep(10)

    #Clear attachment window
    pyautogui.click(ClinLoginCoord["Report_btnClose"])
    time.sleep(2)

    print('\n1. Clinical Sample Login complete.')


############# GLOBAL COORDS #################
def InitVariables():

    InitClinicalSampleLogin()
    InitReceiveByTeam()
    InitPendingTests()
    InitReleaseByPanel()
    InitDeliveryQueue()
    
    return

def InitDeliveryQueue():

    global DeliveryQueueCoord 
    
    DeliveryQueueCoord = {
        "btnSelectReport"                                : (44,255),
        "btnStartDelivery"                               : (77,183),
        "ReportDelivery_btnOk"                           : (969,610)
    }

    return

def InitReleaseByPanel():

    global ReleaseByPanelCoord 
    
    ReleaseByPanelCoord = {
        "btnSelectSample"                                : (25,228),
        "btnReleaseCurrent"                              : (348,151),
        "ReleaseOrder_btnOk"                             : (1160,458)
    }

    return

def InitPendingTests():

    global PendingTestsCoord 
    
    PendingTestsCoord = {
        "drpRunTemplate"                                : (114,180),
        "drpRunTemplate_Copy"                           : (165,242),
        "tabRun"                                        : (27,383),
        "Run_txtSampleNo"                               : (190,471),
        "AssignedTo_FirstSelection"                     : (807,219),
        "AssignedTo_CopyToClipboard"                    : (960,393),
        "AssignedTo_CopyToClipboard_AllRows"            : (1030,393),
        "AssignedTo_CopyToClipboard_SelectedCell"       : (1031,435),
        "Status_FirstSelection"                         : (221,220),
        "Status_CopyToClipboard"                        : (369,393),
        "Status_CopyToClipboard_AllRows"                : (454,393),
        "Status_CopyToClipboard_SelectedCell"           : (451,434),
        "btnAddRun"                                     : (482,173),
        "SelectRunTemplate_btnOk"                       : (1012,629),
        "Run_btnSelectSample"                           : (55,494),
        "Run_btnAddSampleToRun"                         : (505,676),
        "btnCreateRun"                                  : (1648,220),
        "Results_txtResult_FirstSelection"              : (825,483),
        "ResultEntry_btnOk"                             : (1216,673),
        "btnOOSAlert"                                   : (1860,222),
        "Alerts_btnExecute"                             : (616,402),
        "Alerts_btnDismiss"                             : (699,402),
        "Alerts_btnDone"                                : (1286,674),
        "tabResults"                                    : (25,431),
        "btnFinishRun"                                  : (1636,237),
        "Corner"                                        : (26,199),
        "FilterRow"                                     : (92,292),
        "txtAssignedToFilter"                           : (794,220),
        "StatusColHeader"                               : (111,202),
        "Status_FirstSelection_AfterFilter"             : (218,242),
        "StatusAfterFilter_CopyToClipboard"             : (368,414),
        "StatusAfterFilter_CopyToClipboard_AllRows"     : (429,414),
        "StatusAfterFilter_CopyToClipboard_SelectedCell": (431,456),
        "RunnoColHeader"                                : (94,220),
        "btnSelectFirstRun"                             : (24,242)
    }

    return

def InitReceiveByTeam():

    global ReceiveByTeamCoord 
    
    ReceiveByTeamCoord = {
        "btnSelectSamples"                              : (83,174),
        "SelectSamples_drpContainerID"                  : (929,514),
        "SelectSamples_drpLocation"                     : (932,551),
        "SelectSamples_drpLocation_FirstSelection"      : (932,562),
        "SelectSamples_btnOk"                           : (952,591),
        "SelectSamples_btnFinished"                     : (1048,590)
    }

    return

def InitClinicalSampleLogin():

    global ClinLoginCoord 
    
    ClinLoginCoord = {
        "btnAdd"                                        : (45,173),
        "Add_txtSearchPanels"                           : (662,359),
        "Add_txtSearchTests"                            : (841,361),
        "Add_chkboxPanel"                               : (778,383),
        "Add_chkboxTest"                                : (795,400),
        "Add_btnOk"                                     : (1340,746),
        "PanelAndMetadata_btnOk"                        : (1242,796),
        "AccessionInfo_SubmitterElipse"                 : (780,519),
        "SelectSubmitter_drpSubmitter"                  : (1252,382),
        "SelectSubmitter_drpSubmitter_FirstSelection"   : (1252,400),
        "SelectSubmitter_chkboxMyers"                   : (687,526),
        "SelectSubmitter_btnOk"                         : (1157,735),
        "AccessionInfo_txtLastName"                     : (280,703),
        "AccessionInfo_txtFirstName"                    : (791,699),
        "SelectPatient_btnOk"                           : (1169,642),
        "AccessionInfo_drpAuthBy"                       : (1014,976),
        "AccessionInfo_drpAuthBy_SecondSelection"       : (881,882),
        "btnCommit"                                     : (126,172),
        "Commit_btnCancel"                              : (1073,601),
        "tabContainers"                                 : (622,447),
        "Containers_ContainerID_FirstSelection"         : (94,514),
        "Containers_CopyToClipboard"                    : (271,715),
        "Containers_CopyToClipboard_AllRows"            : (351,719),
        "Containers_CopyToClipboard_SelectedCell"       : (294,759),
        "Commit_btnOk"                                  : (993,601),
        "Report_btnClose"                               : (1379,188)
    }

    return


############## MAIN PROCESS #################
def Main():
    print('Press Ctrl-C to quit.')
    testCode = "null"
    
    InitVariables()
    
    global userName
    userName = input("\nEnter your user name: ")
    password = input("Enter your password: ")
    
    while testCode.lower() != "q":
        testCode = input("\nEnter the testcode or Q to quit: ") 
        start_time = time.time()
       
        try:
            testCode = int(testCode)
            testCode = str(testCode)
        except ValueError:
            if testCode.lower() != "q":
                testCode = "x"
        
        if testCode.lower() != "q" and testCode.lower() != "x":   
            panelName = GetPanelName(testCode)
            
            if panelName != "Invalid testcode":
                testName = GetTestName(testCode)
                combo = PromptForAnalyteResults(testCode)

                print("\nBegin test for [" + testCode + ": " + testName + "]...")
                
                ClinicalSampleLogin(testName, panelName)
                ReceiveByTeam()
                MyTeamsPendingTests(testCode, testName, combo)
                ReleaseByPanel(password)
                #ReportDeliveryQueue(password) 
                
                elapsed_time = time.time() - start_time
                elapsed_time = time.strftime("%H:%M:%S", time.gmtime(elapsed_time))
                print("\nDone [Time elapsed: " + elapsed_time + "].")
            else:
                print("\nInvalid testcode.")
        elif testCode.lower() == "x":
            print("\nInvalid entry.")
        else:
            print("\nQuit.")

    return




#************* RUN **************************
Main()
