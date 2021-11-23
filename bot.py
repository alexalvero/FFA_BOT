#   PURPOSE: FFA BOT PROJECT 2021
#   - Import incoming delivery spreadsheet and verify entries into reliance
#   - Catch any discrepancies and log items into a separate spreadsheet
#   - Receive units without any discrepancies for Investigations

#   CREATED   : 08-SEPT-2021
#   LAST EDIT : 23-NOV-2021
#   AUTHOR    : ALEXANDER ALVERO III (personal: alveroalexander@gmail.com || work: alexander.alvero@dexcom.com)

# MODULE REQUIREMENTS
from os import close


try:
    from selenium import webdriver
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.by import By
    from selenium.common.exceptions import TimeoutException,NoSuchElementException
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.options import Options
    import time, getpass,os, csv
    from datetime import date,datetime
    import ctypes
    import sys  



except ImportError or ModuleNotFoundError:
    print("Importing Module failed")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)
 
def get_display_name():
    GetUserNameEx = ctypes.windll.secur32.GetUserNameExW
    NameDisplay = 3
 
    size = ctypes.pointer(ctypes.c_ulong(0))
    GetUserNameEx(NameDisplay, None, size)
 
    nameBuffer = ctypes.create_unicode_buffer(size.contents.value)
    GetUserNameEx(NameDisplay, nameBuffer, size)
    return nameBuffer.value

# VARIABLES
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
chromeDriverPATH = r"C:\Users\aza0220\chromedriver.exe"
# driver = webdriver.Chrome(executable_path= chromeDriverPATH, options=options) # LOCAL FILE 
driver  = webdriver.Chrome(resource_path('./driver/chromedriver.exe'),options=options) # USING FILE
OKTA = "https://dexcom.okta.com"
localPATH = os.path.realpath(__file__)
folderDate = os.path.dirname(localPATH)
date = str(date.today())
folderDatePATH = str(f"{folderDate}\{date}")
user_name = os.getlogin()
user_fullname = get_display_name()

def mkFolder(folderDatePATH): 
    try: 
        os.mkdir(folderDatePATH)
    except FileExistsError or FileNotFoundError:
        pass






Completed = []
devResult2={}

MasterPushlist = []

ManualPushlist = []

discrepancyList = []

keepcode = {'045', '45', '048', '48', '076', '76', '109', '096', '96', '050', '50', '053', '53', '080', '80', '081', '81', '082', '82', '083', '83', '085', '85', '090', '90', '106', '111', '114', '115', '123', '5016', '5018', '5024'}




def openBrowser(OKTA, driver):
    
    driver.get(OKTA)
    driver.maximize_window()
    


def loginCredits():
    try:
        userInput = input("Username: ")
        passInput = getpass.getpass()  # This allow to input password with printing the input
        return userInput, passInput
    except Exception:
        loginCredits()


def login(driver, Keys):

    username, password = loginCredits()
    userField = driver.find_element_by_id("okta-signin-username")
    passwordField = driver.find_element_by_id("okta-signin-password")
    time.sleep(5)
    userField.send_keys(username)
    passwordField.send_keys(password)

    logintBtn = driver.find_element_by_id("okta-signin-submit")
    logintBtn.send_keys(Keys.RETURN)

    try:
        driver.implicitly_wait(2)
        driver.find_element_by_css_selector('[role = "alert"]')
        print("Username/Password Incorrect")
        clearUser = driver.find_element_by_id("okta-signin-username")
        clearPass = driver.find_element_by_id("okta-signin-password")
        # clearUser.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
        # clearPass.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
        clearUser.clear()
        clearPass.clear()
        login(driver, Keys)


    except NoSuchElementException or TimeoutError:
        print("Login Successful!")



def openRelianceApp(driver):

    try: # Wait until all elements are loaded before doing task
        wait = WebDriverWait(driver, 30)
        wait.until(EC.presence_of_element_located((By.ID, "main-content")))
        relianceBtn = driver.find_element_by_css_selector('[alt="ETQ Reliance Prod logo"]')
        relianceBtn.click()
        driver.close() # Closes OKTA tab
        # driver.switch_to.window(driver.window_handles[0]) # Switch focus to Reliance Tab
        driver.switch_to.window(driver.window_handles[0]) # Switch focus to Reliance Tab

    except TimeoutException:
        print("Cannot find Open Relaince App")

def openComplaintApp(driver):

    try:
        wait = WebDriverWait(driver, 5)
        wait.until(EC.presence_of_element_located((By.ID, "HomeApplicationsList")))
        complaintsAppbtn = driver.find_element_by_id("COMPLAINTS_APP")
        complaintsAppbtn.click()
    except TimeoutException:
        print("Cannot find Complaint Handling Tab")


def openDeviceAnalysis(driver, Keys):



    # CHECK IF BUTTON IS DISPLAYED BEFORE CLICKING
    openBtn = driver.find_element_by_id("1._Open_0")
    openBtn.click()

    try:
        value = driver.find_element_by_id("div_1._Open_0").get_attribute("style")
        if value == "display: none;":
            openBtn.click()
    except:
        pass


    deviceAnalysisBtn = driver.find_element_by_id("Device_Analysis_3")
    deviceAnalysisBtn.click()

    try:
        value = driver.find_element_by_id("div_Device_Analysis_3").get_attribute("style")
        if value == "display: none;":
            deviceAnalysisBtn.click()
    except:
        pass

    bynumberBtn = driver.find_element_by_id("BY_NUMBER_22_P")
    bynumberBtn.click()




def searchCOM(RGA,COM,CODE,SN,RETURN_TYPE,UNITYPE):


    incomingRGA = RGA
    incomingCOM = COM
    incomingCODE = CODE
    incomingSN = SN
    incomingTYPE = RETURN_TYPE
    incomingUNITTYPE = UNITYPE

    
    discrepancyDEVList = []
    validatedDEVList = {}    
    PushList = []
    NotRGAProcessingList = []
    resultdevList = {}

    discrepancyDEVList.clear()
    validatedDEVList.clear()
    PushList.clear()
    resultdevList.clear()

    try: 
        comColumn = driver.find_element_by_css_selector('[name = "COMPLAINT_NUMBER_21_P_SEARCH"]')
        comColumn.clear()
    except NoSuchElementException or TimeoutException:
        wait = WebDriverWait(driver,5)
        wait.until(EC.presence_of_element_located((By.ID, "save_and_close_document")))
        driver.find_element_by_id('save_and_close_document').click()


    operatorName = driver.find_element_by_id('USER_MENU').text
    
    try:
        wait = WebDriverWait(driver, 5)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[name = "COMPLAINT_NUMBER_21_P_SEARCH"]')))
        comColumn.send_keys(incomingCOM)
        comColumn.send_keys(Keys.RETURN)
    except NoSuchElementException:
        print("Cannot find Search Element!")
        driver.refresh()

    # DEV DATA FETCH START
    try:
        wait = WebDriverWait(driver, 5)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="div_virtual_container"]/table/tbody')))

        dev = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr/td[2]').text
        comNo = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr[1]/td[3]').text
        comOwner = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr[1]/td[5]').text
        currentPhase = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr[1]/td[6]').text
        code = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr[1]/td[7]').text
        
        try:
            returnType = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr[1]/td[8]').text
            
            rga = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr[1]/td[9]').text


            result = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr')
            result.click()

            if len(returnType)<= 1:
                returnType =  driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody[3]/tr[2]/td/table/tbody/tr/td/span[1]").text
            else:
                pass

            try:
                x = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody[3]/tr[5]/td/table/tbody/tr/td/span[1]').text
                serialnumber = x.upper()
            except NoSuchElementException:
                serialnumber = "NA"


            productLine = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody[3]/tr[4]/td/table/tbody/tr/td/span[1]').text

            try: 
                productTest = driver.find_element_by_xpath('//*[@id="DX_DEVAN_PRODUCT_TO_BE_TESTED_P_AUTOCOMPLETE"]').text
            except NoSuchElementException or TimeoutException:
                productTest = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[3]/tbody/tr/td/table[2]/tbody/tr[1]/td/table/tbody/tr/td[1]/span[1]').text

            try:
                disposition = driver.find_element_by_id("DX_DEVAN_PRODUCT_INFORMATION_P_DX_DEVAN_PRODUCT_DISPOSITION_P_0_AUTOCOMPLETE").get_attribute("value")
                
            except NoSuchElementException or TimeoutException:
                disposition = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody[3]/tr[13]/td/table/tbody/tr/td/span[1]").text

            
            try: 
                lotnumber = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody[3]/tr[6]/td/table/tbody/tr/td/span[1]").text
            except NoSuchElementException:
                pass    


            resultdevList[comNo] = {}

            resultdevList[comNo]["RGA"] = rga
            resultdevList[comNo]["COM"] = comNo
            resultdevList[comNo]["Serial Number"] = serialnumber
            resultdevList[comNo]["Code"] = code
            resultdevList[comNo]['Return Type'] = returnType
            resultdevList[comNo]['Product Line'] = productLine
            resultdevList[comNo]['Current Phase'] = currentPhase
            resultdevList[comNo]['DEV'] = dev
            resultdevList[comNo]['Disposition'] = disposition
            resultdevList[comNo]['COM Owner'] = comOwner
            resultdevList[comNo]['Lot Number'] = lotnumber
            

            
            for p_id, p_info in resultdevList.items():
                resultrga = p_info['RGA']
                resultcom = p_info['COM']
                resultcode = p_info['Code']
                resultSN = p_info['Serial Number']
                resultType = p_info['Return Type']
                resultDisposition = p_info['Disposition']
                
                # DEV VALIDATION PART START
                if incomingRGA in resultrga:
                    pass
                    # print(f'RGA {resultrga} matched!')

                    if incomingCOM in resultcom:
                        pass
                        # print(f'{resultcom} matched!')

                        if incomingCODE in resultcode:
                            

                            if incomingSN in resultSN or incomingSN == "N/A":
                                pass
                                

                                if incomingTYPE in resultType:

                                    validatedDEVList[comNo]={}
                                    validatedDEVList[comNo]["ValidatedRGA"] = incomingRGA
                                    validatedDEVList[comNo]["ValidatedCOM"] = incomingCOM
                                    validatedDEVList[comNo]["ValidateDEV"] = dev
                                    

                                    

                                else:
                                    print(f"DISCREPANCY ADDED: Issue: Return Type doesn't match in DEV! RGA:({incomingRGA})")
                                    discrepancyDEVList.append(date)
                                    discrepancyDEVList.append(incomingRGA)
                                    discrepancyDEVList.append(incomingCOM)
                                    discrepancyDEVList.append(incomingCODE)
                                    discrepancyDEVList.append(incomingSN)
                                    discrepancyDEVList.append(incomingUNITTYPE)
                                    discrepancyDEVList.append(comOwner)
                                    discrepancyDEVList.append(f"Return Type doesn't match in DEV! Physical Return Type : {incomingTYPE} DEV Return Type: {resultType}")
                                    discrepancyList.append(discrepancyDEVList)


                            else:
                                print(f"DISCREPANCY ADDED: Issue: Serial Number doesn't match in DEV! RGA:({incomingRGA})")
                                discrepancyDEVList.append(date)
                                discrepancyDEVList.append(incomingRGA)
                                discrepancyDEVList.append(incomingCOM)
                                discrepancyDEVList.append(incomingCODE)
                                discrepancyDEVList.append(incomingSN)
                                discrepancyDEVList.append(incomingUNITTYPE)
                                discrepancyDEVList.append(comOwner)
                                discrepancyDEVList.append(f"Serial Number doesn't match in DEV! Physical Serial Number : {incomingSN} DEV Serial Number: {resultSN}")
                                discrepancyList.append(discrepancyDEVList)
                                

                        else:
                            print(f"DISCREPANCY ADDED: Issue: CODE doesn't match in DEV! RGA:({incomingRGA})")
                            discrepancyDEVList.append(date)
                            discrepancyDEVList.append(incomingRGA)
                            discrepancyDEVList.append(incomingCOM)
                            discrepancyDEVList.append(incomingCODE)
                            discrepancyDEVList.append(incomingSN)
                            discrepancyDEVList.append(incomingUNITTYPE)
                            discrepancyDEVList.append(comOwner)
                            discrepancyDEVList.append(f"CODE doesn't match in DEV! Physical CODE : {incomingCODE} DEV CODE: {resultcode}")
                            discrepancyList.append(discrepancyDEVList)
                            

                    else:
                        print(f"DISCREPANCY ADDED: Issue: COM doesn't match in DEV! RGA:({incomingRGA})")
                        discrepancyDEVList.append(date)
                        discrepancyDEVList.append(incomingRGA)
                        discrepancyDEVList.append(incomingCOM)
                        discrepancyDEVList.append(incomingCODE)
                        discrepancyDEVList.append(incomingSN)
                        discrepancyDEVList.append(incomingUNITTYPE)
                        discrepancyDEVList.append(comOwner)
                        discrepancyDEVList.append(f"COM doesn't match in DEV! Physical COM : {incomingCOM} DEV COM: {resultcom}")
                        discrepancyList.append(discrepancyDEVList)

                else: 
                    print(f"DISCREPANCY ADDED: Issue: RGA doesn't match in DEV! RGA:({incomingRGA})")
                    discrepancyDEVList.append(date)
                    discrepancyDEVList.append(incomingRGA)
                    discrepancyDEVList.append(incomingCOM)
                    discrepancyDEVList.append(incomingCODE)
                    discrepancyDEVList.append(incomingSN)
                    discrepancyDEVList.append(incomingUNITTYPE)
                    discrepancyDEVList.append(comOwner)
                    discrepancyDEVList.append(f"RGA doesn't match in DEV! Physical RGA : {incomingRGA} DEV RGA: {resultrga}")
                    discrepancyList.append(discrepancyDEVList)
                # DEV VALIDATION PART END    
            

            comLink = driver.find_element_by_id('ETQ$SOURCE_LINK_links_list').click()
            comAllTab = driver.find_element_by_id('ALL_TABS_TAB').click()

               
            # COM DATA FETCH START
            try:
                comresultRGA = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[3]/tbody/tr[2]/td/div/div/div/div/table[3]/tbody/tr/td/table[2]/tbody/tr[3]/td/table/tbody/tr/td/span[1]').text

                if len(comresultRGA) != len(incomingRGA):
                     comresultRGA = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[3]/tbody/tr[2]/td/div/div/div/div/table[4]/tbody/tr/td/table[2]/tbody/tr[3]/td/table/tbody/tr/td/span[1]").text
                
            except NoSuchElementException:
                comresultRGA = incomingRGA

            comresultCOM = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[1]/tbody/tr[2]/td/div/div/div/div/table[2]/tbody/tr/td/table[2]/tbody/tr[1]/td/table/tbody/tr/td[1]/span/span[1]').text
            comresultCode = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[1]/tbody/tr[2]/td/div/div/div/div/table[2]/tbody/tr/td/table[2]/tbody/tr[8]/td/table/tbody/tr/td[1]/span[1]').text
            comresultType = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[3]/tbody/tr[2]/td/div/div/div/div/table[1]/tbody/tr/td/table[2]/tbody/tr[1]/td/table/tbody/tr/td/span[1]').text
            comresultProductLine = driver.find_element_by_xpath('/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[3]/tbody/tr[2]/td/div/div/div/div/table[1]/tbody/tr/td/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/span[1]').text
            # try:
            #     comresultSN = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[3]/tbody/tr[2]/td/div/div/div/div/table[1]/tbody/tr/td/table[2]/tbody/tr[7]/td/table/tbody/tr/td/span[1]").text
            # except NoSuchElementException:
            #     comresultSN = serialnumber
            
            if RETURN_TYPE == "Transmitter" or RETURN_TYPE == "Receiver":
                w = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/table[3]/tbody/tr[2]/td/div/div/div/div/table[1]/tbody/tr/td/table[2]/tbody/tr[7]/td/table/tbody/tr/td/span[1]").text
                comresultSN = w.upper()
            else:
                comresultSN = serialnumber


            # COM DATA FETCH END
                    
            for p2_id, p2_info in validatedDEVList.items():
                validatedrga = p2_info['ValidatedRGA']
                validatedcom = p2_info['ValidatedCOM']

                if incomingCOM in validatedcom:
                    # COM VALIDATION PART START
                    if comresultRGA in validatedrga:
                        pass

                        if comresultCOM in resultcom :
                            pass

                            if comresultCode == resultcode : 
                                pass
                                
                                if comresultType == resultType:
                                    pass

                                    if comresultSN == resultSN or comresultSN == "n/a" or comresultSN == "NA" or comresultSN == "na" or comresultSN == "":
                                        pass
                                        if currentPhase == "RGA Processing":
                                            print(f"ALERT: DEV and COM VALIDATION COMPLETE for RGA:{incomingRGA}")
                                            # ADD VALIDATED DATA IN PUSHABLE LIST TO PUSH 
                                            PushList.append(date)
                                            PushList.append(incomingRGA)
                                            PushList.append(incomingCOM)
                                            PushList.append(incomingCODE)
                                            PushList.append(incomingSN)
                                            PushList.append(incomingTYPE)
                                            PushList.append(incomingUNITTYPE)
                                            PushList.append(productLine)
                                            PushList.append(lotnumber)
                                            PushList.append(comOwner)
                                            PushList.append(operatorName)
                                            PushList.append(disposition)
                                            MasterPushlist.append(PushList)

                                            
                                            


                                            # FOR PRF VALIDATION , ADD VALIDATED DEV IN DEV RESULT DICTIONARY
                                            devResult2[dev] = {}
                                            devResult2[dev]["ValidatedCOM"] = True
                                            devResult2[dev]["ValidatedDEV"] = True

                                            

                                        else:
                                            print(f"ALERT: ISSUE: NOT IN RGA PROCESSING PHASE RGA: {incomingRGA}")
                                            NotRGAProcessingList.append(date)
                                            NotRGAProcessingList.append(incomingRGA)
                                            NotRGAProcessingList.append(incomingCOM)
                                            NotRGAProcessingList.append(incomingCODE)
                                            NotRGAProcessingList.append(incomingSN)
                                            NotRGAProcessingList.append(incomingUNITTYPE)
                                            NotRGAProcessingList.append(" ")
                                            NotRGAProcessingList.append(f"Current Phase: {currentPhase}")
                                            ManualPushlist.append(NotRGAProcessingList)


                                    
                                    
                                    else:
                                        print(f"DISCREPANCY ADDED: Issue: Serial Number doesn't match in COM! RGA:({incomingRGA}) ")
                                        discrepancyDEVList.append(date)
                                        discrepancyDEVList.append(incomingRGA)
                                        discrepancyDEVList.append(incomingCOM)
                                        discrepancyDEVList.append(incomingCODE)
                                        discrepancyDEVList.append(incomingSN)
                                        discrepancyDEVList.append(incomingUNITTYPE)
                                        discrepancyDEVList.append(comOwner)
                                        discrepancyDEVList.append(f"Serial Number doesn't match in COM! Physical Serial Number : {incomingSN} COM Serial Number: {comresultSN}")
                                        discrepancyList.append(discrepancyDEVList)
                                
                                else:
                                    print(f"DISCREPANCY ADDED: Issue: Return Type doesn't match in COM! RGA:({incomingRGA}) ")
                                    discrepancyDEVList.append(date)
                                    discrepancyDEVList.append(incomingRGA)
                                    discrepancyDEVList.append(incomingCOM)
                                    discrepancyDEVList.append(incomingCODE)
                                    discrepancyDEVList.append(incomingSN)
                                    discrepancyDEVList.append(incomingUNITTYPE)
                                    discrepancyDEVList.append(comOwner)
                                    discrepancyDEVList.append(f"RETURN TYPE doesn't match in COM! Physical Return Type : {incomingTYPE} COM COM: {comresultType}")
                                    discrepancyList.append(discrepancyDEVList)



                            else:
                                print(f"DISCREPANCY ADDED: Issue: CODE doesn't match in COM! RGA:({incomingRGA}) ")
                                discrepancyDEVList.append(date)
                                discrepancyDEVList.append(incomingRGA)
                                discrepancyDEVList.append(incomingCOM)
                                discrepancyDEVList.append(incomingCODE)
                                discrepancyDEVList.append(incomingSN)
                                discrepancyDEVList.append(incomingUNITTYPE)
                                discrepancyDEVList.append(comOwner)
                                discrepancyDEVList.append(f"CODE doesn't match in COM! Physical CODE : {incomingCODE} COM COM: {comresultCode}")
                                discrepancyList.append(discrepancyDEVList)

                        else: 
                            print(f"DISCREPANCY ADDED: Issue: COM doesn't match in COM! RGA:({incomingRGA})")
                            discrepancyDEVList.append(date)
                            discrepancyDEVList.append(incomingRGA)
                            discrepancyDEVList.append(incomingCOM)
                            discrepancyDEVList.append(incomingCODE)
                            discrepancyDEVList.append(incomingSN)
                            discrepancyDEVList.append(incomingUNITTYPE)
                            discrepancyDEVList.append(comOwner)
                            discrepancyDEVList.append(f"COM doesn't match in COM! Physical COM : {incomingCOM} COM COM: {comresultCOM}")
                            discrepancyList.append(discrepancyDEVList)

                    else:
                        print(f"DISCREPANCY ADDED: Issue: RGA doesn't match in COM! RGA:({incomingRGA})")
                        discrepancyDEVList.append(date)
                        discrepancyDEVList.append(incomingRGA)
                        discrepancyDEVList.append(incomingCOM)
                        discrepancyDEVList.append(incomingCODE)
                        discrepancyDEVList.append(incomingSN)
                        discrepancyDEVList.append(incomingUNITTYPE)
                        discrepancyDEVList.append(comOwner)
                        discrepancyDEVList.append(f"RGA doesn't match in COM! Physical RGA : {incomingRGA} COM RGA: {comresultRGA}")
                        discrepancyList.append(discrepancyDEVList)
                    # COM VALIDATION PART END 
                else: 
                    pass
            
            try:
                save_n_closeCOM = driver.find_element_by_id('save_and_close_document').click()
            
            except NoSuchElementException:
                cancelCOM = driver.find_element_by_id('cancel_document').click()    


            try:
                    wait = WebDriverWait(driver,5)
                    wait.until(EC.presence_of_element_located((By.ID, "save_and_close_document")))
                    driver.find_element_by_id('save_and_close_document').click()

            except TimeoutException or NoSuchElementException:
                    wait = WebDriverWait(driver,5)
                    wait.until(EC.presence_of_element_located((By.ID, "cancel_document")))
                    driver.find_element_by_id('cancel_document').click()
                    
        except NoSuchElementException:
            print(f'ALERT: ISSUE: Cannot validate RGA: {incomingRGA}!')
            NotRGAProcessingList.append(date)
            NotRGAProcessingList.append(incomingRGA)
            NotRGAProcessingList.append(incomingCOM)
            NotRGAProcessingList.append(incomingCODE)
            NotRGAProcessingList.append(incomingSN)
            NotRGAProcessingList.append(incomingUNITTYPE)
            NotRGAProcessingList.append(" ")
            NotRGAProcessingList.append(f"Cannot validate/Missing Information in DEV/COM : ({incomingRGA})")
            ManualPushlist.append(NotRGAProcessingList)
            try:
                save_n_closeCOM = driver.find_element_by_id('save_and_close_document').click()
            
            except NoSuchElementException:
                cancelCOM = driver.find_element_by_id('cancel_document').click()   

            try:
                    wait = WebDriverWait(driver,5)
                    wait.until(EC.presence_of_element_located((By.ID, "save_and_close_document")))
                    driver.find_element_by_id('save_and_close_document').click()

            except TimeoutException or NoSuchElementException:
                    wait = WebDriverWait(driver,5)
                    wait.until(EC.presence_of_element_located((By.ID, "cancel_document")))
                    driver.find_element_by_id('cancel_document').click() 


        
        
    # DEV DATA FETCH END
        
    except TimeoutException or NoSuchElementException:
        print(f'DISCREPANCY ADDED: Issue: Cannot find {incomingCOM}!')
        discrepancyDEVList.append(date)
        discrepancyDEVList.append(incomingRGA)
        discrepancyDEVList.append(incomingCOM)
        discrepancyDEVList.append(incomingCODE)
        discrepancyDEVList.append(incomingSN)
        discrepancyDEVList.append(incomingUNITTYPE)
        discrepancyDEVList.append(" ")
        discrepancyDEVList.append(f"Cannot find COM: ({incomingCOM})")

        if len(discrepancyDEVList)!= 0:
            discrepancyList.append(discrepancyDEVList)
            
        else:
            pass
        
      

def getFile(): # FUNCTION TO PROCESS THE INCOMING.CSV FILE

        sensor_list = ["SENSOR_GEN6"]
        transmitter_list = ["FIREFLY", "G6_NUEVO", "TX", "TX_GEN6"]
        accessory_list = ["CABLE", "CHARGER_N_CABLE", "CHARGER"]
        receiver_list = ["SCOUT", "SCOUT_GEN6", "RX"]

        try:
            filename = input("FILENAME: ")
            filenamePath  = str(os.path.join(folderDate, filename))
            with open(filenamePath) as infile:
                for rows in infile:
                    y = rows.replace("\n", "")
                    x = y.replace(",", " ")
                    z = x.split(" ")

                    RGA = z[1]
                    COM = "COM-"+z[2]
                    CODE = z[3]
                    SN = z[4]
                    UNITTYPE = z[5]
                    

                    if UNITTYPE in sensor_list:
                        RETURN_TYPE = "Sensor"

                    elif UNITTYPE in transmitter_list:
                        RETURN_TYPE = "Transmitter"
                    
                    elif UNITTYPE in accessory_list:
                        RETURN_TYPE = "Accessory"

                    elif UNITTYPE in receiver_list:
                        RETURN_TYPE = "Receiver"

                    else:
                        RETURN_TYPE = ""

                    searchCOM(RGA,COM,CODE,SN,RETURN_TYPE,UNITTYPE)

        except FileNotFoundError:
            print('File Not Found! Try Again!')
            getFile()




def getDEV():
    deviceAnalbtn = driver.find_element_by_id("Device_Analysis_3")

    try:
        value = driver.find_element_by_id("div_Device_Analysis_3").get_attribute("style")
        if value == "display: none;":
            deviceAnalbtn.click()
    except:
        pass

    byNumberBtn = driver.find_element_by_id("BY_NUMBER_22_P")
    byNumberBtn.click()


def openPrf():
    shippingProfilebtn = driver.find_element_by_id("4._Shipping_Profile_22")
    shippingProfilebtn.click()

    try: 
        value = driver.find_element_by_id("div_4._Shipping_Profile_22").get_attribute("style")
        if value == 'display: none;':
            shippingProfilebtn.click()
    except:
        pass
    
    productReceiptBtn = driver.find_element_by_id("DX_RPR_RGA_PRODUCT_RECEIPT_PROFILE_P")
    productReceiptBtn.click()

def validateDEV(devlist):
    if all(name in devResult2 for name in devlist):
        return True
    else: 
        return False
   
        



def insidePRF(RGA,COM,CODE,SerialNumber,ReturnType,comOwner,LotNumber,UnitType, disposition):
    
    devlist = []
    devlist.clear()
    prfDiscrepancyList = []
    prfDiscrepancyList.clear()
    prfManualPush =[]
    prfManualPush.clear()
    completed_list = []
    
    try:
        wait = WebDriverWait(driver, 5)
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '[name = "RGA_NUMBER_1_P_SEARCH"]')))
        rgaColumn = driver.find_element_by_css_selector('[name = "RGA_NUMBER_1_P_SEARCH"]')
        rgaColumn.clear()
        rgaColumn.send_keys(RGA)
        rgaColumn.send_keys(Keys.RETURN)
    except NoSuchElementException:
        print("Cannot find Search Element!")
        driver.refresh()

    try:
        wait = WebDriverWait(driver, 5)
        wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="div_virtual_container"]/table/tbody')))
        prfSerialNumber = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr/td[6]').text
        result = driver.find_element_by_xpath('//*[@id="div_virtual_container"]/table/tbody/tr')
        result.click()
    except NoSuchElementException:
        print(f"DISCREPANCY ADDED: Issue: PRF not created! RGA:({RGA})")
        prfDiscrepancyList.append(date)
        prfDiscrepancyList.append(RGA)
        prfDiscrepancyList.append(COM)
        prfDiscrepancyList.append(CODE)
        prfDiscrepancyList.append(SerialNumber)
        prfDiscrepancyList.append(ReturnType)
        prfDiscrepancyList.append(comOwner)
        prfDiscrepancyList.append(f"PRF not created for RGA: {RGA}")
        if len(prfDiscrepancyList)!=0:
            discrepancyList.append(prfDiscrepancyList)
        else:
            pass    
        
    html_list = driver.find_element_by_id("DX_RPR_DEVICE_ANALYSIS_MULTIVALUED_LINKS_P_links_list")
    items = html_list.find_elements_by_tag_name("li")
    for item in items:
        text = item.text
        a = text.split(" ")
        dev = a[2]
        devlist.append(dev)
    
        
   
    validateDEV(devlist)
    
    if validateDEV(devlist) == True:
        try:

            prfRGA = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[2]/tbody/tr/td/table[2]/tbody/tr[2]/td/table/tbody/tr/td[1]/span[1]").text
            # prfSerialNumber = driver.find_element_by_xpath('//*[@id="DX_RPR_PRODUCT_INFORMATION_P_DX_RPR_SERIAL_NUMBER_P_0"]').text
            

           
            try:
                prfDisposition = driver.find_element_by_id("DX_RPR_PRODUCT_INFORMATION_P_DX_RPR_PRODUCT_DISPOSITION_P_0_AUTOCOMPLETE").get_attribute("value")
            except NoSuchElementException:
                prfDisposition = driver.find_element_by_css_selector("[name = 'DX_RPR_PRODUCT_INFORMATION_P_DX_RPR_PRODUCT_DISPOSITION_P_0_AUTOCOMPLETE']").get_attribute("value")


                
            if SerialNumber == prfSerialNumber or  SerialNumber == "N/A" or SerialNumber == "NI":
                pass
                if disposition == prfDisposition:
                    pass
                    if RGA == prfRGA:
                        try:
                            investigatorReceivedbtn = driver.find_element_by_xpath('//*[@id="dx_rpr_investigator_received_button_p"]').click()

                            wait = WebDriverWait(driver, 2)
                            wait.until(EC.presence_of_all_elements_located((By.XPATH, "/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td[2]/span[1]")))
                            investigatorReceivedDate = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td[2]/span[1]").text

                            com2 = COM.split("-")
                            finalCOM = str(f'{com2[1]}-{com2[2]}')

                            completed_list.append(date)
                            completed_list.append(RGA)
                            completed_list.append(finalCOM)
                            completed_list.append(CODE)
                            completed_list.append(SerialNumber)
                            completed_list.append(UnitType)
                            completed_list.append("RELIANCE")
                            completed_list.append("INVESTIGATOR RECEIVED")
                            completed_list.append(ReturnType)
                            completed_list.append(f"{user_fullname}({user_name})")
                            Completed.append(completed_list)
                        except NoSuchElementException:
                            investigatorReceivedDate = driver.find_element_by_xpath("/html/body/form/center/table[2]/tbody/tr[4]/td[2]/div[1]/div/div/table[4]/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td[2]/span[1]").text
                            
                            date_obj = datetime.strptime(investigatorReceivedDate, '%b %d, %Y').strftime('%Y-%m-%d')
                            
                            if date_obj != date:
                                print(f"ALERT: ISSUE: PRF ALREADY PUSHED {date_obj} RGA: {RGA}")
                                com2 = COM.split("-")
                                finalCOM = str(f'{com2[1]}-{com2[2]}')
                                prfManualPush.append(date)
                                prfManualPush.append(RGA)
                                prfManualPush.append(finalCOM)
                                prfManualPush.append(CODE)
                                prfManualPush.append(SerialNumber)
                                prfManualPush.append(UnitType)
                                prfManualPush.append(f"{user_fullname}({user_name})")
                                prfManualPush.append(f"ALREADY PUSHED: {investigatorReceivedDate}")
                                prfManualPush.append(ReturnType)
                                Completed.append(completed_list)
                            else:
                                com2 = COM.split("-")
                                finalCOM = str(f'{com2[1]}-{com2[2]}')
                                completed_list.append(date)
                                completed_list.append(RGA)
                                completed_list.append(finalCOM)
                                completed_list.append(CODE)
                                completed_list.append(SerialNumber)
                                completed_list.append(UnitType)
                                completed_list.append("RELIANCE")
                                completed_list.append("INVESTIGATOR RECEIVED")
                                completed_list.append(ReturnType)
                                completed_list.append(f"{user_fullname} ({user_name})")
                                Completed.append(completed_list)

                    else:
                        print(f"DISCREPANCY ADDED: Issue: RGA doesn't match in PRF! RGA: ({RGA})")
                        prfDiscrepancyList.append(date)
                        prfDiscrepancyList.append(RGA)
                        prfDiscrepancyList.append(COM)
                        prfDiscrepancyList.append(CODE)
                        prfDiscrepancyList.append(SerialNumber)
                        prfDiscrepancyList.append(ReturnType)
                        prfDiscrepancyList.append(UnitType)
                        prfDiscrepancyList.append(comOwner)
                        prfDiscrepancyList.append(f"RGA doesn't match in PRF! Physical RGA: {RGA} PRF RGA: {prfRGA} ")
                        if len(prfDiscrepancyList)!=0:
                            discrepancyList.append(prfDiscrepancyList)
                else:
                    print(f"DISCREPANCY ADDED: Issue: Disposition doesn't match in PRF! RGA: ({RGA})")
                    prfDiscrepancyList.append(date)
                    prfDiscrepancyList.append(RGA)
                    prfDiscrepancyList.append(COM)
                    prfDiscrepancyList.append(CODE)
                    prfDiscrepancyList.append(SerialNumber)
                    prfDiscrepancyList.append(UnitType)
                    prfDiscrepancyList.append(comOwner)
                    prfDiscrepancyList.append(f"Serial Number doesn't match in PRF: Physical Disposition: {disposition} PRF Disposition: {prfDisposition}")
                    prfDiscrepancyList.append(ReturnType)
                    if len(prfDiscrepancyList)!=0:
                        discrepancyList.append(prfDiscrepancyList)


            else:
                print(f"DISCREPANCY ADDED: Issue: Serial Number doesn't match in PRF! RGA: ({RGA})")
                prfDiscrepancyList.append(date)
                prfDiscrepancyList.append(RGA)
                prfDiscrepancyList.append(COM)
                prfDiscrepancyList.append(CODE)
                prfDiscrepancyList.append(SerialNumber)
                prfDiscrepancyList.append(UnitType)
                prfDiscrepancyList.append(comOwner)
                prfDiscrepancyList.append(f"Serial Number doesn't match in PRF! Physical SN: {SerialNumber} PRF SN: {prfSerialNumber}")
                prfDiscrepancyList.append(ReturnType)
                if len(prfDiscrepancyList)!=0:
                    discrepancyList.append(prfDiscrepancyList)

        except NoSuchElementException:
            print(f"Something went wrong! RGA: {RGA}")

    else:
        print(f"ALERT: ISSUE: Missing DEV/COM Validation or Missing Label (RGA: {RGA})")
        prfManualPush.append(date)
        prfManualPush.append(RGA)
        prfManualPush.append(COM)
        prfManualPush.append(CODE)
        prfManualPush.append(SerialNumber)
        prfManualPush.append(UnitType)
        prfManualPush.append(comOwner)
        prfManualPush.append("Missing DEV/COM Validation or Missing Label")
        prfManualPush.append(ReturnType)
        ManualPushlist.append(prfManualPush)
    try:
        wait = WebDriverWait(driver, 2)
        wait.until(EC.presence_of_element_located((By.ID, "save_and_close_document")))
        savecloseBtn = driver.find_element_by_id("save_and_close_document").click()
    except NoSuchElementException:
        print("here")


def pushRGA():
    
    for rows in MasterPushlist:
        RGA = rows[1]
        COM = rows[2]
        CODE = rows[3]
        SerialNumber = rows[4]
        ReturnType = rows[5]
        UnitType = rows[6]
        LotNumber = rows[7]
        comOwner = rows[9]
        disposition = rows[11]


       
        insidePRF(RGA,COM,CODE,SerialNumber,ReturnType,comOwner,LotNumber,UnitType, disposition)

def mkSheet():
    discrepancy_List = str(f'Discrepancy_List({date}).csv')
    completedList = str(f'Completed_List({date}).csv')
    Verify_List = str(f'Verify_List({date}).csv')

    x: str = os.path.join(folderDatePATH, discrepancy_List)
    y: str = os.path.join(folderDatePATH, completedList)
    z: str = os.path.join(folderDatePATH, Verify_List)
    sheetExists1 = os.path.exists(x)
    sheetExists2 = os.path.exists(y)
    sheetExists3 = os.path.exists(z)




    if sheetExists1 == 'False':
        with open (x, 'w' ,newline='', encoding='utf-8') as infile:
            writer = csv.writer(infile)
            writer.writerows(discrepancyList)
            infile.close()
    else:
        with open (x , 'a'  , newline='', encoding='utf-8') as infile:
            append = csv.writer(infile)
            append.writerows(discrepancyList)
            infile.close()

    if sheetExists2 == 'False':
        with open (y, 'w' ,newline='', encoding='utf-8') as infile:
            writer = csv.writer(infile)
            writer.writerows(Completed)
            infile.close()
    else:
        with open (y , 'a'  , newline='', encoding='utf-8') as infile:
            append = csv.writer(infile)
            append.writerows(Completed)
            infile.close()
    
    if sheetExists3 == 'False':
        with open (z, 'w' ,newline='', encoding='utf-8') as infile:
            writer = csv.writer(infile)
            writer.writerows(ManualPushlist)
            infile.close()
    else:
        with open (z , 'a'  , newline='', encoding='utf-8') as infile:
            append = csv.writer(infile)
            append.writerows(ManualPushlist)
            infile.close()

# def combinesheets():
#     import pandas as pd
#     import os as os
#     import glob as gl

#     os.chdir(folderDatePATH)

#     Excel_File = pd.ExcelFile("new_excel_filename.xlsx")
#     Sheet_Name = Excel_File.sheet_names
#     length_of_Sheet = len(Excel_File.sheet_names)
#     print("List of sheets in you xlsx file :\n",Sheet_Name)

#     for i in range(0,length_of_Sheet):
#         df = pd.read_excel("new_excel_filename.xlsx", sheet_name = i)
#         df = df.iloc[:,0:3]
#         df.to_csv(Sheet_Name[i]+".csv", index = False)
#         print("Created :",Sheet_Name[i],".csv")
        
#     filenames = [i for i in gl.glob('*.{}'.format('csv'))]
#     combined_csv = pd.concat([pd.read_csv(f) for f in filenames ])
#     combined_csv.to_csv( "combined_csv_filename.csv", index=False, encoding='utf-8-sig')


def main():
    mkFolder(folderDatePATH)
    openBrowser(OKTA, driver)
    login(driver, Keys)
    openRelianceApp(driver)
    openComplaintApp(driver)
    openDeviceAnalysis(driver,Keys)
    getFile()
    openPrf()
    pushRGA()
    mkSheet()

    # combinesheets()

    
    


    time.sleep(5)
    driver.quit()


if __name__ == '__main__':
    main()
