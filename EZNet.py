import openpyxl
import pprint
import time
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select


# Set up File Selection
root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()

# Set up HTML Interactions
driver = webdriver.Firefox()
# driver.get("http://www.python.org")
driver.get("file:///C:/Users/jwhitten/Downloads/EZ%20Net%20Source.html")
# assert "Python" in driver.title
# elem = driver.find_element_by_name("q")
# elem.clear()
# elem.send_keys("pycon")
# elem.send_keys(Keys.RETURN)
# assert "No results found." not in driver.page_source
# driver.close()

# Set up Excel using OpenPyXL

print('Reading in Billing data...')

wb = openpyxl.load_workbook(file_path)
sheet = wb['Data Entry']
eznetData = {}

print('Read in Billing...')
for row in range(2, sheet.max_row):
    # Each row in the spreadsheet has data
    # Member ID	| SSES Provider ID | Diagnosis | Place of Service | Procedure Code | Modifier | Date of Service From |
    # Date of Service to | Units | Total Billed Per Line | | Last Name | First Name | Date of Birth

    # Identifying Information
    memberID = str(sheet['A' + str(row)].value)

    if memberID == "None":
        break

    ssesProviderID = str(sheet['B' + str(row)].value)

    # Specific Invoice Information
    diag = str(sheet['C' + str(row)].value)
    pos = str(sheet['D' + str(row)].value)
    procCode = str(sheet['E' + str(row)].value)
    modifier = str(sheet['F' + str(row)].value)
    serviceFrom = str(sheet['G' + str(row)].value)
    serviceTo = str(sheet['H' + str(row)].value)
    units = str(sheet['I' + str(row)].value)
    totalBilled = str(sheet['J' + str(row)].value)

    # Customer Info
    lastName = str(sheet['L' + str(row)].value)
    firstName = str(sheet['M' + str(row)].value)
    dob = str(sheet['N' + str(row)].value)

    # Make sure the key for this memberID exists
    eznetData.setdefault(memberID, {})
    # Set the default data in case something is missing
    eznetData[memberID].setdefault(ssesProviderID, {'Diagnosis:': '',
                                                    'Place of Service:': '',
                                                    'Procedure Code:': '',
                                                    'Modifier': '',
                                                    'Date of Service From:': '',
                                                    'Date of Service To:': '',
                                                    'Units': '0',
                                                    'Total Billed Per Line:': '$0.00',
                                                    'Last Name:': '',
                                                    'First Name:': '',
                                                    'Date of Birth:': ''})

    # Write the data to billingData
    eznetData[memberID][ssesProviderID]['Diagnosis:'] = diag
    eznetData[memberID][ssesProviderID]['Place of Service:'] = pos
    eznetData[memberID][ssesProviderID]['Procedure Code:'] = procCode
    eznetData[memberID][ssesProviderID]['Modifier:'] = modifier
    eznetData[memberID][ssesProviderID]['Date of Service From:'] = serviceFrom
    eznetData[memberID][ssesProviderID]['Date of Service To:'] = serviceTo
    eznetData[memberID][ssesProviderID]['Units:'] = units
    eznetData[memberID][ssesProviderID]['Total Billed Per Line:'] = totalBilled
    #
    eznetData[memberID][ssesProviderID]['Last Name:'] = lastName
    eznetData[memberID][ssesProviderID]['First Name:'] = firstName
    eznetData[memberID][ssesProviderID]['Date of Birth:'] = dob
    #
    driver.find_element_by_id('MainContent_txtHPMemberID').clear()
    driver.find_element_by_id('MainContent_txtHPMemberID').send_keys(memberID)
    driver.find_element_by_id('MainContent_txtProviderID').clear()
    driver.find_element_by_id('MainContent_txtProviderID').send_keys(ssesProviderID)
    #
    driver.find_element_by_id('MainContent_txtDiagnosiCode').clear()
    driver.find_element_by_id('MainContent_txtDiagnosiCode').send_keys(diag)
    Select(driver.find_element_by_id('MainContent_ddlPlaceOfService')).select_by_value(pos)
    driver.find_element_by_id('MainContent_txtProcedureCode').clear()
    driver.find_element_by_id('MainContent_txtProcedureCode').send_keys(procCode)
    # This is a drop down box.  Need different Code.
    Select(driver.find_element_by_id('MainContent_ddlModifier1')).select_by_value(modifier)
    #
    # This should work on the real page.  But commenting out for now.
    # driver.find_element_by_id('MainContent_wdcDateServiceFrom_clientState').clear()
    # driver.find_element_by_id('MainContent_wdcDateServiceFrom_clientState').send_keys(serviceFrom)
    # driver.find_element_by_id('MainContent_wdcDateServiceTo_clientState').clear()
    # driver.find_element_by_id('MainContent_wdcDateServiceTo_clientState').send_keys(serviceTo)
    driver.find_element_by_id('MainContent_txtQty').clear()
    driver.find_element_by_id('MainContent_txtQty').send_keys(units)
    driver.find_element_by_id('MainContent_txtBillCharges').clear()
    driver.find_element_by_id('MainContent_txtBillCharges').send_keys(totalBilled)

driver.close()

# Open a new text file and write the contents of billingData to it
# Don't really need this, but adding it anyways.  Should be interacting with the webpage now.
print('Writing EzNet information...')
eznetInfo = open('eznetData.py', 'w')
# eznetInfo.write('eznetData = \n')
eznetInfo.write(pprint.pformat(eznetData))
eznetInfo.close()
print('Done.')
