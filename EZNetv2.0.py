import openpyxl
import pprint
from ctypes import *
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains


def get_excel():
    # Set up File Selection
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename()

    return file_path


def confirm():
    # Set up an OK box
    yesno = windll.user32.MessageBoxW(0, 'Is the data correct?', 'Data Confirm', 4)
    if yesno == 6:
        return 'Yes'
    elif yesno == 7:
        return 'No'


def get_website():
    # Set up HTML Interactions
    driver = webdriver.Firefox()

    driver.get("https:///tpawebpcg.pcgus.com/EZ-NET60/Login.aspx")
    # login(driver)
    print('Login')
    while True:
        yn = confirm()
        if yn == 'Yes':
            # driver.find_element_by_id('MainContent_ctl00_btnLogin').click()
            break

    driver.get("https:///tpawebpcg.pcgus.com/EZ-NET60/Webportal/EZNET/ClaimSubmissionEntry.aspx")
    # driver.get("file:///C:/Users/jwhitten/Downloads/EZ%20Net%20Source.html")
    assert "EZ-NET™" in driver.title
    return driver


def import_excel(file_path):
    # Set up Excel using OpenPyXL
    print('Reading in Billing data...')

    wb = openpyxl.load_workbook(file_path)
    sheet = wb['Data Entry']
    eznet_data = []

    print('Read in Billing...')
    for row in range(2, sheet.max_row):
        # Each row in the spreadsheet has data
        # Member ID	| SSES Provider ID | Diagnosis | Place of Service | Procedure Code | Modifier |
        # Date of Service From | Date of Service to | Units | Total Billed Per Line | | Last Name | First Name |
        # Date of Birth

        member_id = str(sheet['A' + str(row)].value)

        # If excel file has blank lines at the end, escape.
        if member_id == "None":
            break

        proc_code = sheet['E' + str(row)].value
        # if 11 or 12 return 16404 or 16390 respectively
        pos = str(sheet['D' + str(row)].value)
        if '11' == pos:
            pos = '16404'
        elif '12' == pos:
            pos = '16390'
        mod = str(sheet['F' + str(row)].value)
        dosf = str(sheet['G' + str(row)].value)
        dost = str(sheet['H' + str(row)].value)
        units = str(sheet['I' + str(row)].value)
        total_billed = str(sheet['J' + str(row)].value)

        eznet_data.append([member_id, pos, proc_code, mod, units, total_billed, dosf, dost])

    return eznet_data


def login(webpage):
    webpage.find_element_by_id('MainContent_ctl00_txtUserName').clear()
    webpage.find_element_by_id('MainContent_ctl00_txtUserName').send_keys('SorryTestingPython')
    webpage.find_element_by_id('MainContent_ctl00_txtPassword').clear()
    webpage.find_element_by_id('MainContent_ctl00_txtPassword').send_keys('12345')

    webpage.find_element_by_id('MainContent_ctl00_btnLogin').click()


def put_data_on_page(webpage, data):
    webpage.find_element_by_id('MainContent_txtHPMemberID').clear()
    webpage.find_element_by_id('MainContent_txtHPMemberID').send_keys(data[0])
    webpage.find_element_by_id('MainContent_txtProviderID').clear()
    webpage.find_element_by_id('MainContent_txtProviderID').send_keys('1427179076-01')
    #
    webpage.find_element_by_id('MainContent_txtDiagnosiCode').clear()
    webpage.find_element_by_id('MainContent_txtDiagnosiCode').send_keys('Z59.9')
    Select(webpage.find_element_by_id('MainContent_ddlPlaceOfService')).select_by_value(data[1])

    click_diag(webpage)


def put_procedure_on_page(webpage, data):
    webpage.find_element_by_id('MainContent_txtProcedureCode').clear()
    webpage.find_element_by_id('MainContent_txtProcedureCode').send_keys(data[2])
    # This is a drop down box.  Need different Code.
    Select(webpage.find_element_by_id('MainContent_ddlModifier1')).select_by_value(data[3])
    #
    webpage.find_element_by_id('MainContent_txtQty').clear()
    webpage.find_element_by_id('MainContent_txtQty').send_keys(data[4])
    webpage.find_element_by_id('MainContent_txtBillCharges').clear()
    webpage.find_element_by_id('MainContent_txtBillCharges').send_keys(data[5])
    # actions = ActionChains(webpage)
    # actions.double_click(webpage.find_element_by_id('MainContent_txtBillCharges'))
    # actions.send_keys(webpage.find_element_by_id('MainContent_txtBillCharges'), data[5])
    # actions.perform()
    # This should work on the real page.  But commenting out for now.
    # webpage.find_element_by_id('MainContent_wdcDateServiceFrom_clientState').clear()
    # webpage.find_element_by_id('MainContent_wdcDateServiceFrom_clientState').send_keys(data[6])
    # webpage.find_element_by_id('MainContent_wdcDateServiceTo_clientState').clear()
    # webpage.find_element_by_id('MainContent_wdcDateServiceTo_clientState').send_keys(data[7])
    #
    click_proc(webpage)


def click_submit(webpage):
    # webpage.find_element_by_id('MainContent_btnSubmitRequest').click()
    print('Submit Clicked')


def click_diag(webpage):
    # webpage.find_element_by_id('MainContent_btnDiagCodes').click()
    print('Diag Clicked')


def click_proc(webpage):
    # webpage.find_element_by_id('MainContent_btnProcCodes').click()
    print('Proc Clicked')


def convert_excel(dataset):
    # Open a new text file and write the contents of billingData to it
    # Don't really need this, but adding it anyways.  Should be interacting with the webpage now.
    print('Writing EzNet information...')
    eznet_info = open('eznetData.py', 'w')
    # eznet_info.write('eznetData = \n')
    eznet_info.write(pprint.pformat(dataset))
    eznet_info.close()
    print('Done.')


def main():
    file_path = get_excel()
    eznet_data = import_excel(file_path)
    eznet_webpage = get_website()
    old_member_id = None
    old_pos = None
    yesno = 'No'

    # print(eznet_data)
    for member_data in eznet_data:
        if old_member_id != member_data[0] or old_pos != member_data[1]:
            if old_member_id is not None:
                print('First run')
                click_submit(eznet_webpage)
            # New member or new place of service
            put_data_on_page(eznet_webpage, member_data)
        put_procedure_on_page(eznet_webpage, member_data)

        old_member_id = member_data[0]
        old_pos = member_data[1]

        yesno = confirm()
        #    if yesno == 'No':
        #       break
        if yesno == 'No':
            break

    if yesno == 'Yes':
        click_submit(eznet_webpage)

    eznet_webpage.close()
    convert_excel(eznet_data)


if __name__ == "__main__":
    main()
