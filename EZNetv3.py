import openpyxl
import pprint
from ctypes import *
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.support.ui import Select


class Tree:
    def __init__(self, root):
        self.root = root
        self.children = {}

    def addNode(self, obj):
        self.children.append(obj)


class Node:
    def __init__(self, data):
        self.data = data
        self.children = []

    def addNode(self, obj):
        self.children.append(obj)


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
    driver.get("file:///C:/Users/jwhitten/Downloads/EZ%20Net%20Source.html")
    assert "EZ-NET™" in driver.title
    return driver


def import_excel(file_path):
    # Set up Excel using OpenPyXL
    print('Reading in Billing data...')

    wb = openpyxl.load_workbook(file_path)
    sheet = wb['Data Entry']
    eznet_data = {}

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

        # eznet_data[member_id] = {pos: [proc_code, mod, units, total_billed, dosf, dost]}
        member = Tree(member_id)
        member.addNode(Node(pos))

        member.children[0].addNode(Node(proc_code))
        member.children[0].children[0] = [mod, units, total_billed, dosf, dost]

        print(member)

    return eznet_data


def put_data_on_page(webpage, data, urn):
    webpage.find_element_by_id('MainContent_txtHPMemberID').clear()
    webpage.find_element_by_id('MainContent_txtHPMemberID').send_keys(urn)
    webpage.find_element_by_id('MainContent_txtProviderID').clear()
    webpage.find_element_by_id('MainContent_txtProviderID').send_keys('SSES')
    #
    webpage.find_element_by_id('MainContent_txtDiagnosiCode').clear()
    webpage.find_element_by_id('MainContent_txtDiagnosiCode').send_keys('R39.2')
    Select(webpage.find_element_by_id('MainContent_ddlPlaceOfService')).select_by_value(data[urn]['Place of Service'])


def put_procedure_on_page(webpage, dataset, urn, code):
    data = dataset[urn][code]

    webpage.find_element_by_id('MainContent_txtProcedureCode').clear()
    webpage.find_element_by_id('MainContent_txtProcedureCode').send_keys(code)
    # This is a drop down box.  Need different Code.
    Select(webpage.find_element_by_id('MainContent_ddlModifier1')).select_by_value(data['Modifier'])
    #
    # This should work on the real page.  But commenting out for now.
    # webpage.find_element_by_id('MainContent_wdcDateServiceFrom_clientState').clear()
    # webpage.find_element_by_id('MainContent_wdcDateServiceFrom_clientState').send_keys(data['Date of Service From'])
    # webpage.find_element_by_id('MainContent_wdcDateServiceTo_clientState').clear()
    # webpage.find_element_by_id('MainContent_wdcDateServiceTo_clientState').send_keys(data['Date of Service To'])
    webpage.find_element_by_id('MainContent_txtQty').clear()
    webpage.find_element_by_id('MainContent_txtQty').send_keys(data['Units'])
    webpage.find_element_by_id('MainContent_txtBillCharges').clear()
    webpage.find_element_by_id('MainContent_txtBillCharges').send_keys(data['Total Billed Per Line'])


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
    # print(eznet_data)
    for member_id, proc_code in eznet_data.items():
        print("proc_code")
        print(proc_code)
        # put_data_on_page(eznet_webpage, eznet_data, member_id, proc_code)
        # for key in proc_code:
        # print("eznet_data[member_id][proc_code]")
        # print(eznet_data[member_id]['Date of Service From'])
        # put_procedure_on_page(eznet_webpage, eznet_data, member_id, key)
        yesno = confirm()
        #    if yesno == 'No':
        #       break
        if yesno == 'No':
            break

    eznet_webpage.close()
    convert_excel(eznet_data)


if __name__ == "__main__":
    main()
