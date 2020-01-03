from openpyxl.styles import Color, PatternFill, Font, Border
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl, time, datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

root = tk.Tk()
root.withdraw()
filename = filedialog.askopenfilename(initialdir="/",title="Open File",filetypes=(("Excel", "*.xlsx"), ("All Files", "*.*")))
wb = openpyxl.load_workbook(filename)
sh = wb['Sheet1']
tran = []
stat = []
amount = []
count = sh.max_row+1
for i in range(2, count):
    tran.append(sh.cell(row=i, column=1).value)

chromedriver = filedialog.askopenfilename(initialdir="/",title="Open File",filetypes=(("Executable", "*.exe"), ("All Files", "*.*")))
browser = webdriver.Chrome(chromedriver)
browser.get('https://fastagcsc.bankofbaroda.com/BOBPOS/Default.aspx')
user_n = browser.find_element_by_id('txtUserName')
user_n.send_keys('sj103402')
pass_w = browser.find_element_by_id('txtPassword')
pass_w.send_keys('Welcome@2716')
result = messagebox.askokcancel('Captcha', 'Navigate to browser window and Enter Captcha.\nPress Ok once Done.')
browser.get('https://fastagcsc.bankofbaroda.com/BOBPOS/pages/Retailer/MakePayment/MPOSDeposit.aspx')
for i in range(0,len(tran)):
    search_bar = browser.find_element_by_id('BodyContent_txtSearchReferenceno')
    search_btn = browser.find_element_by_id('BodyContent_btnSearch')
    search_bar.send_keys(tran[i])
    search_btn.click()
    try:
        status = browser.find_element_by_xpath('//*[@id="BodyContent_gvDeposits"]/tbody/tr[2]/td[5]').text
    except:
        status = 'Nope'
    try:
        amt = int(float(browser.find_element_by_xpath('//*[@id="BodyContent_gvDeposits"]/tbody/tr[2]/td[4]').text))
    except:
        amt = 'NA'
    browser.find_element_by_id('BodyContent_btnSearchReset').click()
    stat.append(status)
    amount.append(amt)
for i in range(2,count+2):
    try:
        sh.cell(row=i, column=4).value = stat[i-2]
        sh.cell(row=i, column=5).value = amount[i-2]
        if sh.cell(row=i, column=4).value == 'Approved':
            sh.cell(row=i, column=4).fill = PatternFill(fill_type='solid', start_color='00ff00', end_color='00ff00')
        if sh.cell(row=i, column=4).value == 'Rejected':
            sh.cell(row=i, column=4).fill = PatternFill(fill_type='solid', start_color='ff0000', end_color='ff0000')
        if sh.cell(row=i, column=4).value == 'Nope':
            sh.cell(row = i, column = 4).value = 'No Entry Found'
            sh.cell(row=i, column=4).fill = PatternFill(fill_type='solid', start_color='ff0000', end_color='ffff00')
    except:
        pass
    wb.save(filename)
    

