import gspread
from oauth2client.service_account import ServiceAccountCredentials
import FUNC
import os
import time
import datetime
from tkinter import *
from tkinter import messagebox
import json
import re
import random
import HTML

# hide main window of tkinter
root = Tk()
root.withdraw()

# current time and time in string format
x = datetime.datetime.now()
timen = str(datetime.datetime.now())
month = x.strftime('%b')
day = x.strftime('%d')
# messagebox.showinfo("month", month )

# working directory path
dir = os.path.dirname(__file__)

# for reporting create HTML file
HTMLFILE = os.path.join(
    'C:\\Python\\Result\\' + 'GSheet_Result' + '_' + x.strftime('%x-%H-%M').replace('/', '-') + '.html')
file = open(HTMLFILE, 'w')

UserReport = os.path.join(
    'C:\\Python\\Result\\' + 'Month_Result' + '_' + x.strftime('%x-%H-%M').replace('/', '-') + '.html')
fileUser = open(UserReport, 'w')

## Below code is to use HTML.py for reporting
t = HTML.Table(header_row=['User Report', month + ' month'])


def Report(str1, str2):
    t.rows.append([str1, str2])


# Create basic structure of HTML file
FUNC.basicHtmlStructure(file)

# update row in report table
FUNC.ReportN(file, 'Process Start', 'Done', str(datetime.datetime.now()))

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
# credentials = ServiceAccountCredentials.from_json_keyfile_name(os.path.join(dir,'Resources/Report Automation-4eb984ef55aa.json'),scope)
try:
    credentials = ServiceAccountCredentials.from_json_keyfile_name('Report Automation-4eb984ef55aa.json', scope)
    gfile = gspread.authorize(credentials)  # authenticate with Google
    time.sleep(.5)
    # Read User detail from json file and start login in console for each user and preparing reporting
    with open(os.path.join(dir, 'Resources/dataUserNew.json')) as json_file:
        datau = json.load(json_file)
        for p in datau['Balance']:
            loginName = p['name']
            login = p['Userid']
            bal = p['Balance']
            bal = (float(re.sub(r"\D", "", bal)) * 1000) + random.randint(300, 800)
            # messagebox.showinfo("balance", loginName)
            # messagebox.showinfo("balance", bal )
            sheetName = '_EngagementC_' + loginName
            mysheet = gfile.open(sheetName).worksheet('Phase2')
            time.sleep(.5)
            Report(loginName, 'Balance')
            try:
                colCurMonth = mysheet.find(month + '19').col
                FUNC.ReportN(file, 'Updating sheet of user: ' + loginName + ' for Month: ' + month + ' 19', 'PASS',
                             str(datetime.datetime.now()))
                if int(day) > 26:
                    ccol = chr(colCurMonth + ord('A'))
                    pcol = chr((colCurMonth - 1) + ord('A'))
                    mysheet.update_cell(2, colCurMonth + 1, '=' + pcol + '4')
                    mysheet.update_cell(5, colCurMonth + 1, '=' + ccol + '4' + '-' + ccol + '2' + '-' + ccol + '3')
                    mysheet.update_cell(6, colCurMonth + 1, '=0.2*' + ccol + '5')
                    mysheet.update_cell(7, colCurMonth + 1, '=' + ccol + '5' + '-' + ccol + '6')
                    mysheet.update_cell(8, colCurMonth + 1, '=' + ccol + '7' + '/(' + ccol + '2' + '+' + ccol + '3)')
                    mysheet.update_cell(9, colCurMonth + 1, '=' + pcol + '9*(1+' + ccol + '8)')
                    mysheet.update_cell(11, colCurMonth + 1, '=' + pcol + '11-' + ccol + '6' + '+' + ccol + '10')

                    mysheet.update_cell(4, colCurMonth, bal)
                    time.sleep(.5)
                    mysheet.update_cell(4, colCurMonth + 1, bal)
                    time.sleep(.5)
                    Report('Initial value', mysheet.acell(ccol + '2').value)
                    Report('Final closing', mysheet.acell(ccol + '4').value)
                    Report(month + ' month Profit', mysheet.acell(ccol + '5').value)
                    Report('Return', mysheet.acell('A1').value)
                    Report('Profit Pool', mysheet.acell(ccol + '11').value)

                    FUNC.ReportN(file, 'Final Balance for Month: ' + month + ' is ' + mysheet.acell(ccol).value + '',
                                 'DONE', str(datetime.datetime.now()))
                else:
                    ccol = chr((colCurMonth - 1) + ord('A'))
                    pcol = chr((colCurMonth - 2) + ord('A'))
                    mysheet.update_cell(2, colCurMonth, '=' + pcol + '4')
                    mysheet.update_cell(5, colCurMonth, '=' + ccol + '4' + '-' + ccol + '2' + '-' + ccol + '3')
                    mysheet.update_cell(6, colCurMonth, '=0.2*' + ccol + '5')
                    mysheet.update_cell(7, colCurMonth, '=' + ccol + '5' + '-' + ccol + '6')
                    mysheet.update_cell(8, colCurMonth, '=' + ccol + '7' + '/(' + ccol + '2' + '+' + ccol + '3)')
                    mysheet.update_cell(9, colCurMonth, '=' + pcol + '9*(1+' + ccol + '8)')
                    mysheet.update_cell(11, colCurMonth, '=' + pcol + '11-' + ccol + '6' + '+' + ccol + '10')

                    # print(mysheet.cell(4, colCurMonth-1))
                    mysheet.update_cell(4, colCurMonth - 1, bal)
                    time.sleep(.5)
                    mysheet.update_cell(4, colCurMonth, bal)
                    # print(mysheet.cell(4, colCurMonth))

                    Report('Initial value', mysheet.acell(pcol + '2').value)
                    Report('Final closing', mysheet.acell(pcol + '4').value)
                    Report(month + ' month Profit', mysheet.acell(pcol + '5').value)
                    Report('Return', mysheet.acell('A1').value)
                    Report('Profit Pool', mysheet.acell(ccol + '11').value)
                    time.sleep(.5)
                    FUNC.ReportN(file,
                                 'Final closing balance for ' + month + ' month of above user is ' + mysheet.acell(
                                     pcol + '4').value + '', 'DONE', str(datetime.datetime.now()))
            except:
                FUNC.ReportN(file, 'Current or last Month entry not found. Please add same and rerun', 'FAIL', timen)
except:
    FUNC.ReportN(file, 'Not able to open gsheet on Drive', 'FAIL', timen)

htmlcode = str(t)
fileUser.write(htmlcode)
fileUser.write('<p>')
time.sleep(.5)
fileUser.close()
time.sleep(.5)
with open(os.path.join(dir, 'Resources/dataUserNew.json')) as json_file:
    datau = json.load(json_file)
    for p in datau['Balance']:
        loginName = p['name']
        searchExp = """<TR bgcolor="snow">
  <TD>{user}</TD>
  <TD>Balance</TD>
 </TR>""".format(user=loginName)
        replaceExp = """<TR bgcolor="khaki">
  <TD><font color=black face=century size=3>{user}</TD>
  <TD><font color=black face=century size=3>Balance</TD>
 </TR>""".format(user=loginName)

        FUNC.replace(UserReport, searchExp, replaceExp)

searchExp1 = 'TD>Return</TD>'
replaceExp1 = 'TD><font color=blue face=century size=3>Return</TD>'

FUNC.replace(UserReport, searchExp1, replaceExp1)

searchExp2 = 'TD>Profit Pool</TD>'
replaceExp2 = 'TD><font color=red face=century size=3>Profit Pool</TD>'

FUNC.replace(UserReport, searchExp2, replaceExp2)

# mysheet.delete_row(100)
# print(mysheet.row_count)
# #list_of_lists = worksheet.get_all_values()
#
# mysheet.append_row(['moredata','56','N'])
#
# mysheet.delete_row(2)
#
# print(mysheet.acell('A2'))
# print(mysheet.acell('A2').value)
# print(mysheet.cell(2,1))
# print(mysheet.cell(2,2).value)
# print(mysheet.acell('A2').col)
# print(mysheet.cell(2,2).row)
#
# mysheet.update_acell('B2','NewV')
# mysheet.update_cell(1,2,'NewVm')
# mysheet.update_acell('B3','NewV')
#
# print(mysheet.rowcount)

# print(mysheet.findall('NewV'))
# list_Cells = mysheet.findall('NewV')
# # Find a cell matching a regular expression
# #amount_re = re.compile(r'(Big|Enormous) dough')
# #cell = worksheet.find(amount_re)
#
# list_Cells[1].value = 'newdata'
#
# print(mysheet.acell('B3').value)
#
# mysheet.update_cells(list_Cells)
#
# for cell in list_Cells:
#     cell.value = 'loopV'
#
# cell_list = mysheet.range('A1:B2')
#
# #sh2 = gfile.create('Newsheet')
# #sh2.share('neerajtoday@gmail.com', perm_type='user', role='writer')
# #sh2.share('neeraj.ce@gmail.com', perm_type='user', role='writer')
# mysheet1 = sh2.add_worksheet(title="Nworksheet", rows="50", cols="20")
# #sh.del_worksheet(worksheet)
#
# # Select worksheet by index. Worksheet indexes start from zero
# #worksheet = sh.get_worksheet(0)
#
# # By title
# #worksheet = sh.worksheet("January")
#
# # Most common case: Sheet1
# #worksheet = sh.sheet1
#
# # Get a list of all worksheets
# #worksheet_list = sh.worksheets()


# start_col1 = colCurMonth
# end_col1 = colCurMonth
# start_row1 = 2
# end_row1 = 11
#
# cell_range1 = '{col_i}{row_i}:{col_f}{row_f}'.format(
#     col_i=chr((start_col1 - 1) + ord('A')),  # converts number to letter
#     col_f=chr((end_col1 - 1) + ord('A')),  # subtract 1 because of 0-indexing
#     row_i=start_row1,
#     row_f=end_row1)
#
# range1 = mysheet.range(cell_range1)
# # print (range1)
#
# start_col2 = colCurMonth + 1
# end_col2 = colCurMonth + 1
# start_row2 = 2
# end_row2 = 11
#
# cell_range2 = '{col_i}{row_i}:{col_f}{row_f}'.format(
#     col_i=chr((start_col2 - 1) + ord('A')),  # converts number to letter
#     col_f=chr((end_col2 - 1) + ord('A')),  # subtract 1 because of 0-indexing
#     row_i=start_row2,
#     row_f=end_row2)
#
# range2 = mysheet.range(cell_range2)
# print(range2)
#
# # for range1_cell,range2_cell in zip(range1,range2):
# #    range2_cell.value = range2_cell.value
#
# for cell in range2:
#     cell.value = '=V4'
#
# mysheet.update_cells(range2, value_input_option='USER_ENTERED')
# # formula = wks.acell('A2', value_render_option='FORMULA').value
# # wks.update_acell('B3', formula)
# # mysheet.update_cells(range2)
# # print(range2)
#
# # formula = mysheet.acell('W2').input_value
#
# # then paste the value to a new cell
# # mysheet.update_acell('W2', '=V4')
#
# messagebox.showinfo("Column count", "ok")
# # print(mysheet.cell(4, colCurMonth))
# mysheet.update_cell(4, colCurMonth, bal)
# time.sleep(.5)
# mysheet.update_cell(4, colCurMonth + 1, bal)
# # print(mysheet.cell(4, colCurMonth))

# print(mysheet.get_all_records())
# print(mysheet.row_count)
# print(mysheet.col_count)
# messagebox.showinfo("Column count", mysheet.col_count)
# mysheet.update_acell('W2', '=V4')

# sheetName = 'EngagementC'
# messagebox.showinfo("a",sheetName)
# mysheet = gfile.open(sheetName).Phase2 # open sheet
