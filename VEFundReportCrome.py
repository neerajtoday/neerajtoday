# -------------------------------------------------------------------------------
# Name:        module1
# Author:      Neeraj
# Created:     05/11/2018
# Copyright:   (c) bunti 2018
# Licence:     <your licence>
# -------------------------------------------------------------------------------

#import HTML
import FUNC
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import datetime
from tkinter import *
from tkinter import messagebox
import json
import re
import random
import requests
import pymsgbox
#import Config

# hide main window of tkinter
root = Tk()
root.withdraw()

#current time and time in string format
x = datetime.datetime.now()
timen = str(datetime.datetime.now())

#working directory path
dir = os.path.dirname(__file__)

#Create blank json file for balance update
data = {}
data['Balance'] = []

#for reporting create HTML file
HTMLFILE = os.path.join('C:\\Python\\Result\\' +'Report_Result'+'_'+x.strftime('%x-%H-%M').replace('/','-')+'.html')
file = open(HTMLFILE, 'w')

#Create basic structure of HTML file
FUNC.basicHtmlStructure(file)

#update row in report table
FUNC.ReportN(file,'Process Start','Done',timen)

## Below code is to use HTML.py for reporting
# t = HTML.Table(header_row=['Step1', 'Observation1', 'Pass/Fail'])
# def Report(step,observation,result):
#     t.rows.append([step, observation, result])
#
# htmlcode = str(t)
# f.write(htmlcode)
# f.write('<p>')

#Clear existing content before writting
with open(os.path.join(dir, 'Resources/dataUserNew.txt'), 'a+') as outfiletxt:
    outfiletxt.seek(0)
    outfiletxt.truncate()
    outfiletxt.close()

#Read User detail from json file and start login in console for each user and preparing reporting
with open(os.path.join(dir, 'Resources/dataUser.json')) as json_file:
    datau = json.load(json_file)
    for p in datau['User']:
        loginName = p['name']
        login = p['Userid']
        password = p['Pass']
        #messagebox.showinfo("User", login)
        #messagebox.showinfo("password", password)

        # get the path of ChromeDriverServer
        chrome_driver_path = os.path.join(dir, 'Resources/chromedriver.exe') #dir + "\chromedriver.exe"

        # create a new Chrome session
        driver = webdriver.Chrome(chrome_driver_path)
        driver.implicitly_wait(3)
        driver.maximize_window()

        # Navigate to the application home page
        driver.get("https://console.zerodha.com/dashboard/")
        resp_code = requests.get("https://console.zerodha.com/dashboard/").status_code
        if int(resp_code) >=100 and int(resp_code) <200 :
            FUNC.ReportN(file, 'response from zerodha : request received and under process', 'DONE', timen)
        elif int(resp_code) >=200 and int(resp_code) <300 :
            FUNC.ReportN(file, 'response from zerodha : Success.. request processed', 'PASS', timen)
        elif int(resp_code) >=300 and int(resp_code) <400 :
            FUNC.ReportN(file, 'response from zerodha : Redirecting.. error in processing request', 'FAIL', timen)
        elif int(resp_code) >=400 and int(resp_code) <500 :
            FUNC.ReportN(file, 'response from zerodha : Client error', 'FAIL', timen)
        elif int(resp_code) >=500 and int(resp_code) <600 :
            FUNC.ReportN(file, 'response from zerodha : Server error', 'FAIL', timen)

        #driver.get("https://kite.zerodha.com/")
        time.sleep(5)
        #FUNC.ReportN(file, 'Kite login page displayed', 'PASS', timen)
        #messagebox.showinfo("ans2", "stop")

        # identifying User ID text box field by X Path and entering user id
        try:
            #e_user = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[1]/input")
            e_user = driver.find_element_by_css_selector(".uppercase > input")
                #"/html/body/div[1]/div/div/div[1]/div/div/div/form/div[2]/input")
            e_user.send_keys(login)
            FUNC.ReportN(file, 'Kite login page displayed', 'PASS', timen)
        except:# Exception('ElementNotVisibleException')
            FUNC.ReportN(file, 'Kite login page not displayed', 'FAIL', timen)
        time.sleep(.5)

        # identifying Password text box field by X Path and entering Password
        #e_password = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[2]/input")
        e_password = driver.find_element_by_css_selector(".su-input-group:nth-child(2) > input")
        e_password.send_keys(password)
        time.sleep(.5)

        # identifying Submit button by X Path and submit same
        e_submit = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[4]/button")
        e_submit.submit()
        time.sleep(4)

        FUNC.ReportN(file, 'Kite User login done for console', 'PASS', timen)

        #Read security question 1
        try:
            eQone = driver.find_element_by_xpath(
                "/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[2]/div/label")
            FUNC.ReportN(file, 'Kite Security Question page displayed', 'PASS', timen)
        except:# Exception('ElementNotVisibleException')
            FUNC.ReportN(file, 'Kite Security Question page not displayed', 'FAIL', timen)
        time.sleep(.5)

        #eQone = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[2]/div/label")
        textQone = eQone.text
        #get answer for Q1
        ans1 = FUNC.Answer(textQone)
        #textQone = 'vv' + '.png'
        #messagebox.showinfo("ans1", ans1)

        # Read security question 1
        eQtwo = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[3]/div/label")
        textQtwo = eQtwo.text
        # get answer for Q2
        ans2 = FUNC.Answer(textQtwo)
        #messagebox.showinfo("ans2", ans2)

        FUNC.ReportN(file, 'Security question page displayed and web elements found', 'PASS', str(datetime.datetime.now()))

        # Enter answer from above for user DN0280 otherwise use default value
        if login == "DN0280":
            e_Aone = driver.find_element_by_xpath(
                "/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[2]/div/input")
            e_Aone.send_keys(ans1)
            time.sleep(.5)

            e_Atwo = driver.find_element_by_xpath(
                "/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[3]/div/input")
            e_Atwo.send_keys(ans2)
            time.sleep(.5)
        else:
            e_Aone = driver.find_element_by_xpath(
                "/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[2]/div/input")
            e_Aone.send_keys("aaa")
            time.sleep(.5)

            e_Atwo = driver.find_element_by_xpath(
                "/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[3]/div/input")
            e_Atwo.send_keys("aaa")
            time.sleep(.5)

        e_submit = driver.find_element_by_xpath("/html/body/div[1]/div/div[2]/div[1]/div/div/div[2]/form/div[4]/button")
        e_submit.submit()

        FUNC.ReportN(file, 'Security answer submited', 'PASS', str(datetime.datetime.now()))
        time.sleep(6)

        try:
            eUserID = driver.find_element_by_xpath('//*[@id="userDetails"]')
            FUNC.ReportN(file, 'Console Dashboard page displayed', 'PASS', str(datetime.datetime.now()))
            UserId = eUserID.text
        except:# Exception('ElementNotVisibleException')
            FUNC.ReportN(file, 'Console Dashboard page not displayed', 'FAIL', str(datetime.datetime.now()))
        time.sleep(.5)

        #driver.save_screenshot(textQone)
        var = UserId+'_'+x.strftime('%x').replace('/','-')
        #messagebox.showinfo("Title", var)
        #driver.get_screenshot_as_file('C:\\Users\\bunti\Documents\\python\\ss\\'+var+'.png')
        #time.sleep(.5)

        eAcctBal = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div/div[1]/div/div[2]/div[1]/div[1]/h1')
        AcctBal = eAcctBal.text
        #messagebox.showinfo("Title", re.sub(r"\D","", AcctBal))
        AccBalNum = (float(re.sub(r"\D","", AcctBal))*1000)+random.randint(300, 800)
        # Highlight element
        FUNC.highlight(eAcctBal, var)
        time.sleep(.5)
        FUNC.ReportNwithSS(file, 'account balance of user: ' +UserId+ ' is ' +AcctBal, 'DONE', str(datetime.datetime.now()), var)

        time.sleep(.5)

        # close the browser window
        driver.quit()

        data['Balance'].append({
            'name': loginName,
            'Userid': login,
            'Balance': AcctBal
        })
        #enter balance and user detail in text file for cross verification
        with open(os.path.join(dir, 'Resources/dataUserNew.txt'), 'a+') as outfiletxt:
            outfiletxt.truncate()
            outfiletxt.write('Userid' + ':' + login + '\t')
            outfiletxt.write('Balance' + ':' + AcctBal + '\t')
            outfiletxt.write('BalanceNumber' + ':' + str(AccBalNum) + '\n')

#Update balance for each user in json file
with open(os.path.join(dir, 'Resources/dataUserNew.json'), 'w') as outfile:
    json.dump(data, outfile)

FUNC.ReportN(file, 'Details updated in json file', 'PASS', str(datetime.datetime.now()))


#messagebox.showinfo("Title", 'Process completed..')
#root.after(3000, lambda: root.destroy()) # Destroy the widget after 30 seconds
#root.mainloop()

#pymsgbox.alert('Process Completed.. click on ok','Message')

#Close HTML file
file.close()