# import webdriver
#This version is without the flask webpage
from selenium import webdriver
from selenium.webdriver.chrome import options
from selenium.webdriver.common.by import By
import re
import pygsheets
from selenium.common.exceptions import NoSuchElementException
import time
#headless mode

options = webdriver.ChromeOptions()
options.add_argument('headless')



# get url from user
site ="site:"+ input("Enter the site you want to search:")
keywords = input("Enter Keywords seperated by spaces:")
dateRange= " "+str(input("Enter a date range e.g 2005..2012 or 2015..2016:"))
reason= input("And this is for (emails,Leads) or leave blank:").lower()
#if reason=="leads":
    #emailCount=int(input("How many emails do you need:"))

keyLists = keywords.split()
nkeyLists=[]
for x in keyLists:
        x = '"'+ x + '"'
        nkeyLists.append(x)
searchString=""
for r in nkeyLists:
        searchString+=" "+r

#search for searchTerm
if reason == "" or reason == " " or reason == "research":
        searchTerm="http://www.google.com/search?q="+searchString+site+dateRange
elif reason == "emails" or reason == "leads":
        searchTerm="http://www.google.com/search?q="+'"@gmail.com" OR "@yahoo.com" OR "@icloud.com" OR "@hotmail.com"'+searchString+site+dateRange
else:
        print("You may have entered a wrong value for a field, Try again.")

#carry out search and copy the page to a new file

driver = webdriver.Chrome(executable_path = "C:\\Users\\Your_Username\\projects\\pj\\chromedriver.exe")#,options=options
driver.get(searchTerm)
print(searchTerm)

while True: 
    #check if google recaptcha is visible the continue with code
    try:
        result= driver.find_element(By.CLASS_NAME,"GyAeWb")
    except NoSuchElementException:
        recaptcha= driver.find_element(By.ID, "captcha-form")
        if recaptcha.is_displayed() == True:
            print("Google recaptcha is affecting program, try using a vpn or waiting for some time.\n The current emails have been added to the spreadsheet")
            with open("finalemails.txt", "w") as clearfinals:
                clearfinals.write("")
            time.sleep(30)
            continue
#continue to getting emails
    with open("pagetext.txt","w",encoding="utf8") as pagetext:
        pagetext.write(result.text)

#filter emails, clear duplicates and other irrelevant texts
    newfile= open("pagetext.txt", "r", encoding="latin-1")
    currentEmailList=[]

    for lines in newfile:
        text = lines
        emails = re.findall(r"[\w\.-]+@[\w\.-]+\.\w+", text)
        currentEmailList.append(emails)

    for emailList in currentEmailList[:]:
        if emailList == []:
            currentEmailList.remove(emailList)
    newEmailList=[]
    for cleanEmailList in currentEmailList:
        for email in cleanEmailList:
            if email in newEmailList:
                continue
            elif email not in newEmailList:
                newEmailList.append(email)
    newfile.close()
    with open("finalemails.txt",'a',encoding='utf8') as finalEmails:
        for finalemail in newEmailList:
            if finalemail[0]=="-":
                finalemail=finalemail.replace("-","")
            finalEmails.write(finalemail+"\n")
    
        

    try:
        nextButton= driver.find_element(By.ID,"pnnext")
        nextButton.click()

    except NoSuchElementException: 
            with open("pagetext.txt", "w", encoding="utf8") as clearFile:
                clearFile.write("")
            driver.delete_all_cookies()
            driver.quit()
            break

authorizeJson= pygsheets.authorize(service_file="C:\\Users\\Your_Username\\projects\\pj\\service_account.json")
sh=authorizeJson.open("spreadsheet")
with open("finalemails.txt","r")as lencheck:
    sh.add_worksheet("emails",rows=len(lencheck.readlines())+1,cols=2)
wks=sh.worksheet_by_title("emails")
wks.update_value("A1","Emails")
wks.cell("A1").set_text_format("bold",True)
wks.update_value("B1","profession")
wks.cell("B1").set_text_format("bold",True)
anum=2
with open("finalemails.txt", "r") as getemails:
    for finishedemail in getemails:
            wks.update_value("A{}".format(anum), finishedemail)
            wks.update_value("B{}".format(anum), keywords)
            anum+=1


        







