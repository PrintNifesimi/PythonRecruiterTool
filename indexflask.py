from flask import Flask, render_template, request
from selenium import webdriver
from selenium.webdriver.chrome import options
from selenium.webdriver.common.by import By
import re
import pygsheets
import random
import string
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException
import time
import os


#app start
app = Flask(__name__)
save_path = 'C:\\Users\\Your_Username\\projects'
file_name = "secretkeyfile.txt"
completeName = os.path.join(save_path, file_name)
seckeyfile=open(completeName,"w")
seckeyfile.close()
@app.route('/', methods=['GET'])
def home():
    return render_template('home.html')

'''@app.route('/emails/',methods=['GET'])
def emailsscraped():
    return render_template('emails.html')'''

@app.route('/sent/', methods=["POST"])
def sent():
    site="site:"+request.form['sitestr']
    keywords=request.form['keywords']
    daterange=" "+str(request.form['date'])
    reason=request.form['reason'].lower()
    keyLists = keywords.split()
    nkeyLists=[]
    for x in keyLists:
        x = '"'+ x + '"'
        nkeyLists.append(x)
    searchString=""
    for r in nkeyLists:
        searchString+=" "+r
#to avoid error of same file name and excel sheet, random id gen
    def id_generator(size=4, chars=string.ascii_uppercase + string.digits):
        return ''.join(random.choice(chars) for _ in range(size))
#search for searchTerm
    if reason == "" or reason == " " or reason == "research":
        searchTerm="http://www.google.com/search?q="+searchString+site+daterange
    elif reason == "emails" or reason == "leads":
        searchTerm="http://www.google.com/search?q="+'"@gmail.com" OR "@yahoo.com" OR "@icloud.com" OR "@hotmail.com"'+searchString+site+daterange
    else:
        print("You may have entered a wrong value for a field, Try again.")

#carry out search and copy the page to a new file

    driver = webdriver.Chrome(executable_path = "C:\\Users\\Your_Username\\projects\\pj\\chromedriver.exe")#,options=options
    driver.get(searchTerm)
    while True: 
    #check if google recaptcha is visible the continue with code
        try:
            result= driver.find_element(By.CLASS_NAME,"GyAeWb")
        except NoSuchElementException:
            recaptcha= driver.find_element(By.ID, "captcha-form")
            if recaptcha.is_displayed() == True:
                print("Google recaptcha is affecting program, try using a vpn or waiting for some time.\n The current emails have been added to the spreadsheet")
                time.sleep(15)
                continue
        except ElementClickInterceptedException:
            time.sleep(60)
            exit(0)
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
        
        with open(completeName, 'r')as keyfile:
            if os.stat(completeName).st_size==0:
                with open(completeName,"a") as keyfilewrite:
                    seckey=id_generator()
                    keyfilewrite.write(seckey)
        with open(completeName, "r")as readkeyfile:
            keysKey=readkeyfile.readline()
        with open("finalemails({}).txt".format(keysKey),'a',encoding='utf8') as finalEmails:
            for finalemail in newEmailList:
                for letter in finalemail:
                    if letter=="-":
                        hyphenindex=finalemail.index(letter)+1
                        finalemail=finalemail.replace(finalemail[0:hyphenindex],"")
                finalEmails.write(finalemail+"\n")
    
        

        try:
            nextButton= driver.find_element(By.ID,"pnnext")
            if nextButton.is_displayed() == True:
                nextButton.click()
            elif nextButton.is_displayed() == False:
                driver.implicitly_wait(10)
                nextButton.click()

        except NoSuchElementException:
            print("All emails have been scraped") 
            with open("pagetext.txt", "w", encoding="utf8") as clearFile:
                clearFile.write("")
            driver.delete_all_cookies()
            driver.quit()
            break
        htlist=[]
        with open("finalemails({}).txt".format(keysKey), "r") as htmlshow:
            for htemails in htmlshow:
                htlist.append(htemails)
        newhtlist=htlist
    authorizeJson= pygsheets.authorize(service_file="C:\\Users\\Your_Username\\projects\\pj\\service_account.json")
    sh=authorizeJson.open("spreadsheet")
    with open("finalemails({}).txt".format(keysKey),"r")as lencheck:
        excelRows=len(lencheck.readlines())+1
        sh.add_worksheet("emails({})".format(keysKey),rows=excelRows,cols=3)
    wks=sh.worksheet_by_title("emails({})".format(keysKey))
    wks.update_value("A1","Emails")
    wks.cell("A1").set_text_format("bold",True)
    wks.update_value("B1","profession")
    wks.cell("B1").set_text_format("bold",True)
    wks.update_value("C1","site")
    wks.cell("C1").set_text_format("bold",True)
    anum=2
    with open("finalemails({}).txt".format(keysKey), "r") as getemails:
        print("Emails Are Being added to the spreadsheet")
        for finishedemail in getemails:
            wks.update_value("A{}".format(anum), finishedemail)
            wks.update_value("B{}".format(anum), keywords)
            wks.update_value("C{}".format(anum), site)
            anum+=1
    with open(completeName, "w") as clearkeyfile:
                clearkeyfile.write("")
    return render_template('home.html',
            newhtlist=newhtlist)
            
@app.route('/directions/')
def directions():
    return render_template('directions.html')
if __name__ == '__main__':
    app.run(debug=True)