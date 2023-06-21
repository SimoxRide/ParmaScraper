
from asyncio.windows_events import NULL
from glob import glob
import re
import selenium
from selenium import webdriver
import time as t
import json
import os
from sqlalchemy import null

import xlsxwriter

import chromedriver_autoinstaller



columitem=NULL

cardsitem=NULL
driverpath="Driver\\"+str(chromedriver_autoinstaller.get_chrome_version()).split(".")[0]+"\\chromedriver.exe"
chromedriver_autoinstaller.install(False,"Driver\\")

options = webdriver.ChromeOptions()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = webdriver.Chrome(executable_path=driverpath, options=options)

driver.get("https://www.prosciuttodiparma.com/tutti-i-produttori-del-consorzio-del-prosciutto-di-parma/")
driver.maximize_window()



def SaveToExcell():
    workbook = xlsxwriter.Workbook('demo.xlsx')
    worksheet = workbook.add_worksheet()

    # Widen the first column to make the text clearer.
    worksheet.set_column('A:C', 20)

    # Add a bold format to use to highlight cells.
    bold = workbook.add_format({'bold': True})
     
    
    # Write some simple text.
    worksheet.write('A1', 'Hello')

    # Text with formatting.
    worksheet.write('A2', 'World', bold)

    # Write some numbers, with row/column notation.
    worksheet.write(2, 0, 123)
    worksheet.write(3, 0, 123.456)

    # Insert an image.
    #worksheet.insert_image('B5', 'logo.png')

    workbook.close()






#acceptcookie
try:
    driver.find_element_by_xpath("//*[@id=\"metisCookieTopBar\"]/div[3]/a").click()
except:
    print("[LOG]NoCookies")


def ExtractCollumItem():
    global columitem
    try:
        columitem=driver.find_element_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div")
        
        
    except:
        print("[LOG]CardCollum not found")
    ExtractAllCard()

def ExtractAllCard():
    global columitem
    global cardsitem
    workbook = xlsxwriter.Workbook('scrapeddata.xlsx')
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:C', 20)
    bold = workbook.add_format({'bold': True})
    worksheet.write('A1', 'Name',bold)
    worksheet.write('B1', 'Emails', bold)
    worksheet.write('C1', 'Numbers', bold)
    
   
    worksheet.set_column(0,2,25)
    worksheet.set_column(1,1,120)
    worksheet.set_column(2,2,95)
    




    try:
        cardsitem =columitem.find_elements_by_xpath("/html/body/div[1]/div[2]/div[2]/div/div/div[2]/div/div/div")
    except:
        print("[LOG]Cards not found")
    kz=1
    for item in cardsitem:
        
        Nome=item.find_element_by_tag_name("h4").text
        print(Nome)
        worksheet.write(kz, 0, Nome)
        emails=item.find_elements_by_tag_name("a")
        a=""
        remails={}
        
        k=0
        for item2 in emails:
            a=item2.text
            if(not str(a).startswith("www")):
                remails[k]=a
                k+=1
        
        print(remails)
        txt=""
        for ic in remails:
            txt+=remails[ic]+" , "
        worksheet.write(kz, 1, txt)
        data={}
        c=0

        try:
            nm1=item.find_elements_by_tag_name("p")
            for tagi in nm1:
                try:
                    tagi.find_elements_by_css_selector("i.fas.fa-phone.fontGold")
                    if(tagi.text!=""):
                        data[c]=tagi.text
                        c+=1
                    
                    
                except:
                    print("")
        except:
            print("can't find i tag")

        
        
        numbers=""
        for item3 in data:
            lista=str(data[item3]).split("\n")
            for item4 in lista:
                if item4.startswith("0") and item4.__contains__(r"/"):
                    print(item4)
                    numbers+=item4+" , "
        worksheet.write(kz, 2, numbers)
        kz+=1
    
    workbook.close()

            
        
                




        
        
        


    
    



t.sleep(2)

ExtractCollumItem()

    







t.sleep(1000)









