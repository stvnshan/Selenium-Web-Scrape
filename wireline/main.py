#---------------------------------------------------------------------------------------------
#user interface
from sys import excepthook
from selenium.webdriver import Edge
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By 
from selenium.common.exceptions import NoSuchElementException
import time
import os
import re
import openpyxl
from openpyxl import load_workbook 
import subprocess


SLEEPTIME = 10
ROOT = "C:\\Users\\Steven.Shan\\mypython\\tableau\\webscrape\\"
#ROOT = "C:\\Users\\Wei.Li1\\Documents\\DailyCableTSU\\"
options = webdriver.EdgeOptions()
options.use_chromium = True

#make the screen max to make elements apprear on te screen 
options.add_argument("--start-maximized")
  
print("Enter the year:")
y = int(input())
print("Enter the month(or 0 to download the entire year):")
m = int(input())


#get current year and month
web = Edge(options=options)
web.get("https://bianalytics.rci.rogers.com/t/BI/views/DailyCableTSU/DailyCableTSU?:embed=yes&:display_count=no&:showVizHome=no&:origin=viz_share_link#1") 
time.sleep(15)
el = WebDriverWait(web,30).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#tableau_base_widget_LegacyCategoricalQuickFilter_5 > div > div.CFContent > span')))
el.click()
time.sleep(15)
mydate = web.find_element(By.CSS_SELECTOR, '#FI_federated\.07uceah05f3pzg12vx7ye10ec8da\,none\:Calculation_783907863373090816\:ok1863049714700099538_14391167526842317961_0 > div.facetOverflow > a').get_attribute("title")
curMonth = int(re.match('^[^/]*', mydate).group(0))
curyear = mydate[-4:]
print("current month is: " +str(curMonth)+" current year is: "+curyear)
web.quit()



#if user wants to download current year
if m == 0 and int(curyear) == y:
    startMonth = curMonth-1
    print("downloading from month "+str(startMonth)+" to 1 "+curyear)
    DOWNLOAD_DIRECTORY = ROOT+curyear
    os.makedirs(DOWNLOAD_DIRECTORY)
    DOWNLOAD_DIRECTORY = DOWNLOAD_DIRECTORY+"\\"
    i=1
    while i<=startMonth:

        subprocess.Popen(["python.exe", "runScrape.py", str(curMonth-i), DOWNLOAD_DIRECTORY], creationflags=subprocess.CREATE_NEW_CONSOLE)
        i = i+1
        time.sleep(SLEEPTIME)
        if i<=startMonth:
            subprocess.Popen(["python.exe", "runScrape.py", str(curMonth-i), DOWNLOAD_DIRECTORY], creationflags=subprocess.CREATE_NEW_CONSOLE)
            i = i+1
            time.sleep(SLEEPTIME)
        if i<=startMonth:
            #"py runScrape.py 0 "C:\\Users\\Wei.Li1\\Documents\\DailyCableTSU\\""
            os.system("start /wait cmd /c "+"runScrape.py "+str(curMonth-i)+" "+DOWNLOAD_DIRECTORY)
            i = i+1
            time.sleep(SLEEPTIME)

    print("successfully_scrapped from month " +str(startMonth)+" to 1 "+curyear)
#else if user wants to download other year
elif m == 0:
    print("downloading from month 12 to 1 " + str(y))
    DOWNLOAD_DIRECTORY = ROOT+str(y)
    os.makedirs(DOWNLOAD_DIRECTORY)
    DOWNLOAD_DIRECTORY = DOWNLOAD_DIRECTORY+"\\"
    i=curMonth + (int(curyear)-1-y)*12
    while i< curMonth + (int(curyear)-1-y)*12 + 12:
        subprocess.Popen(["python.exe", "runScrape.py", str(i), DOWNLOAD_DIRECTORY], creationflags=subprocess.CREATE_NEW_CONSOLE)
        i = i+1
        time.sleep(50)
        if i< curMonth + (int(curyear)-1-y)*12 + 12:
            subprocess.Popen(["python.exe", "runScrape.py", str(i), DOWNLOAD_DIRECTORY], creationflags=subprocess.CREATE_NEW_CONSOLE)
            i = i+1
            time.sleep(50)
        if i< curMonth + (int(curyear)-1-y)*12 + 12:
            os.system("runScrape.py "+str(i)+" "+DOWNLOAD_DIRECTORY)
            i = i+1
            time.sleep(SLEEPTIME)

    print("successfully_scrapped from month 12" +" to 1 "+str(y))

else:
    DOWNLOAD_DIRECTORY = ROOT

    if int(curyear) == y:
        num = curMonth - m
    else:
        num = curMonth + 12*(int(curyear) - 1 - y) + (12 - m)
    os.system("runScrape.py "+str(num)+" "+DOWNLOAD_DIRECTORY)

    


print("Have A Great Day!")