Selenium Web Scrape Wireless postpaid


# Instruction for Wei:

1. Open Move Mouse and click start

2. Navigate to the folder and click runScrape file

3. Input the year and month 

4. It will print Successfully scraped XXXX if everything is downloaded and modified successfully.

5. A map file is always created that records all the info about what has been downloaded.

#----------------------------------------------------------------------------------------------

Environment Setup

pip install selenium

pip install openpyxl

It takes a long time for the script to finish running, so while the script is running, it is necessary to keep mouse moving in order to enable the computer to never sleep or lock the screen. 
To do that, an app called move mouse is recommended. Move mouse can ensure the screen is never locked.


#----------------------------------------------------------------------------------------------


Notice

ROOT(a global variable inside runScrape.py) needs to be changed if one wants to run script on their own local computer since ROOT records the download directory and webdriver will download files to that download directory
 
runScrape.py stands alone which is different from wireline webscrape

The data for one month will be downloaded inside the folder of that month. If folder name for the month already exists for some reason, one needs to rename existing folder before downloading

A map will be created for each month which records the information of what has been downloaded

For the Date filter, script will input the date to the Date textbox. And the script will loop through the filters of Segment, Brand and Province. 

In the script, elements are located by css selector or xpath. Most of the elements work fine if they are located by css selector. However, if one wants to click the body of the website (for example click_x('/html/body/div[5]',web))
this element has to be located by xpath. The potential problem is that there might be a loadingGlassPane on top of the website body when you try to click the website body and the loadingGlassPane will interrupt with clicking. Some times 
it goes away after waiting for a few seconds and sometimes it doesn't, so removing the loadingGlassPane manually through script is required if it does not go away after waiting for some time.  

If the website changes the filter format of Segment, Brand and Province, script needs to be modified since naming part is hard coded.

