#===============================================================================================================================================
from bs4 import BeautifulSoup
import os
import re
import csv
import datetime
import sys
from hkextools import nettools
#===============================================================================================================================================
folderPath = ''
#===============================================================================================================================================
#Obtain Date values
def obtDate ():
    today = str(datetime.date.today())
    thisDay = today[8:]
    thisMonth = today[5:7]
    thisYear = today[0:4]
    return thisDay, thisMonth, thisYear
#===============================================================================================================================================
#Function to obtain current path, subfolders, files
def dirWalk(callPath, mode):
    dirpath = []
    dirnames = []
    filenames  = []
    for dirpath, dirnames, filenames in os.walk(callPath):
        if (mode == 1):
            return dirpath
        if (mode == 2):
            return dirnames
        if (mode == 3):
            for f in filenames: 
                if (f.startswith('.')): filenames.remove(f)
            return filenames
#===============================================================================================================================================
def readCoList(pathToRead):

    coRecord = [[], [], [], [], [], [], [], [], [], [], [], [], [], []]
    recordCount = 0

    #Reads CSV file and return
    with open(pathToRead, "r", encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
        for row in csvreader:
            #Read all from CSV
            if (recordCount > 0):
                for i in range (0, 13):
                    coRecord[i].append(row[i])
            else:
                recordCount = 1

    return coRecord
#===============================================================================================================================================
#Function in scraping company info on HKEx, mode 0 is for English, mode 1 is for Chinese
def extractCo(code, mode, csvoutput):
    temp = ''
    printInfo = ''
    tagCount = 1
    opNo = 0
    fieldNo = [-1, -1, 10, -1, 16, 18, 12, 22, 28, 32, 30, 34, 36, 26]

    if mode == 1:
        coContent = nettools.tryConnect('http://www.hkex.com.hk/eng/invest/company/profile_page_e.asp?WidCoID=' + code + '&WidCoAbbName=&Month=&langcode=e')
        coSoup = BeautifulSoup(coContent, 'html.parser')
    else:
        coContent = nettools.tryConnect('http://www.hkex.com.hk/chi/invest/company/profile_page_c.asp?WidCoID=' + code + '&WidCoAbbName=&Month=&langcode=c')
        #chiContent = coContent.read().decode('Big5-HKSCS', 'ignore').encode('cp950')
        chiContent = coContent.read().decode('Big5-HKSCS', 'ignore')
        coSoup = BeautifulSoup(chiContent, 'html.parser')

    if (re.search('Company Website', str(coSoup), re.M | re.I) == None) and (mode == 1):
        fieldNo[:] = [ fieldI - 1 for fieldI in fieldNo[:]]
    if (re.search('Secondary Listing', str(coSoup), re.M | re.I) != None) and (mode == 1):
        fieldNo[7:13] = [ fieldI + 2 for fieldI in fieldNo[7:13]]
    
    for info in coSoup.find_all('font'):
        temp = info.get_text()
        if (mode == 1 and tagCount in fieldNo):
            charToRe = ['\r', '\n', '\t']
            for nlChar in charToRe:
                temp = temp.replace(nlChar, '')
            #temp = [temp.replace(nlChar, '') for nlChar in charToRe]
            csvoutput[fieldNo.index(tagCount)] = temp.strip()
        elif (mode == 2) and (tagCount == fieldNo[2]):
            csvoutput[3] = temp.strip()
        tagCount += 1
    return csvoutput, coSoup
#===============================================================================================================================================
#Validate the stock code
def validCode(code):
    codeIsValid = True
    if (code > 8999):
        codeIsValid = False
    elif (code < 8000):
        if (code > 4000):
            codeIsValid = False
    return codeIsValid
#===============================================================================================================================================
#Function to compare list with CSV and output to new CSV file
def outToCSV(mode, coRecord, archivePath):
    #Obtain html file, and parse with BeautifulSoup
    content = nettools.tryConnect('http://www.hkexnews.hk/listedco/listconews/advancedsearch/stocklist_active_main.htm')
    #content = urllib.request.urlopen('http://www.hkexnews.hk/listedco/listconews/advancedsearch/stocklist_active_main.htm')
    soup = BeautifulSoup(content, 'html.parser')

    #Open files for output, create if not exist
    csvfile = open(os.path.join(folderPath, 'coList.csv'), "w+", newline='', encoding='utf-8')
    #csvfile = open(folderPath + "\\coList.csv", "w+", encoding='cp950')
    csvwriter = csv.writer(csvfile)
    excludeLog  = open (os.path.join(folderPath, 'excludeLog.txt'), "w+", encoding='utf-8')

    #Initiates variables
    startOutput = False
    count = 0
    coSoup = ''
    #curList = ['HKD', 'USD', 'RMB']
    #curFlag = False
    needParse = False
    #csvoutput = ['','','','','','','','','','','','','','','']

    #Print csv head
    csvoutput = ['Code', 'Short', 'Name (Eng)', 'Name (Chi)', 'Principal Office', 'Place Incorporated', 'Principal Activities', 'Industry Classification', 'Currency', 'Issued Shares', 'Authorised Shares', 'Par Value', 'Board Lot', 'Listing Date', 'Remark']
    csvwriter.writerow(csvoutput)
    csvoutput = ['','','','','','','','','','','','','','','']

    #Loop for all info
    for info in soup.find_all('td'):
        #Start to read only after reaching stock code 00001
        if startOutput == True:
            #Obtain text inside tag <td>
            temp = info.get_text()
            #When reading should be company name
            if count == 1:
                csvoutput[1] = temp
                if (validCode(int(csvoutput[0]))):
                    if (mode == 1):
                        try:
                            if (coRecord[1][coRecord[0].index(csvoutput[0])] == csvoutput[1]):
                                iPos = int(coRecord[0].index(csvoutput[0]))
                                for i in range (0, 13):
                                    csvoutput[i] = coRecord[i][iPos]
                                csvoutput[14] = ""
                                csvwriter.writerow(csvoutput)
                                needParse = False
                            else:
                                needParse = True
                                csvoutput[14] = 'Changed'
                        except ValueError:
                            needParse = True
                            csvoutput[14] = 'New'
                    if (needParse == True or mode != 1):
                        print ('Now on: ' + '\t' + csvoutput[0])
                        csvoutput, coSoup = extractCo(csvoutput[0], 1, csvoutput)
                        csvoutput = [s.replace('\xa0',' ') for s in csvoutput]
                        #curFlag = False
                        needParse = False
                        if (re.search('industry', str(coSoup), re.M | re.I) != None) and (re.search('HSIC', csvoutput[7], re.M | re.I) != None) and (re.search('Preference Share', str(coSoup), re.M | re.I) == None):
                        #for cur in curList:
                            #if (csvoutput[8][0:3] == cur): curFlag = True
                        #if (curFlag == True):
                            csvoutput, coSoup = extractCo(csvoutput[0], 2, csvoutput)
                            csvwriter.writerow(csvoutput)
                        else:
                            print (csvoutput[0] +' is not valid')
                            excludeLog.write(csvoutput[0] + '\n')
                csvoutput = ['','','','','','','','','','','','','','','']
                count = 0
            #When reading should be stock code
            else:
                csvoutput[0] = temp
                count = 1
        #First 2 readings not relevant
        else:
            count += 1
            if count == 3:
                startOutput = True
                count = 0

    #Output results
    csvfile.close()
    excludeLog.close()
#===============================================================================================================================================