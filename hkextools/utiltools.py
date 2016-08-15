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
    coRecord = []
    with open(pathToRead, 'r', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
        coRecord = list(csvreader)
    return coRecord

#===============================================================================================================================================
#Validate the stock code
def validCode(code):
    if (code > 8999): return False
    if (code < 6030) and (code > 4000): return False
    return True

#===============================================================================================================================================
def extractCo(code, mode, csvoutput):
    modeHead = 2
    modeTail = 14
    searchTag = ['coCode', 'ShortName', 'Company/Securities Name:', '公司/證券名稱:', 'Principal Office:', 'Place Incorporated:', 'Principal Activities:', 'Industry Classification:', 'Trading Currency:', 'Issued Shares:', 'Authorised Shares:', 'Par Value:', 'Board Lot:', 'Listing Date :']

    if mode == 1:
        coContent = nettools.tryConnect('http://www.hkex.com.hk/eng/invest/company/profile_page_e.asp?WidCoID=' + code + '&WidCoAbbName=&Month=&langcode=e')
        coSoup = BeautifulSoup(coContent, 'html.parser')
    else:
        coContent = nettools.tryConnect('http://www.hkex.com.hk/chi/invest/company/profile_page_c.asp?WidCoID=' + code + '&WidCoAbbName=&Month=&langcode=c')
        chiContent = coContent.read()#.decode('Big5-HKSCS')
        coSoup = BeautifulSoup(chiContent, 'html.parser')
        modeTail = 4

    for i in range (modeHead, modeTail):
        tagFound = False
        for info in coSoup.find_all('font'):
            if (tagFound == True):
                csvoutput[i] = info.get_text().strip(' \r\n\t').replace('\xa0',' ').strip()
                if (i==9): csvoutput[i] = csvoutput[i].split(' ')[0]
                break

            if (re.search(searchTag[i], info.get_text().strip(), re.M | re.I) != None) and (csvoutput[i]==''): tagFound = True

    if (re.search(csvoutput[2], 'pref', re.M | re.I) != None): csvoutput[4] == ''

    return csvoutput

#===============================================================================================================================================
#Function to compare list with CSV and output to new CSV file
def outToCSV(mode, archivePath):
    #Obtain html file, and parse with BeautifulSoup
    content = nettools.tryConnect('http://www.hkexnews.hk/listedco/listconews/advancedsearch/stocklist_active_main.htm')
    coSoup = BeautifulSoup(content, 'html.parser')

    #Initiates variables
    startOutput = False
    needParse = False
    #Reading records from CSV
    if (mode == 1): coRecord = readCoList(archivePath)

    #Read csv for writing
    csvfile = open(os.path.join(folderPath, 'Output', 'coList.csv'), "w+", newline='', encoding='utf-8')
    csvwriter = csv.writer(csvfile)
    csvoutput = ['Code', 'Short', 'Name (Eng)', 'Name (Chi)', 'Principal Office', 'Place Incorporated', 'Principal Activities', 'Industry Classification', 'Currency', 'Issued Shares', 'Authorised Shares', 'Par Value', 'Board Lot', 'Listing Date', 'Remark']
    csvwriter.writerow(csvoutput)
    csvoutput = ['','','','','','','','','','','','','','','']
    excludeLog = open (os.path.join(folderPath, 'Output', 'excludeLog.txt'), "w+", encoding='utf-8')
    bankLog = open (os.path.join(folderPath, 'Output', 'Banks'), "w+", encoding='utf-8')

    #Loop for all info
    for info in coSoup.find_all('td'):
        #Exclude info not related to stocks
        if (str(info.get_text())=='00001'): startOutput=True
        if (startOutput):
            #Checking if it's stock code info depending on contains number and empty list or not
            if (info.get_text()[0].isdigit() and csvoutput[0]==''):
                csvoutput[0] = str(info.get_text())
            else:
                #Checks if code valid
                if (validCode(int(csvoutput[0]))):
                    #Start checking from archive if current company exists in our record
                    if (mode==1):
                        print (str(info.get_text()))
                        try:
                            thisCo = [x for x in coRecord if csvoutput[0] in x]
                            if (thisCo[0][1] == str(info.get_text())):
                                print ('Now on: ' + '\t' + csvoutput[0])
                                if (re.search('Banks', thisCo[0][7], re.M | re.I) != None) and (csvoutput[0] != 222): bankLog.write(csvoutput[0] + '\n')
                                csvwriter.writerow(thisCo[0])
                            else:
                                #Treat as a new company if company new has changed
                                needParse = True
                        except:
                            #Take a new parse search if the stock code is not found in our record
                            needParse = True
                    #Handling the parse request
                    if (needParse or mode != 1):
                        print ('Now on: ' + '\t' + csvoutput[0] + '\t[New]')
                        csvoutput[1] = str(info.get_text())
                        csvoutput[14] = 'New'
                        csvoutput = extractCo(csvoutput[0], 2, extractCo(csvoutput[0], 1, csvoutput))
                        if (re.search('Banks', csvoutput[7], re.M | re.I) != None) and (csvoutput[0] != 222): bankLog.write(csvoutput[0] + '\n')
                        #Normally an equity will have field [Principal Office] and [Place Incorporated]
                        if (csvoutput[4] != '' or csvoutput[5] != ''):
                            csvwriter.writerow(csvoutput)
                        else:
                            print (csvoutput[0] + ' is not equity')
                            excludeLog.write(csvoutput[0] + '\n')
                        needParse = False
                csvoutput = ['','','','','','','','','','','','','','','']

    csvfile.close()
    excludeLog.close()
    bankLog.close()

#===============================================================================================================================================