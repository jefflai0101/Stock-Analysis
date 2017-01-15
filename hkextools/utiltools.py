#===============================================================================================================================================
from bs4 import BeautifulSoup
import os
import re
import sys
import csv
import json
import datetime
from hkextools import nettools
#===============================================================================================================================================
folderPath = ''
#===============================================================================================================================================
def readSettings(pathToRead):
    with open(pathToRead, 'r') as settingsFile:
        jSet = json.loads(settingsFile.read())
        return jSet
#===============================================================================================================================================
def readCSV(pathToRead):
    #Reading all industries
    with open(pathToRead, 'r', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile)
        readInfo = list(csvreader)
        return readInfo
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
    with open(pathToRead, "r", encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
        return list(csvreader)[1:]
#===============================================================================================================================================
def extractCo(items, code, mode):
    linkArg = ['eng','e','e'] if mode == 1 else ['chi','c','c']
    coContent = nettools.tryConnect('http://www.hkex.com.hk/'+linkArg[0]+'/invest/company/profile_page_'+linkArg[1]+'.asp?WidCoID=' + code + '&WidCoAbbName=&Month=&langcode='+linkArg[2])
    coSoup = BeautifulSoup(coContent.read().decode('Big5-HKSCS', 'ignore'), 'html.parser')

    coLabels = [tdSoup.get_text().replace('\r', '').replace('\n', '').replace('\t', '').replace('\xa0','').lstrip().rstrip() for tdSoup in coSoup.find_all('td', {'colspan' : '0', 'valign' : 'top', 'align' : 'left', 'height' : '18', 'width' : '150'})]
    coDetails = [tdSoup.get_text().replace('\r', '').replace('\n', '').replace('\t', '').replace('\xa0','').lstrip().rstrip() for tdSoup in coSoup.find_all('td', {'colspan' : '3', 'align' : 'left', 'height' : '18', 'width' : '300'})]

    coInfo = dict(zip(coLabels,coDetails))

    for iKey, item in enumerate(items):
        if (item in coInfo):  items[iKey] = coInfo[item]

    return items
#===============================================================================================================================================
#Validate the stock code
def validCode(code):
    return (code <= 4000) or ((code >= 8000) and (code < 9000))
#===============================================================================================================================================
#Function to compare list with CSV and output to new CSV file
def outToCSV(mode):

    onlineList = []
    coRecords = []
    temp = ''
    #Obtain all full list of listed companies
    listSoup = BeautifulSoup(nettools.tryConnect('http://www.hkexnews.hk/listedco/listconews/advancedsearch/stocklist_active_main.htm'), 'html.parser')
    for codeInfo in listSoup.find_all('td'):
        if (codeInfo['align'] == 'Center'):
            temp = codeInfo.get_text() if (validCode(int(codeInfo.get_text()))) else ''
        elif (codeInfo['align'] == 'left' and temp != ''):
            onlineList.append([temp, codeInfo.get_text()])

    onlineList = dict(zip([coRecord[0] for coRecord in onlineList], [coRecord[1] for coRecord in onlineList]))
    targetList = sorted(set(onlineList))

    if mode:
        #Check if folder 'Archive' exists. Shell cmd to archive the current version of CSV file into Archive folder
        if (os.path.isdir(os.path.join(folderPath, 'Archive')) == False): os.system ('mkdir Archive')
        #archivePath = str(os.path.join(folderPath, 'Archive', str(datetime.date.today())) + '.csv')
        #coListPath = str(os.path.join(folderPath, 'coList.csv'))
        archivePath = str(os.path.join('Archive', str(datetime.date.today())) + '.csv')
        coListPath = str('coList.csv')
        #shellCommand = 'move ' if os.name == 'nt' else 'mv '
        shellCommand = 'mv '
        os.system(shellCommand + coListPath + ' ' + archivePath)
        coRecords = readCoList(archivePath)

        currentList = dict(zip([coRecord[0] for coRecord in coRecords], [coRecord[1] for coRecord in coRecords]))                       #currentList    [1,2,3  ]
                                                                                                                                        #onlineList     [  2,3,4]
        targetList = sorted(set(onlineList) ^ set(currentList))                                                                         #targetList     [1,    4]
        changeList = sorted(set([key for key in (set(currentList) & set(onlineList)) if currentList[key] != onlineList[key]]))          #changeList     [  2    ]
        removeList = sorted(((set(onlineList) ^ set(currentList)) & set(currentList)) | set(changeList))                                #removeList     [1,2    ]
        targetList = sorted((set(targetList) ^ set(removeList)) | set(changeList))                                                      #targetList     [  2   4]

        coRecords = [coRecord for coRecord in coRecords if not (coRecord[0] in removeList)]
        for coRecord in coRecords: coRecord[14] = ''

    try:
        for key in targetList:
            print('Working on : ' + key)
            items = extractCo(extractCo([key, onlineList[key], 'Company/Securities Name:', '公司/證券名稱:', 'Principal Office:', 'Place Incorporated:', 'Principal Activities:', 'Industry Classification:', 'Trading Currency:', 'Issued Shares:(Click here for important notes)', 'Authorised Shares:', 'Par Value:',  'Board Lot:', 'Listing Date :', '', '',''], key, 1), key, 2)
            if (sum([item[-1]==':' for item in items if item!='']) <= 0):
                items[9] = items[9].split('(')[0]
                if mode: items[16] == 'New'
                coRecords.append(items)
        writeCoList(coRecords)
    except:
        writeCoList(coRecords)

#   Add paused progress status for resume
#   Check if the exclude list exists, if so read in all and don't check info, otherwise write
#   Also make sure update exclude list if new is found

#    excludeLog = open(os.path.join(folderPath, 'excludeLog.txt'), "w+", encoding='utf-8')
#    excludeLog.close()
#===============================================================================================================================================
def writeCoList(coRecords):

    csvwriter = csv.writer(open(os.path.join(folderPath, 'coList.csv'), "w+", newline='', encoding='utf-8'))
    csvwriter.writerow(['Code', 'Short', 'Name (Eng)', 'Name (Chi)', 'Principal Office', 'Place Incorporated', 'Principal Activities', 'Industry Classification', 'Currency', 'Issued Shares', 'Authorised Shares', 'Par Value', 'Board Lot', 'Listing Date', 'Year End', 'Data Type', 'Remark'])
    for coRecord in sorted(coRecords): csvwriter.writerow(coRecord)

#===============================================================================================================================================
def main():
    outToCSV(1) if (os.path.isfile(os.path.join(folderPath,'coList.csv')) == True) else outToCSV(0)

#===============================================================================================================================================