#===============================================================================================================================================
import os
import re
import csv
import time
import socket
import urllib.request
from bs4 import BeautifulSoup
#===============================================================================================================================================
folderPath = ''
thisDay = ''
thisMonth = ''
thisYear = ''
keywords = []
keywordCount = 0
#===============================================================================================================================================
def tryConnect(link):
    networkOK = False
    while networkOK == False:
        try:
            content = urllib.request.urlopen(link, timeout=10)
            networkOK = True
            return content
        except urllib.error.URLError:
            time.sleep(1)
        except socket.timeout:
            time.sleep(1)

#===============================================================================================================================================
def firstCheck(coCode, tableSoup):
    #Checking for listed banks and separates from other companies
    if os.path.isfile(os.path.join(folderPath, 'Companies', coCode, 'index.csv')) == False: return False
    try:
        aInfo = tableSoup.find('tr', {'class' : None}).find('a', {'class' : 'news'})
        with open(os.path.join(folderPath, 'Companies', coCode, 'index.csv'), 'r', encoding='utf-8') as indexFile:
            csvreader = csv.reader(indexFile, delimiter=',', quotechar='\"')
            return list(csvreader)[1][2] == 'http://www.hkexnews.hk'+aInfo['href']
    except:
        return False

#===============================================================================================================================================
#Search for announcements
def aSearch(coCode):

    #Create folder for the company if it doesn't exist
    if (os.path.isdir(os.path.join(folderPath, 'Companies')) == False): os.system ('mkdir Companies')
    if (os.path.isdir(os.path.join(folderPath, 'Companies', coCode)) == False): os.system ('mkdir ' + os.path.join('Companies', coCode))

    #Obtain the viewState value for the current page
    soup = BeautifulSoup(urllib.request.urlopen('http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx'), 'html.parser')
    #Obtains the viewstate key of the current page
    viewState = soup.find('input', {"id": "__VIEWSTATE"}).get('value')

    #Setting for the post
    values = {
    '__VIEWSTATE' : viewState,
    '__VIEWSTATEENCRYPTED' : '',
    'ctl00$ddlTierTwo' : '23,1,3',
    'ctl00$ddlTierTwoGroup' : '10,2',
    'ctl00$hfAlert' : '',
    'ctl00$hfStatus' : 'AEM',
    'ctl00$rdo_SelectDateOfRelease' : 'rbManualRange',
    'ctl00$rdo_SelectDocType' : 'rbAll',
    'ctl00$rdo_SelectSortBy' : 'rbDateTime',
    'ctl00$sel_DateOfReleaseFrom_d' : '01',
    'ctl00$sel_DateOfReleaseFrom_m' : thisMonth,
    'ctl00$sel_DateOfReleaseFrom_y' : str(int(thisYear) - 3),
    'ctl00$sel_DateOfReleaseTo_d' : thisDay,
    'ctl00$sel_DateOfReleaseTo_m' : thisMonth,
    'ctl00$sel_DateOfReleaseTo_y' : thisYear,
    'ctl00$sel_DocTypePrior2006' : '-1',
    'ctl00$sel_defaultDateRange' : 'SevenDays',
    'ctl00$sel_tier_1' : '-2',
    'ctl00$sel_tier_2' : '-2',
    'ctl00$sel_tier_2_group' : '-2',
    'ctl00$txtKeyWord' : '',
    'ctl00$txt_stock_code' : coCode,
    'ctl00$txt_stock_name' : '',
    'ctl00$txt_today' : (thisYear + thisMonth + thisDay)
    }

    print ('Now Searching: ' + coCode)
    cnW(values, coCode, 1)

#===============================================================================================================================================
# Make HTML request
def cnW(values, coCode, mode):
    headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.97 Safari/537.36'}
    link = 'http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx'
    data = urllib.parse.urlencode(values)
    req = urllib.request.Request(link, data.encode('Big5'), headers = headers)
    coSoup = BeautifulSoup(tryConnect(req).read(), 'html.parser')
    tableSoup = coSoup.find('table', {'align' : 'Left'})
    fileState = 'w+'

    if (mode > 1):
        fileState = 'a'
    else:
        if (firstCheck(coCode, tableSoup)): return

    with open(os.path.join(folderPath, 'Companies', coCode, 'index.csv'), fileState, newline='', encoding='utf-8') as indexFile:

        csvwriter = csv.writer(indexFile)
        if (mode == 1): csvwriter.writerow(['Category', 'Headline', 'link'])
        for infoSoup in tableSoup.find_all('tr', {'class' : None}):
            spanInfo = infoSoup.find('span',{'id': re.compile(r'lbShortText')})
            aInfo = infoSoup.find('a', {'class' : 'news'})
            csvwriter.writerow([spanInfo.get_text(), aInfo.get_text(), 'http://www.hkexnews.hk'+aInfo['href']])

#        if not (temp[0:8] == '...More'):

#        for i in range(0, keywordCount):
#            matchKey = re.search(keywords[0][i], temp.strip(), re.M | re.I)
#            if (matchKey != None):# and (mode == 1):
#                pdfDown(coCode, pdfLink)

    #Stop parsing if no 'Next' button was found on page
    if coSoup.find('input', {"id": "ctl00_btnNext"}) != None:
        viewState = coSoup.find('input', {"id": "__VIEWSTATE"}).get('value')
        values = {'__VIEWSTATE' : viewState, '__VIEWSTATEENCRYPTED' : '', 'ctl00$btnNext.x' : '27', 'ctl00$btnNext.y' : '8'}
        cnW(values, coCode, mode+1)
    
#===============================================================================================================================================
#Download PDF
def pdfDown(coCode, pdfLink):

    #Check if folder exists, create otherwise
    if (os.path.isdir(os.path.join(folderPath, 'Companies', coCode, 'PDF')) == False):
        os.system ('mkdir ' + os.path.join('Companies', coCode, 'PDF')) 

    #Obtain the file name for the download target and check whether the document is 'pdf' file
    fileName = pdfLink.split('/')[-1].split('.')

    if (fileName[1] == 'pdf'):
        if (os.path.isfile(os.path.join(folderPath, 'Companies', coCode, 'PDF', fileName[0]) + '.pdf') == False):

            dlFile = open(os.path.join(folderPath, 'Companies', coCode, 'PDF', fileName[0]) + '.pdf', 'wb+')

            response = tryConnect(pdfLink)
            data = response.read()

            dlFile.write (data)
            dlFile.close()
#===============================================================================================================================================