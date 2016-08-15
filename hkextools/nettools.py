#===============================================================================================================================================
from bs4 import BeautifulSoup
import urllib.request
import os
import re
import socket
import sys
import datetime
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
		networkOK = True
		try:
			content = urllib.request.urlopen(link, timeout=5)
		except urllib.error.URLError:
			networkOK = False
		except socket.timeout:
			networkOK = False
	return content

#===============================================================================================================================================
#Check if need to update announcement list
def checkFirstA(coCode, coSoup):
	if os.path.isfile(os.path.join(folderPath, 'Companies', coCode, 'index.txt')) == False:
		return False
	else:
		try:
			with open(os.path.join(folderPath, 'Companies', coCode, 'index.txt'), 'r', encoding='utf-8') as IndexFile:
				rows = IndexFile.readlines()
				if ((rows[1][len('http://www.hkexnews.hk'):].strip()) == (coSoup.find_all('a')[1].get('href')).strip()):
					#print ('No update needed!')
					return True
				else:
					#print ('Updating...')
					return False
		except UnicodeDecodeError:
			return False

#===============================================================================================================================================
#Search for announcements
def aSearch(coCode):

	#Create folder for the company if it doesn't exist
	if (os.path.isdir(os.path.join(folderPath, 'Output', 'Companies')) == False):
		os.system ('mkdir ' + str(os.path.join('Output', 'Companies')))
	if (os.path.isdir(os.path.join(folderPath, 'Output', 'Companies', coCode)) == False):
			os.system ('mkdir ' + os.path.join('Output', 'Companies', coCode))

	#Obtain the viewState value for the current page
	viewState = ''
	content = urllib.request.urlopen('http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx')
	coSoup = BeautifulSoup(content, 'html.parser')
	#Obtains the viewstate key of the current page
	for info in coSoup.find_all('input', {"id": "__VIEWSTATE"}):
		viewState = info.get('value')

	#Setting for the post
	headers = {}
	headers['User-Agent'] = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.97 Safari/537.36"
	url = 'http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx'
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

	# do POST
	viewState, noMore = cnW(values, headers, url, coCode, 1)

	#Continue for the company until older than 3 years or no more page available
	while (noMore == False):
		viewState, noMore = nextPage(viewState, coCode)

#===============================================================================================================================================
# Continue search for next page
def nextPage(viewState, coCode):
	#Post settings
	headers = {}
	headers['User-Agent'] = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.97 Safari/537.36"
	url = 'http://www.hkexnews.hk/listedco/listconews/advancedsearch/search_active_main.aspx'
	values = {
	'__VIEWSTATE' : viewState,
	'__VIEWSTATEENCRYPTED' : '',
	'ctl00$btnNext.x' : '27',
	'ctl00$btnNext.y' : '8'
	}

	# do POST
	viewState, noMore = cnW(values, headers, url, coCode, 2)

	return viewState, noMore
#===============================================================================================================================================
# Make HTML request
def cnW(values, headers, url, coCode, mode):
	#Post http request
	data = urllib.parse.urlencode(values)
	req = urllib.request.Request(url, data.encode('Big5'), headers = headers)
	rsp = tryConnect(req)
	#rsp = urllib.request.urlopen(req)
	content = rsp.read()
	#content = content.decode('Big5')

	#Initiate variables
	coSoup = BeautifulSoup(content, 'html.parser')
	count = 0
	temp = ''
	charToRe = ['\r', '\n', '\t']
	viewState = ''
	fileState = 'w+'
	noUpdates = False

	#Switches the w/r/a status for opening files
	if (mode == 2): fileState = 'a'

	if (mode == 1): 
		if (checkFirstA(coCode, coSoup) == True): return '', True

	#Opens text file for output for the current company
	with open(os.path.join(folderPath, 'Output', 'Companies', coCode, 'index.txt'), fileState, encoding='utf-8') as indexFile:

		for info in coSoup.find_all('a'):
			if count >= 1:
				temp = info.get_text()
				if not (temp[0:8] == '...More'):
					if not (info.get('href') == 'search_active_main.aspx'):
						pdfLink = 'http://www.hkexnews.hk' + info.get('href')
						for nlChar in charToRe:
							temp = temp.replace(nlChar, '')
						indexOut = '[' + temp.strip() + ']\n' + pdfLink + '\n'
						indexFile.write (indexOut)

#						for i in range(0, keywordCount):
#							matchKey = re.search(keywords[0][i], temp.strip(), re.M | re.I)
#							if (matchKey != None):# and (mode == 1):
#								pdfDown(coCode, pdfLink)
			count += 1

	#Continue parsing Flag
	noMore = False

	#Obtains the viewstate key of the current page
	for info in coSoup.find_all('input', {"id": "__VIEWSTATE"}):
		viewState = info.get('value')

	#Stop parsing if no 'Next' button was found on page
	if coSoup.find_all('input', {"id": "ctl00_btnNext"}) == []:
		noMore = True
	
	return viewState, noMore
#===============================================================================================================================================
#Download PDF
def pdfDown(coCode, pdfLink):

	#Check if folder exists, create otherwise
	if (os.path.isdir(os.path.join(folderPath, 'Output', 'Companies', coCode, 'PDF')) == False):
		os.system ('mkdir ' + os.path.join('Output','Companies', coCode, 'PDF'))	

	#Obtain the file name for the download target and check whether the document is 'pdf' file
	fileName = pdfLink.split('/')[-1].split('.')

	if (fileName[1] == 'pdf'):
		if (os.path.isfile(os.path.join(folderPath, 'Output', 'Companies', coCode, 'PDF', fileName[0]) + '.pdf') == False):

			dlFile = open(os.path.join(folderPath, 'Output', 'Companies', coCode, 'PDF', fileName[0]) + '.pdf', 'wb+')

			response = tryConnect(pdfLink)
			data = response.read()

			dlFile.write (data)
			dlFile.close()
#===============================================================================================================================================