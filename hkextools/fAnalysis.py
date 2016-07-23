#imports
from bs4 import BeautifulSoup
import urllib.request
import os
import re
import csv
import socket
import datetime
import xlsxwriter
import threading
#===============================================================================================================================================
#Obtain current folder path
folderPath = os.path.dirname(os.path.abspath(__file__))
#folderPath = ''
thisDay = ''
thisMonth = ''
thisYear = ''
#===============================================================================================================================================
#Imports from HKEx Tools
from hkextools import nettools

#===============================================================================================================================================
#***********************************							Connect and Scrape							***********************************
#===============================================================================================================================================
#Function in scraping company financial info on Aastocks
def cNS(code):

	if (os.path.isdir(os.path.join(folderPath, 'Companies', code)) == False):
		os.system ('mkdir ' + os.path.join(folderPath, 'Companies', code))
	if (os.path.isdir(os.path.join(folderPath, 'Companies', code, 'Financial')) == False):
		os.system ('mkdir ' + os.path.join('Companies', code, 'Financial'))

	cNSCount = 1
	link = ''
	writeFlag = True
	buff = []

	with open(os.path.join(folderPath,'Companies', code, 'Financial', 'Figures.csv'), "w+", newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)

		while (cNSCount < 4):
			if (cNSCount == 1):
				link = 'http://www.aastocks.com/en/stocks/analysis/company-fundamental/profit-loss?symbol=' + code
				fileName = 'PL.txt'
			if (cNSCount == 2):
				link = 'http://www.aastocks.com/en/stocks/analysis/company-fundamental/balance-sheet?symbol=' + code
				fileName = 'BS.txt'
			if (cNSCount == 3):
				link = 'http://www.aastocks.com/en/stocks/analysis/company-fundamental/cash-flow?symbol=' + code
				fileName = 'CF.txt'

			content = nettools.tryConnect(link)
			coSoup = BeautifulSoup(content, 'html.parser')
			
			for info in coSoup.find_all('td'):
				try:
					if (re.search('fieldWithoutBorder', str(info['class']), re.M | re.I) != None):
						if ((cNSCount != 1) and (info.get_text().strip() == 'Closing Date')):
							pass
						elif ((cNSCount != 3) and (info.get_text().strip() in ['Auditor\'s Opinion','Currency','Unit'])):
							pass
						else:
							writeFlag = False
							buff.append(info.get_text().strip())
					elif ((re.search('cls', str(info['class']), re.M | re.I) != None) and (writeFlag == False)):
						buff.append(info.get_text().strip())
						if (re.search('bold', str(info['class']), re.M | re.I) != None):
							writeFlag = True
				except KeyError:
					pass

				if (writeFlag == True and buff != []):
					csvwriter.writerow(buff)
					del buff[:]

			cNSCount += 1

#===============================================================================================================================================
#Scraping 3 month stock prices
def threeAverage(code):
	link = 'http://www.hkex.com.hk/eng/invest/stock_data/cache/pricetable_page_e_' + code.lstrip('0') + '_3.htm'
	#link = 'http://www.hkex.com.hk/eng/invest/company/pricetable_page_e.asp?WidCoID=' + code.strip('0') + '&Month=3&SRC=HP&SRC1=97701663'
	readFlag = False
	priceList = [[],[]]
	readCount = 0
	priceSum = 0
	recordCount = 0
	headers = {
	'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
	'Upgrade-Insecure-Requests': 1,
	'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36',
	'Referer': 'http://www.hkex.com.hk/eng/invest/company/chart_page_e.asp?WidCoID=' + code.strip('0') + '&WidCoAbbName=&Month=3&langcode=e',
	'Accept-Encoding': 'gzip, deflate, sdch',
	'Accept-Language': 'en-US,en;q=0.8,zh-TW;q=0.6,zh;q=0.4',
	'Cookie': '_ga=GA1.3.981083393.1448469839; ASPSESSIONIDQARRDRAD=DNDMEOABPNPAFBPKJDFHKMKD; HKEx=1347528896.20480.0000; TS0175cbd8=01e8725c7a5645eaa094af50bd40b768ab7075355ebebee2eebf7aaba91f821a63ccb9bf0932e4a05ee5ab68a46f5fd660462628d862b32ee1dac258f46ae9473e7b3c4f55'
	}
	values = {
	'WidCoID': code,
	'Month': 3,
	'SRC': 'HP',
	'SRC1': '97701663'
	}
	data = urllib.parse.urlencode(values)
	req = urllib.request.Request(link, data.encode('Big5'), headers = headers)
	rsp = nettools.tryConnect(link)
	content = rsp.read()
	#content = nettools.tryConnect(link)
	coSoup = BeautifulSoup(content, 'html.parser')

	#with open(os.path.join(folderPath, 'sample.txt'), "w+", encoding='utf-8') as htmlCode:
	#	htmlCode.write(coSoup.prettify())

	for info in coSoup.find_all('tr'):
		for dInfo in info.find_all('td'):
			if (dInfo.get_text() == 'Date'):
				readFlag = True
			if (readFlag == True):
				if (readCount == 3):
					matchKey = re.search('[a-zA-Z]', dInfo.get_text(), re.M | re.I)
					if (matchKey == None):
						priceList[0].append(recordCount)
						priceList[1].append(float(dInfo.get_text()))
						#priceSum += float(dInfo.get_text())
						#print (dInfo.get_text())
						recordCount += 1
					readCount = 0
				else:
					readCount += 1

	for i in range (0, recordCount - 66):
		priceList[0].pop(0)
		priceList[1].pop(0)

	if (len(priceList[1]) == 0): return priceForSus(code)

	for dailyPrice in priceList[1]: priceSum += dailyPrice

	return (priceSum/recordCount) if recordCount != 0 else 0
#===============================================================================================================================================
#Scrape for last closing price if suspended				Need to improve the suspension check
def priceForSus(code):
	#<td bgcolor="#CCCCCC" align="center" width="53">
	#<font face="Verdana, Arial, Helvetica, sans-serif" size="2">
	i = 0
	link = 'http://www.hkex.com.hk/eng/invest/company/quote_page_e.asp?WidCoID=' + code + '&WidCoAbbName=&Month=1&langcode=e'
	content = nettools.tryConnect(link)
	coSoup = BeautifulSoup(content, 'html.parser')

	for info in coSoup.find_all('font'):
		if (i == 18):
			if (info.get_text() == 'N/A'):
				return (0)
			else:
				return float(info.get_text())
		else:
			i += 1
	if (i < 18):
		return 0
#===============================================================================================================================================
#***********************************							Calculation									***********************************
#===============================================================================================================================================
#Makes all calls for calculations, and output for the company
def cNOut(coCode, issuedShares):

	figureList = []
	with open(os.path.join(folderPath, 'Companies', coCode, 'Financial', 'Figures.csv'), 'r', encoding='utf-8') as csvfile:
		csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
		figureList = list(csvreader)
		figUnit = 1
		if (figureList[79][1] == 'Thousand'): figUnit = 1000
		if (figureList[79][1] == 'Million'): figUnit = 1000000
	maxRange = min(3, len(figureList[0]) - 2)
	startLine = int(len(figureList[0]) - 1)

	with open(os.path.join(folderPath, 'Companies', coCode, 'Financial', 'Ratios.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for i in range(0, maxRange):
			allRatios = []
			
			#------------------------------------------------------------------------------------------------
			#[0]				Currency - Unit
			allRatios.append(figureList[80][1])
			#------------------------------------------------------------------------------------------------
			#[1]				Unit
			allRatios.append(figureList[79][1])
			#------------------------------------------------------------------------------------------------
			#[2]				Ratio for Year ended
			allRatios.append(figureList[0][startLine-i][0:4])
			#------------------------------------------------------------------------------------------------
			#[3]				Gross Profit Margin
			allRatios.append(calDivRatios(figureList[5][startLine-i], figureList[1][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[4]				EBITDA Margin
			allRatios.append(calDivRatios(figureList[14][startLine-i], figureList[1][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[5]				Net Profit Margin
			allRatios.append(calDivRatios(figureList[13][startLine-i], figureList[1][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[6]				EBITDA Coverage
			allRatios.append(calDivRatios(figureList[14][startLine-i], figureList[18][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[7]				Current Ratio
			allRatios.append(calDivRatios(figureList[27][startLine-i], figureList[36][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[8]				Quick Ratio
			allRatios.append(calDivRatios(calSumFigures(figureList[27][startLine-i],0,0) - calSumFigures(figureList[29][startLine-i],figureList[30][startLine-i],figureList[31][startLine-i]), figureList[36][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[9]				NAV
			allRatios.append(calSumFigures(figureList[42][startLine-i], 0, 0))
			#------------------------------------------------------------------------------------------------
			#[10]				Debt to Assets
			allRatios.append(calDivRatios(figureList[48][startLine-i], figureList[32][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[11]				Debt to Equity
			allRatios.append(calDivRatios(figureList[48][startLine-i], figureList[42][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[12]				Average Total Assets
			allRatios.append(calDivRatios(calSumFigures(figureList[32][startLine-i-1],figureList[32][startLine-i],0), 2, 2))
			#------------------------------------------------------------------------------------------------
			#[13]				Average Total Equity
			allRatios.append(calDivRatios(calSumFigures(figureList[47][startLine-i-1],figureList[47][startLine-i],0), 2, 2))
			#------------------------------------------------------------------------------------------------
			#[14]				Assets Turnover
			allRatios.append(calDivRatios(figureList[13][startLine-i], allRatios[12], 2))
			#------------------------------------------------------------------------------------------------
			#[15]				Leverage Ratio
			allRatios.append(calDivRatios(allRatios[12], allRatios[13], 2))
			#------------------------------------------------------------------------------------------------
			#[16]				ROE
			allRatios.append(round(allRatios[4] * allRatios[14] * allRatios[15], 4))
			#------------------------------------------------------------------------------------------------
			if (i < 1): 

				#[17]				Z-Score
				threeMonthAverage = threeAverage(coCode)
				zScoreA = calDivRatios(calSumFigures(figureList[27][startLine-i],0,0) - calSumFigures(figureList[36][startLine-i],0,0), allRatios[12], 2)
				zScoreB = calDivRatios(figureList[45][startLine-i], allRatios[12], 2)
				zScoreC = calDivRatios(calSumFigures(figureList[16][startLine-i],0,0) - calSumFigures(figureList[17][startLine-i],0,0), allRatios[12], 2)
				zScoreD = calDivRatios((calSumFigures(issuedShares,0,0) * threeMonthAverage / figUnit), figureList[48][startLine-i], 2) if (threeMonthAverage != 'N/A') else 0
				zScoreE = calDivRatios(figureList[1][startLine-i], allRatios[12], 2)
				zScore = round((1.2 * zScoreA) + (1.4 * zScoreB) + (3.3 * zScoreC) + (0.6 * zScoreD) + zScoreE, 3)
			
				allRatios.append(zScore)
				#For debug
				#print ('zScoreA is ' + str(zScoreA) + '\nzScoreB is ' + str(zScoreB) + '\nzScoreC is ' + str(zScoreC) + '\nzScoreD is ' + str(zScoreD) + '\nzScoreE is ' + str(zScoreE))
				#print ('zScoreD components are: ' + str(int(issuedShares.replace(',', '')) * threeMonthAverage / figUnit) + ' | ' + str(figureList[45][startLine-i]))
				#print ('Issued Shares: ' + str(int(issuedShares.replace(',', ''))))
				#print ('3 Month Average: ' + str(threeMonthAverage))
				#print ('Figure Unit: ' + str(figUnit))
				#print ('Z-score for ' + coCode + ' is ' + str(zScore))
			#------------------------------------------------------------------------------------------------
				#[18]				PE Ratio
				#print (threeMonthAverage)
				#print (calDivRatios(figureList[13][startLine-i], calSumFigures(issuedShares,0,0), 2) * figUnit)
				#print (calDivRatios(threeMonthAverage,calDivRatios(figureList[13][startLine-i], calSumFigures(issuedShares,0,0), 2), 2)/figUnit)
				#print (threeMonthAverage)
				peRatio = calDivRatios(threeMonthAverage,calDivRatios(figureList[13][startLine-i], calSumFigures(issuedShares,0,0), 2), 2)/figUnit
				allRatios.append(peRatio)
				#print('PE ratio for ' + coCode + ' is ' + str(peRatio))
				#------------------------------------------------------------------------------------------------
				#[19]				3 Months Average
				allRatios.append(threeMonthAverage)
				#------------------------------------------------------------------------------------------------
				#[20]				Latest Price
				allRatios.append(priceForSus(coCode))
				#------------------------------------------------------------------------------------------------
			csvwriter.writerow(allRatios)

#===============================================================================================================================================
#Calculate Ratios for Banks
def cBNOut(coCode):

	carRatio = []
	writeFlag = False
	link = 'http://www.aastocks.com/en/stocks/analysis/company-fundamental/financial-ratios?symbol=' + coCode
	content = nettools.tryConnect(link)
	coSoup = BeautifulSoup(content, 'html.parser')
	for info in coSoup.find_all('td'):
		try:
			if (re.search('fieldWithoutBorder', str(info['class']), re.M | re.I) != None) and (info.get_text().strip() == 'Capital Adequacy (%)'):
				writeFlag = True
			elif ((re.search('cls', str(info['class']), re.M | re.I) != None) and (writeFlag == True)):
					carRatio.append(info.get_text().strip())
					if (re.search('bold', str(info['class']), re.M | re.I) != None):
						writeFlag = False
						break
		except KeyError:
			pass

	figureList = []
	with open(os.path.join(folderPath, 'Companies', coCode, 'Financial', 'Figures.csv'), 'r', encoding='utf-8') as csvfile:
		csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
		figureList = list(csvreader)
		figUnit = 1
		if (figureList[75][1] == 'Thousand'): figUnit = 1000
		if (figureList[75][1] == 'Million'): figUnit = 1000000
	maxRange = min(3, len(figureList[0]) - 2)
	startLine = int(len(figureList[0]) - 1)

	with open(os.path.join(folderPath, 'Companies', coCode, 'Financial', 'Ratios.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for i in range(0, maxRange):
			allRatios = []

			#------------------------------------------------------------------------------------------------
			#[0]				Currency
			allRatios.append(figureList[76][1])
			#------------------------------------------------------------------------------------------------
			#[1]				Unit
			allRatios.append(figureList[75][1])
			#------------------------------------------------------------------------------------------------
			#[2]				Ratio for Year ended
			allRatios.append(figureList[0][startLine-i][0:4])
			#------------------------------------------------------------------------------------------------
			#[3]				NAV
			allRatios.append(calSumFigures(figureList[40][startLine-i],0,0))
			#------------------------------------------------------------------------------------------------
			#[4]				Average TA
			allRatios.append(calDivRatios(calSumFigures(figureList[31][startLine-i], figureList[31][startLine-i-1],0), 2,2))
			#------------------------------------------------------------------------------------------------
			#[5]				Average Equity
			allRatios.append(calDivRatios(calSumFigures(figureList[40][startLine-i], figureList[40][startLine-i-1],0), 2,2))
			#------------------------------------------------------------------------------------------------
			#[6]				EBITDA Margin
			allRatios.append(calDivRatios(figureList[12][startLine-i], figureList[1][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[7]				Asset Turnover
			allRatios.append(calDivRatios(figureList[5][startLine-i], allRatios[4], 2))
			#------------------------------------------------------------------------------------------------
			#[8]				Leverage Ratio
			allRatios.append(calDivRatios(allRatios[4], allRatios[5], 2))
			#------------------------------------------------------------------------------------------------
			#[9]				ROE
			allRatios.append(round(allRatios[6] * allRatios[7] * allRatios[8], 4))
			#------------------------------------------------------------------------------------------------
			#[10]				ROA
			allRatios.append(calDivRatios(figureList[15][startLine-i], allRatios[4], 2))
			#------------------------------------------------------------------------------------------------
			#[11]				Loan/Deposit Ratio
			allRatios.append(calDivRatios(figureList[24][startLine-i], figureList[36][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[12]				Efficiency Ratio
			allRatios.append(calDivRatios(figureList[5][startLine-i], figureList[1][startLine-i], 2))
			#------------------------------------------------------------------------------------------------
			#[13]				Bad Debts Provision
			allRatios.append(calSumFigures(figureList[6][startLine-i], 0, 0))
			#------------------------------------------------------------------------------------------------
			#[14]				Net Interest Spread
			allRatios.append(calDivRatios(figureList[2][startLine-i], calSumFigures(figureList[24][startLine-i],figureList[28][startLine-i],figureList[30][startLine-i]), 2))
			#------------------------------------------------------------------------------------------------
			#[15]				Capital Adequacy Ratio (CAR)
			allRatios.append(carRatio[len(figureList[0]) - 2 - i])
			#------------------------------------------------------------------------------------------------
			if (i < 1): 
				#[16]				3 Months Average
				allRatios.append(threeAverage(coCode))
				#------------------------------------------------------------------------------------------------
				#[17]				Latest Price
				allRatios.append(priceForSus(coCode))
				#------------------------------------------------------------------------------------------------
				#print (allRatios[16], allRatios[17])
			csvwriter.writerow(allRatios)

#===============================================================================================================================================
#Return summed figures
def calSumFigures(figA, figB, figC):
	tempA = figA
	tempB = figB
	tempC = figC
	if (type(figA) is str): 
		if (figA == '-'or figA == ''):
			tempA = 0
		else:
			tempA = int(figA.replace(',', ''))
	if (type(figB) is str):
		if (figB == '-'or figB == ''):
			tempB = 0
		else:
			tempB = int(figB.replace(',', ''))
	if (type(figC) is str): 
		if (figC == '-'or figC == ''):
			tempC = 0
		else:
			tempC = int(figC.replace(',', ''))
	return (tempA + tempB + tempC)
#===============================================================================================================================================
#Returns divided ratios in percentage or decimals, depending on the operating mode
def calDivRatios(figA, figB, mode):
	figADone = False
	temp = 0
	if (figA == '-' or figA == '' or figA == 0):
		tempA = 0
		figADone = True
	else:
		tempA = figA
	tempB = figB
	if (type(figA) is str and figADone == False): tempA = int(figA.replace(',', ''))
	if (type(figB) is str): 
		if (figB == '-' or figB == ''):
			return 0
		else:
			tempB = int(figB.replace(',', ''))
	else:
		if (figB == 0):
			return 0
	if (not tempB == 0): temp = tempA / tempB
	if (mode == 1):
		return (str(round(temp ,4) * 100) + '%')
	elif (mode == 2):
		return temp#(round(temp, 4))

#===============================================================================================================================================
def eachCompany(coCode, noOfShares, bankList):
	print ('Calculating Ratios for : ' + coCode)
	cNS(coCode)
	#Calculate Ratios for the company, based on whether it's a bank or not
	if coCode not in bankList:
		if (coCode != '02277' and coCode != '08346'):
			cNOut(coCode, noOfShares)
	else:
		cBNOut(coCode)

#===============================================================================================================================================
#*********************************							Classification										********************************
#===============================================================================================================================================
#Creates industry classification list
def classifyList(classList, stockList):
	uniqueList = sorted(set(classList))
	stockList.pop(0)
	#print (uniqueList)
	listToWrite = []
	with open(os.path.join(folderPath, 'IndustryIndex.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for i in range(0, len(uniqueList)):
			listToWrite = [j for j, x in enumerate(classList) if x == uniqueList[i]]
			for indexPlace in range (0, len(listToWrite)):
				listToWrite[indexPlace] = stockList[listToWrite[indexPlace]]
			listToWrite.insert(0, uniqueList[i].replace('/', ' &'))
			csvwriter.writerow(listToWrite)
			listToWrite = []

#===============================================================================================================================================
#Extracts ratios and export as one industry
def extractRatios(listToWrite, ratioLabels, stockList, shortNameList, mode):

#	if (os.name == 'nt'):
#		outputPath = 'C:\\Users\\Jeff\\Dropbox\\Station\\HKEx\\Industries'
#	else:
#		outputPath = os.path.join(folderPath, 'Industries', listToWrite[0][0:-8] + '.csv')

	outputPath = os.path.join(folderPath, 'Industries')

	itemList = [[5,6,7,12,13,16,18], [8,9,10,17,18,19,20,21,22],[11,14,15]] if (mode == 1) else [[8,9,11,12,13,14,16],[10,17,18,19], [5,6,7,15]]
	indRatios = [[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]] if (mode == 1) else [[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]]
	maxRange = 23 if (mode == 1) else 20
	for i in range(1, len(listToWrite)):
		with open(os.path.join(folderPath, 'Companies', listToWrite[i], 'Financial', 'Ratios.csv'), 'r', encoding='utf-8') as csvfile:
			csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
			for row in csvreader:
				#indRatios[0].append(row[0])
				indRatios[0].append('HKD')
				#indRatios[1].append(row[1])
				indRatios[1].append('Million')
				indRatios[2].append(row[2])
				indRatios[3].append(listToWrite[i])
				indRatios[4].append(shortNameList[stockList.index(listToWrite[i]) + 1])
				adjFactor = (1/1000) if row[1] == 'Thousand' else 1
				adjFactor *= 7.78 if row[0] == 'USD' else 1
				adjFactor *= 1.2 if row[0] == 'RMB' else 1
				for j in range (5, maxRange):
					if (j in itemList[0]):
						#outFormat = str(round(float(row[j-2])*100,2))+'%'
						outFormat = round(float(row[j-2])*100,2)
					elif (j in itemList[1]):
						#outFormat = str(round(float(row[j-2]),2))
						outFormat = round(float(row[j-2]),2)
					elif (j in itemList[2]):
						#outFormat = str(round(float(row[j-2]) * adjFactor, 2))
						outFormat = round(float(row[j-2]) * adjFactor, 2)
					else:
						outFormat = row[j-2]
					indRatios[j].append(outFormat)
				break

	for i in range (0, maxRange):
		indRatios[i].insert(0, ratioLabels[i])

	with open(os.path.join(outputPath, listToWrite[0][0:-8] + '.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for i in range(0, maxRange):
			csvwriter.writerow(indRatios[i])

#===============================================================================================================================================
#Extracts ratios and export as one industry, in excel format
def ratiosToExcel(listToWrite, ratioLabels, stockList, shortNameList, mode):

	if (os.path.isdir(os.path.join(folderPath, 'ExcelRatios')) == False):
		os.system ('mkdir ' + os.path.join(folderPath, 'ExcelRatios'))

	outputPath = os.path.join(folderPath, 'ExcelRatios')

	workbook = xlsxwriter.Workbook(os.path.join(outputPath, listToWrite[0][0:-8] + '.xlsx'))
	worksheet = workbook.add_worksheet()

	bold = workbook.add_format({'bold': True})
	percent_format = workbook.add_format({'num_format': '0.00"%"'}) 
	twodp_format = workbook.add_format({'num_format': '0.00""'})  

	for i in range(0, len(ratioLabels)):
		worksheet.write(i, 0, ratioLabels[i], bold)

	itemList = [[5,6,7,12,13,16,18], [8,9,10,17,18,19,20,21,22],[11,14,15]] if (mode == 1) else [[8,9,11,12,13,14,16],[10,17,18,19], [5,6,7,15]]
	maxRange = 23 if (mode == 1) else 20
	for i in range(1, len(listToWrite)):
		with open(os.path.join(folderPath, 'Companies', listToWrite[i], 'Financial', 'Ratios.csv'), 'r', encoding='utf-8') as csvfile:
			csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
			for row in csvreader:
				#worksheet.write(0, i, row[0])
				worksheet.write(0, i, 'HKD')
				#worksheet.write(1, i, row[1])
				worksheet.write(1, i, 'Million')
				worksheet.write(2, i, row[2])
				worksheet.write(3, i, listToWrite[i])
				worksheet.write(4, i, shortNameList[stockList.index(listToWrite[i]) + 1])
				adjFactor = (1/1000) if row[1] == 'Thousand' else 1
				adjFactor *= 7.78 if row[0] == 'USD' else 1
				adjFactor *= 1.2 if row[0] == 'RMB' else 1
				for j in range (5, maxRange):
					if (j in itemList[0]):
						worksheet.write(j, i, round(float(row[j-2])*100,2), percent_format)
					elif (j in itemList[1]):
						worksheet.write_number(j, i, round(float(row[j-2]),2), twodp_format)
					elif (j in itemList[2]):
						worksheet.write_number(j, i, round(float(row[j-2]) * adjFactor, 2))
					else:
						worksheet.write(j, i, row[j-2])
				break

	worksheet.set_column(0, 0, 20)
	worksheet.set_column(1, len(listToWrite), 15)
	worksheet.freeze_panes(0,1)
	workbook.close()
#===============================================================================================================================================
#Extracts ratios and export as one industry, in excel format
def consolExcel(listToWrite):

	#outputPath = folderPath
	outputPath = 'C:\\Users\\Jeff\\Dropbox\\Station\\HKEx'

	workbook = xlsxwriter.Workbook(os.path.join(outputPath, 'consolRatios.xlsx'))
	bold = workbook.add_format({'bold': True})
	percent_format = workbook.add_format({'num_format': '0.00"%"'}) 
	twodp_format = workbook.add_format({'num_format': '0.00""'})

	for eachIndustry in listToWrite:

		#worksheet = workbook.add_worksheet(eachIndustry[0:-8])
		#print (eachIndustry.rstrip(' (HSIC*)')[0:30])
		worksheet = workbook.add_worksheet(eachIndustry.rstrip(' (HSIC*)')[0:30])

		with open(os.path.join(folderPath, 'Industries', eachIndustry[0:-8] + '.csv'), 'r', encoding='UTF-8') as csvfile:
			csvreader = csv.reader(csvfile)
			csvRatios = list(csvreader)
		
		#print ('len(csvRatios[0] is: ', len(csvRatios[0]))
		#print ('len(csvRatios is: ', len(csvRatios))

		for i in range(0, len(csvRatios[0])):
			for j in range(0, len(csvRatios)):
				if (i == 0):
					worksheet.write(j, i, csvRatios[j][i], bold)
				else:
					if (j in [5,6,7,11,12,13,16,18]):
						worksheet.write_number(j, i, round(float(csvRatios[j][i]),2), percent_format)
					elif (j in [8,14,15,18,19,20,21,22]):
						worksheet.write_number(j, i, round(float(csvRatios[j][i]),2), twodp_format)
					elif (j in [9,10,17]):
						worksheet.write_number(j, i, round(float(csvRatios[j][i]),2))
					else:
						worksheet.write(j, i, csvRatios[j][i])

		worksheet.set_column(0, 0, 20)
		worksheet.set_column(1, len(csvRatios[0]), 15)
		worksheet.freeze_panes(0,1)
	workbook.close()
#===============================================================================================================================================
def copyFiles():
	outputPath = 'C:\\Users\\Jeff\\Dropbox\\Station\\HKEx'

	#[copy coList.csv and IndustryIndex.csv to dropbox]
	os.system ('copy ' + 'coList.csv ' + os.path.join(outputPath, 'coList.csv'))
	os.system ('copy ' + 'IndustryIndex.csv ' + os.path.join(outputPath, 'IndustryIndex.csv'))

#===============================================================================================================================================
#***********************************							Main Part									***********************************
#===============================================================================================================================================
#Check if folder 'Archive' exists. Shell cmd to archive the current version of CSV file into Archive folder
def main(mode):

	bankList = []
	stockList = []
	shortNameList = []
	issuedShares = []
	classList = []
	consolList = []
	ratioLabels = ['Currency', 'Unit', 'Year End', 'Stock Code', 'Name', 'Gross Profit Margin','EBITDA Margin','Net Profit Margin','EBITDA Coverage','Current Ratio','Quick Ratio','NAV','Debt to Assets','Debt to Equity','Average Total Assets','Average Total Equity','Assets Turnover','Leverage Ratio','ROE','Z-Score', 'PE', '3 Months Average', 'Latest Price']
	bRatioLabels = ['Currency', 'Unit', 'Year End', 'Stock Code', 'Name', 'NAV', 'Average TA', 'Average Equity', 'EBITDA Margin', 'Asset Turnover', 'Leverage Ratio', 'ROE', 'ROA', 'Loan/Deposit', 'Efficieny Ratio', 'Bad Debts Provision', 'Net Interest Spread', 'CAR', '3 Months Average', 'Latest Price']

	#Checking for listed banks and separates from other companies
	with open(os.path.join(folderPath, 'Banks'), 'r', encoding='utf-8') as bankLog:
		rows = bankLog.readlines()
		for row in rows:
			bankList.append(row[0:5])

	#Reading all info from coList
	with open(os.path.join(folderPath, 'coList.csv'), 'r', encoding='UTF-8') as csvfile:
		csvreader = csv.reader(csvfile)
		coInfo = list(csvreader)
		#Extracting info needed
		stockList = [item[0] for item in coInfo]
		shortNameList = [item[1] for item in coInfo]
		issuedShares = [item[9].split(' ')[0] for item in coInfo]
		classList = [item[7].split(' - ')[-1][0:item[7].split(' - ')[-1].find('    ')] if (item[7].split(' - ')[-1].find('    ') != -1) else item[7].split(' - ')[-1] for item in coInfo if item[7] != 'Industry Classification']

	toPop = stockList.index('08346')
	stockList.pop(toPop)
	shortNameList.pop(toPop)
	issuedShares.pop(toPop)
	classList.pop(toPop)

	#Industry classification
	classifyList(classList, stockList)

	if (mode == 0):
		#Scrap and Download Financials
		for i in range(0, len(stockList)):
			scrapThread = threading.Thread(target=eachCompany(stockList[i], issuedShares[i+1], bankList))
			scrapThread.start()


	if (os.path.isdir(os.path.join(folderPath, 'Industries')) == False):
			os.system ('mkdir ' + os.path.join(folderPath, 'Industries'))
	
	with open(os.path.join(folderPath, 'IndustryIndex.csv'), 'r', encoding='utf-8') as csvfile:
		csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
		for row in csvreader:
			listToWrite = row

			print ('Working on : ' + listToWrite[0])
			if (listToWrite[0] == 'Banks (HSIC*)'):
				listToWrite.pop(listToWrite.index('00222'))
				extractRatios(listToWrite, bRatioLabels, stockList, shortNameList, 2)
				ratiosToExcel(listToWrite, bRatioLabels, stockList, shortNameList, 2)
			else:
				extractRatios(listToWrite, ratioLabels, stockList, shortNameList, 1)
				ratiosToExcel(listToWrite, ratioLabels, stockList, shortNameList, 1)
			consolList.append(listToWrite[0])

	consolExcel(consolList)

	copyFiles()

#===============================================================================================================================================