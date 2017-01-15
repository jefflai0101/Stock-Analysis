#imports
from bs4 import BeautifulSoup
import urllib.request
import os
import re
import csv
import socket
import requests
import threading
import xlsxwriter
#===============================================================================================================================================
#Obtain current folder path
thisDay = ''
thisMonth = ''
thisYear = ''
folderPath = ''
outputPath = ''
fxRate = {}
#outputPath = 'D:\\Dropbox\\Station\\HKEx'
#fxRate = {'USD' : 7.78, 'RMB' : 1.1, 'HKD' : 1, 'RM' : 1.74, 'SGD' : 5.43, 'JPY' : 0.067, 'EUR' : 8.23, 'CAD' : 5.90, 'GBP' : 9.47}
labelDict = {'NFI' : ['Currency', 'Unit', 'Year End', 'Gross Profit Margin','EBITDA Margin','Net Profit Margin','EBITDA Coverage','Current Ratio','Quick Ratio','NAV','Debt to Assets','Debt to Equity','Average Total Assets','Average Total Equity','Assets Turnover','Leverage Ratio','ROE','Z-Score', 'PE', '3 Months Average', 'Latest Price'], 'FI' : ['Currency', 'Unit', 'Year End', 'EBITDA Margin', 'NAV', 'Average Total Assets','Average Total Equity', 'Assets Turnover', 'Leverage Ratio', 'ROE', 'ROA', 'Loan/Deposit', 'Efficieny Ratio', 'Bad Debts Provision', 'Net Interest Spread', 'CAR', '3 Months Average', 'Latest Price']}
#===============================================================================================================================================
#Imports from HKEx Tools
from hkextools import nettools
from hkextools import utiltools
#===============================================================================================================================================
#***********************************							Connect and Scrape							***********************************
#===============================================================================================================================================
#Function in scraping company financial info on Aastocks
def cNS(coInfo):
	print('Collecting Financials for %s' %coInfo[0])
	if (os.path.isdir(os.path.join(folderPath, 'Companies', coInfo[0])) == False): os.system ('mkdir ' + os.path.join(folderPath, 'Companies', coInfo[0]))
	if (os.path.isdir(os.path.join(folderPath, 'Companies', coInfo[0], 'Financial')) == False): os.system ('mkdir ' + os.path.join('Companies', coInfo[0], 'Financial'))

	mode = ['profit-loss?symbol=', 'balance-sheet?symbol=', 'cash-flow?symbol=']
	coData = []
	
	for i in range(0,3):
		link = 'http://www.aastocks.com/en/stocks/analysis/company-fundamental/' + mode[i] + coInfo[0]
		content = nettools.tryConnect(link)
		coSoup = BeautifulSoup(content, 'html.parser')

		for trInfo in coSoup.find_all('tr', {'ref' : re.compile(r'[A-Z]+_Field(.*)+')}):
			fData = []
			fData.append(trInfo.find('td', {'class' : re.compile(r'fieldWithoutBorder')}).get_text().strip())
			if not ((i != 2) and (fData[0] in ['Auditor\'s Opinion', 'Unit', 'Currency'])) and not ((i != 0) and (fData[0] == 'Closing Date')):
				for tdInfo in trInfo.find_all('td', {'class' : re.compile(r'(.*)cls(.*)')}):
					if (tdInfo.get_text() != ''): fData.append(tdInfo.get_text().split('/')[0]) if (fData[0] == 'Closing Date') else fData.append(tdInfo.get_text())
			if (fData[0] == 'Closing Date') and (fData[-1] == coInfo[14] and fData[-1] != ''): return coInfo
			if len(fData) > 1 : coData.append(fData)

	with open(os.path.join(folderPath,'Companies', coInfo[0], 'Financial', 'Figures.csv'), "w+", newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for fData in coData: csvwriter.writerow(fData)

	coInfo[14] = coData[0][-1]
	coInfo[15] = 'FI' if len(coData) == 78 else 'NFI'

	return coInfo

	#Add case for adding onto existing financial data e.g. [2011~2015] + 2016

#===============================================================================================================================================
def stockPriceInfo(coCode):
	link = 'https://hk.finance.yahoo.com/q/hp?s='+coCode[1:]+'.HK&a=00&b=4&c='+str(int(thisYear)-2)+'&d=00&e=13&f='+thisYear+'&g=d'
	content = nettools.tryConnect(link)
	coSoup = BeautifulSoup(content, 'html.parser')

	priceList = []

	for trInfo in coSoup.find_all('tr'):
		tdInfo = trInfo.find_all('td', {'class' : 'yfnc_tabledata1'})
		if tdInfo==None or len(tdInfo)!=7: continue
		priceList.append(float(tdInfo[4].get_text()))

	if (priceList==[]):
		lastPrice = float(coSoup.find('span', {'class' : 'time_rtq_ticker'}).get_text())
		threeMonthAverage = float(lastPrice)
	else:
		lastPrice = priceList[0]
		threeMonthAverage = round(sum(priceList)/len(priceList),3)

	return threeMonthAverage, lastPrice

#===============================================================================================================================================
def collectData(coList):
	for coInfo in coList[1:]: coInfo = cNS(coInfo)

	with open(os.path.join(folderPath, 'coList.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for coInfo in coList: csvwriter.writerow(coInfo)

	return coList

#===============================================================================================================================================
def collectCAR(coCode):
	link = 'http://www.aastocks.com/en/stocks/analysis/company-fundamental/financial-ratios?symbol=' + coCode
	content = nettools.tryConnect(link)
	coSoup = BeautifulSoup(content, 'html.parser')
	
	for trInfo in coSoup.find_all('tr', {'ref' : re.compile(r'[A-Z]+_Field(.*)+')}):
		fData = []
		fData.append(trInfo.find('td', {'class' : re.compile(r'fieldWithoutBorder')}).get_text().strip())
		if (fData[0] == 'Capital Adequacy (%)'):
			for tdInfo in trInfo.find_all('td', {'class' : re.compile(r'(.*)cls(.*)')}):
				fData.append(tdInfo.get_text().split('/')[0]) if (fData[0] == 'Closing Date') else fData.append(tdInfo.get_text())
			return fData

#===============================================================================================================================================
def copyFiles():
	#[copy coList.csv and IndustryIndex.csv to dropbox]
	os.system ('copy ' + 'coList.csv ' + os.path.join(outputPath, 'coList.csv'))
	os.system ('copy ' + 'IndustryIndex.csv ' + os.path.join(outputPath, 'IndustryIndex.csv'))

#===============================================================================================================================================
#***********************************							Calculation									***********************************
#===============================================================================================================================================
def calRatio(coList):
	#Exclude '02277' and '08346'?
	for coInfo in coList:
		print('Calculating ratios for %s' %coInfo[0])
		calThread = threading.Thread(target=calNOut(coInfo))
		calThread.start()

#===============================================================================================================================================
def calNOut(coInfo):
	allRatios = []
	fData = utiltools.readCSV(os.path.join(folderPath, 'Companies', coInfo[0], 'Financial', 'Figures.csv'))
	if (fData[18][0]=='Interest Paid'): fData[18][0]='Interest Expense'
	fKey = {dataPoint[0] : key for key, dataPoint in enumerate(fData)}
	fData = [row[1:] for row in fData]

	figUnit = 1
	if (fData[fKey['Unit']][0]=='Thousand'): figUnit = 1000
	if (fData[fKey['Unit']][0]=='Million'): figUnit = 1000000

	#----------------------------------------------------------------------------------------------------------
	#General
	#----------------------------------------------------------------------------------------------------------
	#[0]	Currency - Unit
	allRatios.append(fData[fKey['Currency']])
	#[1]	Unit
	allRatios.append(fData[fKey['Unit']])
	#[2]	Ratio for Year ended
	allRatios.append(fData[fKey['Closing Date']])
	#[4]	EBITDA Margin
	if (coInfo[15]=='FI'): allRatios.append(ratioFormula(1, a=fData[fKey['Profit Before Taxation']], b=fData[fKey['Total Turnover']]))
	if (coInfo[15]=='NFI'): allRatios.append(ratioFormula(1, a=fData[fKey['EBITDA']], b=fData[fKey['Total Turnover']]))
	#[12]	Average Total Assets
	allRatios.append(ratioFormula(2, a=fData[fKey['Total Assets']]))
	#[13]	Average Total Equity
	allRatios.append(ratioFormula(2, a=fData[fKey['Total Equity']]))
	#[14]	Assets Turnover					
	allRatios.append(ratioFormula(1, a=fData[fKey['Net Profit']], b=allRatios[4]))
	#[15]	Leverage Ratio						
	allRatios.append(ratioFormula(1, a=allRatios[4], b=allRatios[5]))
	#[16]	ROE	
	allRatios.append(ratioFormula(3, a=allRatios[3], b=allRatios[6], c=allRatios[7]))
	#[19]	3 Months Average
	threeMonthAverage, lastPrice = stockPriceInfo(coInfo[0]) 
	allRatios.append(toList(len(allRatios[0]), threeMonthAverage))
	#[20]	Latest Price
	allRatios.append(toList(len(allRatios[0]), lastPrice))

	#----------------------------------------------------------------------------------------------------------
	#Non-FI
	#----------------------------------------------------------------------------------------------------------
	if (coInfo[15]=='NFI'):
		#[3]	Gross Profit Margin
		allRatios.insert(3, ratioFormula(1, a=fData[fKey['Gross Profit']], b=fData[fKey['Total Turnover']]))
		#[5]	Net Profit Margin							
		allRatios.insert(5, ratioFormula(1, a=fData[fKey['Net Profit']], b=fData[fKey['Total Turnover']]))
		#[6]	EBITDA Coverage
		allRatios.insert(6, ratioFormula(1, a=fData[fKey['EBITDA']], b=fData[fKey['Interest Expense']]))
		#[7]	Current Ratio								
		allRatios.insert(7, ratioFormula(1, a=fData[fKey['Current Assets']], b=fData[fKey['Current Liabilities']]))
		#[8]	Quick Ratio																	
		allRatios.insert(8, ratioFormula(1, a=fData[fKey['Cash On Hand']], b=fData[fKey['Current Liabilities']]))
		#[9]	NAV
		allRatios.insert(9, fData[fKey['Owner\'s Equity']])
		#[10]	Debt to Assets								
		allRatios.insert(10, ratioFormula(1, a=fData[fKey['Total Liabilities']], b=fData[fKey['Total Assets']]))
		#[11]	Debt to Equity
		allRatios.insert(11, ratioFormula(1, a=fData[fKey['Total Liabilities']], b=fData[fKey['Owner\'s Equity']]))
		#[17]	Z-Score
		CA = cleanVar(fData[fKey['Current Assets']])[-1]
		CL = cleanVar(fData[fKey['Current Liabilities']])[-1]
		reserves = cleanVar(fData[fKey['Reserves']])[-1]
		EBITDA = cleanVar(fData[fKey['EBITDA']])[-1]
		depN = cleanVar(fData[fKey['Depreciation']])[-1]
		averageTD = sum(cleanVar(fData[fKey['Total Liabilities']][-2:]))/2
		netProfit = cleanVar(fData[fKey['Net Profit']])[-1]
		allRatios.insert(17, findZscore(len(allRatios[0]), CA, CL, allRatios[12][-1], reserves, EBITDA, depN, int(coInfo[9].replace(',', '')), allRatios[17][-1], figUnit, averageTD, netProfit))

		#[18]	PE Ratio
#		allRatios.append(toList(len(allRatios[0]), 0 if (int(fData[fKey['Net Profit']][-1].replace(',', '')) < 0) else ((allRatios[9] * int(coInfo[9].replace(',', ''))) / (int(fData[fKey['Net Profit']][-1]) * figUnit))))
		allRatios.insert(18, toList(len(allRatios[0]), ((allRatios[18][-1] * int(coInfo[9].replace(',', ''))) / (int(netProfit) * figUnit * float(fxRate[coInfo[8]])))))

	#----------------------------------------------------------------------------------------------------------
	#FI
	#----------------------------------------------------------------------------------------------------------
	if (coInfo[15]=='FI'):
		#[3]	NAV
		allRatios.insert(6, fData[fKey['Equity']])
		#[10]	ROA
		allRatios.insert(10, ratioFormula(1, a=fData[fKey['Net Profit']], b=allRatios[4]))
		#[11]	Loan/Deposit Ratio
		allRatios.insert(11, ratioFormula(1, a=fData[fKey['Loans']], b=fData[fKey['Deposits']]))
		#[12]	Efficiency Ratio
		allRatios.insert(12, ratioFormula(1, a=fData[fKey['Operating Expenses']], b=fData[fKey['Total Turnover']]))
		#[13]	Bad Debts Provision
		allRatios.insert(13, fData[fKey['Bad Debt Provisions']])
		#[14]	Net Interest Spread
		allRatios.insert(14, ratioFormula(1, a=fData[fKey['Net Interest Income']], b=ratioFormula(4, a=fData[fKey['Loans']], b=fData[fKey['Financial Asset']], c=fData[fKey['Other Assets']])))
		#[15]	Capital Adequacy Ratio (CAR)
		allRatios.insert(15, collectCAR(coInfo[0])[1:])

	allRatios = [[labelDict[coInfo[15]][i]]+allRatio for i, allRatio in enumerate(allRatios)]

	with open(os.path.join(folderPath, 'Companies', coInfo[0], 'Financial', 'Ratios.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for allRatio in allRatios: csvwriter.writerow(allRatio)

	del allRatio
	#Fix Positions

#===============================================================================================================================================
def ratioFormula(mode, a=None, b=None, c=None):
	fRatio = []
	if (a != None): a = cleanVar(a)
	if (b != None): b = cleanVar(b)
	if (c != None): c = cleanVar(c)

	if (mode==1): fRatio = ['-' if b==0 else a/b for a,b in zip(a,b)]
	if (mode==2): fRatio = ['-'] + [(a+b)/2 for a,b in zip(a[1:], a[:-1])]
	if (mode==3): fRatio = ['-' if (a*b*c)==0 else (a*b*c) for a,b,c in zip(a,b,c)]
	if (mode==4): fRatio = [a+b+c for a,b,c in zip(a,b,c)]

	return fRatio

#===============================================================================================================================================
def findZscore(listLen, CA, CL, averageTA, reserves, EBITDA, depN, issuedShares, threeMonthAverage, figUnit, averageTD, netProfit):
	zScoreA = (CA-CL)/averageTA
	zScoreB = reserves/averageTA
	zScoreC = (EBITDA - depN)/averageTA
	zScoreD = (issuedShares * threeMonthAverage)/(figUnit * averageTD)
	zScoreE = netProfit/averageTA
	zScore = round((1.2 * zScoreA) + (1.4 * zScoreB) + (3.3 * zScoreC) + (0.6 * zScoreD) + zScoreE, 3)

	return toList(listLen, zScore)

#===============================================================================================================================================
def cleanVar(dataSet):
	return [0 if (fData=='0' or fData=='' or fData=='-') else fData if ((type(fData) is float) or (type(fData) is int)) else int(fData.replace(',', '')) for fData in dataSet]

#===============================================================================================================================================
def toList(listLen, inputValue):
	theList = ['-'] * (listLen - 1)
	theList.append(inputValue)
	return theList

#===============================================================================================================================================
#*********************************							Classification										********************************
#===============================================================================================================================================
def classifyList(coList):
#	classList = {item[0] : item[7].split(' (')[0].split(' - ')[-1] for item in coList if item[15] != 'FI'}
#	classList.update({item[0] : 'Financial Services' for item in coList if item[15] == 'FI'})
	classList = {item[0] : item[7].split(' (')[0].split(' - ')[-1] if (item[7].split(' (')[0].split(' - ')[-1] != 'Banks') else item[7].split(' (')[0].split(' - ')[-1] if (item[15] == 'FI') else 'Other Financials' for item in coList}
	uniqueList = sorted(set(classList.values()))
	with open(os.path.join(folderPath, 'IndustryIndex.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for industry in uniqueList:
			listToWrite = [industry] + [coCode for coCode in classList if classList[coCode] == industry]
			csvwriter.writerow(listToWrite)

#===============================================================================================================================================
#*************************************							Output									***************************************
#===============================================================================================================================================
#Extracts ratios and export as one industry
def industryRatio(sectorList, fileName, mode):
	labelKey = {0 : 'NFI', 1 : 'FI'}
	sectorRatios = list(labelDict[labelKey[mode]])
	sectorRatios.insert(2, 'Code')
	sectorRatios.insert(3, 'Name')
	keyDict = {label : i for i, label in enumerate(sectorRatios)}
	sectorRatios = [sectorRatios]

	for coInfo in sectorList:
		adjustFactor = 1

		coRatios = utiltools.readCSV(os.path.join(folderPath, 'Companies', coInfo[0], 'Financial', 'Ratios.csv'))
		coRatios = [coRatio[-1] for coRatio in coRatios]
		coRatios.insert(2, coInfo[0])
		coRatios.insert(3, coInfo[1])
		
		fxAdjust = float(fxRate[coRatios[keyDict['Currency']]])
		coRatios[keyDict['Currency']] = 'HKD'
		unitAdjust = 1000 if coRatios[keyDict['Unit']]=='Thousand' else 1
		coRatios[keyDict['Unit']] = 'Million'

		toHKD = ['NAV', 'Average Total Assets', 'Average Total Equity']
		if (mode): toHKD.extend(['Bad Debts Provision'])
		for item in toHKD: coRatios[keyDict[item]] = round(float(coRatios[keyDict[item]].replace(',', '')) * fxAdjust / unitAdjust, 2)

		dtoP = ['EBITDA Margin', 'Assets Turnover', 'Leverage Ratio', 'ROE']
		dtoP.extend(['ROA', 'Loan/Deposit', 'Efficieny Ratio', 'Net Interest Spread']) if (mode) else dtoP.extend(['Gross Profit Margin', 'Net Profit Margin', 'Debt to Assets', 'Debt to Equity'])
		for item in dtoP: coRatios[keyDict[item]] = 0 if coRatios[keyDict[item]] == '-' else round(float(coRatios[keyDict[item]].replace(',',''))*100, 2)

		toRound = [] if (mode) else ['EBITDA Coverage','Current Ratio','Quick Ratio', 'PE']
		for item in toRound: coRatios[keyDict[item]] = 0 if coRatios[keyDict[item]] == '-' else round(float(coRatios[keyDict[item]].replace(',','')), 2)

		coRatios = [0 if item=='-' else round(float(item), 2) if (type(item) is float or type(item) is int) else round(float(item.replace(',','')), 2) if (',' in item) else item for i, item in enumerate(coRatios)]
		
		sectorRatios.append(coRatios)

	sectorRatios = list(map(list, zip(*sectorRatios)))

	with open(os.path.join(folderPath, 'Industries',  fileName + '.csv'), 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		for allRatio in sectorRatios: csvwriter.writerow(allRatio)

#Separate 'FI' companies
#Group others by industry

	if (os.path.isdir(os.path.join(folderPath, 'ExcelRatios')) == False): os.system ('mkdir ' + os.path.join(folderPath, 'ExcelRatios'))

	workbook = xlsxwriter.Workbook(os.path.join(folderPath, 'ExcelRatios', fileName + '.xlsx'))
	worksheet = workbook.add_worksheet()

	bold = workbook.add_format({'bold': True})

	for i, rowItems in enumerate(sectorRatios):
		for j, item in enumerate(rowItems):
			worksheet.write(i, j, item, bold) if (j==0) else worksheet.write(i, j, item)

	worksheet.set_column(0, 0, 20)
	worksheet.set_column(1, len(sectorRatios[0])-1, 15)
	worksheet.freeze_panes(0,1)
	workbook.close()

#===============================================================================================================================================
#Extracts ratios and export as one industry, in excel format
def consolExcel(sectorNames):

	#outputPath = folderPath

	workbook = xlsxwriter.Workbook('consolRatios.xlsx')
	bold = workbook.add_format({'bold': True})
	percent_format = workbook.add_format({'num_format': '0.00"%"'}) 
	twodp_format = workbook.add_format({'num_format': '0.00""'})

	for sectorName in sectorNames:
		worksheet = workbook.add_worksheet(sectorName[0:30])

		sectorRatios = utiltools.readCSV(os.path.join('Industries', sectorName + '.csv'))

		for i, rowItems in enumerate(sectorRatios):
			for j, item in enumerate(rowItems):
				worksheet.write(i, j, item, bold) if (j==0) else worksheet.write(i, j, item)

		worksheet.set_column(0, 0, 20)
		worksheet.set_column(1, len(sectorRatios[0]), 15)
		worksheet.freeze_panes(0,1)
	workbook.close()

#===============================================================================================================================================
#***********************************							Main Part									***********************************
#===============================================================================================================================================
def main(mode):

	coList = utiltools.readCSV(os.path.join(folderPath, 'coList.csv'))

	if (mode == 1) or (sum([True if item[15] == '' else False for item in coList])): coList = collectData(coList)
	calRatio(coList[1:])
	classifyList(coList[1:])

	if (os.path.isdir(os.path.join(folderPath, 'Industries')) == False): os.system ('mkdir ' + os.path.join(folderPath, 'Industries'))

	sectorLists = utiltools.readCSV(os.path.join(folderPath, 'IndustryIndex.csv'))

	for sectorList in sectorLists:
		sectorName = sectorList[0].replace('/', '').replace('(HSIC*)','')
		print ('Working on : ' + sectorName)
		sectorCoInfo = [coInfo[0:2] for coInfo in coList if coInfo[0] in sectorList]
		industryRatio(sectorCoInfo, sectorName, 1) if (sectorName == 'Banks') else industryRatio(sectorCoInfo, sectorName, 0)

	consolExcel([sectorList[0].replace('/', '').replace('(HSIC*)','') for sectorList in sectorLists])
	copyFiles()

#===============================================================================================================================================