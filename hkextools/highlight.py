#imports
import os
import sys
import re
import csv
import xlsxwriter
#===============================================================================================================================================
from hkextools import utiltools
#===============================================================================================================================================
#Obtain current folder path
folderPath = ''
outputPath = ''
criteriaPath = ''
ratioLabels = ['Currency', 'Unit', 'Code', 'Name', 'Year End', 'Gross Profit Margin','EBITDA Margin','Net Profit Margin','EBITDA Coverage','Current Ratio','Quick Ratio','NAV','Debt to Assets','Debt to Equity','Average Total Assets','Average Total Equity','Assets Turnover','Leverage Ratio','ROE','Z-Score', 'PE', '3 Months Average', 'Latest Price']
#===============================================================================================================================================
#***********************************							Main Part									***********************************
#===============================================================================================================================================
def genCrit():
	inPath = os.path.join(criteriaPath, 'Criteria.txt')
	coFilters = utiltools.readSettings(inPath)
	items = [[coFilter, coFilters[coFilter]['Value'], coFilters[coFilter]['Mode']] for coFilter in coFilters]

	return items

#===============================================================================================================================================
def startFilter(thisSector, coFilters):
	coRatios = utiltools.readCSV(os.path.join(folderPath, 'Industries', thisSector + '.csv'))
	matchCo = []
	minMatch = [True for coFilter in coFilters if coFilter[2]!='N']

	#For each company
	for i in range(1, len(coRatios[0])):
		coInfo = {coRatio[0] : coRatio[i] for coRatio in coRatios}
		matchList = []
		resultDict = {}

		for coFilter in coFilters:
			if (coFilter[2]=='M'): matchList.append(float(coInfo[coFilter[0]]) >= float(coFilter[1]))
			if (coFilter[2]=='L'): matchList.append(float(coInfo[coFilter[0]]) <= float(coFilter[1]))

		if sum(matchList) == sum(minMatch): matchCo.append(coInfo)

	return matchCo

#===============================================================================================================================================
def main():
	coFilters = genCrit()
	allSectors = [item[0].replace('/', '').replace('(HSIC*)','') for item in utiltools.readCSV('IndustryIndex.csv') if item[0] != 'Banks']

	workbook = xlsxwriter.Workbook(os.path.join(outputPath, 'highlight.xlsx'))
	worksheet = workbook.add_worksheet()

	bold = workbook.add_format({'bold': True})
	percent_format = workbook.add_format({'num_format': '0.00"%"'})
	twodp_format = workbook.add_format({'num_format': '0.00""'})

	for i, ratioLabel in enumerate(ratioLabels): worksheet.write(i, 0, ratioLabel)

	foundCount = 1
	for thisSector in allSectors:
		coFound = startFilter(thisSector, coFilters)
		for coInfo in coFound:
			for i, ratioLabel in enumerate(ratioLabels):
				worksheet.write(i, foundCount, coInfo[ratioLabel])
			foundCount += 1

	worksheet.set_column(0, 0, 20)
	worksheet.set_column(1, foundCount, 15)
	worksheet.freeze_panes(0,1)
	workbook.close()

#===============================================================================================================================================