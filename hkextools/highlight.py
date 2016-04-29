#imports
import os
import sys
import re
import csv
import xlsxwriter
#===============================================================================================================================================
#Obtain current folder path
folderPath = ''
dbPath = ''
ratioLabels = ['Original Currency', 'Original Unit', 'Year End', 'Stock Code', 'Name', 'Gross Profit Margin','EBITDA Margin','Net Profit Margin','EBITDA Coverage','Current Ratio','Quick Ratio','NAV','Debt to Assets','Debt to Equity','Average Total Assets','Average Total Equity','Assets Turnover','Leverage Ratio','ROE','Z-Score', 'PE', '3 Months Average', 'Latest Price']

#===============================================================================================================================================
#***********************************							Main Part									***********************************
#===============================================================================================================================================
def readCSV(pathToRead):
	#Reading all industries
	with open(pathToRead, 'r', encoding='utf-8') as csvfile:
		csvreader = csv.reader(csvfile)
		readInfo = list(csvreader)
		return readInfo


#===============================================================================================================================================
#Generate a csv file based on the txt file on dropbox
def genCrit():
	inPath = os.path.join(dbPath, 'Criteria.txt')
	outPath = os.path.join(folderPath, 'Settings', 'Criteria.csv')

	with open(outPath, 'w+', newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		with open(inPath, 'r', encoding='utf-8') as criteriaT:
			rows = criteriaT.readlines()
			i = 5
			for row in rows:
				items = row.strip('\n').split(',')
				items.append(i)
				csvwriter.writerow(items)
				i += 1

#===============================================================================================================================================
#Filtering for companies that matches the criteria specified
def startFilter(indC, coCriteria):
	coInfo = readCSV(os.path.join(folderPath, 'Industries', indC + '.csv'))

	listToHighlight = []
	listToReturn = [[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[],[]]
	for i in range(1, len(coInfo[0])):
		matchFlags = [False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False]
		for j in range(5, 22):
			if (j == 21):
				if (coCriteria[j-5][2] == 'M') and (float(coInfo[j+1][i]) > float(coInfo[j][i])): matchFlags[j-5] = True
				elif (coCriteria[j-5][2] == 'L') and (float(coInfo[j+1][i]) < float(coInfo[j][i])): matchFlags[j-5] = True
				elif (coCriteria[j-5][2] == 'N'): matchFlags[j-5] = True
			else:
				if (coCriteria[j-5][2] == 'M') and (float(coInfo[j][i]) > float(coCriteria[j-5][1])): matchFlags[j-5] = True
				elif (coCriteria[j-5][2] == 'L') and (float(coInfo[j][i]) < float(coCriteria[j-5][1])): matchFlags[j-5] = True
				elif (coCriteria[j-5][2] == 'N'): matchFlags[j-5] = True

		if (False in matchFlags):
			pass
		else:
			listToHighlight.append(coInfo[3][i])

	for item in listToHighlight:
		for i in range(0, 23):
			listToReturn[i].append(coInfo[i][coInfo[3].index(item)])

	return listToReturn

#===============================================================================================================================================
#Execution for the highlight process
def main():

	genCrit()

	if (os.name == 'nt'):
		outputPath = dbPath
	else:
		outputPath = folderPath

	itemList = [[5,6,7,12,13,16,18], [8,9,10,17,18,19,20,21,22,11,14,15], [0,1,2,3,4]]

	#Reading all industries
	industClasses = [item[0][0:-8] for item in readCSV(os.path.join(folderPath, 'IndustryIndex.csv'))]

	#Reading all filters
	coCriteria = readCSV(os.path.join(folderPath, 'Settings', 'Criteria.csv'))

	workbook = xlsxwriter.Workbook(os.path.join(outputPath, 'highlight.xlsx'))
	worksheet = workbook.add_worksheet()

	bold = workbook.add_format({'bold': True})
	percent_format = workbook.add_format({'num_format': '0.00"%"'})
	twodp_format = workbook.add_format({'num_format': '0.00""'})

	for i in range (0, 23): worksheet.write(i, 0, ratioLabels[i])

	itemsCount = 1
	for indC in industClasses:
		foundCo = []
		if (not indC in ['Banks', 'Cayman', 'Listin']):
			coInfo = startFilter(indC, coCriteria)
			for j in range(0, len(coInfo[0])):
				for i in range(0,23):
					if (i in itemList[0]):
						worksheet.write(i, (j + itemsCount), round(float(coInfo[i][j]),2), percent_format)
					elif (i in itemList[1]):
						worksheet.write_number(i, (j + itemsCount), round(float(coInfo[i][j]),2), twodp_format)
					elif (i in itemList[2]):
						worksheet.write(i, (j + itemsCount), coInfo[i][j])
					else:
						worksheet.write(i, (j + itemsCount), round(float(coInfo[i][j])))
					#worksheet.write(i, (j + itemsCount), coInfo[i][j])
			itemsCount += len(coInfo[0])

	worksheet.set_column(0, 0, 20)
	worksheet.set_column(1, itemsCount, 15)
	worksheet.freeze_panes(0,1)
	workbook.close()

#===============================================================================================================================================