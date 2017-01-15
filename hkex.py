#===============================================================================================================================================
#imports
import os
import csv
import argparse
import datetime
import threading

#===============================================================================================================================================
#Imports from HKEx Tools
from hkextools import highlight
from hkextools import nettools
from hkextools import statusSum
from hkextools import utiltools
from hkextools import fAnalysis
#===============================================================================================================================================
#Assign folderPath to other modules
folderPath = os.path.dirname(os.path.abspath(__file__))
#Load settings
allSettings = utiltools.readSettings(os.path.join('Settings', 'settings.txt'))
fAnalysis.outputPath = highlight.outputPath = allSettings['Output Path']
highlight.criteriaPath = allSettings['Criteria Path']
fAnalysis.fxRate = allSettings['Currency']
#statusSum.folderPath = utiltools.folderPath = nettools.folderPath = fAnalysis.folderPath = highlight.folderPath = folderPath
statusSum.folderPath = utiltools.folderPath = nettools.folderPath = fAnalysis.folderPath = folderPath
#Assign date values to other modules
thisDay, thisMonth, thisYear = utiltools.obtDate()
fAnalysis.thisDay = nettools.thisDay = thisDay
fAnalysis.thisMonth = nettools.thisMonth = thisMonth
fAnalysis.thisYear = nettools.thisYear = thisYear
#Assign keyword values to other modules
keywords = statusSum.keywordsRead()
nettools.keywords = keywords
#===============================================================================================================================================
#***********************************								Main									***********************************
#===============================================================================================================================================

def main():
	os.system('clear')
	getData = 0

	parser = argparse.ArgumentParser()
	parser.add_argument('-d', '--data', help='Collect financial data for ratio calculations and analysis', action='store_true', required=False)
	parser.add_argument('-l', '--list', help='Collect list of companies', action='store_true', required=False)
	parser.add_argument('-a', '--analysis', help='Calculate ratios', action='store_true', required=False)
	parser.add_argument('-s', '--status', help='Collect announcements and summarise current status', action='store_true', required=False)
	parser.add_argument('-f', '--filter', help='Filtering companies with criteria', action='store_true', required=False)

	args = parser.parse_args()

	if (args.data): getData = 1
	if (os.path.isfile(os.path.join(folderPath,'coList.csv'))==False) or (args.list): utiltools.main()

	coRecords = utiltools.readCoList(os.path.join(folderPath, 'coList.csv'))
	coDict = {coRecord[0] : coRecord[2] for coRecord in coRecords}

	if (args.status): 
		for coCode in sorted(set(coDict)):
			aST = threading.Thread(target=nettools.aSearch(coCode))
			aST.start()
		statusSum.main(coDict)
	
	if (args.analysis): fAnalysis.main(getData)

	if (args.filter): highlight.main()

#===============================================================================================================================================
main()