#===============================================================================================================================================
#imports
import csv
import linecache
import os
import sys
import datetime
import threading
import ctypes
import time

#===============================================================================================================================================
#Imports from HKEx Tools
from hkextools import fAnalysis
from hkextools import highlight
from hkextools import nettools
from hkextools import statusSum
from hkextools import utiltools

#===============================================================================================================================================
#Assign folderPath to other modules
folderPath = os.path.dirname(os.path.abspath(__file__))
statusSum.folderPath = utiltools.folderPath = nettools.folderPath = fAnalysis.folderPath = highlight.folderPath = folderPath
#Assign date values to other modules
thisDay, thisMonth, thisYear = utiltools.obtDate()
fAnalysis.thisDay = nettools.thisDay = thisDay
fAnalysis.thisMonth = nettools.thisMonth = thisMonth
fAnalysis.thisYear = nettools.thisYear = thisYear
#Assign keyword values to other modules
keywords, keywordCount = statusSum.keywordsRead()
nettools.keywords = keywords
nettools.keywordCount = keywordCount
#===============================================================================================================================================
#***********************************								Main									***********************************
#===============================================================================================================================================
#coRecord = [[], [], [], [], [], [], [], [], [], [], [], [], [], []]
shellCommand = ''

shellCommand = 'cls' if os.name == 'nt' else 'clear'
if os.name == 'nt': ctypes.windll.kernel32.SetConsoleTitleA(b'HKEx')
os.system("mode con cols=30")
os.system(shellCommand)

print (folderPath)

letsSkip = False
#letsSkip = True

if letsSkip == False:
	#Check if CSV file exists
	if (os.path.isfile(os.path.join(folderPath,'coList.csv')) == True):

		#Check if folder 'Archive' exists. Shell cmd to archive the current version of CSV file into Archive folder
		if (os.path.isdir(os.path.join(folderPath, 'Archive')) == False): os.system ('mkdir Archive')
		archivePath = str(os.path.join(folderPath, 'Archive', str(datetime.date.today())) + '.csv')
		shellCommand = 'move ' if os.name == 'nt' else 'mv '
		os.system(shellCommand + os.path.join(folderPath, 'coList.csv') + ' ' + archivePath)

		#Open the archived CSV ready to compare
		coRecord = utiltools.readCoList(archivePath)

		#Merges the updated company info and the archived CSV to a new CSV file
		utiltools.outToCSV(1, coRecord, archivePath)
	else:
		#Creates and collect info from scratch if no CSV file was found
		utiltools.outToCSV(0, '', '')

	#Start to search annoucements for each company on the list
	coRecord = utiltools.readCoList(os.path.join(folderPath, 'coList.csv'))
	for i in range (0, len(coRecord[0])):
		aST = threading.Thread(target=nettools.aSearch(coRecord[0][i]))
		aST.start()

	#Periodically Annoucement Check goes here
	#[										]

	#Checks all announcements for brief status
	statusSum.statusTag()
	statusSum.statusSummary (coRecord[0], coRecord[2])

	#Obtain updated financial figures goes here
	#[										]

	fAnalysis.main(0)

	highlight.main()

else:

	fAnalysis.main(0)

	highlight.main()

#===============================================================================================================================================
