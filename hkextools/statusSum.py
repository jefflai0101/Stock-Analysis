#===============================================================================================================================================
import os
import re
import csv
from hkextools import utiltools
#===============================================================================================================================================
keywords = [[],[]]
keywordCount = 0
folderPath = ''
#===============================================================================================================================================
#Reading keywords from csv
def keywordsRead():
	global keywords
	global keywordCount
	with open(os.path.join(folderPath, 'Settings', 'keywords.csv'), "r") as csvfile:
		csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
		for row in csvreader:
			if (row[1] == 'Y'): 
				keywords[0].append(row[0])
				keywords[1].append(row[2])
				keywordCount += 1
	csvfile.close()
	return keywords, keywordCount

#===============================================================================================================================================
#Checks all announcements for brief status
def statusTag():
	docName = dnames = []
	dnames = utiltools.dirWalk(os.path.join(folderPath, 'Output', 'Companies'), 2)
	for d in dnames:
	    dPath = os.path.join(folderPath, 'Output', 'Companies', d )
	    lineCount = 0
	    print ('Summarising : ' + str(d))
	    with open(os.path.join(dPath, 'index.txt'), 'r', encoding='utf-8') as indexRead:
	        with open(os.path.join(dPath, 'statusTag.txt'), 'w+', encoding='utf-8') as indexWrite:
	            lines = indexRead.readlines()
	            for line in lines:
	                #if (lineCount < 40 and (not (line[0:3] == 'htt')) and (not (line[0:2] == '[]'))):
	                if ((not (line[0:3] == 'htt')) and (not (line[0:2] == '[]'))):
	                    for i in range(0, keywordCount):
	                        if (re.search(keywords[0][i], line, re.M | re.I) != None):
	                            docName.append(keywords[1][i])
	                            break
	                lineCount += 1
	            indexWrite.write(str(set(docName)))
	            docName = []
	    indexRead.close()

#===============================================================================================================================================
#Collects info from all the stock codes
def statusSummary(coRecord):
	topLine = ['Code', 'Name', 'Agreement', 'Auditor', 'Change', 'Debt', 'Delist', 'GEM', 'Halt', 'Insurance', 'Litigation', 'Loan', 'Potential Issue', 'Report', 'Resignation', 'Restriction', 'Restructure', 'Resumption', 'RTO', 'Share', 'Transaction', 'Long-stop', 'Offshore', 'Special Dividend', 'Specific Mandate', 'Fluctuation', 'Voting', 'IPO']
	with open(os.path.join(folderPath, 'Output', 'Status Summary.csv'), "w+", newline='', encoding='utf-8') as csvfile:
		csvwriter = csv.writer(csvfile)
		csvwriter.writerow (topLine)
		for i in range(1, len(coRecord)):
			lineContent = [coRecord[i][0], coRecord[i][2], 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N', 'N']
			with open(os.path.join(folderPath, 'Output', 'Companies', coRecord[i][0], 'statusTag.txt'), 'r') as indexRead:
				buffer = indexRead.readline()
			for j in range (2, 28):
				if (re.search(topLine[j], buffer, re.M | re.I) != None): lineContent[j] = 'Y'
			csvwriter.writerow (lineContent)

#===============================================================================================================================================