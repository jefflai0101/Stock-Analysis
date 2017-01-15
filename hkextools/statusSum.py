#===============================================================================================================================================
import os
import re
import csv
#===============================================================================================================================================
#Imports from HKEx Tools
from hkextools import nettools
from hkextools import utiltools
#===============================================================================================================================================
keywords = {}
folderPath = ''
#===============================================================================================================================================
#Reading keywords from csv
def keywordsRead():
    global keywords
    with open(os.path.join(folderPath, 'Settings', 'keywords.csv'), "r") as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar='\"')
        keywords = {key:category for (key, required, category) in list(csvreader) if required == 'Y'}
    return keywords

#===============================================================================================================================================
#Checks all announcements for brief status
def statusTag(coDict):
    for coCode in sorted(set(coDict)):
        docName = []
        dPath = os.path.join(folderPath, 'Companies', coCode)
        with open(os.path.join(dPath, 'index.csv'), 'r', encoding='utf-8') as indexRead:
            csvreader = csv.reader(indexRead, delimiter=',', quotechar='\"')
            coNews = list(csvreader)[1:]
        with open(os.path.join(dPath, 'statusTag.txt'), 'w+', encoding='utf-8') as indexWrite:
            for news in coNews:
                for keyword in keywords:
                    matchKey = re.search(keyword, ' '.join(news[0:1]), re.M | re.I)
                    if (matchKey != None): docName.append(keywords[keyword])
            indexWrite.write(str(set(docName)))

#===============================================================================================================================================
#Collects info from all the stock codes
def statusSummary(coDict):
    topLine = ['Code', 'Name', 'Agreement', 'Auditor', 'Change', 'Debt', 'Delist', 'GEM', 'Halt', 'Insurance', 'Litigation', 'Loan', 'Potential Issue', 'Report', 'Resignation', 'Restriction', 'Restructure', 'Resumption', 'RTO', 'Share', 'Transaction', 'Long-stop', 'Offshore', 'Special Dividend', 'Specific Mandate', 'Fluctuation', 'Voting', 'IPO']
    with open(os.path.join(folderPath, 'Status Summary.csv'), "w+", newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile)
        csvwriter.writerow (topLine)
        for coCode in sorted(set(coDict)):
            lineContent = [coCode, coDict[coCode]]
            with open(os.path.join(folderPath, 'Companies', coCode, 'statusTag.txt'), 'r') as indexRead:
                buff = indexRead.readline()
            lineContent.extend(['Y' if key in buff else 'N' for key in topLine[2:]])
            csvwriter.writerow (lineContent)

#===============================================================================================================================================
def main(coDict):
	statusTag(coDict)
	statusSummary(coDict)

#===============================================================================================================================================