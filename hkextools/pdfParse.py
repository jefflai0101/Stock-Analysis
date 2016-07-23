#===============================================================================================================================================
#Imports for PDFParser
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
#===============================================================================================================================================
#***********************************								Parse PDF								***********************************
#===============================================================================================================================================
def parsePDF(pathPDF, pathText, fname):

	outfile = open(str(os.path.join(pathText, fname))[0:-4] + '.txt', 'w+', encoding='utf-8')

	fp = open(str(os.path.join(pathPDF, fname)), 'rb')
	parser = PDFParser(fp)
	doc = PDFDocument()
	parser.set_document(doc)
	doc.set_parser(parser)
	doc.initialize('')
	rsrcmgr = PDFResourceManager()
	laparams = LAParams()
	device = PDFPageAggregator(rsrcmgr, laparams=laparams)
	interpreter = PDFPageInterpreter(rsrcmgr, device)
	# Process each page contained in the document.
	for page in doc.get_pages():
	    interpreter.process_page(page)
	    layout = device.get_result()
	    for lt_obj in layout:
	        if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
	            #print(lt_obj.get_text())
	            outfile.write(lt_obj.get_text())
	            #outfile.write(lt_obj.get_text())
	    outfile.write ('=' * 100 + '\n')
#===============================================================================================================================================

#Codes for converting PDF into txt

'''
fnames = dnames = []
fndList = [[],[]]
dnames = dirWalk(os.path.join(folderPath, 'Companies'), 2)

for d in dnames:
	dPath = os.path.join(folderPath, 'Companies', d)
	fnames = dirWalk(os.path.join(dPath, 'PDF'), 3)

	for f in fnames:
		fndList[0].append(d)
		fndList[1].append(f)

fileCount = len(fndList[0])

for i in range(0, (fileCount - 1)):
	print ('Now on ' + str(fndList[0][i]) + ' - file : ' + str(fndList[1][i]))
	parsePDF(os.path.join(folderPath, 'Companies', str(fndList[0][i]), 'PDF'), os.path.join(folderPath, 'Companies', str(fndList[0][i]), 'txt'), str(fndList[1][i]))
'''

#===============================================================================================================================================