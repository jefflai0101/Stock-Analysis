#HKEx Stock Analysis

A system to provide aid in analysing companies and industries listed in Hong Kong Stock Market

If you are interested to execute this code on a computer remotely, please take a look at my other tool <a href="https://github.com/jefflai0101/Listener"> listener.py </a>

##Targets

- To obtain a complete list of listed company on HKEx, and archive the list from the past

- To scrap company information from the HKEx official website

- To scrap financial information from Aastocks

- To calculate ratios for the companies, and show by industries based on classification from HKEx

- To filter companies with criteria

- To help understand the recent development of the company with company announcements (Incomplete)

##Setup for the system

On the computer executing the code:

	|->	hkex.py
	|>>	Settings
		|->	keyword.csv
		|->	settings.txt
	|>>	hkextools
		|->	__init__.py
		|->	fAnalysis.py
		|->	highlight.py
		|->	nettools.py
		|->	statusSum.py
		|->	utiltools.py

At the output folder:

	|->	Criteria.txt

##How to use

1)	Set the desired output path in **settings.txt**
	
	Output Folder='C:\Dropbox\HKEx'

Note that you don't need Dropbox to execute this code
You could specify the output folder to another location, such as the same folder where you put **hkex.py** at

2)	Modify 'keywords.csv' for company announcement analysis, in the following format

**You may skip this part as it is incomplete**

	[Announcement keyword]		[Required? (Y/N)]		[Category]
	ASSET RESTRUCTURING				Y					Restructure

3)	Modify the 'Criteria.txt' for company financial filter

	[Criteria Type]			[Threshold]		[M/L/N]		[Position in Excel]
	Gross Profit Margin		30				M					5

For example, the above will mean a filter for companies with **more** than **30%** in **Gross Profit Margin**

Modifying [Criteria Type] and [Position in Excel] **will crash** the system
	
	Notes:
	M = More
	L = Less
	N = Not required

4)	Run hkex.py

	python hkex.py

##Expect Outcomes

- IndustryIndex.csv		-	Industry Classification according to Information from HKEx

- consolRatios.xlsx		-	Showing all financial information based on Industry Classification

- highlight.xlsx		-	Showing financial information for filtered companies

Please refer to the **Sample Output** folder for outcome samples

Note: The program was written as a learning project for Python, and thus there are huge rooms to improve on the codes itself.
