#HKEx Stock Analysis

A system to provide aid in analysing companies and industries listed in Hong Kong Stock Market

If you are interested to execute this code on a computer remotely, please take a look at my other tool <a href="https://github.com/jefflai0101/Listener"> listener.py </a>

##Targets

- To obtain a complete list of listed company on HKEx, and archive the list from the past

- To scrap company information from the HKEx official website

- To scrap financial information from Aastocks

- To calculate ratios for the companies, and show by industries based on classification from HKEx

- To filter companies with criteria

- To help understand the recent development of the company with company announcements (In Progress)

##Setup for the system

On the computer executing the code:

	|->	hkex.py
	|>>	Settings
		|-> Criteria.txt
		|->	keyword.csv
		|->	settings.txt
	|>>	hkextools
		|->	__init__.py
		|->	fAnalysis.py
		|->	highlight.py
		|->	nettools.py
		|->	statusSum.py
		|->	utiltools.py

Alternatively, you can move the **Criteria.txt** file to the output folder, if you decide to specify the output folder in the **settings.txt**

##How to use

1)	Set the desired output path in **settings.txt**
	
	Output Folder="C:/Dropbox/HKEx"

Note that you will need to ensure the specified location path exists
By default, if no path was specified, an folder **Output** will be created in the same level as **hkex.py**

2)	Modify 'keywords.csv' for company announcement analysis, in the following format

	[Announcement keyword]		[Required? (Y/N)]		[Category]
	ASSET RESTRUCTURING				Y					Restructure

3)	Modify the 'Criteria.txt' for company financial filter

	[Criteria Type]			[Threshold]		[M/L/N]
	Gross Profit Margin		30				M		

For example, the above will mean a filter for companies with **more** than **30%** in **Gross Profit Margin**

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

This iss designed under the Python 3.5 environment, and have not tested under other versions of Python.
