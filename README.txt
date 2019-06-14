This script takes multiple Logix Designer tags and logix comments exports, sorts the rows and puts them in a single spreadsheet with differences highlighted and each of the options from the projects in columns on the right.

1.	For each of the projects you want to combine export the tags & logic comments from RSLogix 5000 as a CSV file using the export tool: Tools > Export > Tags and Logic Comments…
2.	Open each of the CSV files in turn and save them as Excel 2010 onwards files (.xlsx)
3.	Install Python 3 https://www.python.org/downloads/
4.	Install openpyxl https://docs.python.org/3/installing/index.html (python -m pip install openpyxl)
5.	Open the Windows command prompt and run the Python script with each of the Excel files are command line arguments. E.g., C:\directory\where\excel\files\and\script\are>python TagsAndLogicCommentsMerge.py first.xlsx second.xlsx third.xlsx
6.	This will run for a minute and create a file TagsAndLogicCommentsMerge.xlsx which you open. There will be red fill in the DESCRIPTION column wherever there is a difference between the projects and a comparison of each of the projects in columns I onwards. All you have to do is copy the one you want into the DESCRIPTION column.
7.	Once you have finished fixing all the differences, delete all columns from I onwards and then save the file in CSV format (File > Save As).
8.	Now import this CSV file into one of the projects (do a backup of the project before of course).
