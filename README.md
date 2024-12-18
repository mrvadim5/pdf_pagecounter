# pdf_pagecounter
Page Count for PDF and then export to .xls

File path to pull PDFs will need to be updated depending on what is being used (is it pulling from a flashdrive, desktop, folder, etc?). Be sure to use forward slashes in the filepath as backslashes can create errors in python. In Windows 11, the filepath can be copied by right clicking on a file or folder (this option does use backslashes though).

Program can be run in cmd line or any interpreter. I’ve been using Anaconda and Anaconda's Jupyter notebook.

The program will only open the first folder it is given, it will not go deeper (typically this means that each year will need to be pulled independently, as each large folder on the drives will yield no PDFs, as defined by this program.

Page totals can be auto-summed in Excel and then added to a master spreadsheet. These should then be saved in the Archive Page Count folder on the share drive.
The program will make its own Excel sheet on the desktop with a generic Page Count name. These can either be renamed or moved to a folder so duplicate issues don’t occur as more are generated.

The first four lines of code are the packages installed to make this run. If using a different interpreter or terminal they may need to be reinstalled.

Notes: Possible improvements to the code- ability to create sequential excel sheet names, ability to scan whole USB drive or deeper in folders, ability to merge data from each spreadsheet into a master spreadsheet (ideally this would be a second program).

PDF_Page_Counter2 appears to work with most IDEs. PDF_Page_Counter functions best with Jupyter. In 2, both the folder to be searched for page count and destination folder for export must be set manually. 

PDF_Counter_Request_Filepath replaces editing the code with the filepath with a user input. The file path for the export will still need to be entered.
