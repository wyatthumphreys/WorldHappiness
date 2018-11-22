#Wyatt Humphreys | wlhumphreys@student.rtc.edu
#CNA330 11/13/18
#Sources used: https://www.sitepoint.com/using-python-parse-spreadsheet-data/ && https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/

import os, xlrd, xlwt, sqlite3

workbook = xlrd.open_workbook('WHR2018Chapter2OnlineData.xls')
Table21 = workbook.sheet_by_name('Table2.1')
Figure22 = workbook.sheet_by_name('Figure2.2')
Figure23 = workbook.sheet_by_name('Figure2.3')
Figure24 = workbook.sheet_by_name('Figure2.4')
SFactors = workbook.sheet_by_name('SupportingFactors')

#I am NOT smart enough to have written this block, code was copied from here and modified to work: https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/
num_cols = Figure23.ncols   # Number of columns
for row_idx in range(0, Figure23.nrows):    # Iterate through rows
    print ('-'*40)
    print ('Row: %s' % row_idx)   # Print row number
    for col_idx in range(0, num_cols):  # Iterate through columns
        cell_obj = Figure23.cell(row_idx, col_idx)  # Get cell object by row, col
        print ('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))

#con = sqlite3.connect("Happiness.db")
#cur = con.cursor()


