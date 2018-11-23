#Wyatt Humphreys | wlhumphreys@student.rtc.edu
#CNA330 11/13/18
#Sources used: https://www.sitepoint.com/using-python-parse-spreadsheet-data/ && https://blogs.harvard.edu/rprasad/2014/06/16/reading-excel-with-python-xlrd/ && https://github.com/ZennyBaff/CNA330/blob/master/World_Happiness_Report.py

import os, xlrd, xlwt, sqlite3, matplotlib.pyplot as plt; plt.rcdefaults()
import pandas as pd
import numpy as np

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

#Borrowed some code from Ian with permission to do conversion with Figure 2.2. Modified in line 46. Original code found at: https://github.com/ZennyBaff/CNA330/blob/master/World_Happiness_Report.py

#Creating the lists that are read off of the Excel spreadsheet.
Page_2_A = (pd.read_excel('WHR2018Chapter2OnlineData.xls', 'Figure2.2', index_col=None,  na_values=['NA'], usecols= "A"))
Page_2_B = (pd.read_excel('WHR2018Chapter2OnlineData.xls', 'Figure2.2', index_col=None, converters={'Happiness':float}, na_values=['NA'], usecols= "B"))
Page_2_A_List = list(Page_2_A['Country'])
Page_2_B_List = list(Page_2_B['Happiness score'])
Page_2_B_List_Float = [float(i) for i in Page_2_B_List]
Passed_List = Page_2_A_List

#Creating the Bar Graph from the lists.
y_pos = np.arange(len(Page_2_A_List))
plt.bar(y_pos, Page_2_B_List_Float, align='center', alpha=0.5)
plt.xticks(y_pos,Page_2_A_List)
plt.ylabel('Happiness')
plt.xlabel('Countries')
plt.title("World Happiness Report")
plt.show()
print(Passed_List)


