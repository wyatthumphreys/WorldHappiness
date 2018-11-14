#Wyatt Humphreys | wlhumphreys@student.rtc.edu
#CNA330 11/13/18
#Sources used: https://www.sitepoint.com/using-python-parse-spreadsheet-data/

import os, xlrd, xlwt, sqlite3

workbook = xlrd.open_workbook('WHR2018Chapter2OnlineData.xls')

con = sqlite3.connect("Happiness.db")
cur = con.cursor()


