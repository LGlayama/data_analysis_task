import xlrd
from xlwt import *
import csv
import datetime

cleandatafile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\0. Clean Data.xls'

outputfile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\'

cleandata = xlrd.open_workbook(cleandatafile)
table = cleandata.sheet_by_index(0) 


apps=list(set(table.col_values(0)[2:]))

print apps