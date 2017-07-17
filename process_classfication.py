# -*- coding: utf-8 -*-

import xlrd
from xlwt import *

inputfile1='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\App Classification List from Huawei.xlsx'
inputfile2='E:\\datawang\\HuaweiData-20170615\\0. Template\\Software\\2. Software Overall\\0. Process Classification.xlsx'
outputfilepath='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\2. Software Overall\\'

#copy the first three col from the first file

data1 = xlrd.open_workbook(inputfile1)
table1 = data1.sheet_by_index(1) 
cols1 = table1.col_values(0)
cols2 = table1.col_values(1)
cols3 = table1.col_values(2)
classes = table1.col_values(3)

#build the dictionary from Chinese to English

data2=xlrd.open_workbook(inputfile2)
table2 = data2.sheet_by_index(0)
Chinese= table2.col_values(5)
English= table2.col_values(4)
English[0]='others'



#produce new excel file

book = Workbook(encoding='utf-8')
sheet = book.add_sheet('Sheet1')


for i in range(0,len(cols1)):
    sheet.write(i,0,cols1[i])

for i in range(0,len(cols2)):
    sheet.write(i,1,cols2[i])

for i in range(0,len(cols3)):
    sheet.write(i,2,cols3[i])

for i in range(1,len(classes)):
	element=English[Chinese.index(classes[i])]
	sheet.write(i,4,element)

book.save(outputfilepath+'0. Process Classification.xls')


