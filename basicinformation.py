
import xlrd
from xlwt import *
import csv
import datetime

outputfile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\2. Software Overall\\'

cleandatafile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\0. Clean Data.xls'

bibook = Workbook(encoding = 'utf-8')
bisheet = bibook.add_sheet('basic')

cleandata = xlrd.open_workbook(cleandatafile)
table = cleandata.sheet_by_index(0) 
nrows=table.nrows
current_app =table.cell(2,0).value
sum_ft=0.0
sum_bt=0.0
sum_f=0.0
sum_b=0.0
sum_fb=0.0
sum_a=0.0
row=2
for i in range(2,nrows):
	if table.cell(i,0).value==current_app:
		sum_ft=sum_ft+ float(table.cell(i,3).value)
		sum_bt=sum_bt+ float(table.cell(i,4).value)
		if float(table.cell(i,3).value) != 0:
			sum_f=sum_f+1
		if float(table.cell(i,4).value) != 0:
			sum_b=sum_b+1
		if float(table.cell(i,3).value) != 0 and float(table.cell(i,4).value) != 0: 
			sum_fb=sum_fb+1
		sum_a=sum_a+1
	else:
		bisheet.write(row,0,current_app)
		bisheet.write(row,1,sum_ft/sum_a)
		bisheet.write(row,2,sum_bt/sum_a)
		bisheet.write(row,3,sum_f/sum_a)
		bisheet.write(row,4,sum_b/sum_a)
		bisheet.write(row,5,sum_fb/sum_a)
		row=row+1
		current_app=table.cell(i,0).value

		sum_ft=0.0
		sum_bt=0.0
		sum_f=0.0
		sum_b=0.0
		sum_fb=0.0
		sum_a=0.0

		sum_ft=sum_ft+ float(table.cell(i,3).value)
		sum_bt=sum_bt+ float(table.cell(i,4).value)
		if float(table.cell(i,3).value) != 0:
			sum_f=sum_f+1
		if float(table.cell(i,4).value) != 0:
			sum_b=sum_b+1
		if float(table.cell(i,3).value) != 0 and float(table.cell(i,4).value) != 0: 
			sum_fb=sum_fb+1
		sum_a=sum_a+1

bisheet.write(row,0,current_app)
bisheet.write(row,1,sum_ft/sum_a)
bisheet.write(row,2,sum_bt/sum_a)
bisheet.write(row,3,sum_f/sum_a)
bisheet.write(row,4,sum_b/sum_a)
bisheet.write(row,5,sum_fb/sum_a)

bibook.save(outputfile+'1. Basic Information.xls')




