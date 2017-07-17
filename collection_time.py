import xlrd
from xlwt import *
import csv
import datetime
import random
import numpy as np


cleandatafile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\0. Clean Data.xls'

outputfile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\3. User\\'

pabook = Workbook(encoding = 'utf-8')
pasheet = pabook.add_sheet('basic')

cleandata = xlrd.open_workbook(cleandatafile)
table = cleandata.sheet_by_index(0) 

nrows=table.nrows
collection_users=[]
collection_times=[]

for i in range(2,nrows):
	userid=table.cell(i,1).value
	period=table.cell(i,2).value
	if userid not in collection_users:
		collection_users.append(userid)	
		times=[]
		times.append(period)
		collection_times.append(times)
	else:
		index=collection_users.index(userid)
		collection_times[index].append(period)
	
for a in range(0,len(collection_users)):

	row=a+1
	pasheet.write(row,0,collection_users[a])
	pasheet.write(row,1,len(collection_times[a]))


pabook.save(outputfile+'0. Collection times.xls')
