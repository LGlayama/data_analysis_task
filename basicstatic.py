
import xlrd
from xlwt import *
import csv
import datetime


rawdata_cpu1='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Cpu_Total_preprocessed_data\\Get1015Cpu_Total_preprocessed_data0.csv'
rawdata_cpu2='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Cpu_Total_preprocessed_data\\Get1015Cpu_Total_preprocessed_data1.csv'
rawdata_cpu3='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Cpu_Total_preprocessed_data\\Get1015Cpu_Total_preprocessed_data2.csv'

cleandatafile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\0. Clean Data.xls'

outputfile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\'


bsbook = Workbook(encoding = 'utf-8')
bssheet = bsbook.add_sheet('basic')

bssheet.write(0, 0, label = 'Number of collection times')
bssheet.write(0, 1, label = 'Number of valid data lines')
bssheet.write(0, 2, label = 'Number of valid Users')
bssheet.write(0, 3, label = 'Number of valid Apps')
bssheet.write(0, 4, label = 'Number of valid periods')
collection_times=['0']
with open(rawdata_cpu1) as cpu1f:
	cpu1 = csv.reader(cpu1f)
	headers = next(cpu1)
	uindex=headers.index('IMEIorMEID')
	stindex=headers.index('STARTTIME')
	etindex=headers.index('ENDTIME')

	for line in cpu1:
		nstime=datetime.datetime.strptime(line[stindex],"%Y-%m-%d %H:%M:%S")
		netime=datetime.datetime.strptime(line[etindex],"%Y-%m-%d %H:%M:%S")  	
		strperiod=line[stindex]+'-'+line[etindex]		
		gap=(netime-nstime).total_seconds()
		if (gap<82800) | (gap>90000):
			continue
		label= line[uindex]+'-'+strperiod

		if label in collection_times:
			
			continue
		collection_times.append(label)
with open(rawdata_cpu2) as cpu2f:
	cpu2 = csv.reader(cpu2f)
	headers = next(cpu2)
	uindex=headers.index('IMEIorMEID')
	stindex=headers.index('STARTTIME')
	etindex=headers.index('ENDTIME')
	for line in cpu2:
		nstime=datetime.datetime.strptime(line[stindex],"%Y-%m-%d %H:%M:%S")
		netime=datetime.datetime.strptime(line[etindex],"%Y-%m-%d %H:%M:%S")  	
		strperiod=line[stindex]+'-'+line[etindex]		
		gap=(netime-nstime).total_seconds()
		if (gap<82800) | (gap>90000):
			continue
		label= line[uindex]+'-'+strperiod

		if label in collection_times:
			
			continue

		collection_times.append(label)
with open(rawdata_cpu3) as cpu3f:
	cpu3 = csv.reader(cpu3f)
	headers = next(cpu3)
	uindex=headers.index('IMEIorMEID')
	stindex=headers.index('STARTTIME')
	etindex=headers.index('ENDTIME')
	for line in cpu3:
		nstime=datetime.datetime.strptime(line[stindex],"%Y-%m-%d %H:%M:%S")
		netime=datetime.datetime.strptime(line[etindex],"%Y-%m-%d %H:%M:%S")  	
		strperiod=line[stindex]+'-'+line[etindex]		
		gap=(netime-nstime).total_seconds()
		if (gap<82800) | (gap>90000):
			continue
		label= line[uindex]+'-'+strperiod

		if label in collection_times:
			
			continue
		collection_times.append(label)

collections=len(collection_times)-1
bssheet.write(1, 0, label = collections)

cleandata = xlrd.open_workbook(cleandatafile)
table = cleandata.sheet_by_index(0) 

nrows=table.nrows-2
bssheet.write(1,1,label=nrows)

users=table.col_values(1)
num_user=len(list(set(users)))-2

bssheet.write(1,2,label=num_user)

app=table.col_values(0)
num_app=len(list(set(app)))-2

bssheet.write(1,3,label=num_app)

date=table.col_values(2)
num_date=len(list(set(date)))-2

bssheet.write(1,4,label=num_date)

bsbook.save(outputfile+'0. Basic Statistics.xls')