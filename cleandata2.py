
import xlrd
from xlwt import *
import csv
import datetime
from xlutils.copy import copy



rawdata_wake_lock1='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Wakelock_Total_preprocessed_data\\Get1015Wakelock_Total_preprocessed_data0.csv'
rawdata_wake_lock2='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Wakelock_Total_preprocessed_data\\Get1015Wakelock_Total_preprocessed_data0.csv'
rawdata_bright='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Brightness_Total_preprocessed_data.csv'
rawdata_gps='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Gps_Total_preprocessed_data.csv'
rawdata_gpu='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Gpu_Total_preprocessed_data.csv'
rawdata_sensor='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Sensor_Total_preprocessed_data.csv'
rawdata_wake_up='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Wakeup_Total_preprocessed_data.csv'



inputfilename='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\0. Clean Data.xls'

ccleandata=xlrd.open_workbook(inputfilename)

csheet=ccleandata.sheets()[0]
nrows = csheet.nrows

wcleandata=copy(ccleandata)

sheet=wcleandata.get_sheet(0)



for i in range(2,nrows):

	napp=csheet.row(i)[0].value
	nuser=csheet.row(i)[1].value
	nperiod=csheet.row(i)[2].value


	with open(rawdata_wake_lock1) as wlf1:
		wl1 = csv.reader( wlf1)
		headers = next(wl1)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		wl1index=headers.index('WAKELOCKPOWER.P4')
		for line in wl1:
			strperiod=line[stindex]+'-'+line[etindex]

			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,10, line[wl1index])

	with open(rawdata_wake_lock2) as wlf2:
		wl2 = csv.reader( wlf2)
		headers = next(wl2)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		wl2index=headers.index('WAKELOCKPOWER.P4')
		
		for line in wl2:
			strperiod=line[stindex]+'-'+line[etindex]
			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,10, line[wl2index])

	with open(rawdata_bright) as brightf:
		bright = csv.reader(brightf)
		headers = next(bright)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		bpindex=headers.index('BRIGHTNESSPOWER.P1')
		for line in bright:
			strperiod=line[stindex]+'-'+line[etindex]
			
			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,8, line[bpindex])

	with open(rawdata_gps) as gpsf:
		gps = csv.reader(gpsf)
		headers = next(gps)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		gpindex=headers.index('GPSPOWER.P1')
		for line in gps:

			strperiod=line[stindex]+'-'+line[etindex]

			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,9, line[gpindex])
										
	with open(rawdata_gpu) as gpuf:
		gpu = csv.reader(gpuf)
		headers = next(gpu)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		guindex=headers.index('GPUPOWER.P1')
		for line in gpu:
			strperiod=line[stindex]+'-'+line[etindex]
			
			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,12, line[guindex])
									

	with open(rawdata_sensor) as sensorf:
		sensor = csv.reader(sensorf)
		headers = next(sensor)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		ssindex=headers.index('SENSORPOWER.P4')
		for line in sensor:

			strperiod=line[stindex]+'-'+line[etindex]
			
			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,11, line[ssindex])
									

	with open(rawdata_wake_up) as wakeupf:
		wakeup = csv.reader(wakeupf)
		headers = next(wakeup)
		atrrindex=headers.index('APPNAME.P1')
		uindex=headers.index('IMEIorMEID')
		stindex=headers.index('STARTTIME')
		etindex=headers.index('ENDTIME')
		wuindex=headers.index('WAKEUPPOWER.P2')
		for line in wakeup:
			
			strperiod=line[stindex]+'-'+line[etindex]
			
			if napp==line[atrrindex] and nuser==line[uindex] and nperiod==strperiod:
				sheet.write(i,7, line[wuindex])

os.remove(inputfilename)
wcleandata.save(inputfilename)


									