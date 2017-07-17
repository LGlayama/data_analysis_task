

import xlrd
from xlwt import *
import csv
import datetime

rawdata_cpu1='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Cpu_Total_preprocessed_data\\Get1015Cpu_Total_preprocessed_data0.csv'
rawdata_cpu2='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Cpu_Total_preprocessed_data\\Get1015Cpu_Total_preprocessed_data1.csv'
rawdata_cpu3='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Cpu_Total_preprocessed_data\\Get1015Cpu_Total_preprocessed_data2.csv'
rawdata_wake_lock1='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Wakelock_Total_preprocessed_data\\Get1015Wakelock_Total_preprocessed_data0.csv'
rawdata_wake_lock2='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Wakelock_Total_preprocessed_data\\Get1015Wakelock_Total_preprocessed_data0.csv'
rawdata_bright='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Brightness_Total_preprocessed_data.csv'
rawdata_gps='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Gps_Total_preprocessed_data.csv'
rawdata_gpu='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Gpu_Total_preprocessed_data.csv'
rawdata_sensor='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Sensor_Total_preprocessed_data.csv'
rawdata_wake_up='E:\\datawang\\HuaweiData-20170615\\1. Raw Data\\Get1015Wakeup_Total_preprocessed_data.csv'


classficationfile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\2. Software Overall\\0. Process Classification.xls'
outputfilepath='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\'


class appinformation(object):
	def __init__(self,name):
		self.name = name     
       

class period(object):
	def __init__(self,name):
		self.name=name
        

class userinformation(object):
	def __init__(self,name):
		self.name = name
		

class realinformation(object):
    def __init__(self):        
        self.cpufrontime=None
        self.cpubacktime=None
        self.cpufrontpower=None
        self.cpubackpower=None
        self.wakeuppower=0
        self.brightnesspower=0
        self.gpspower=0
        self.wakelockpower=0
        self.sensorpower=0
        self.gpupower=0



apps=[]
#get the package of different app

classesbook = xlrd.open_workbook(classficationfile)
classessheet = classesbook .sheet_by_index(0)
classes= classessheet.col_values(0)
with open(rawdata_cpu1) as cpu1f,open(rawdata_cpu2) as cpu2f,open(rawdata_cpu3) as cpu3f:
	cpu1 = csv.reader(cpu1f)
	cpu2 = csv.reader(cpu2f)
	cpu3 = csv.reader(cpu3f)
#get the index of different value in the csv
	headers = next(cpu1)
	atrrindex=headers.index('APPNAME.P1')
	uindex=headers.index('IMEIorMEID')
	stindex=headers.index('STARTTIME')
	etindex=headers.index('ENDTIME')
	cftindex=headers.index('CPUPOWER.P1')
	cfpindex=headers.index('CPUPOWER.P2')
	cbtindex=headers.index('CPUPOWER.P3')
	cbpindex=headers.index('CPUPOWER.P4')

	for line in cpu1:
		if line[atrrindex] not in classes:

			continue

		ninfo=realinformation()
		ninfo.cpufrontime=line[cftindex] if line[cftindex]!='' else '0'
		ninfo.cpubacktime=line[cbtindex] if line[cbtindex]!='' else '0'
		ninfo.cpufrontpower=line[cfpindex] if line[cfpindex]!='' else '0'
		ninfo.cpubackpower=line[cbpindex] if line[cbpindex]!='' else '0'
		if float(ninfo.cpufrontime)+float(ninfo.cpubacktime)<15:

			continue
    		
		nstime=datetime.datetime.strptime(line[stindex],"%Y-%m-%d %H:%M:%S")
		netime=datetime.datetime.strptime(line[etindex],"%Y-%m-%d %H:%M:%S")  	
		strperiod=line[stindex]+'-'+line[etindex]		
		gap=(netime-nstime).total_seconds()
		if (gap<82800) | (gap>90000):

			continue

		flaga = 0

		if apps:

			for x in apps:
					
				if x.name == line[atrrindex]:
					flaga=1						
					tempa=apps.index(x)

			
		if flaga==0:
			napp=appinformation(line[atrrindex])
			napp.users=[]
			nuser=userinformation(line[uindex])
			nuser.periods=[]
			nper=period(strperiod)
            	   			
			nper.information=ninfo            	
			nuser.periods.append(nper)
			napp.users.append(nuser)
			apps.append(napp)
				
		else:
			flagb=0
			if apps[tempa].users:
				for p in apps[tempa].users:
					if p.name==line[uindex]:
						flagb=1
						tempb=apps[tempa].users.index(p)
			if flagb==0:
				nuser=userinformation(line[uindex])
				nuser.periods=[]
				nper=period(strperiod)

				nper.information=ninfo            	
				nuser.periods.append(nper)
				apps[tempa].users.append(nuser)
					
			else:
				nper=period(strperiod)

				nper.information=ninfo   
				apps[tempa].users[tempb].periods.append(nper)

	for line in cpu2:
		if line[atrrindex] not in classes:

			continue

		ninfo=realinformation()
		ninfo.cpufrontime=line[cftindex] if line[cftindex]!='' else '0'
		ninfo.cpubacktime=line[cbtindex] if line[cbtindex]!='' else '0'
		ninfo.cpufrontpower=line[cfpindex] if line[cfpindex]!='' else '0'
		ninfo.cpubackpower=line[cbpindex] if line[cbpindex]!='' else '0'
		if float(ninfo.cpufrontime)+float(ninfo.cpubacktime)<15:

			continue
    		
		nstime=datetime.datetime.strptime(line[stindex],"%Y-%m-%d %H:%M:%S")
		netime=datetime.datetime.strptime(line[etindex],"%Y-%m-%d %H:%M:%S")  	
		strperiod=line[stindex]+'-'+line[etindex]		
		gap=(netime-nstime).total_seconds()
		if (gap<82800) | (gap>90000):
			continue

		flaga = 0

		if apps:

			for x in apps:
					
				if x.name == line[atrrindex]:
					flaga=1						
					tempa=apps.index(x)

			
		if flaga==0:
			napp=appinformation(line[atrrindex])
			napp.users=[]
			nuser=userinformation(line[uindex])
			nuser.periods=[]
			nper=period(strperiod)
            	   			
			nper.information=ninfo            	
			nuser.periods.append(nper)
			napp.users.append(nuser)
			apps.append(napp)
				
		else:
			flagb=0
			if apps[tempa].users:
				for p in apps[tempa].users:
					if p.name==line[uindex]:
						flagb=1
						tempb=apps[tempa].users.index(p)
			if flagb==0:
				nuser=userinformation(line[uindex])
				nuser.periods=[]
				nper=period(strperiod)

				nper.information=ninfo            	
				nuser.periods.append(nper)
				apps[tempa].users.append(nuser)
					
			else:
				nper=period(strperiod)

				nper.information=ninfo   
				apps[tempa].users[tempb].periods.append(nper)

	for line in cpu3:
		if line[atrrindex] not in classes:

			continue

		ninfo=realinformation()
		ninfo.cpufrontime=line[cftindex] if line[cftindex]!='' else '0'
		ninfo.cpubacktime=line[cbtindex] if line[cbtindex]!='' else '0'
		ninfo.cpufrontpower=line[cfpindex] if line[cfpindex]!='' else '0'
		ninfo.cpubackpower=line[cbpindex] if line[cbpindex]!='' else '0'
		if float(ninfo.cpufrontime)+float(ninfo.cpubacktime)<15:

			continue
    		
		nstime=datetime.datetime.strptime(line[stindex],"%Y-%m-%d %H:%M:%S")
		netime=datetime.datetime.strptime(line[etindex],"%Y-%m-%d %H:%M:%S")  	
		strperiod=line[stindex]+'-'+line[etindex]		
		gap=(netime-nstime).total_seconds()
		if (gap<82800) | (gap>90000):

			continue

		flaga = 0

		if apps:

			for x in apps:
				if x.name == line[atrrindex]:
					flaga=1						
					tempa=apps.index(x)

			
		if flaga==0:
			napp=appinformation(line[atrrindex])
			napp.users=[]
			nuser=userinformation(line[uindex])
			nuser.periods=[]
			nper=period(strperiod)
            	   			
			nper.information=ninfo            	
			nuser.periods.append(nper)
			napp.users.append(nuser)
			apps.append(napp)
				
		else:
			flagb=0
			if apps[tempa].users:
				for p in apps[tempa].users:
					if p.name==line[uindex]:
						flagb=1
						tempb=apps[tempa].users.index(p)
						
			if flagb==0:
				nuser=userinformation(line[uindex])
				nuser.periods=[]
				nper=period(strperiod)

				nper.information=ninfo            	
				nuser.periods.append(nper)
				apps[tempa].users.append(nuser)
					
			else:
				nper=period(strperiod)

				nper.information=ninfo   
				apps[tempa].users[tempb].periods.append(nper)

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.wakelockpower=line[wl1index]

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.wakelockpower=line[wl2index]

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.brightnesspower=line[bpindex]

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.gpspower=line[gpindex]

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.gpupower=line[guindex]

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.sensorpower=line[ssindex]

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
		for x in apps:
			if line[atrrindex] == x.name:
				for y in x.users:
					if line[uindex] == y.name:
						for z in y.periods:
							if strperiod == z.name:
								
								z.information.wakeuppower=line[wuindex]

book = Workbook(encoding='utf-8')
sheet = book.add_sheet('Sheet1')

row=2
for i in range(0,len(apps)):
	for j in range(0,len(apps[i].users)):
		for k in range (0,len(apps[i].users[j].periods)):
			sheet.write(row,0 ,apps[i].name )
			sheet.write(row,1,apps[i].users[j].name)
			sheet.write(row,2,apps[i].users[j].periods[k].name)
			sheet.write(row,3,apps[i].users[j].periods[k].information.cpufrontime)
			sheet.write(row,4, apps[i].users[j].periods[k].information.cpubacktime)
			sheet.write(row,5, apps[i].users[j].periods[k].information.cpufrontpower)
			sheet.write(row,6, apps[i].users[j].periods[k].information.cpubackpower)

			sheet.write(row,7, apps[i].users[j].periods[k].information.wakeuppower)
			sheet.write(row,8, apps[i].users[j].periods[k].information.brightnesspower)
			sheet.write(row,9, apps[i].users[j].periods[k].information.gpspower)
			sheet.write(row,10, apps[i].users[j].periods[k].information.wakelockpower)
			sheet.write(row,11, apps[i].users[j].periods[k].information.sensorpower)
			sheet.write(row,12, apps[i].users[j].periods[k].information.gpupower)
			
			row=row+1



book.save(outputfilepath + '0. Clean Data.xls')









    		
    			
    			
            	



    		
    		

    		

            






