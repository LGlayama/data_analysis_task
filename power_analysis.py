import xlrd
from xlwt import *
import csv
import datetime
import random
import numpy as np


cleandatafile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\1. Preprocessing\\0. Clean Data.xls'

outputfile='E:\\datawang\\HuaweiData-20170615\\3. Result\\Software\\2. Software Overall\\'






pabook = Workbook(encoding = 'utf-8')
pasheet = pabook.add_sheet('basic')

cleandata = xlrd.open_workbook(cleandatafile)
table = cleandata.sheet_by_index(0) 

nrows=table.nrows
current_app =table.cell(2,0).value

sum_ft=0.0
sum_bt=0.0
sum_fp=0.0
sum_bp=0.0
sum_a=0.0
sum_wu=0.0
sum_br=0.0
sum_gs=0.0
sum_wl=0.0
sum_ss=0.0
sum_gu=0.0
tempdata=[]
temp=[]

row=2

for i in range(2,nrows):
	if table.cell(i,0).value==current_app:
		sum_ft=sum_ft+ float(table.cell(i,3).value)
		sum_bt=sum_bt+ float(table.cell(i,4).value)		
		sum_fp=sum_fp+ float(table.cell(i,5).value)
		sum_bp=sum_bp+ float(table.cell(i,6).value)
		
		sum_wu=sum_wu+ float(table.cell(i,7).value)
		sum_br=sum_br+ float(table.cell(i,8).value)
		sum_gs=sum_gs+ float(table.cell(i,9).value)
		sum_wl=sum_wl+ float(table.cell(i,10).value)
		sum_ss=sum_ss+ float(table.cell(i,11).value)
		sum_gu=sum_gu+ float(table.cell(i,12).value)

		temp=[float(table.cell(i,3).value),float(table.cell(i,4).value),float(table.cell(i,7).value),float(table.cell(i,8).value),float(table.cell(i,9).value),float(table.cell(i,10).value),float(table.cell(i,11).value),float(table.cell(i,12).value)]
		tempdata.append(temp)
		sum_a=sum_a+1
	else:
		pasheet.write(row,0,current_app)
		pasheet.write(row,1,sum_bt/(sum_ft+sum_bt))
		if sum_ft!= 0:
			pasheet.write(row,2,sum_fp/sum_ft)
		else:
			pasheet.write(row,2,'NAN')
		if sum_bt!= 0:
			pasheet.write(row,3,sum_bp/sum_bt)
		else:
			pasheet.write(row,3,'NAN')

		if(sum_ft==0 and sum_bt!=0):
			pasheet.write(row,4,'NAN')
			pasheet.write(row,5,sum_wu/sum_bt)
			pasheet.write(row,6,'NAN')
			pasheet.write(row,7,sum_br/sum_bt)
			pasheet.write(row,8,'NAN')
			pasheet.write(row,9,sum_gs/sum_bt)
			pasheet.write(row,10,'NAN')
			pasheet.write(row,11,sum_wl/sum_bt)
			pasheet.write(row,12,'NAN')
			pasheet.write(row,13,sum_ss/sum_bt)
			pasheet.write(row,14,'NAN')
			pasheet.write(row,15,sum_gu/sum_bt)
		elif(sum_ft!=0 and sum_bt==0):
			pasheet.write(row,4,sum_wu/sum_ft)
			pasheet.write(row,5,'NAN')
			pasheet.write(row,6,sum_br/sum_ft)
			pasheet.write(row,7,'NAN')
			pasheet.write(row,8,sum_gs/sum_ft)
			pasheet.write(row,9,'NAN')
			pasheet.write(row,10,sum_wl/sum_ft)
			pasheet.write(row,11,'NAN')
			pasheet.write(row,12,sum_ss/sum_ft)
			pasheet.write(row,13,'NAN')
			pasheet.write(row,14,sum_gu/sum_ft)
			pasheet.write(row,15,'NAN')
		else:



			x1_s=0
			y1_s=0
			x2_s=0
			y2_s=0
			x3_s=0
			y3_s=0
			x4_s=0
			y4_s=0
			x5_s=0
			y5_s=0
			x6_s=0
			y6_s=0

			for ele in tempdata:
				M=ele[0]
				N=ele[1]
				# if((M-sum_ft*N/sum_bt)==0):
				# 	print M
				# 	print sum_ft
				# 	print N
				# 	print sum_bt
				# 	print "zero"

				x1=(ele[2]-sum_wu*N/sum_bt)/((M-sum_ft*N/sum_bt)+0.001)
				y1=sum_wu-(sum_ft*x1/sum_bt)
				x1_s=x1_s+x1
				y1_s=y1_s+y1

				x2=(ele[3]-sum_br*N/sum_bt)/((M-sum_ft*N/sum_bt)+0.001)
				y2=sum_br-(sum_ft*x2/sum_bt)
				x2_s=x2_s+x2
				y2_s=y2_s+y2

				x3=(ele[4]-sum_gs*N/sum_bt)/((M-sum_ft*N/sum_bt)+0.001)
				y3=sum_gs-(sum_ft*x3/sum_bt)
				x3_s=x3_s+x3
				y3_s=y3_s+y3

				x4=(ele[5]-sum_wl*N/sum_bt)/((M-sum_ft*N/sum_bt)+0.001)
				y4=sum_wl-(sum_ft*x4/sum_bt)
				x4_s=x4_s+x4
				y4_s=y4_s+y4

				x5=(ele[6]-sum_ss*N/sum_bt)/((M-sum_ft*N/sum_bt)+0.001)
				y5=sum_ss-(sum_ft*x5/sum_bt)
				x5_s=x5_s+x5
				y5_s=y5_s+y5

				x6=(ele[7]-sum_gu*N/sum_bt)/((M-sum_ft*N/sum_bt)+0.001)
				y6=sum_gu-(sum_ft*x6/sum_bt)
				x6_s=x6_s+x6
				y6_s=y6_s+y6

			pasheet.write(row,4,x1_s/sum_a)
			pasheet.write(row,5,y1_s/sum_a)
			pasheet.write(row,6,x2_s/sum_a)
			pasheet.write(row,7,y2_s/sum_a)
			pasheet.write(row,8,x3_s/sum_a)
			pasheet.write(row,9,y3_s/sum_a)
			pasheet.write(row,10,x4_s/sum_a)
			pasheet.write(row,11,y4_s/sum_a)
			pasheet.write(row,12,x5_s/sum_a)
			pasheet.write(row,13,y5_s/sum_a)
			pasheet.write(row,14,x6_s/sum_a)
			pasheet.write(row,15,y6_s/sum_a)

		row=row+1
		current_app=table.cell(i,0).value

		sum_ft=0.0
		sum_bt=0.0
		sum_fp=0.0
		sum_bp=0.0
		sum_a=0.0
		sum_wu=0.0
		sum_br=0.0
		sum_gs=0.0
		sum_wl=0.0
		sum_ss=0.0
		sum_gu=0.0
		tempdata=[]
		temp=[]
		

		sum_ft=sum_ft+ float(table.cell(i,3).value)
		sum_bt=sum_bt+ float(table.cell(i,4).value)
		sum_fp=sum_fp+ float(table.cell(i,5).value)
		sum_bp=sum_bp+ float(table.cell(i,6).value)
		sum_a=sum_a+1

		sum_wu=sum_wu+ float(table.cell(i,7).value)
		sum_br=sum_br+ float(table.cell(i,8).value)
		sum_gs=sum_gs+ float(table.cell(i,9).value)
		sum_wl=sum_wl+ float(table.cell(i,10).value)
		sum_ss=sum_ss+ float(table.cell(i,11).value)
		sum_gu=sum_gu+ float(table.cell(i,12).value)
		temp=[float(table.cell(i,3).value),float(table.cell(i,4).value),float(table.cell(i,7).value),float(table.cell(i,8).value),float(table.cell(i,9).value),float(table.cell(i,10).value),float(table.cell(i,11).value),float(table.cell(i,12).value)]
		tempdata.append(temp)

pasheet.write(row,0,current_app)
pasheet.write(row,1,sum_bt/(sum_ft+sum_bt))
pasheet.write(row,2,sum_fp/sum_ft)
pasheet.write(row,3,sum_bp/sum_bt)
pasheet.write(row,4,x1_s/sum_a)
pasheet.write(row,5,y1_s/sum_a)
pasheet.write(row,6,x2_s/sum_a)
pasheet.write(row,7,y2_s/sum_a)
pasheet.write(row,8,x2_s/sum_a)
pasheet.write(row,9,y2_s/sum_a)
pasheet.write(row,10,x2_s/sum_a)
pasheet.write(row,11,y2_s/sum_a)
pasheet.write(row,12,x2_s/sum_a)
pasheet.write(row,13,y2_s/sum_a)
pasheet.write(row,14,x2_s/sum_a)
pasheet.write(row,15,y2_s/sum_a)



pabook.save(outputfile+'2. Power Analysis.xls')
