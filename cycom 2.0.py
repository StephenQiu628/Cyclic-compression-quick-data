#Coder: Stephen Qiu
#Date: 2022.7.13~2022.7.17
#This program works with instron exported data files, the format of which can be found in the sample provided on github. Users can modify this code or adjust the instron exported data file to suit their situation.

import xlrd
import xlwt
import numpy as np

#Determin the baseline (Divide-and-Conquer)
def Min(n):
	if len(n) == 1 :
		return n[0]
	if len(n) == 2:
		if n[0] <= n[1]:
			return n[0]
		return n[1]
	middle = int(len(n)/2)
	m1 = n[:middle]
	m2 = n[middle:]
	x1 = Min(m1)
	x2 = Min(m2)
	if x1 <= x2:
		return x1
	return x2

#Determin the maximum stress (Divide-and-Conquer)
def Max(n):
	if len(n) == 1 :
		return n[0]
	if len(n) == 2:
		if n[0] >= n[1]:
			return n[0]
		return n[1]
	middle = int(len(n)/2)
	m1 = n[:middle]
	m2 = n[middle:]
	x1 = Max(m1)
	x2 = Max(m2)
	if x1 >= x2:
		return x1
	return x2

#Create a new result table
xlsbook = xlwt.Workbook(encoding='utf-8')
sheet = xlsbook.add_sheet('Results1')
sheet.write(0,0,'循环次数')
sheet.write(0,1,'能量回复率/%')
sheet.write(0,2,'能量损耗系数/%')
sheet.write(0,3,'弹性模量/MPa')
sheet.write(0,4,'压缩应力/MPa')
sheet.write(0,5,'最大应力/MPa')
xlsbook.save('Results1.xls')
#循环次数 means the number of cycle.
# 能量回复率 means the energy return rate.
# 能量损耗系数 means the energy loss coefficient.
# 弹性模量 means the elastic modulus.
# 压缩应力 means the compressive stress (20%) .
# 最大应力 means the maximum stress (50%) .

#Repeat the cycle from 1 to 20
time = 0
while time <= 19:
	time = time + 1
	
	#Read the file
	rawfile = xlrd.open_workbook('./Specimen_RawData_1.xls')
	table = rawfile.sheet_by_name("Specimen_RawData_1")
	nrows = table.nrows
	ncols = table.ncols
	
	#Add a table header
	workbook = xlwt.Workbook(encoding='utf-8')
	new_sheet = workbook.add_sheet('Chosen Data')

	cycletime = time - 1

	rank_list = []
	for i in range (1,nrows):
		if table.row_values(i)[5] == cycletime: 
	#Change the number in [] for the number of the column that you want to choose
			rank_list.append(i)
	#print(rank_list)

	#Write the header
	for i in range(3, 5):
	#for i in range(ncols):
		new_sheet.write(0,i-3,table.cell(0,i).value)
		#new_sheet.write(0,i,table.cell(0,i).value)
	
	for i in range(len(rank_list)):
		for j in range(3, 5):
		#for j in range(ncols):
			new_sheet.write(i+1, j-3, table.cell(rank_list[i],j).value)
			#new_sheet.write(i+1, j, table.cell(rank_list[i],j).value)
	workbook.save('Chosen Data.xls')

	#Read the chosen data
	rawfile = xlrd.open_workbook('./Chosen Data.xls')
	table = rawfile.sheet_by_name("Chosen Data")
	nrows1 = table.nrows
	ncols1 = table.ncols
	strain_array = np.zeros(nrows1-1)
	stress_array = np.zeros(nrows1-1)
	#print(strain_array)

	for i in range(1, nrows1):
		strain_array[i-1] = table.row_values(i)[0]
		stress_array[i-1] = table.row_values(i)[1]
	#print(strain_array,stress_array)
	
	#Trapezoidal approximation integral
	totals1 = 0
	for i in range(len(strain_array)-1):
		deltas = (stress_array[i] + stress_array[i+1] - 2 * Min(stress_array)) * np.abs((strain_array[i+1]-strain_array[i])) / 2
		totals1 =  totals1 + deltas 
	#print(totals1)

	#To find the modulus of elasticity, take the slope between 6% and 7% of strain, and take the middle value
	t = 0
	s = 0
	m = 0
	while strain_array[t] <= 4:
		t = t + 1
		s = t
	while strain_array[t] <= 6:
		t = t + 1
		m = t
	modulus_one = (stress_array[m]-stress_array[s])/(strain_array[m]-strain_array[s])
	modulus_two = (stress_array[m+1]-stress_array[s+1])/(strain_array[m+1]-strain_array[s+1])
	modulus_three = (stress_array[m+2]-stress_array[s+2])/(strain_array[m+2]-strain_array[s+2])
	modulus_four = (stress_array[m+3]-stress_array[s+3])/(strain_array[m+3]-strain_array[s+3])
	modulus_five = (stress_array[m+4]-stress_array[s+4])/(strain_array[m+4]-strain_array[s+4])
	modulus = 100 * np.median([modulus_one,modulus_two,modulus_three,modulus_four,modulus_five])

	#Find the stress at 10% strain and the corresponding stress when the strain is just greater than 10%
	t = 0
	while strain_array[t] <= 10:
		t = t + 1 
		ten_stress = stress_array[t-1]

	#Find the stress at 50% strain (maximum stress)
	max_stress_1 = Max(stress_array)

	#Calculate the integral area of deload stage
	rawfile = xlrd.open_workbook('./Specimen_RawData_1.xls')
	table = rawfile.sheet_by_name("Specimen_RawData_1")
	nrows = table.nrows
	ncols = table.ncols
	workbook = xlwt.Workbook(encoding='utf-8')
	new_sheet = workbook.add_sheet('Chosen Data')
	cycletime = cycletime + 0.5

	rank_list = []
	for i in range (2,nrows):
		if table.row_values(i)[5] == cycletime : 
			rank_list.append(i)

	for i in range(3, 5):
		new_sheet.write(0,i-3,table.cell(0,i).value)
	
	for i in range(len(rank_list)):
		for j in range(3, 5):
			new_sheet.write(i+1, j-3, table.cell(rank_list[i],j).value)

	workbook.save('Chosen Data.xls')
	rawfile = xlrd.open_workbook('./Chosen Data.xls')
	table = rawfile.sheet_by_name("Chosen Data")
	nrows1 = table.nrows
	ncols1 = table.ncols
	strain_array = np.zeros(nrows1 - 1)
	stress_array = np.zeros(nrows1 - 1)

	for i in range(1, nrows1):
		strain_array[i-1] = table.row_values(i)[0]
		stress_array[i-1] = table.row_values(i)[1]

	totals2 = 0
	for i in range(len(strain_array)-1):
		deltas = (stress_array[i] + stress_array[i+1] - 2 * Min(stress_array)) * np.abs((strain_array[i+1]-strain_array[i])) / 2
		totals2 =  totals2 + deltas 
	Resilience = np.round(totals2 / totals1 * 100, 1)
	Energy_loss_coefficient = 100 - Resilience

	#Find the stress at 50% strain (maximum stress)
	max_stress_2 = Max(stress_array)
	
	max_stress = Max([max_stress_1, max_stress_2])

	#Import the energy return rate into the results table
	new_file = xlrd.open_workbook('./Results1.xls')
	sheet1 = new_file.sheet_by_name("Results1")
	sheet.write(time, 0, time)
	sheet.write(time, 1, Resilience)
	sheet.write(time, 2, Energy_loss_coefficient)
	sheet.write(time, 3, modulus)
	sheet.write(time, 4, ten_stress)
	sheet.write(time, 5, max_stress)
	xlsbook.save('Results1.xls')
