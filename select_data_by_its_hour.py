# -*- coding=UTF-8 -*-

import xlrd 
import xlutils.copy
import pylab
import sys

def open_workbook_and_wordsheet(fname, reference_sheet_index):
	rb = xlrd.open_workbook(fname)
	sh_reference = rb.sheet_by_index(reference_sheet_index-1)
	wb = xlutils.copy.copy(rb)
	ws_1 = wb.get_sheet(0)
	ws_2 = wb.get_sheet(1)
	ws_3 = wb.get_sheet(2)
	return sh_reference, wb, ws_1, ws_2, ws_3

def elect_data_by_its_hour(sh_reference, wb, ws_editted, hour, fname):
	rows = 1949
	rows_have_written = 0
	electricity_consumption = []
	temperature = []
	for i in xrange(1, rows-1):
		cell_value = sh_reference.cell_value(i, 0)
		hour_num = int(list(cell_value)[-3] + list(cell_value)[-2])
		if hour_num == hour:
			ws_editted.write(rows_have_written, 0, sh_reference.cell_value(i, 0))
			ws_editted.write(rows_have_written, 1, sh_reference.cell_value(i, 2))
			ws_editted.write(rows_have_written, 2, sh_reference.cell_value(i, 1))
			electricity_consumption.append(float(sh_reference.cell_value(i, 1)))
			temperature.append(float(sh_reference.cell_value(i, 2))/10.0)
			rows_have_written += 1
	wb.save(fname)	
	return electricity_consumption, temperature

def rSquare(estimated, measured):
	"""measured: one dimensional array of measured values
	   estimate: one dimensional array of predicted values"""
	SEE = ((estimated - measured)**2).sum()
	mMean = measured.sum()/float(len(measured))
	MV = ((mMean - measured)**2).sum()
	#SEE2 = []
	#MV2 = []
	#for i in xrange(len(measured)):
	#	print 'estimated - measured' ,(estimated[i] - measured[i])**2
	#	SEE2.append((estimated[i] - measured[i])**2)
	#	MV2.append((mMean - measured[i])**2)
	#print sum(SEE2)/sum(MV2)
	#print 1 - sum(SEE2)/sum(MV2)
	return 1 - SEE/MV

def tryFits(electricity_consumption, temperature, hour):
	electricity = pylab.array(electricity_consumption)
	temp = pylab.array(temperature)
	pylab.title('hour ' + str(hour) )
	pylab.xlabel('temperature')
	pylab.ylabel('electricity_consumption')
	pylab.plot(temp, electricity, 'bo')
	a,b = pylab.polyfit(temp,electricity, 1)
	electricity_sim = a*temp + b
	pylab.plot(temp, electricity_sim, 'r', label = 'Linear Fit' + ', R2 = ' + str(round(rSquare(electricity_sim, electricity), 4)))
	a,b,c = pylab.polyfit(temp,electricity, 2)
	electricity_sim = a*(temp**2) + b*temp + c
	pylab.plot(temp, electricity_sim, 'g', label = 'Quadratic Fit' + ', R2 = ' + str(round(rSquare(electricity_sim ,electricity), 4)) + ' equation: ' +'%.2f x^2 + %.2f x + %.2f' %(a, b, c))
	print temp
	print electricity
	print electricity_sim
	#a,b,c,d = pylab.polyfit(temp,electricity, 3)
	#electricity_sim = a*(temp**3) + b*(temp**2) + c*temp + d
	#pylab.plot(temp, electricity_sim, 'k', label = 'Cubic Fit' + ', R2 = ' + str(round(rSquare(electricity ,electricity_sim), 4)))
	pylab.legend()

def elect_data(sh_reference, wb, ws_editted, fname):
	rows = 163
	electricity_consumption = []
	temperature = []
	for i in xrange(0, rows-1):
			electricity_consumption.append(float(sh_reference.cell_value(i, 1)))
			temperature.append(float(sh_reference.cell_value(i, 2))/10.0)
	wb.save(fname)	
	return electricity_consumption, temperature

if __name__ == '__main__':
	print sys.argv[1]
	hour = int(sys.argv[1])
	sh_reference, wb, ws_1, ws_2, ws_editted = open_workbook_and_wordsheet('/home/tangent13/Documents/9-20.xlsx',2)
	electricity_consumption, temperature = elect_data_by_its_hour(sh_reference, wb, ws_editted, hour, '/home/tangent13/Documents/9-20.xlsx')
	#sh_reference, wb, ws_1, ws_2, ws_editted = open_workbook_and_wordsheet('/home/tangent13/Documents/9-20.xlsx',3)
	#electricity_consumption, temperature = elect_data(sh_reference, wb, ws_editted, '/home/tangent13/Documents/9-20.xlsx')
	#print len(electricity_consumption)
	#print len(temperature)
	tryFits(electricity_consumption, temperature, hour)
	pylab.show()	


