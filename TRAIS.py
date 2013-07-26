# http://www.python-excel.org/
import xlrd
wb = xlrd.open_workbook('201305_TRAIScurrent.xls')
sh = wb.sheet_by_name(u'Public Data')
dataset = {}
for rownum in range(1, sh.nrows):
	#print (sh.row_values(rownum))
	row = sh.row_values(rownum)
	NPRIID = row[0]
	if (not (NPRIID in dataset)):
		dataset[NPRIID] = NPRIID
	else:
		dataset[NPRIID]["Toxic"]
print len(dataset)