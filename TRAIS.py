# http://www.python-excel.org/
'''
NPRI ID	0
Organization Name	1
Facility ID	2
Facility Name	3
Relationship Type	4
MOE REG127 Number	5
NAICS	6
Number of Employees	7
Street Address (Physical Address)	8
"Municipality/
City (Physical Address)"	9
Province (Physical Address)	10
PostalZip (Physical Address)	11
Country (Physical Address)	12
Additional Information (Physical Address)	13
UTM Zone	14
UTM Easting	15
UTM Northing	16
Latitude	17
Longitude	18
Public Contact	19
Contact Telephone Number	20
Contact Telephone Extension	21
Contact Fax Number	22
Contact Email	23
Contact Language Correspondence	24
Contact Position	25
Parent Legal Name	26
Parent Percentage Owned	27
Substance Name	28
CAS Number	29
Units	30
Use (Amount Entered Facility)	31
Creation  (Amount Created)	32
Contained In Product (Amount In Product)	33
Report Sum of All Media	34
Stack or Point (Releases to Air)	35
Storage or Handling (Releases to Air)	36
Fugitive (Releases to Air)	37
Spills (Releases to Air)	38
Other Non Point (Releases to Air)	39
Direct Discharges (Releases to Water)	40
Spills (Releases to Water Bodies)	41
Leaks (Releases to Water Bodies)	42
Spills (Releases to Land)	43
Leaks (Releases to Land)	44
Other (Releases to Land)	45
Landfill (Onsite)	46
Land Treatment (Onsite)	47
Underground Injection (Onsite)	48
Landfill (Offsite)	49
Land Treatment (Offsite)	50
Underground Injection (Offsite)	51
Storage (Offsite)	52
Physical (Offsite Treatment)	53
Chemical (Offsite Treatment)	54
Biological (Offsite Treatment)	55
Incineration Thermal (Offsite Treatment)	56
Municipal Sewage Treatment Plant (Offsite Treatment)	57
Tailings Management (Onsite)	58
Waste Rock Management (Onsite)	59
Tailings Management (Offsite)	60
Waste Rock Management (Offsite)	61
Recovery of Energy (Recycling)	62
Recover of Solvents (Recycling)	63
Recovery of Organic Substances (Recycling)	64
Recovery of Metals and Metal Compounds (Recycling)	65
Recovery of Inorganic Materials (Recycling)	66
Recovery of Acids or Bases (Recycling)	67
Recovery of Catalysts (Recycling)	68
Recovery of Pollution Abatement Residue (Recycling)	69
Refining of Reuse of Used Oil (Recycling)	70
Other (Recycling)	71
Highest Ranking Employee	72
Other Sources (VOC) (Releases to Air)	73
'''

class Facility:
	def __init__(self, row):
		self.row = row
		self.toxic = {row[28] : row[28:]}
import xlrd
wb = xlrd.open_workbook('201305_TRAIScurrent.xls')
sh = wb.sheet_by_name(u'Public Data')
dataset = {}
for rownum in range(1, sh.nrows):
	#print (sh.row_values(rownum))
	row = sh.row_values(rownum)
	NPRIID = row[0]
	if (not (NPRIID in dataset)):		
		facility = Facility(row)
		dataset[NPRIID] = facility
	else:
		facility = dataset[NPRIID]
		facility.toxic[row[28]] = row[28:]
for key, value in dataset.iteritems():
	if type(key) is unicode and len(key) == 0:
		continue
	NPRIID = int(key)
	print str(type(NPRIID)) + ", " + str(NPRIID)