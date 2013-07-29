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
"Municipality/City (Physical Address)"	9
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
import sys
reload(sys)
sys.setdefaultencoding("latin-1")

class Substance:
	def __init__(self, row):
		self.row = row
	def parse(self, item):
		#print type(item)
		if (type(item) is unicode or type(item) is str) and len(item) == 0:
			return 0
		elif type(item) is unicode or type(item) is str :
			return float(item)
		else:
			return item
	def __str__(self):
		result = "{\n"
		result = result + "\t\t\t\tName: \"" + self.row[28] + "\","
		result = result + "\n\t\t\t\tUnits: \"" + self.row[30] + "\","
		result = result + "\n\t\t\t\tUsed: \"" + str(self.row[31]) + "\","
		result = result + "\n\t\t\t\tCreated: \"" + str(self.row[32]) + "\","
		result = result + "\n\t\t\t\tContained: \"" + str(self.row[33]) + "\","
		air = self.parse(self.row[35]) + self.parse(self.row[36]) + self.parse(self.row[37]) + self.parse(self.row[38])  + self.parse(self.row[39]) + self.parse(self.row[73])  #AJ, AK, AL, AM, AN, BV
		result = result + "\n\t\t\t\tAir: " + str(air) + ","
		water = self.parse(self.row[40]) + self.parse(self.row[41]) + self.parse(self.row[42])  # AO, AP, AQ
		result = result + "\n\t\t\t\tWater: " + str(water) + ","
		land = self.parse(self.row[43]) + self.parse(self.row[44]) + self.parse(self.row[45])   # AR, AS, AT
		result = result + "\n\t\t\t\tLand: " + str(land) + ","
		disposalOnSite = self.parse(self.row[46]) + self.parse(self.row[47]) + self.parse(self.row[48])   # AU, AV, AW
		result = result + "\n\t\t\t\tDOnSite: " + str(disposalOnSite) + ","
		disposalOffSite = self.parse(self.row[49]) + self.parse(self.row[50]) + self.parse(self.row[51]) + self.parse(self.row[52]) + self.parse(self.row[53]) + self.parse(self.row[54]) + self.parse(self.row[55])  + self.parse(self.row[56]) + self.parse(self.row[57])   # AX, AY, AZ, BA, BB, BC, BD, BE, BF
		#print disposalOffSite
		result = result + "\n\t\t\t\tDOffSite: " + str(disposalOffSite) + ","
		recycleOffSite = self.parse(self.row[62]) + self.parse(self.row[63]) + self.parse(self.row[64]) + self.parse(self.row[65]) + self.parse( self.row[66]) + self.parse(self.row[67]) + self.parse(self.row[68] ) + self.parse(self.row[69]) + self.parse(self.row[70]) + self.parse(self.row[71])   # BK, BL, BM, BN, BO, BP, BQ, BR, BS, BT
		result = result + "\n\t\t\t\tROffSite: " + str(recycleOffSite)
		result = result + "\n\t\t\t}"
		return result
class Facility:
	def __init__(self, row):
		self.row = row
		self.substances = [Substance(row)]
	def __str__(self):
		result = "var info = {\n"
		result = result + "\t\t\tFacilityName: \"" + self.row[3] + "\","
		result = result + "\n\t\t\tCompanyName: \"" + self.row[1] + "\","
		result = result + "\n\t\t\tAddress: \"" + self.row[8] + " / " + self.row[9] + "\","
		result = result + "\n\t\t\tSector: \"" + str(int(self.row[6])) + " - " + NAICSDictionary[str(int(self.row[6]))] + "\","
		result = result + "\n\t\t\tNPRIID: \"" + str(int(self.row[0])) + "\","
		result = result + "\n\t\t\tPublicContact: \"" + (self.row[19]) + "\","
		if len(str(self.row[20])) == 0:
			phone = ""
		elif type(self.row[20]) is float:
			phone = str(int(self.row[20]))
		else:
			phone = self.row[20]
		result = result + "\n\t\t\tPublicContactPhone: \"" + phone + "\","
		result = result + "\n\t\t\tPublicContactEmail: \"" + self.row[23] + "\","
		result = result + "\n\t\t\tHighestRankingEmployee: \"" + self.row[72] + "\","
		result = result + "\n\t\t\tSubstances: ["
		for substance in self.substances:
			result = result + str(substance) + ","
		result = result[:-1] + "]\n\t\t};";
		return result

NAICSDictionary = {}
import fileinput
for line in fileinput.input('NAICS.txt'):
	items = line.strip().split("\t")
	code = (items[0])
	name = (items[1])[1:-1]
	NAICSDictionary[code] = name
	
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
		facility.substances.append(Substance(row))
for key, value in dataset.iteritems():
	if type(key) is unicode and len(key) == 0:
		continue
	NPRIID = int(key)
	data = str(value)
	languages = ["EN", "FR"]
	for lang in languages:
		text_file = open("template_" + lang + ".html", "r")
		template = text_file.read()
		text_file.close()
		template = template.replace("${TRAIS_DATA}", data)	
		handle = open("json/" + lang + "/annual" + str(NPRIID) + ".html",'w+')
		handle.write(template)
		handle.close();
	