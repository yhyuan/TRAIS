# http://www.python-excel.org/
'''

'''
import sys
reload(sys)
sys.setdefaultencoding("latin-1")
import fileinput
#Create NAICS Dictionary
NAICSDictionary = {}
for line in fileinput.input('NAICS.txt'):
	items = line.strip().split("\t")
	code = (items[0])
	name = (items[1])[1:-1]
	NAICSDictionary[code] = name

FieldIndexDict = {}
for line in fileinput.input('FieldIndex.txt'):
	items = line.strip().split("\t")
	name = (items[0])
	index = (items[1])
	FieldIndexDict[name] = index

#Create substance Dictionary French
substanceDictionary = {}
newsubstanceDictionary = {}
i = 0
for line in fileinput.input('substance_codes.txt'):
	i = i + 1
	if i == 1:
		continue  # skip the first line
	items = line.strip().split("\t")
	code = (items[1])[1:-1]
	substanceDictionary[code] = items  # CODE	SUBSTANCE_EN	SUBSTANCE_FR	CAS
	
class Substance:
	def __init__(self, row):
		self.row = row
	def isEmpty(self):
		return len(self.row[28]) == 0
	def parse(self, item):
		if (type(item) is unicode or type(item) is str) and len(item) == 0:
			return 0
		elif type(item) is unicode or type(item) is str :
			return float(item)
		else:
			return item
	def __str__(self):
		if len(self.row[28]) == 0:
			return ""	
		result = "{\n"
		result = result + "\t\t\t\tName: \"" + self.row[28] + "\","
		#result = result + "\t\t\t\tCode: \"" + self.getCode() + "\","
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
		result = result + "\n\t\t\t\tDOffSite: " + str(disposalOffSite) + ","
		recycleOffSite = self.parse(self.row[62]) + self.parse(self.row[63]) + self.parse(self.row[64]) + self.parse(self.row[65]) + self.parse( self.row[66]) + self.parse(self.row[67]) + self.parse(self.row[68] ) + self.parse(self.row[69]) + self.parse(self.row[70]) + self.parse(self.row[71])   # BK, BL, BM, BN, BO, BP, BQ, BR, BS, BT
		result = result + "\n\t\t\t\tROffSite: " + str(recycleOffSite)
		result = result + "\n\t\t\t}"
		return result
	def getCode(self):
		if len(self.row[28]) == 0:
			return ""
		# Find the code with substance name
		if self.row[28] in substanceDictionary:
			return substanceDictionary[self.row[28]][0][1:-1]
		else:
			# If failed, try to find the code with CAS Number
			CASNumber = self.row[29]
			for key, value in substanceDictionary.iteritems():
				if (CASNumber == value[3][1:-1]):
					return substanceDictionary[key][0][1:-1]
			# if failed again, try to add it to the dictionary. 
			substanceDictionary[self.row[28]] = ["\"S" + str(len(substanceDictionary) + 1) + "\"", "\"" + self.row[28] + "\"", "\"" + self.row[28] + "\"", "\"" + self.row[29] + "\""]
			newsubstanceDictionary[self.row[28]] = ["\"S" + str(len(substanceDictionary)) + "\"", "\"" + self.row[28] + "\"", "\"" + self.row[28] + "\"", "\"" + self.row[29] + "\""]
			return substanceDictionary[self.row[28]][0][1:-1]
	def getFeatureClassString(self):
		if len(self.row[28]) == 0:
			return ""
		#result = "{Name:\"" + self.row[28] + "\","
		result = "{A:" + str(int(self.getCode()[1:])) + ","  #Name
		if len(self.row[30]) > 0:
			#result = result + "Units:\"" + self.row[30] + "\"," 
			result = result + "B:" + str(UnitsDict[self.row[30]]) + ","  # Units
		if len(str(self.row[31])) > 0:
			if str(self.row[31]) in RangeDict:
				result = result + "C:\"I" + str(RangeDict[str(self.row[31])]) + "\","  # Used
			else:
				result = result + "C:\"" + str(self.row[31]) + "\","  # Used
		if len(str(self.row[32])) > 0:
			if str(self.row[32]) in RangeDict:
				result = result + "D:\"I" + str(RangeDict[str(self.row[32])]) + "\","  # Created
			else:		
				result = result + "D:\"" + str(self.row[32]) + "\","  # Created
		if len(str(self.row[33])) > 0:
			if str(self.row[33]) in RangeDict:
				result = result + "E:\"I" + str(RangeDict[str(self.row[33])]) + "\","  # Contained
			else:				
				result = result + "E:\"" + str(self.row[33]) + "\"," # Contained
		air = self.parse(self.row[35]) + self.parse(self.row[36]) + self.parse(self.row[37]) + self.parse(self.row[38])  + self.parse(self.row[39]) + self.parse(self.row[73])  #AJ, AK, AL, AM, AN, BV
		if air > 0:
			result = result + "F:" + str(air) + ","  # Air 
		water = self.parse(self.row[40]) + self.parse(self.row[41]) + self.parse(self.row[42])  # AO, AP, AQ
		if water > 0:
			result = result + "G:" + str(water) + "," # Water
		land = self.parse(self.row[43]) + self.parse(self.row[44]) + self.parse(self.row[45])   # AR, AS, AT
		if land > 0:
			result = result + "H:" + str(land) + "," #Land
		disposalOnSite = self.parse(self.row[46]) + self.parse(self.row[47]) + self.parse(self.row[48])   # AU, AV, AW
		if disposalOnSite > 0:
			result = result + "I:" + str(disposalOnSite) + ","  # DonSite
		disposalOffSite = self.parse(self.row[49]) + self.parse(self.row[50]) + self.parse(self.row[51]) + self.parse(self.row[52]) + self.parse(self.row[53]) + self.parse(self.row[54]) + self.parse(self.row[55])  + self.parse(self.row[56]) + self.parse(self.row[57])   # AX, AY, AZ, BA, BB, BC, BD, BE, BF
		if disposalOffSite > 0:
			result = result + "J:" + str(disposalOffSite) + "," # DoffSite
		recycleOffSite = self.parse(self.row[62]) + self.parse(self.row[63]) + self.parse(self.row[64]) + self.parse(self.row[65]) + self.parse( self.row[66]) + self.parse(self.row[67]) + self.parse(self.row[68] ) + self.parse(self.row[69]) + self.parse(self.row[70]) + self.parse(self.row[71])   # BK, BL, BM, BN, BO, BP, BQ, BR, BS, BT
		if recycleOffSite > 0:
			result = result + "I:" + str(recycleOffSite)  # ROffSite
		if result[-1] == ",":
			result = result[:-1]
		result = result + "}"
		return result
class Facility:
	def __init__(self, row, id):
		self.row = row
		self.id = id
		self.substances = [Substance(row)]
		
		self.FacilityName =  self.row[FieldIndexDict["Facility Name"]]
		self.OrganizationName =  self.row[FieldIndexDict["Organization Name"]]
		self.Address =  self.row[FieldIndexDict["Street Address (Physical Address)"]] + " / " + self.row[FieldIndexDict["Municipality/City (Physical Address)"]]
		self.Sector =  str(self.row[FieldIndexDict["NAICS"]]) + " - " + NAICSDictionary[str(int(self.row[FieldIndexDict["NAICS"]]))]
		self.NPRIID =  str(int(self.row[FieldIndexDict["NPRI ID"]]))
		self.PublicContact =  self.row[FieldIndexDict["Public Contact"]]
		self.ContactTelephoneNumber =  self.formatPhone(self.row[FieldIndexDict["Contact Telephone Number"]])
		self.ContactEmail =  self.row[FieldIndexDict["Contact Email"]]
		self.HighestRankingEmployee =  self.row[FieldIndexDict["Highest Ranking Employee"]]
		self.Latitude = str(self.row[FieldIndexDict["Latitude"]])
		self.Longitude = str(self.row[FieldIndexDict["Longitude"]])
		self.FacilityID =  str(int(self.row[FieldIndexDict["Facility ID"]]))
		self.NAICS = str(int(self.row[FieldIndexDict["NAICS"]]))
	def formatPhone(self, phoneInnput):
		if len(str(phoneInnput)) == 0:
			return ""
		elif type(phoneInnput) is float:
			return str(int(phoneInnput))
		return phoneInnput
	def getSubstancesString(self):
		def getString(x): return str(x)
		substancesStringList = map(getString, self.substances)
		return ",".join(map(getString, self.substances))
	def getSubstancesCodeList(self):
		def getCode(substance): return substance.getCode()
		return "_".join(map(getCode, self.substances)) + "_"
	def getSubstancesNumber(self):
		substances_number = len(self.substances)
		if substances_number == 1 and self.substances[0].isEmpty():
			substances_number = 0
		return substances_number
	def __str__(self):
		result = "var info = {\n" + "\t\t\tFacilityName: \"" + self.FacilityName + "\"," + "\n\t\t\tCompanyName: \"" + self.OrganizationName + "\"," + "\n\t\t\tAddress: \"" + self.Address + "\"," + "\n\t\t\tSector: \"" + self.Sector + "\"," + "\n\t\t\tNPRIID: \"" + self.NPRIID + "\"," + "\n\t\t\tPublicContact: \"" + self.PublicContact + "\"," + "\n\t\t\tPublicContactPhone: \"" + self.ContactTelephoneNumber + "\"," + "\n\t\t\tPublicContactEmail: \"" + self.ContactEmail + "\"," + "\n\t\t\tHighestRankingEmployee: \"" + self.HighestRankingEmployee + "\","
		substanceResult = self.getSubstancesString()
		if len(substanceResult) > 0:
			result = result + "\n\t\t\tSubstances: [" + substanceResult[:-1] + "]";
		else:
			result = result[:-1]
		result = result + "\n\t\t};"
		return result

	def getFeatureClassString(self):
		return self.Latitude + "\t" + self.Longitude  + "\t" + self.FacilityID + "\t\"" +  self.FacilityName + "\"\t\"" + self.Address + "\"\t\"" + self.OrganizationName + "\"\t" +  self.NPRIID + "\t" + self.NAICS + "\t\"" + NAICSDictionary[self.NAICS] + "\"\t" + str(self.getSubstancesNumber() ) + "\t\""	+ self.getSubstancesCodeList() + "\"" + "\t\"" + self.PublicContact + "\"" + "\t\"" + self.ContactTelephoneNumber + "\"\t\"" + self.ContactEmail + "\"\t\"" + self.HighestRankingEmployee + "\"\t"
		# Substances
		#substanceResult = ""
		#for substance in self.substances:
		#	substanceString = substance.getFeatureClassString()
		#	if len(substanceString) > 0:
		#		substanceResult = substanceResult + substanceString + ","
		#if len(substanceResult) > 0:
		#	result = result + "[" + substanceResult[:-1] + "]"
		#result = result[:-1];		
		#return result

	def getFirstLetterCompanyName(self):
		return self.OrganizationName[0]
	def getODAstr(self):
		return "{CompanyName:\"" + self.row[1] + "\",FacilityName:\"" + self.row[3] + "\"," + "NPRIID:\"" + str(int(self.row[0])) + "\"," + "City:\"" + self.row[9] + "\"," + "Substances:" + str(len(self.substances)) + "}" 

	


#Read Excel File	
import xlrd
wb = xlrd.open_workbook('201305_TRAIScurrent.xls')
sh = wb.sheet_by_name(u'Public Data')
dataset = {}
UnitsDict = {}
RangeDict = {}

for rownum in range(1, sh.nrows):
	#print (sh.row_values(rownum))
	row = sh.row_values(rownum)
	NPRIID = row[0]
	if (not (NPRIID in dataset)):		
		facility = Facility(row, len(dataset) + 1)
		dataset[NPRIID] = facility
	else:
		facility = dataset[NPRIID]
		facility.substances.append(Substance(row))
	Units = row[30]
	if (not (Units in UnitsDict)):		
		UnitsDict[Units] = len(UnitsDict)
	Used = str(row[31])
	if (len(Used) > 1) and (Used[0] == ">") and (not (Used in RangeDict)):		
		RangeDict[Used] = len(RangeDict)
	Created = str(row[32])
	if (len(Created) > 1) and (Created[0] == ">") and (not (Created in RangeDict)):		
		RangeDict[Created] = len(RangeDict)
	Contained = str(row[33])
	if (len(Contained) > 1) and (Contained[0] == ">") and (not (Contained in RangeDict)):		
		RangeDict[Contained] = len(RangeDict)
	
#Generate Reports
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

#Generate Data for ODA version
ODADict = {}
for key, facility in dataset.iteritems():
	if type(key) is unicode and len(key) == 0:
		continue
	NPRIID = int(key)
	firstLetter = facility.getFirstLetterCompanyName()
	if (not (firstLetter in ODADict)):		
		ODADict[firstLetter] = [facility]
	else:
		facilityList = ODADict[firstLetter]
		facilityList.append(facility)

result = "var facilityDict = [\n"
for firstLetter in sorted(ODADict):
	facilityList  = ODADict[firstLetter]
	facilityList = sorted(facilityList, key= lambda fac: (fac.row[1], fac.row[3], fac.row[9]))   # sort by Company Name - Facility Name
	result = result + "\t\t\t{index: \"" + firstLetter + "\",\n"
	result = result + "\t\t\tfacilityList: ["
	for facility in facilityList:
		result = result + facility.getODAstr() + ","
	result = result[:-1] + "]},\n"
result = result[:-2] + "];\n"

languages = ["EN", "FR"]
for lang in languages:
	text_file = open("ODAtemplate_" + lang + ".html", "r")
	template = text_file.read()
	text_file.close()
	template = template.replace("${TRAIS_DATA}", result)	
	handle = open("json/" + lang + "/oda.html",'w+')
	handle.write(template)
	handle.close();
	
# Generate txt file for feature class
result = "FacilGeogrLatitude\tFacilGeogrLongitude\tFacilityID\tFacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tSectorDesc\tNUMsubst\tSubstance_List\tContact\tPhone\tEmail\tHREmploy\tSubs2010\n"
for key, facility in dataset.iteritems():
	if type(key) is unicode and len(key) == 0:
		continue
	#print str(facility)
	result = result + facility.getFeatureClassString() + "\n"
handle = open("json/TRAIS.txt",'w+')
handle.write(result)
handle.close();

result = "{\n\t\"type\": \"FeatureCollection\",\n\t\"features\": ["
for key, facility in dataset.iteritems():
	result = result + facility.getGeoJson() + ",\n"
result = result[:-2] + "\n\t]\n}"

handle = open("TRAIS.json",'w+')
handle.write(result)
handle.close();

for key, substance in newsubstanceDictionary.iteritems():
	print substance[0] + "\t" + substance[1] + "\t" + substance[2] + "\t" + substance[3]

for Unit, index in UnitsDict.iteritems():
	print Unit + "\t" + str(index)

for Range, index in RangeDict.iteritems():
	print Range + "\t" + str(index)

	