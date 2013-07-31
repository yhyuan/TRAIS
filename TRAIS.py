# http://www.python-excel.org/
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
	index = int(items[1])
	FieldIndexDict[name] = index

#Create substance Dictionary French
substanceDictionary = {}
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
		self.SubstanceName =  row[FieldIndexDict["Substance Name"]]
		self.CASNumber =  row[FieldIndexDict["CAS Number"]]
		self.Units =  row[FieldIndexDict["Units"]]
		self.Use =  str(row[FieldIndexDict["Use (Amount Entered Facility)"]])
		self.Creation =  str(row[FieldIndexDict["Creation  (Amount Created)"]])
		self.Contained =  str(row[FieldIndexDict["Contained In Product (Amount In Product)"]])
		air = self.parse(row[FieldIndexDict["Stack or Point (Releases to Air)"]]) + self.parse(row[FieldIndexDict["Storage or Handling (Releases to Air)"]]) + self.parse(row[FieldIndexDict["Fugitive (Releases to Air)"]]) + self.parse(row[FieldIndexDict["Spills (Releases to Air)"]])  + self.parse(row[FieldIndexDict["Other Non Point (Releases to Air)"]]) + self.parse(row[FieldIndexDict["Other Sources (VOC) (Releases to Air)"]])  #AJ, AK, AL, AM, AN, BV
		self.ReleasestoAir =  str(air)
		water = self.parse(row[FieldIndexDict["Direct Discharges (Releases to Water)"]]) + self.parse(row[FieldIndexDict["Spills (Releases to Water Bodies)"]]) + self.parse(row[FieldIndexDict["Leaks (Releases to Water Bodies)"]])  # AO, AP, AQ		
		self.ReleasestoWater = str(water)
		land = self.parse(row[FieldIndexDict["Spills (Releases to Land)"]]) + self.parse(row[FieldIndexDict["Leaks (Releases to Land)"]]) + self.parse(row[FieldIndexDict["Other (Releases to Land)"]])   # AR, AS, AT		
		self.ReleasestoLand =  str(land)
		disposalOnSite = self.parse(row[FieldIndexDict["Landfill (Onsite)"]]) + self.parse(row[FieldIndexDict["Land Treatment (Onsite)"]]) + self.parse(row[FieldIndexDict["Underground Injection (Onsite)"]])   # AU, AV, AW
		self.DisposalOnSite = str(disposalOnSite)
		disposalOffSite = self.parse(row[FieldIndexDict["Landfill (Offsite)"]]) + self.parse(row[FieldIndexDict["Land Treatment (Offsite)"]]) + self.parse(row[FieldIndexDict["Underground Injection (Offsite)"]]) + self.parse(row[FieldIndexDict["Storage (Offsite)"]]) + self.parse(row[FieldIndexDict["Physical (Offsite Treatment)"]]) + self.parse(row[FieldIndexDict["Chemical (Offsite Treatment)"]]) + self.parse(row[FieldIndexDict["Biological (Offsite Treatment)"]])  + self.parse(row[FieldIndexDict["Incineration Thermal (Offsite Treatment)"]]) + self.parse(row[FieldIndexDict["Municipal Sewage Treatment Plant (Offsite Treatment)"]])   # AX, AY, AZ, BA, BB, BC, BD, BE, BF
		self.DisposalOffSite = str(disposalOffSite)		
		recycleOffSite = self.parse(row[FieldIndexDict["Recovery of Energy (Recycling)"]]) + self.parse(row[FieldIndexDict["Recover of Solvents (Recycling)"]]) + self.parse(row[FieldIndexDict["Recovery of Organic Substances (Recycling)"]]) + self.parse(row[FieldIndexDict["Recovery of Metals and Metal Compounds (Recycling)"]]) + self.parse( row[FieldIndexDict["Recovery of Inorganic Materials (Recycling)"]]) + self.parse(row[FieldIndexDict["Recovery of Acids or Bases (Recycling)"]]) + self.parse(row[FieldIndexDict["Recovery of Catalysts (Recycling)"]] ) + self.parse(row[FieldIndexDict["Recovery of Pollution Abatement Residue (Recycling)"]]) + self.parse(row[FieldIndexDict["Refining of Reuse of Used Oil (Recycling)"]]) + self.parse(row[FieldIndexDict["Other (Recycling)"]])   # BK, BL, BM, BN, BO, BP, BQ, BR, BS, BT
		self.RecycleOffSite = str(recycleOffSite)			
	def isEmpty(self):
		return len(self.SubstanceName) == 0
	def parse(self, item):
		if (type(item) is unicode or type(item) is str) and len(item) == 0:
			return 0
		elif type(item) is unicode or type(item) is str :
			return float(item)
		else:
			return item
	def __str__(self):
		if self.isEmpty():
			return ""	
		return "{\n" + "\t\t\t\tName: \"" + self.SubstanceName + "\"," + "\n\t\t\t\tUnits: \"" + self.Units + "\"," + "\n\t\t\t\tUsed: \"" + self.Use + "\"," + "\n\t\t\t\tCreated: \"" + self.Creation + "\"," + "\n\t\t\t\tContained: \"" + self.Contained + "\"," + "\n\t\t\t\tAir: " + self.ReleasestoAir + "," + "\n\t\t\t\tWater: " + self.ReleasestoWater + "," + "\n\t\t\t\tLand: " + self.ReleasestoLand + "," + "\n\t\t\t\tDOnSite: " + self.DisposalOnSite + "," + "\n\t\t\t\tDOffSite: " + self.DisposalOffSite + "," + "\n\t\t\t\tROffSite: " + self.RecycleOffSite + "\n\t\t\t}"
	def getCode(self):
		if self.isEmpty():
			return ""
		# Find the code with substance name
		if self.SubstanceName in substanceDictionary:
			return substanceDictionary[self.SubstanceName][0][1:-1]
		else:
			# If failed, try to find the code with CAS Number
			for key, value in substanceDictionary.iteritems():
				if (self.CASNumber == value[3][1:-1]):
					return substanceDictionary[key][0][1:-1]
			# if failed again, try to add it to the dictionary. 
			substanceDictionary[self.SubstanceName] = ["\"S" + str(len(substanceDictionary) + 1) + "\"", "\"" + self.SubstanceName + "\"", "\"" + self.SubstanceName + "\"", "\"" + self.CASNumber + "\""]
			return substanceDictionary[self.SubstanceName][0][1:-1]
	def getFeatureClassString(self):
		if self.isEmpty():
			return "\t\t\t\t\t\t\t\t\t\t"
		return "\"" + self.SubstanceName + "\"\t\"" + self.Units + "\"\t\"" + self.Use + "\"\t\"" + self.Creation + "\"\t\"" + self.Contained + "\"\t\"" + self.ReleasestoAir + "\"\t\"" + self.ReleasestoWater + "\"\t\"" + self.ReleasestoLand + "\"\t\"" + self.DisposalOnSite + "\"\t\"" + self.DisposalOffSite + "\"\t\"" + self.RecycleOffSite + "\""
class Facility:
	def __init__(self, row, id):
		self.id = id
		self.substances = [Substance(row)]
		
		self.FacilityName =  row[FieldIndexDict["Facility Name"]]
		self.OrganizationName =  row[FieldIndexDict["Organization Name"]]
		self.Address =  (row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]]
		self.City = row[FieldIndexDict["Municipality/City (Physical Address)"]]
		self.Sector =  str(row[FieldIndexDict["NAICS"]]) + " - " + NAICSDictionary[str(int(row[FieldIndexDict["NAICS"]]))]
		NPRIID = row[FieldIndexDict["NPRI ID"]]
		if (len(str(NPRIID)) > 0):
			self.NPRIID = str(int(NPRIID))
		else:
			self.NPRIID = ""
		self.PublicContact =  row[FieldIndexDict["Public Contact"]]
		self.ContactTelephoneNumber =  self.formatPhone(row[FieldIndexDict["Contact Telephone Number"]])
		self.ContactEmail =  row[FieldIndexDict["Contact Email"]]
		self.HighestRankingEmployee =  row[FieldIndexDict["Highest Ranking Employee"]]
		self.Latitude = str(row[FieldIndexDict["Latitude"]])
		self.Longitude = str(row[FieldIndexDict["Longitude"]])
		FacilityID = row[FieldIndexDict["Facility ID"]]
		if (len(str(NPRIID)) > 0):
			self.FacilityID = str(int(FacilityID))
		else:
			self.FacilityID = ""
		#self.FacilityID =  str(int(row[FieldIndexDict["Facility ID"]]))
		self.NAICS = str(int(row[FieldIndexDict["NAICS"]]))
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
			result = result + "\n\t\t\tSubstances: [" + substanceResult + "]";
		else:
			result = result[:-1]
		result = result + "\n\t\t};"
		return result

	def getFeatureClassString(self):
		return self.Latitude + "\t" + self.Longitude  + "\t" + self.FacilityID + "\t\"" +  self.FacilityName + "\"\t\"" + self.Address + "\"\t\"" + self.OrganizationName + "\"\t" +  self.NPRIID + "\t" + self.NAICS + "\t\"" + NAICSDictionary[self.NAICS] + "\"\t" + str(self.getSubstancesNumber() ) + "\t\""	+ self.getSubstancesCodeList() + "\"" + "\t\"" + self.PublicContact + "\"" + "\t\"" + self.ContactTelephoneNumber + "\"\t\"" + self.ContactEmail + "\"\t\"" + self.HighestRankingEmployee + "\"\t"
		
	def getReportString(self):
		return "\"" +  self.FacilityName + "\"\t\"" + self.Address + "\"\t\"" + self.OrganizationName + "\"\t" +  self.NPRIID + "\t" + self.Sector + "\t\"" + self.PublicContact + "\"\t\"" + self.ContactTelephoneNumber + "\"\t\"" + self.ContactEmail + "\"\t\"" + self.HighestRankingEmployee + "\"\t"

	def getFirstLetterCompanyName(self):
		return self.OrganizationName[0]

	def getODAstr(self):
		return "{CompanyName:\"" + self.OrganizationName + "\",FacilityName:\"" + self.FacilityName + "\"," + "NPRIID:\"" + self.NPRIID + "\"," + "City:\"" + self.City + "\"," + "Substances:" + str(self.getSubstancesNumber()) + "}" 

	


#Read Excel File	
import xlrd
wb = xlrd.open_workbook('201305_TRAIScurrent.xls')
sh = wb.sheet_by_name(u'Public Data')
dataset = {}

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
	facilityList = sorted(facilityList, key= lambda fac: (fac.OrganizationName, fac.FacilityName, fac.City))   # sort by Company Name - Facility Name
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
result = "FacilGeogrLatitude\tFacilGeogrLongitude\tFacilityID\tFacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tSectorDesc\tNUMsubst\tSubstance_List\tContact\tPhone\tEmail\tHREmploy\n"
for key, facility in dataset.iteritems():
	if type(key) is unicode and len(key) == 0:
		continue
	#print str(facility)
	result = result + facility.getFeatureClassString() + "\n"
handle = open("json/TRAIS.txt",'w+')
handle.write(result)
handle.close();

# Generate txt file for feature class
result = "FacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tContact\tPhone\tEmail\tHREmploy\tSubstanceName\tUnits\tUse\tCreation\tContained\tReleasestoAir\tReleasestoWater\tReleasestoLand\tDisposalOnSite\tDisposalOffSite\tRecycleOffSite\n"
for key, facility in dataset.iteritems():
	if type(key) is unicode and len(key) == 0:
		continue
	for substance in facility.substances:
		result = result + facility.getReportString() + "\t" + substance.getFeatureClassString() + "\n"
handle = open("json/Substances.txt",'w+')
handle.write(result)
handle.close();

for key, substance in substanceDictionary.iteritems():
	print substance[0] + "\t" + substance[1] + "\t" + substance[2] + "\t" + substance[3]

	