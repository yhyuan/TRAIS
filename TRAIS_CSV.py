# http://www.python-excel.org/
# input:
#	NAICS.txt: the relationship between sector number and sector description
#	FieldIndex.txt: the relationship between the field name and column id in Excel file
#	substance_codes.txt: the substance English/French name, Code, and CAS number
#	201305_TRAIScurrent.xls: the main input excel file. 
#	
# output:
#	Facilities.txt: the facility location
#	Substances.txt: the substance information
#	substance_codes_output.txt: the substance information, which include the ones from substance_codes.txt and the ones found in 201305_TRAIScurrent.xls

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

def parseNumber(item):
	if (type(item) is unicode or type(item) is str) and len(item) == 0:
		return 0
	elif type(item) is unicode or type(item) is str :
		return float(item)
	else:
		return item	
def getCode(SubstanceName, CASNumber):
	if len(SubstanceName) == 0:
		return ""
	# Find the code with substance name
	if SubstanceName in substanceDictionary:
		return substanceDictionary[SubstanceName][0][1:-1]
	else:
		# If failed, try to find the code with CAS Number
		for key, value in substanceDictionary.iteritems():
			if (CASNumber == value[3][1:-1]):
				return substanceDictionary[key][0][1:-1]
		# if failed again, try to add it to the dictionary. 
		substanceDictionary[SubstanceName] = ["\"S" + str(len(substanceDictionary) + 1) + "\"", "\"" + SubstanceName + "\"", "\"" + SubstanceName + "\"", "\"" + CASNumber + "\""]
		return substanceDictionary[SubstanceName][0][1:-1]

#Read Excel File	
import xlrd
wb = xlrd.open_workbook('201305_TRAIScurrent.xls')
sh = wb.sheet_by_name(u'Public Data')
traisExcelList = map(lambda rownum: sh.row_values(rownum), range(1, sh.nrows))
def facilityListWithSubstancesfunc(row):
	latitude = str(row[FieldIndexDict["Latitude"]])
	if (len(latitude) == 0):
		latitude = "0.0000" 
	longitude = str(row[FieldIndexDict["Longitude"]])
	if (len(longitude) == 0):
		longitude = "0.0000" 
	longitude = float(longitude)
	if (longitude > 0):
		longitude = -longitude;
	longitude = str(longitude)
	FacilityID = str(row[FieldIndexDict["Facility ID"]])
	if (len(FacilityID) == 0):
		FacilityID = ""
	else:
		FacilityID = str(int(float(FacilityID)))
	FacilityName = row[FieldIndexDict["Facility Name"]]
	Address = ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]])
	OrganizationName = row[FieldIndexDict["Organization Name"]]
	NPRIID = str(row[FieldIndexDict["NPRI ID"]])
	if(len(NPRIID) > 0):
		NPRIID = (str(int(float(NPRIID)) + 10000000000)[1:])
	else:
		NPRIID = ""
	NAICS = str(row[FieldIndexDict["NAICS"]])
	SectorDesc = NAICSDictionary[str(int(float(NAICS)))]
	code = getCode(row[FieldIndexDict["Substance Name"]], row[FieldIndexDict["CAS Number"]])
	return [latitude, longitude, FacilityID, FacilityName, Address, OrganizationName, NPRIID, NAICS, SectorDesc, code]
facilityDuplicatedListWithSubstances = map(facilityListWithSubstancesfunc, traisExcelList)
facilityList = list(set(map(lambda fac: "\t".join(fac[:-1]), facilityDuplicatedListWithSubstances)))
facilityListDict = dict((fac,[]) for fac in facilityList)
#Add Elements to the lists in facilityListDict
def addFacilityDict(fac):
	facilityListDict["\t".join(fac[:-1])].append(fac[-1])
map(addFacilityDict, facilityDuplicatedListWithSubstances)

def addFacilityID(i):
	subNum = 0
	subList = "_".join(facilityListDict[facilityList[i]])
	if len(subList) != 0:
		subNum = len(facilityListDict[facilityList[i]])
		subList = subList + "_"
	return facilityList[i].split("\t") + [str(i), str(subNum), subList]	
facilityListID = map(addFacilityID, range(len(facilityList)))

def outputFacilityID(fac):
	# [latitude 0, longitude 1, FacilityID 2, FacilityName 3, Address 4, OrganizationName 5, NPRIID 6, NAICS 7, SectorDesc 8, id 9, subNum 10, subList 11]
	return fac[9] + "\t" + fac[0] + "\t" + fac[1] + "\t" + fac[2] + "\t\"" + fac[3] + "\"\t\"" + fac[4] + "\"\t\"" + fac[5] + "\"\t\"" + fac[6] + "\"\t" + fac[7] + "\t\"" + fac[8]  + "\"\t" + fac[10] + "\t\"" + fac[11] + "\""
handle = open("tmp/Facilities.txt",'w+')
head = "ID\tLatitude\tLongitude\tFacilityID\tFacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tSectorDesc\tNUMsubst\tSubstance_List\n"
handle.write(head + "\n".join(map(outputFacilityID, facilityListID)))
handle.close()

# Create a dictionary using the facility as key and id as value
keys = map(lambda fac: fac[6] + " - " + fac[3], facilityListID)
values = map(lambda fac: fac[9], facilityListID)
facilityListIDDict = dict(zip(keys, values))
#print facilityListIDDict
def substanceListfunc(row):
	latitude = str(row[FieldIndexDict["Latitude"]])
	if (len(latitude) == 0):
		latitude = "0.0000" 
	longitude = str(row[FieldIndexDict["Longitude"]])
	if (len(longitude) == 0):
		longitude = "0.0000" 	
	longitude = float(longitude)
	if (longitude > 0):
		longitude = -longitude;
	longitude = str(longitude)
	FacilityName = row[FieldIndexDict["Facility Name"]]
	Address = ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]])
	OrganizationName = row[FieldIndexDict["Organization Name"]]
	NPRIID = str(row[FieldIndexDict["NPRI ID"]])
	if(len(NPRIID) > 0):
		NPRIID = (str(int(float(NPRIID)) + 10000000000)[1:])
	else:
		NPRIID = ""	
	id = str(facilityListIDDict[NPRIID + " - " + FacilityName])
	NAICS = str(row[FieldIndexDict["NAICS"]])
	NAICS = str(int(float(NAICS)))
	Sector = NAICS + " - " + NAICSDictionary[NAICS]
	PublicContact = row[FieldIndexDict["Public Contact"]]
	ContactTelephoneNumber = row[FieldIndexDict["Contact Telephone Number"]]
	if(len(str(ContactTelephoneNumber)) == 0):
		ContactTelephoneNumber = ""
	if (type(ContactTelephoneNumber) is float):
		ContactTelephoneNumber = str(int(ContactTelephoneNumber))
	ContactEmail = row[FieldIndexDict["Contact Email"]]
	HighestRankingEmployee = row[FieldIndexDict["Highest Ranking Employee"]]
	SubstanceName = row[FieldIndexDict["Substance Name"]]
	Units = row[FieldIndexDict["Units"]]
	Use = str(row[FieldIndexDict["Use (Amount Entered Facility)"]])
	Creation = str(row[FieldIndexDict["Creation  (Amount Created)"]])
	Contained = str(row[FieldIndexDict["Contained In Product (Amount In Product)"]])
	ReleasestoAir = str(parseNumber(row[FieldIndexDict["Stack or Point (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Storage or Handling (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Fugitive (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Spills (Releases to Air)"]])  + parseNumber(row[FieldIndexDict["Other Non Point (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Other Sources (VOC) (Releases to Air)"]]))
	ReleasestoWater = str(parseNumber(row[FieldIndexDict["Direct Discharges (Releases to Water)"]]) + parseNumber(row[FieldIndexDict["Spills (Releases to Water Bodies)"]]) + parseNumber(row[FieldIndexDict["Leaks (Releases to Water Bodies)"]]))
	ReleasestoLand = str(parseNumber(row[FieldIndexDict["Spills (Releases to Land)"]]) + parseNumber(row[FieldIndexDict["Leaks (Releases to Land)"]]) + parseNumber(row[FieldIndexDict["Other (Releases to Land)"]]))
	DisposalOnSite = str(parseNumber(row[FieldIndexDict["Landfill (Onsite)"]]) + parseNumber(row[FieldIndexDict["Land Treatment (Onsite)"]]) + parseNumber(row[FieldIndexDict["Underground Injection (Onsite)"]]))
	DisposalOffSite = str(parseNumber(row[FieldIndexDict["Landfill (Offsite)"]]) + parseNumber(row[FieldIndexDict["Land Treatment (Offsite)"]]) + parseNumber(row[FieldIndexDict["Underground Injection (Offsite)"]]) + parseNumber(row[FieldIndexDict["Storage (Offsite)"]]) + parseNumber(row[FieldIndexDict["Physical (Offsite Treatment)"]]) + parseNumber(row[FieldIndexDict["Chemical (Offsite Treatment)"]]) + parseNumber(row[FieldIndexDict["Biological (Offsite Treatment)"]])  + parseNumber(row[FieldIndexDict["Incineration Thermal (Offsite Treatment)"]]) + parseNumber(row[FieldIndexDict["Municipal Sewage Treatment Plant (Offsite Treatment)"]]))
	RecycleOffSite = str(parseNumber(row[FieldIndexDict["Recovery of Energy (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recover of Solvents (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Organic Substances (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Metals and Metal Compounds (Recycling)"]]) + parseNumber( row[FieldIndexDict["Recovery of Inorganic Materials (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Acids or Bases (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Catalysts (Recycling)"]] ) + parseNumber(row[FieldIndexDict["Recovery of Pollution Abatement Residue (Recycling)"]]) + parseNumber(row[FieldIndexDict["Refining of Reuse of Used Oil (Recycling)"]]) + parseNumber(row[FieldIndexDict["Other (Recycling)"]]))
	return "\t".join([id, latitude, longitude]) + "\t\"" + "\"\t\"".join([FacilityName,Address,OrganizationName,NPRIID,Sector,PublicContact,ContactTelephoneNumber,ContactEmail,HighestRankingEmployee,SubstanceName,Units,Use,Creation,Contained,ReleasestoAir,ReleasestoWater,ReleasestoLand,DisposalOnSite,DisposalOffSite,RecycleOffSite]) + "\""
	
substanceList = map(substanceListfunc, traisExcelList)
handle = open("tmp/Substances.txt",'w+')
head = "ID\tLatitude\tLongitude\tFacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tContact\tPhone\tEmail\tHREmploy\tSubstanceName\tUnits\tUse\tCreation\tContained\tReleasestoAir\tReleasestoWater\tReleasestoLand\tDisposalOnSite\tDisposalOffSite\tRecycleOffSite\n"
handle.write(head + "\n".join(substanceList))
handle.close();

substanceCodeList = map(lambda list: "0.0\t0.0\t" + "\t".join(list), substanceDictionary.values())
handle = open("tmp/substance_codes_output.txt",'w+')
head = "Latitude\tLongitude\tCODE\tSUBSTANCE_EN\tSUBSTANCE_FR\tCAS\n"
handle.write(head + "\n".join(substanceCodeList))
handle.close();

