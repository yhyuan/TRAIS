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
	latitude = "0.0000" if (len(latitude) == 0)
	longitude = str(row[FieldIndexDict["Longitude"]])
	longitude = "0.0000" if (len(longitude) == 0)
	FacilityID = str(row[FieldIndexDict["Facility ID"]])
	FacilityID = "" if (len(FacilityID) == 0)
	FacilityName = row[FieldIndexDict["Facility Name"]]
	Address = ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]])
	OrganizationName = row[FieldIndexDict["Organization Name"]]
	NPRIID = str(row[FieldIndexDict["NPRI ID"]])
	NPRIID = ((str(int(NPRIID) + 10000000000)[1:]) if (len(NPRIID) > 0) else "")
	NAICS = str(row[FieldIndexDict["NAICS"]])
	SectorDesc = NAICSDictionary[str(int(NAICS))]
	code = getCode(row[FieldIndexDict["Substance Name"]], row[FieldIndexDict["CAS Number"]])
	return [FacilityID, FacilityName, Address, OrganizationName, NPRIID, NAICS, SectorDesc, code]
#	coorList = [latitude, longitude]
#	otherList = [FacilityID, FacilityName, Address, OrganizationName, NPRIID, NAICS, SectorDesc, code]
#	return "\t".join(coorList) + "\t\"" + "\"\t\"".join(otherList) + "\""
facilityDuplicatedListWithSubstances = map(facilityListWithSubstancesfunc, traisExcelList)
'''
def facilityListfunc(row):
	latitude = str(row[FieldIndexDict["Latitude"]])
	latitude = "0.0000" if (len(latitude) == 0)
	longitude = str(row[FieldIndexDict["Longitude"]])
	longitude = "0.0000" if (len(longitude) == 0)
	FacilityID = str(row[FieldIndexDict["Facility ID"]])
	FacilityID = "" if (len(FacilityID) == 0)
	FacilityName = row[FieldIndexDict["Facility Name"]]
	Address = ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]])
	OrganizationName = row[FieldIndexDict["Organization Name"]]
	NPRIID = str(row[FieldIndexDict["NPRI ID"]])
	NPRIID = ((str(int(NPRIID) + 10000000000)[1:]) if (len(NPRIID) > 0) else "")
	NAICS = str(row[FieldIndexDict["NAICS"]])
	SectorDesc = NAICSDictionary[str(int(NAICS))]
	coorList = [latitude, longitude]
	otherList = [FacilityID, FacilityName, Address, OrganizationName, NPRIID, NAICS, SectorDesc]
	return "\t".join(coorList) + "\t\"" + "\"\t\"".join(otherList) + "\""
'''
#facilityListfunc = lambda row: ("0.0000" if (len(str(row[FieldIndexDict["Latitude"]])) == 0) else str(row[FieldIndexDict["Latitude"]])) + "\t" + ("0.0000" if (len(str(row[FieldIndexDict["Longitude"]])) == 0) else str(row[FieldIndexDict["Longitude"]])) + "\t" + (str(int(row[FieldIndexDict["Facility ID"]])) if (len(str(row[FieldIndexDict["Facility ID"]])) > 0) else "") + "\t\"" +  row[FieldIndexDict["Facility Name"]] + "\"\t\"" + ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]]) + "\"\t\"" + row[FieldIndexDict["Organization Name"]] + "\"\t\"" +  ((str(int(row[FieldIndexDict["NPRI ID"]]) + 10000000000)[1:]) if (len(str(row[FieldIndexDict["NPRI ID"]])) > 0) else "") + "\"\t" + (str(row[FieldIndexDict["NAICS"]]) + "\t\"" + NAICSDictionary[str(int(row[FieldIndexDict["NAICS"]]))]) + "\"" 
#facilityDuplicatedList = map(facilityListfunc, traisExcelList)
#facilityDuplicatedList = map(lambda list: list[:-1], facilityDuplicatedListWithSubstances)

facilityList = list(set(map(lambda list: list[:-1], facilityDuplicatedListWithSubstances)))
facilityListDict = dict((fac,[]) for fac in facilityList)
#facilityListWithSubstancesfunc = lambda row: ("0.0000" if (len(str(row[FieldIndexDict["Latitude"]])) == 0) else str(row[FieldIndexDict["Latitude"]])) + "\t" + ("0.0000" if (len(str(row[FieldIndexDict["Longitude"]])) == 0) else str(row[FieldIndexDict["Longitude"]])) + "\t" + (str(int(row[FieldIndexDict["Facility ID"]])) if (len(str(row[FieldIndexDict["Facility ID"]])) > 0) else "") + "\t\"" +  row[FieldIndexDict["Facility Name"]] + "\"\t\"" + ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]]) + "\"\t\"" + row[FieldIndexDict["Organization Name"]] + "\"\t\"" +  ((str(int(row[FieldIndexDict["NPRI ID"]]) + 10000000000)[1:]) if (len(str(row[FieldIndexDict["NPRI ID"]])) > 0) else "") + "\"\t" + (str(row[FieldIndexDict["NAICS"]]) + "\t\"" + NAICSDictionary[str(int(row[FieldIndexDict["NAICS"]]))]) + "\"\t"  + getCode(row[FieldIndexDict["Substance Name"]], row[FieldIndexDict["CAS Number"]]) + ""

#def addFacilityListDict(line):
		
#Add Elements to the lists in facilityListDict
#map(lambda line: (facilityListDict["\t".join(line.split("\t")[:-1])]).append(line.split("\t")[-1]), facilityDuplicatedListWithSubstances)
def addFacilityDict(fac):
	facilityListDict[fac[:-1]].append(fac[-1])
map(addFacilityDict, facilityDuplicatedListWithSubstances)

def addFacilityID(i):
subNum = 0
	subNum = 0
	subList = "_".join(facilityListDict[facilityList[i]])
	if len(subList) != 0:
		subNum = len(facilityListDict[facilityList[i]]))
		subList = subList + "_"
	facilityList[i].append(str(i))
	facilityList[i].append(str(subNum))
	facilityList[i].append(str(subList))
	
facilityListID = map(lambda i: str(i) + "\t" + facilityList[i] + "\t" + str(0 if (len("_".join(facilityListDict[facilityList[i]])) == 0) else len(facilityListDict[facilityList[i]])) + "\t\"" + "_".join(facilityListDict[facilityList[i]]) + ("" if (len("_".join(facilityListDict[facilityList[i]])) == 0) else "_") + "\"", range(len(facilityList)))
for i in range(len(facilityListID)):
	row = facilityListID[i]
	items = row.split("\t")
	longitude = items[2]
	if(len(longitude) == 0):
		longitude = 0
	else:
		longitude = float(longitude)
	if (longitude > 0):
		longitude = -longitude;
	facilityListID[i] = "\t".join(items[0:2])  + "\t" + str(longitude) + "\t" + "\t".join(items[3:])
#facilityListID = map(lambda row: , facilityListID)
# Create a dictionary using the facility as key and id as value
#facilityListIDDict = dict((str(int(facility.split("\t")[6])) + " - " + facility.split("\t")[3],0) for facility in facilityList)
#map(lambda i: facilityListIDDict[(str(int((facilityList[i]).split("\t")[6])) + " - " + (facilityList[i]).split("\t")[3],0)], range(len(facilityList)))
facilityListIDDict = {}
for i in range(len(facilityList)):
	items = (facilityList[i]).split("\t")
	NPRIID = items[6][1:-1]
	key = ""
	if len(NPRIID) != 0:
		key = str(int(NPRIID))
	facilityListIDDict[key + " - " + items[3]] = i

handle = open("tmp/Facilities.txt",'w+')
head = "ID\tLatitude\tLongitude\tFacilityID\tFacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tSectorDesc\tNUMsubst\tSubstance_List\n"
handle.write(head + "\n".join(facilityListID))
#handle.write(head + "\n".join(facilityDuplicatedListWithSubstances))
handle.close();

substanceListfunc = lambda row: ("0.0000" if (len(str(row[FieldIndexDict["Latitude"]])) == 0) else str(row[FieldIndexDict["Latitude"]])) + "\t" + ("0.0000" if (len(str(row[FieldIndexDict["Longitude"]])) == 0) else str(row[FieldIndexDict["Longitude"]])) + "\t\"" +  row[FieldIndexDict["Facility Name"]] + "\"\t\"" + ((row[FieldIndexDict["Street Address (Physical Address)"]]).strip() + " / " + row[FieldIndexDict["Municipality/City (Physical Address)"]]) + "\"\t\"" + row[FieldIndexDict["Organization Name"]] + "\"\t" +  (str(int(row[FieldIndexDict["NPRI ID"]])) if (len(str(row[FieldIndexDict["NPRI ID"]])) > 0) else "") + "\t\"" + (str(row[FieldIndexDict["NAICS"]]) + " - " + NAICSDictionary[str(int(row[FieldIndexDict["NAICS"]]))]) + "\"\t\"" + row[FieldIndexDict["Public Contact"]] + "\"\t\"" + ("" if (len(str(row[FieldIndexDict["Contact Telephone Number"]])) == 0) else (str(int(row[FieldIndexDict["Contact Telephone Number"]])) if (type(row[FieldIndexDict["Contact Telephone Number"]]) is float) else row[FieldIndexDict["Contact Telephone Number"]])) + "\"\t\"" + row[FieldIndexDict["Contact Email"]] + "\"\t\"" + row[FieldIndexDict["Highest Ranking Employee"]] + "\"\t\"" + row[FieldIndexDict["Substance Name"]] + "\"\t\"" + row[FieldIndexDict["Units"]] + "\"\t\"" + str(row[FieldIndexDict["Use (Amount Entered Facility)"]]) + "\"\t\"" + str(row[FieldIndexDict["Creation  (Amount Created)"]]) + "\"\t\"" + str(row[FieldIndexDict["Contained In Product (Amount In Product)"]]) + "\"\t\"" + str(parseNumber(row[FieldIndexDict["Stack or Point (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Storage or Handling (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Fugitive (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Spills (Releases to Air)"]])  + parseNumber(row[FieldIndexDict["Other Non Point (Releases to Air)"]]) + parseNumber(row[FieldIndexDict["Other Sources (VOC) (Releases to Air)"]])) + "\"\t\"" + str(parseNumber(row[FieldIndexDict["Direct Discharges (Releases to Water)"]]) + parseNumber(row[FieldIndexDict["Spills (Releases to Water Bodies)"]]) + parseNumber(row[FieldIndexDict["Leaks (Releases to Water Bodies)"]])) + "\"\t\"" + str(parseNumber(row[FieldIndexDict["Spills (Releases to Land)"]]) + parseNumber(row[FieldIndexDict["Leaks (Releases to Land)"]]) + parseNumber(row[FieldIndexDict["Other (Releases to Land)"]])) + "\"\t\"" + str(parseNumber(row[FieldIndexDict["Landfill (Onsite)"]]) + parseNumber(row[FieldIndexDict["Land Treatment (Onsite)"]]) + parseNumber(row[FieldIndexDict["Underground Injection (Onsite)"]])) + "\"\t\"" + str(parseNumber(row[FieldIndexDict["Landfill (Offsite)"]]) + parseNumber(row[FieldIndexDict["Land Treatment (Offsite)"]]) + parseNumber(row[FieldIndexDict["Underground Injection (Offsite)"]]) + parseNumber(row[FieldIndexDict["Storage (Offsite)"]]) + parseNumber(row[FieldIndexDict["Physical (Offsite Treatment)"]]) + parseNumber(row[FieldIndexDict["Chemical (Offsite Treatment)"]]) + parseNumber(row[FieldIndexDict["Biological (Offsite Treatment)"]])  + parseNumber(row[FieldIndexDict["Incineration Thermal (Offsite Treatment)"]]) + parseNumber(row[FieldIndexDict["Municipal Sewage Treatment Plant (Offsite Treatment)"]]))+ "\"\t\"" + str(parseNumber(row[FieldIndexDict["Recovery of Energy (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recover of Solvents (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Organic Substances (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Metals and Metal Compounds (Recycling)"]]) + parseNumber( row[FieldIndexDict["Recovery of Inorganic Materials (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Acids or Bases (Recycling)"]]) + parseNumber(row[FieldIndexDict["Recovery of Catalysts (Recycling)"]] ) + parseNumber(row[FieldIndexDict["Recovery of Pollution Abatement Residue (Recycling)"]]) + parseNumber(row[FieldIndexDict["Refining of Reuse of Used Oil (Recycling)"]]) + parseNumber(row[FieldIndexDict["Other (Recycling)"]])) + "\""
substanceList = map(substanceListfunc, traisExcelList)
#substanceList = map(lambda row: str(facilityListIDDict[(str(int(row.split("\t")[5])) if len(row.split("\t")[5]) > 0 else "") + " - " + row.split("\t")[2]]) + "\t" + row, substanceList)
for i in range(len(substanceList)):
	row = substanceList[i]
	items = row.split("\t")
	NPRIID = ""
	if len(items[5]) > 0:
		NPRIID = str(int(items[5]))
	id = str(facilityListIDDict[NPRIID + " - " + items[2]])
	
	longitude = items[1]
	if(len(longitude) == 0):
		longitude = 0
	else:
		longitude = float(longitude)
	if (longitude > 0):
		longitude = -longitude;
	substanceList[i] = id + "\t" + items[0] + "\t" + str(longitude) + "\t" + "\t".join(items[2:])
	
handle = open("tmp/Substances.txt",'w+')
head = "ID\tLatitude\tLongitude\tFacilityName\tAddress\tOrganizationName\tNPRI_ID\tSector\tContact\tPhone\tEmail\tHREmploy\tSubstanceName\tUnits\tUse\tCreation\tContained\tReleasestoAir\tReleasestoWater\tReleasestoLand\tDisposalOnSite\tDisposalOffSite\tRecycleOffSite\n"
handle.write(head + "\n".join(substanceList))
handle.close();

substanceCodeList = map(lambda list: "0.0\t0.0\t" + "\t".join(list), substanceDictionary.values())
handle = open("tmp/substance_codes_output.txt",'w+')
head = "Latitude\tLongitude\tCODE\tSUBSTANCE_EN\tSUBSTANCE_FR\tCAS\n"
handle.write(head + "\n".join(substanceCodeList))
handle.close();

