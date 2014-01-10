import sys
reload(sys)
sys.setdefaultencoding("latin-1")

import xlrd, arcpy, string, os, zipfile, fileinput, time
from datetime import date
start_time = time.time()

INPUT_PATH = "input"
OUTPUT_PATH = "output"
if arcpy.Exists(OUTPUT_PATH + "\\TRAIS.gdb"):
	os.system("rmdir " + OUTPUT_PATH + "\\TRAIS.gdb /s /q")
os.system("del " + OUTPUT_PATH + "\\*TRAIS*.*")
arcpy.CreateFileGDB_management(OUTPUT_PATH, "TRAIS", "9.3")
arcpy.env.workspace = OUTPUT_PATH + "\\TRAIS.gdb"

def createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields):
	print "Create " + featureName + " feature class"
	featureNameNAD83 = featureName + "_NAD83"
	featureNameNAD83Path = arcpy.env.workspace + "\\"  + featureNameNAD83
	arcpy.CreateFeatureclass_management(arcpy.env.workspace, featureNameNAD83, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
	# Process: Define Projection
	arcpy.DefineProjection_management(featureNameNAD83Path, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
	# Process: Add Fields	
	for featrueField in featureFieldList:
		arcpy.AddField_management(featureNameNAD83Path, featrueField[0], featrueField[1], featrueField[2], featrueField[3], featrueField[4], featrueField[5], featrueField[6], featrueField[7], featrueField[8])
	# Process: Append the records
	cntr = 1
	try:
		with arcpy.da.InsertCursor(featureNameNAD83, featureInsertCursorFields) as cur:
			for rowValue in featureData:
				cur.insertRow(rowValue)
				cntr = cntr + 1
	except Exception as e:
		print "\tError: " + featureName + ": " + e.message
	# Change the projection to web mercator
	arcpy.Project_management(featureNameNAD83Path, arcpy.env.workspace + "\\" + featureName, "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
	arcpy.Delete_management(featureNameNAD83Path, "FeatureClass")
	print "Finish " + featureName + " feature class."

featureName = "sectorNames"
featureData = []
for line in fileinput.input('input\\sectorNames.txt'):
	items = line.strip().split("\t")
	code = int(items[0])
	featureData.append([(0.0, 0.0), code, items[1], items[2]])
#print featureData
#print featureData
featureFieldList = [["ID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["sectorNameEn", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["sectorNameFr", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "ID", "sectorNameEn", "sectorNameFr")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

featureName = "NAICS"
featureData = []
NAICSDict = {}
for line in fileinput.input('input\\NAICS.txt'):
	items = line.strip().split("\t")
	code = int(items[0])
	name = (items[1])[1:-1]
	NAICSDict[code] = name		
	featureData.append([(0.0, 0.0), code, name])
featureFieldList = [["NAICS", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["Name", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "NAICS", "Name")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

def parseNumber(item):
	if (type(item) is unicode or type(item) is str) and len(item) == 0:
		return 0
	elif type(item) is unicode or type(item) is str :
		return float(item)
	else:
		return item	

def calculateTotal(row, annualReportFieldIndexDict, fields):
	total = 0
	for field in fields:
		total = total + parseNumber(row[annualReportFieldIndexDict[field]])
	return str(total)

featureName = "AnnualReport"
annualReportFieldIndexDict = {}
featureFieldList = [["UniqueFacilityID", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""]]
annualReportFieldIndexDict[0] = "UniqueFacilityID"
index = 1
fieldList = ["NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "OffSiteTransfersReportingPeriodofLastReportedQuantity", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription", "DisposedOnSiteAnnualPercentageChangeHTMLOnly", "DisposedOffSiteAnnualPercentageChangeHTMLOnly"]
longTextFieldList = ["PlanObjectives", "Descriptionofreductionstepstaken", "DescriptionofAdditionalAction", "Comparisonofthesteps", "Comments", "AmendmentsDescription"]
for field in fieldList:
	if (field in longTextFieldList):
		featureFieldList.append([field, "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""])
	else:
		featureFieldList.append([field, "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""])
	annualReportFieldIndexDict[field] = index
	index = index + 1

fieldList = ["ReleasestoAir", "ReleasestoWater", "ReleasestoLand", "DisposalOnSite", "DisposalOffSite", "RecycleOffSite"]
for field in fieldList:
	featureFieldList.append([field, "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""])

#featureInsertCursorFields = ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription")
featureInsertCursorFields = ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "OffSiteTransfersReportingPeriodofLastReportedQuantity", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription", "DisposedOnSiteAnnualPercentageChangeHTMLOnly", "DisposedOffSiteAnnualPercentageChangeHTMLOnly", "ReleasestoAir", "ReleasestoWater", "ReleasestoLand", "DisposalOnSite", "DisposalOffSite", "RecycleOffSite")
featureData = []
substanceList = []
substanceCASNumberDict = {}
#AnnualReportXLSList = ['TRA - Annual Report - 2010 - 20130815 - Amended.xls', 'TRA - Annual Report - 2011 - 20130815 - Amended.xls']
AnnualReportXLSList = ['TRA - Annual Report - 2010 - 20131220 - Final.xls', 'TRA - Annual Report - 2011 - 20131220 - Final.xls', 'TRA - Annual Report - 2012 - 20131220 - Final.xls']
for AnnualReportXLS in AnnualReportXLSList:
	#print "Process: " + AnnualReportXLS
	wb = xlrd.open_workbook('input\\Data\\' + AnnualReportXLS)
	sh = wb.sheet_by_name(u'Data')
	for rownum in range(1, sh.nrows):
		row = sh.row_values(rownum)
		ReleasestoAir = calculateTotal(row, annualReportFieldIndexDict, ["StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir"])
		ReleasestoWater = calculateTotal(row, annualReportFieldIndexDict, ["DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater"])
		ReleasestoLand = calculateTotal(row, annualReportFieldIndexDict, ["SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand"])
		DisposalOnSite = calculateTotal(row, annualReportFieldIndexDict, ["LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite"])
		DisposalOffSite = calculateTotal(row, annualReportFieldIndexDict, ["LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite"])
		RecycleOffSite = calculateTotal(row, annualReportFieldIndexDict, ["EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite"])
		SubstanceName = row[30]
		if (len(SubstanceName) > 0):
			substanceList.append(SubstanceName)
			substanceCASNumberDict[SubstanceName] = row[31] # CASNumber
		#rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62], row[63], row[64], row[65], row[66], row[67], row[68], row[69], row[70], row[71], row[72], row[73], row[74], row[75], row[76], row[77], row[78], row[79], row[80], row[81], row[82], row[83], row[84], row[85], row[86], row[87], row[88], row[89], row[90], row[91], row[92], row[93], row[94], row[95], row[96], row[97], row[98], row[99], row[100], row[101], row[102], row[103], row[104], row[105], row[106], row[107], row[108], row[109], row[110], row[111], row[112], row[113], row[114], row[115], row[116], row[117], row[118], row[119], row[120], row[121], row[122], row[123], row[124], row[125], row[126], row[127], row[128], row[129], row[130], row[131], row[132], row[133], row[134], row[135], row[136], ID, Address, NPRI_ID, Sector, Contact, Phone, Email, HREmploy, Use, Creation, Contained, ReleasestoAir,ReleasestoWater,ReleasestoLand,DisposalOnSite,DisposalOffSite,RecycleOffSite]
		rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62], row[63], row[64], row[65], row[66], row[67], row[68], row[69], row[70], row[71], row[72], row[73], row[74], row[75], row[76], row[77], row[78], row[79], row[80], row[81], row[82], row[83], row[84], row[85], row[86], row[87], row[88], row[89], row[90], row[91], row[92], row[93], row[94], row[95], row[96], row[97], row[98], row[99], row[100], row[101], row[102], row[103], row[104], row[105], row[106], row[107], row[108], row[109], row[110], row[111], row[112], row[113], row[114], row[115], row[116], row[117], row[118], row[119], row[120], row[121], row[122], row[123], row[124], row[125], row[126], row[127], row[128], row[129], row[130], row[131], row[132], row[133], row[134], row[135], row[136], row[137], row[138], row[139]]
		rowValue = rowValue + [ReleasestoAir, ReleasestoWater, ReleasestoLand, DisposalOnSite, DisposalOffSite, RecycleOffSite]
		featureData.append(rowValue)

createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

substanceList = list(set(substanceList))
substancesLanguageDict = {}
cntr = 1
for line in fileinput.input('input\\substance_codes.txt'):
	cntr = cntr + 1
	if cntr == 2:
		continue  # skip the first line		
	items = line.strip().split("\t")
	substancesLanguageDict[(items[1])[1:-1]] = (items[2])[1:-1]

featureName = "SubstanceCodes"
featureData = []
substancesCodeDict = {}
cntr = 1
for substance in substanceList:
	if cntr < 10:
		code = "S00" + str(cntr)
	elif cntr < 100:
		code = "S0" + str(cntr)
	elif cntr < 1000:
		code = "S" + str(cntr)
	else:
		print "More than 1000 substances are found. Please update the scripts."
	#print substance
	substancesCodeDict[substance] = code
	substanceFR = substance
	if substance in substancesLanguageDict:
		substanceFR = substancesLanguageDict[substance]
	CASNumber = ""
	if substance in substanceCASNumberDict:
		CASNumber = substanceCASNumberDict[substance]
	rowValue = [(0.0, 0.0), code, substance, substanceFR, CASNumber]
	featureData.append(rowValue)
	cntr = cntr + 1
#print featureData
featureFieldList = [["CODE", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["SUBSTANCE_EN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SUBSTANCE_FR", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["CASNumber", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "CODE", "SUBSTANCE_EN", "SUBSTANCE_FR", "CASNumber")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

featureName = "PlanSummary"
featureData = []
#wb = xlrd.open_workbook('input\\Data\\TRA - Plan Summary - 2011 - 20130815 - Final.xls')
#wb = xlrd.open_workbook('input\\Data\\TRA - Plan Summary - 2011 - 20130815 - Amended (SAMPLE ONLY).xls')
wb = xlrd.open_workbook('input\\Data\\TRA - Plan Summary - 2011 - 20131220 - Final.xls')
sh = wb.sheet_by_name(u'Data')
planSummaryDict = {}
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if row[0] in planSummaryDict:
		planSummaryDict[row[0]] = planSummaryDict[row[0]] + 1
	else:
		planSummaryDict[row[0]] = 1
	rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62], row[63], row[64]]
	featureData.append(rowValue)	
featureFieldList = [["UniqueFacilityID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NPRIID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ReportingPeriod", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["OrganizationName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["FacilityName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NAICS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NumberofEmployees", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["StreetAddressPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MunicipalityCityPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ProvincePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PostalCodePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactFullName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactPosition", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactTelephone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactEMail", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["HighestRankingEmployee", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SubstanceName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SubstanceCAS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["IntenttoReduceUseYN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["StatementofIntenttoReduceUseText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonforNoIntenttoReduceUseText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["IntenttoReduceCreationYN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["StatementofIntenttoReduceCreationText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonforNoIntenttoReduceCreationText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["PlanObjectives", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["UseReductionQuantityTargetValue", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UseReductionQuantityTargetUnit", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UseReductionTimelineTargetYears", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UseReductionTargetDescription", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["CreationReductionQuantityTargetValue", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["CreationReductionQuantityTargetUnit", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["CreationReductionTimelineTargetYears", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["CreationReductionTargetDescription", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonsforUse", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonsforUseSummary", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonsforCreation", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonsforCreationSummary", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["StatementNoOptionImplementedYN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ReasonsNoOptionImplemented", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["OptionReductionCategory", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ActivityTaken", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DescriptionofOption", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedUseReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedCreationReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedContainedinProductReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedAirReleasesReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedAirReleasesReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedWaterReleasesReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedWaterReleasesReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedLandReleasesReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedLandReleasesReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedOnsiteDisposalsReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedOnsiteDisposalsReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedOffsiteDisposalsReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedOffsiteDisposalsReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedOffsiteRecyclingReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["EstimatedOffsiteRecyclingReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["AnticipatedTimelinesforAchievingReductionsinUse", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["AnticipatedTimelinesforAchievingReductionsinCreation", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["RationaleWhyOptionImplemented", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["DescriptionofAnyAdditionalActionsTaken", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", ""], ["VersionofthePlan", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactEMail", "HighestRankingEmployee", "SubstanceName", "SubstanceCAS", "IntenttoReduceUseYN", "StatementofIntenttoReduceUseText", "ReasonforNoIntenttoReduceUseText", "IntenttoReduceCreationYN", "StatementofIntenttoReduceCreationText", "ReasonforNoIntenttoReduceCreationText", "PlanObjectives", "UseReductionQuantityTargetValue", "UseReductionQuantityTargetUnit", "UseReductionTimelineTargetYears", "UseReductionTargetDescription", "CreationReductionQuantityTargetValue", "CreationReductionQuantityTargetUnit", "CreationReductionTimelineTargetYears", "CreationReductionTargetDescription", "ReasonsforUse", "ReasonsforUseSummary", "ReasonsforCreation", "ReasonsforCreationSummary", "StatementNoOptionImplementedYN", "ReasonsNoOptionImplemented", "OptionReductionCategory", "ActivityTaken", "DescriptionofOption", "EstimatedUseReductionPercent", "EstimatedCreationReductionPercent", "EstimatedContainedinProductReductionPercent", "EstimatedAirReleasesReduction", "EstimatedAirReleasesReductionPercent", "EstimatedWaterReleasesReduction", "EstimatedWaterReleasesReductionPercent", "EstimatedLandReleasesReduction", "EstimatedLandReleasesReductionPercent", "EstimatedOnsiteDisposalsReduction", "EstimatedOnsiteDisposalsReductionPercent", "EstimatedOffsiteDisposalsReduction", "EstimatedOffsiteDisposalsReductionPercent", "EstimatedOffsiteRecyclingReduction", "EstimatedOffsiteRecyclingReductionPercent", "AnticipatedTimelinesforAchievingReductionsinUse", "AnticipatedTimelinesforAchievingReductionsinCreation", "RationaleWhyOptionImplemented", "DescriptionofAnyAdditionalActionsTaken", "VersionofthePlan")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

featureName = "ExitRecords"
featureData = []
exitRecordDict = {}

#wb = xlrd.open_workbook('input\\Data\\TRA - Exit Records - 2011 - 20130815 - Final.xls')
wb = xlrd.open_workbook('input\\Data\\TRA - Exit Records - 2011 - 20131220 - Final.xls')
sh = wb.sheet_by_name(u'Data')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if row[0] in exitRecordDict:
		exitRecordDict[row[0]] = exitRecordDict[row[0]] + 1
	else:
		exitRecordDict[row[0]] = 1

	year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row[21], wb.datemode)
#	py_date = datetime.datetime(year, month, day, hour, minute, nearest_second)
	monthStr = str(month)
	if len(monthStr) == 1:
		monthStr = "0" + monthStr
	dayStr = str(day)
	if len(dayStr) == 1:
		dayStr = "0" + dayStr
	# print (str(year) + "/" + str(month) + "/" + str(day))
	row[21] = str(year) + "/" + monthStr + "/" + dayStr
	rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23]]
	featureData.append(rowValue)
wb = xlrd.open_workbook('input\\Data\\TRA - Exit Records - 2012 - 20131220 - Final.xls')
sh = wb.sheet_by_name(u'Data')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	if row[0] in exitRecordDict:
		exitRecordDict[row[0]] = exitRecordDict[row[0]] + 1
	else:
		exitRecordDict[row[0]] = 1

	year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row[21], wb.datemode)
#	py_date = datetime.datetime(year, month, day, hour, minute, nearest_second)
	monthStr = str(month)
	if len(monthStr) == 1:
		monthStr = "0" + monthStr
	dayStr = str(day)
	if len(dayStr) == 1:
		dayStr = "0" + dayStr
	# print (str(year) + "/" + str(month) + "/" + str(day))
	row[21] = str(year) + "/" + monthStr + "/" + dayStr
	rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23]]
	featureData.append(rowValue)
featureFieldList = [["UniqueFacilityID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NPRIID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ReportingYear", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["OrganizationName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["FacilityName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NAICS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NumberofEmployees", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["StreetAddressPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MunicipalityCityPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ProvincePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PostalCodePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactFullName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactPosition", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactTelephone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactEMail", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["HighestRankingEmployee", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SubstanceName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SubstanceCAS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DateofSubmission", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Reason", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""], ["DescriptionofCircumstances", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingYear", "OrganizationName", "FacilityName", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactEMail", "HighestRankingEmployee", "SubstanceName", "SubstanceCAS", "DateofSubmission", "Reason", "DescriptionofCircumstances")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

featureName = "ExemptionRecords"
featureData = []
#wb = xlrd.open_workbook('input\\Data\\TRA - Exemption Records - 2012 - 20130815 - V2 (SAMPLE ONLY).xls')
wb = xlrd.open_workbook('input\\Data\\TRA - Exemption Records - 2012 - 20131220 - Final.xlsx')
sh = wb.sheet_by_name(u'Data')
exemptionRecordDict = {}
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	#print rownum
	#print row
	if row[0] in exemptionRecordDict:
		exemptionRecordDict[row[0]] = exemptionRecordDict[row[0]] + 1
	else:
		exemptionRecordDict[row[0]] = 1

	year, month, day, hour, minute, second = xlrd.xldate_as_tuple(row[21], wb.datemode)
	monthStr = str(month)
	if len(monthStr) == 1:
		monthStr = "0" + monthStr
	dayStr = str(day)
	if len(dayStr) == 1:
		dayStr = "0" + dayStr
	# print (str(year) + "/" + str(month) + "/" + str(day))
	row[21] = str(year) + "/" + monthStr + "/" + dayStr
	rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24]]
	featureData.append(rowValue)
featureFieldList = [["UniqueFacilityID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NPRIID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ReportingPeriod", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["OrganizationName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["FacilityName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NAICS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NumberofEmployees", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["StreetAddressPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MunicipalityCityPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ProvincePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PostalCodePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactFullName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactPosition", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactTelephone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PublicContactEMail", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["HighestRankingEmployee", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SubstanceName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SubstanceCAS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DateofSubmission", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["ApplicableArea", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["DescriptionofCircumstances", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["RecordRank", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactEMail", "HighestRankingEmployee", "SubstanceName", "SubstanceCAS", "DateofSubmission", "ApplicableArea", "DescriptionofCircumstances", "RecordRank")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)


substanceListDict = {}
#wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2010 - 20130815 - Final.xls')
wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2010 - 20131220 - Final.xls')
sh = wb.sheet_by_name(u'Data')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	SubstanceName = row[30]
	if (len(SubstanceName) == 0):
		continue
	code = substancesCodeDict[SubstanceName]
	UniqueFacilityID = row[0]
	if UniqueFacilityID in substanceListDict:
		if not (code in substanceListDict[UniqueFacilityID]): 
			substanceListDict[UniqueFacilityID].append(code)
	else:
		substanceListDict[UniqueFacilityID] = [code]
#wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2011 - 20130815 - Final.xls')
wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2011 - 20131220 - Final.xls')
sh = wb.sheet_by_name(u'Data')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	SubstanceName = row[30]
	if (len(SubstanceName) == 0):
		continue
	code = substancesCodeDict[SubstanceName]
	UniqueFacilityID = row[0]
	if UniqueFacilityID in substanceListDict:
		if not (code in substanceListDict[UniqueFacilityID]): 
			substanceListDict[UniqueFacilityID].append(code)
	else:
		substanceListDict[UniqueFacilityID] = [code]
wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2012 - 20131220 - Final.xls')
sh = wb.sheet_by_name(u'Data')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	SubstanceName = row[30]
	if (len(SubstanceName) == 0):
		continue
	code = substancesCodeDict[SubstanceName]
	UniqueFacilityID = row[0]
	if UniqueFacilityID in substanceListDict:
		if not (code in substanceListDict[UniqueFacilityID]): 
			substanceListDict[UniqueFacilityID].append(code)
	else:
		substanceListDict[UniqueFacilityID] = [code]

featureName = "Facilities"
featureData = []
#wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2010 and 2011 (M) - 20130815 - Final.xls')
#wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2012 - 20130815 - Draft.xls')
#wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2012 - 20130815 - Draft 2.xls')
#wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2012 - 20131220 - Draft 1.xls')
wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2012 - 20131220 - Final.xls')

sh = wb.sheet_by_name(u'Main')
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	#print rownum
	#print row
	NPRI_ID = str(row[1]).strip()
	if isinstance(row[1], float):
		NPRI_ID = str(int(row[1])).strip()
	if len(NPRI_ID) == 0:
		NPRI_ID = ""
	elif len(NPRI_ID) == 10:
		NPRI_ID = NPRI_ID
	elif len(NPRI_ID) < 10:
		zeroList = ["0"] * (10 - len(NPRI_ID))
		NPRI_ID = "".join(zeroList) + NPRI_ID
	if len(str(row[5]).strip()) == 0:
		Sector = None
		SectorDesc = ""
	else:
		Sector = row[5]
		SectorDesc = ""
		if Sector in NAICSDict:
			SectorDesc = NAICSDict[Sector]
		else:
			print "Unknown Sector number"
	NUMsubst = 0
	Substance_List = ""
	if row[0] in substanceListDict:
		NUMsubst = len(substanceListDict[row[0]])
		Substance_List = "_".join(substanceListDict[row[0]]) + "_"
	NUMPlanSummary = 0
	if row[0] in planSummaryDict:
		NUMPlanSummary = planSummaryDict[row[0]]
	NUMRecord = 0
	if row[0] in exitRecordDict:
		NUMRecord = exitRecordDict[row[0]]
	if row[0] in exemptionRecordDict:
		NUMRecord = exemptionRecordDict[row[0]] + NUMRecord

	if len(str(row[14]).strip()) == 0:
		row[14] = 0.0
	if len(str(row[13]).strip()) == 0:
		row[13] = 0.0
	if len(str(row[5]).strip()) == 0:
		row[5] = None
	if row[14] > 0:
		row[14] = -row[14]
	rowValue = [(row[14], row[13]), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], NPRI_ID, Sector, SectorDesc, NUMsubst, Substance_List, NUMPlanSummary, NUMRecord]
	featureData.append(rowValue)
featureFieldList = [["UniqueID", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", ""], ["NPRIId", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["MOEId", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Organisation", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Facility", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NAICS", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Year", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["StreetAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["City", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["PostalCode", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Latitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Longitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SourceDataset", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SourceXMLID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NPRI_ID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Sector", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["SectorDesc", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NUMsubst", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["Substance_List", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", ""], ["NUMPlanSummary", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""], ["NUMRecord", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", ""]]
featureInsertCursorFields = ("SHAPE@XY", "UniqueID", "NPRIId", "MOEId", "Organisation", "Facility", "NAICS", "Year", "StreetAddress", "City", "PostalCode", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "SourceDataset", "SourceXMLID", "NPRI_ID", "Sector", "SectorDesc", "NUMsubst", "Substance_List", "NUMPlanSummary", "NUMRecord")
createFeatureClass(featureName, featureData, featureFieldList, featureInsertCursorFields)

elapsed_time = time.time() - start_time
print elapsed_time