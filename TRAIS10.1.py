import sys
reload(sys)
sys.setdefaultencoding("latin-1")

import xlrd, arcpy, string, os, zipfile, fileinput, time
from datetime import date
start_time = time.time()

INPUT_PATH = "input"
#INPUT_GDB = INPUT_PATH + "\\TRAIS.gdb"
OUTPUT_PATH = "output"
if arcpy.Exists(OUTPUT_PATH + "\\TRAIS.gdb"):
	os.system("rmdir " + OUTPUT_PATH + "\\TRAIS.gdb /s /q")
os.system("del " + OUTPUT_PATH + "\\*TRAIS*.*")
arcpy.CreateFileGDB_management(OUTPUT_PATH, "TRAIS", "9.3")
arcpy.env.workspace = OUTPUT_PATH + "\\TRAIS.gdb"
'''
wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2010 and 2011 (M) - 20130815 - Final.xls')
sh = wb.sheet_by_name(u'Main')
cntr = 1
IDDict = {}
for rownum in range(1, sh.nrows):
	row = sh.row_values(rownum)
	IDDict[row[0]] = cntr
	cntr = cntr + 1
'''

print "Create sectorNames feature class"
sectorNames = "sectorNames_NAD83"
NAICSFeatureClass = arcpy.env.workspace + "\\"  + sectorNames
arcpy.CreateFeatureclass_management(arcpy.env.workspace, sectorNames, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(NAICSFeatureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
# Process: Add Fields	
arcpy.AddField_management(NAICSFeatureClass, "ID", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
arcpy.AddField_management(NAICSFeatureClass, "sectorNameEn", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(NAICSFeatureClass, "sectorNameFr", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
try:
	with arcpy.da.InsertCursor(sectorNames, ("SHAPE@XY", "ID", "sectorNameEn", "sectorNameFr")) as cur:
		cntr = 1
		for line in fileinput.input('input\\sectorNames.txt'):
			items = line.strip().split("\t")
			code = int(items[0])
			rowValue = [(0.0, 0.0), code, items[1], items[2]]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print e.message
arcpy.Project_management(NAICSFeatureClass, arcpy.env.workspace + "\\sectorNames", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(NAICSFeatureClass, "FeatureClass")
print "Finish sectorNames feature class."

print "Create NAICS feature class"
NAICS = "NAICS_NAD83"
NAICSFeatureClass = arcpy.env.workspace + "\\"  + NAICS
arcpy.CreateFeatureclass_management(arcpy.env.workspace, NAICS, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(NAICSFeatureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
# Process: Add Fields	
arcpy.AddField_management(NAICSFeatureClass, "NAICS", "LONG", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
arcpy.AddField_management(NAICSFeatureClass, "Name", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
NAICSDict = {}
try:
	with arcpy.da.InsertCursor(NAICS, ("SHAPE@XY", "NAICS", "Name")) as cur:
		cntr = 1
		for line in fileinput.input('input\\NAICS.txt'):
			items = line.strip().split("\t")
			code = int(items[0])
			name = (items[1])[1:-1]
			NAICSDict[code] = name		
			rowValue = [(0.0, 0.0), code, name]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print e.message
arcpy.Project_management(NAICSFeatureClass, arcpy.env.workspace + "\\NAICS", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(NAICSFeatureClass, "FeatureClass")
print "Finish NAICS feature class."

print "Create Annual Report feature class"
AnnualReport = "AnnualReport_NAD83"
AnnualReportFeatureClass = arcpy.env.workspace + "\\"  + AnnualReport
arcpy.CreateFeatureclass_management(arcpy.env.workspace, AnnualReport, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(AnnualReportFeatureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
# Process: Add Fields	
#arcpy.AddField_management(AnnualReportFeatureClass, "UniqueID", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")

arcpy.AddField_management(AnnualReportFeatureClass, "UniqueFacilityID", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
fieldList = ["NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription"]
for field in fieldList:
	arcpy.AddField_management(AnnualReportFeatureClass, field, "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")

'''arcpy.AddField_management(AnnualReportFeatureClass, "ID", "LONG", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Address", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "NPRI_ID", "LONG", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Sector", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Contact", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Phone", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Email", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "HREmploy", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Use", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Creation", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "Contained", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "ReleasestoAir", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "ReleasestoWater", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "ReleasestoLand", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "DisposalOnSite", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "DisposalOffSite", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")
arcpy.AddField_management(AnnualReportFeatureClass, "RecycleOffSite", "TEXT", "", "", "", "", "NULLABLE", "REQUIRED", "")

def parseStringToFloat(item):
	if (type(item) is unicode or type(item) is str) and len(item) == 0:
		return 0
	elif type(item) is unicode or type(item) is str :
		return float(item)
	else:
		return item

def calculateTotal(items):
	return str(reduce(lambda a,d: a + d , map(parseStringToFloat, items), 0))
'''
substanceList = []
substanceCASNumberDict = {}
wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2010 - 20130815 - Final.xls')
sh = wb.sheet_by_name(u'Data')
cntr = 1
try:
	#with arcpy.da.InsertCursor(AnnualReport, ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription", "ID", "Address", "NPRI_ID", "Sector", "Contact", "Phone", "Email", "HREmploy", "Use", "Creation", "Contained", "ReleasestoAir", "ReleasestoWater", "ReleasestoLand", "DisposalOnSite", "DisposalOffSite", "RecycleOffSite")) as cur:
	with arcpy.da.InsertCursor(AnnualReport, ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription")) as cur:
		for rownum in range(1, sh.nrows):
			row = sh.row_values(rownum)
			#ID = IDDict[row[0]]
			'''Address = row[9] + " / " + row[10]
			NPRI_ID = 0
			NPRI_ID = str(row[1]).strip()
			if len(NPRI_ID) == 0:
				NPRI_ID = None
			else:
				NPRI_ID = int(float(NPRI_ID))
			 
			if len(str(row[7]).strip()) == 0:
				Sector = None
				SectorDesc = ""
			else:
				Sector = row[7]
				SectorDesc = ""
				if Sector in NAICSDict:
					SectorDesc = NAICSDict[Sector]
				else:
					print "Unknown Sector number"
			Sector = str(Sector) + " - " + SectorDesc
			
			Contact = row[20]
			
			Phone  = row[22]
			
			Email  = row[24]
			
			HREmploy = row[26]
			
			Use = ""
			Creation = ""
			Contained = ""
			ReleasestoAir = ""
			ReleasestoWater = ""
			ReleasestoLand = ""
			DisposalOnSite = ""
			DisposalOffSite = ""
			RecycleOffSite = ""
			
			
			Use = calculateTotal([row[33]])
			Creation = calculateTotal([row[34]])
			Contained = calculateTotal([row[35]])
			ReleasestoAir = calculateTotal([row[37], row[38], row[39], row[40], row[41], row[42]])
			ReleasestoWater = calculateTotal([row[43], row[44], row[45]])
			ReleasestoLand = calculateTotal([row[46], row[47], row[48]])
			DisposalOnSite = calculateTotal([row[49], row[50], row[51], row[52], row[53], row[54], row[55]])
			DisposalOffSite = calculateTotal([])
			RecycleOffSite = calculateTotal([])
			'''
			SubstanceName = row[30]
			if (len(SubstanceName) > 0):
				substanceList.append(SubstanceName)
				substanceCASNumberDict[SubstanceName] = row[31] # CASNumber
			#rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62], row[63], row[64], row[65], row[66], row[67], row[68], row[69], row[70], row[71], row[72], row[73], row[74], row[75], row[76], row[77], row[78], row[79], row[80], row[81], row[82], row[83], row[84], row[85], row[86], row[87], row[88], row[89], row[90], row[91], row[92], row[93], row[94], row[95], row[96], row[97], row[98], row[99], row[100], row[101], row[102], row[103], row[104], row[105], row[106], row[107], row[108], row[109], row[110], row[111], row[112], row[113], row[114], row[115], row[116], row[117], row[118], row[119], row[120], row[121], row[122], row[123], row[124], row[125], row[126], row[127], row[128], row[129], row[130], row[131], row[132], row[133], row[134], row[135], row[136], ID, Address, NPRI_ID, Sector, Contact, Phone, Email, HREmploy, Use, Creation, Contained, ReleasestoAir,ReleasestoWater,ReleasestoLand,DisposalOnSite,DisposalOffSite,RecycleOffSite]
			rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62], row[63], row[64], row[65], row[66], row[67], row[68], row[69], row[70], row[71], row[72], row[73], row[74], row[75], row[76], row[77], row[78], row[79], row[80], row[81], row[82], row[83], row[84], row[85], row[86], row[87], row[88], row[89], row[90], row[91], row[92], row[93], row[94], row[95], row[96], row[97], row[98], row[99], row[100], row[101], row[102], row[103], row[104], row[105], row[106], row[107], row[108], row[109], row[110], row[111], row[112], row[113], row[114], row[115], row[116], row[117], row[118], row[119], row[120], row[121], row[122], row[123], row[124], row[125], row[126], row[127], row[128], row[129], row[130], row[131], row[132], row[133], row[134], row[135], row[136]]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print "Error 2010: " + e.message + ": " + str(cntr)

wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2011 - 20130815 - Final.xls')
sh = wb.sheet_by_name(u'Data')
cntr = 1
test = ""
try:
	with arcpy.da.InsertCursor(AnnualReport, ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "RelationshipType", "MOEREG127Number", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "CountryPhysicalAddress", "AdditionalInformationPhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactFaxNumber", "PublicContactEmail", "PublicContactLanguageCorrespondence", "HighestRankingEmployee", "ParentLegalName", "ParentBusinessNumber", "ParentPercentageOwned", "SubstanceName", "CASNumber", "Units", "EnteredtheFacilityUsed", "Created", "ContainedinProduct", "ReportSumofAllMedia", "StackorPointReleasestoAir", "StorageorHandlingReleasestoAir", "FugitiveReleasestoAir", "SpillsReleasestoAir", "OtherNonPointReleasestoAir", "OtherReleasestoAir", "DirectDischargesReleasestoWater", "SpillsReleasestoWater", "LeaksReleasestoWater", "SpillsReleasestoLand", "LeaksReleasestoLand", "OtherReleasestoLand", "LandfillDisposedOnSite", "LandTreatmentDisposedOnSite", "UndergroundInjectionDisposedOnSite", "LandfillDisposedOffSite", "LandTreatmentDisposedOffSite", "UndergroundInjectionDisposedOffSite", "StorageDisposedOffSite", "PhysicalTreatmentOffSiteTransfers", "ChemicalTreatmentOffSiteTransfers", "BiologicalTreatmentOffSiteTransfers", "IncinerationThermalOffSiteTransfers", "MunicipalSewageTreatmentPlantOffsiteTransfers", "TailingsManagementDisposedOnSite", "WasteRockManagementDisposedOnSite", "TailingsManagementDisposedOffSite", "WasteRockManagementDisposedOffSite", "EnergyRecoveryRecycledOffSite", "RecoveryofSolventsRecycledOffSite", "RecoveryofOrganicSubstancesRecycledOffSite", "RecoveryofMetalsandMetalCompoundsRecycledOffSite", "RecoveryofInorganicMaterialsRecycledOffSite", "RecoveryofAcidsandBasesRecycledOffSite", "RecoveryofCatalystsRecycledOffSite", "RecoveryofPollutionAbatementResiduesRecycledOffSite", "RefiningofReuseofUsedOilRecycledOffSite", "OtherRecycledOffSite", "UseEnteredtheFacilityAnnualPercentageChange", "UseReportingPeriodofLastReportedQuantity", "CreatedAnnualPercentageChange", "CreatedReportingPeriodofLastReportedQuantity", "ContainedinProductAnnualPercentageChange", "ContainedinProductReportingPeriodofLastReportedQuantity", "ReasonsforChangeTRAQuantifications", "AllMediaAnnualPercentageChange", "AllMediaReportingPeriodofLastReportedQuantity", "ReleasestoAirAnnualPercentageChange", "ReleasestoAirReportingPeriodofLastReportedQuantity", "ReleasestoWaterAnnualPercentageChange", "ReleasestoWaterReportingPeriodofLastReportedQuantity", "ReleasestoLandAnnualPercentageChange", "ReleasestoLandReportingPeriodofLastReportedQuantity", "ReasonsforChangeAllMedia", "DisposedOnSiteAnnualPercentageChange", "DisposedOnSiteReportingPeriodofLastReportedQuantity", "DisposedOffSiteAnnualPercentageChange", "DisposedOffSiteReportingPeriodofLastReportedQuantity", "OffSiteTransfersAnnualPercentageChange", "ReasonsforChangeDisposals", "RecycledOffSiteAnnualPercentageChange", "RecycledReportingPeriodofLastReportedQuantity", "ReasonsForChangeRecycling", "PlanObjectives", "UseReductionTargetQuantity", "UseReductionTargetUnits", "UseReductionTargetTimeline", "UseDescriptionofTargets", "CreationReductionTargetQuantity", "CreationReductionTargetUnits", "CreationReductionTargetTimeline", "CreationDescriptionofTargets", "NoOptionsIdentifiedforUseorCreation", "Option", "ToxicsReductionCategory", "OptionActivityTaken", "Descriptionofreductionstepstaken", "Comparisonofthesteps", "OptionsImplementedAmountofreductioninuse", "OptionsImplementedAmountofreductionincreation", "OptionsImplementedAmountofreductionincontainedinproduct", "OptionsImplementedAmountofreductioninreleasetoair", "OptionsImplementedAmountofreductioninreleasetowater", "OptionsImplementedAmountofreductioninreleasetoland", "OptionsImplementedAmountofreductionindisposedonsite", "OptionsImplementedAmountofreductioninthesubstancedisposedoffsite", "OptionsImplementedAmountofreductioninrecycled", "Willthetimelinesbemet", "Comments", "DescriptionofAdditionalAction", "AdditionalActionsAmountofreductioninuse", "AdditionalActionsAmountofreductionincreation", "AdditionalActionsAmountofreductionincontainedinproduct", "AdditionalActionsAmountofreductioninreleasetoair", "AdditionalActionsAmountofreductioninreleasetowater", "AdditionalActionsAmountofreductioninreleasetoland", "AdditionalActionsAmountofreductionindisposedonsite", "AdditionalActionsAmountofreductioninthesubstancedisposedoffsite", "AdditionalActionsAmountofreductioninrecycled", "AmendmentsDescription")) as cur:
		for rownum in range(1, sh.nrows):
			row = sh.row_values(rownum)
			#ID = IDDict[row[0]]
			test = row[0]
			SubstanceName = row[30]
			if (len(SubstanceName) > 0):
				substanceList.append(SubstanceName)
				substanceCASNumberDict[SubstanceName] = row[31] # CASNumber
			rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62], row[63], row[64], row[65], row[66], row[67], row[68], row[69], row[70], row[71], row[72], row[73], row[74], row[75], row[76], row[77], row[78], row[79], row[80], row[81], row[82], row[83], row[84], row[85], row[86], row[87], row[88], row[89], row[90], row[91], row[92], row[93], row[94], row[95], row[96], row[97], row[98], row[99], row[100], row[101], row[102], row[103], row[104], row[105], row[106], row[107], row[108], row[109], row[110], row[111], row[112], row[113], row[114], row[115], row[116], row[117], row[118], row[119], row[120], row[121], row[122], row[123], row[124], row[125], row[126], row[127], row[128], row[129], row[130], row[131], row[132], row[133], row[134], row[135], row[136]]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print "Error 2011: " + e.message + ": " + str(cntr) + ", " + test
arcpy.Project_management(AnnualReportFeatureClass, arcpy.env.workspace + "\\AnnualReport", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(AnnualReportFeatureClass, "FeatureClass")
print "Finish Creating Annual Report feature class."
substanceList = list(set(substanceList))
#print substanceList

print "Create Substance Codes feature class"
substancesLanguageDict = {}
cntr = 1
for line in fileinput.input('input\\substance_codes.txt'):
	cntr = cntr + 1
	if cntr == 2:
		continue  # skip the first line		
	items = line.strip().split("\t")
	substancesLanguageDict[(items[1])[1:-1]] = (items[2])[1:-1]

featureName = "SubstanceCodes_NAD83"
featureClass = arcpy.env.workspace + "\\"  + featureName
arcpy.CreateFeatureclass_management(arcpy.env.workspace, featureName, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(featureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
# Process: Add Fields	
arcpy.AddField_management(featureClass, "CODE", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
arcpy.AddField_management(featureClass, "SUBSTANCE_EN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(featureClass, "SUBSTANCE_FR", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(featureClass, "CASNumber", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")

print "substance codes length: " + str(len(substanceList))
#print substanceList
substancesCodeDict = {}
cntr = 1
try:
	with arcpy.da.InsertCursor(featureName, ("SHAPE@XY", "CODE", "SUBSTANCE_EN", "SUBSTANCE_FR", "CASNumber")) as cur:
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
			cur.insertRow(rowValue)	
			#print cntr
			cntr = cntr + 1			
except Exception as e:
	print "Error: " + e.message + ", " + str(cntr)
arcpy.Project_management(featureClass, arcpy.env.workspace + "\\SubstanceCodes", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(featureClass, "FeatureClass")
print "Finish Substance Codes feature class."

substanceListDict = {}
wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2010 - 20130815 - Final.xls')
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
wb = xlrd.open_workbook('input\\Data\\TRA - Annual Report - 2011 - 20130815 - Final.xls')
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

print "Create Plan Summary feature class"
PlanSummary = "PlanSummaryNAD83"
PlanSummaryFeatureClass = arcpy.env.workspace + "\\"  + PlanSummary
arcpy.CreateFeatureclass_management(arcpy.env.workspace, PlanSummary, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(PlanSummaryFeatureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.AddField_management(PlanSummaryFeatureClass, "UniqueFacilityID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "NPRIID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReportingPeriod", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "OrganizationName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "FacilityName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "NAICS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "NumberofEmployees", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "StreetAddressPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "MunicipalityCityPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ProvincePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "PostalCodePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "PublicContactFullName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "PublicContactPosition", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "PublicContactTelephone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "PublicContactEMail", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "HighestRankingEmployee", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "SubstanceName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "SubstanceCAS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "IntenttoReduceUseYN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "StatementofIntenttoReduceUseText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonforNoIntenttoReduceUseText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "IntenttoReduceCreationYN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "StatementofIntenttoReduceCreationText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonforNoIntenttoReduceCreationText", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "PlanObjectives", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UseReductionQuantityTargetValue", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UseReductionQuantityTargetUnit", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UseReductionTimelineTargetYears", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "UseReductionTargetDescription", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "CreationReductionQuantityTargetValue", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "CreationReductionQuantityTargetUnit", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "CreationReductionTimelineTargetYears", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "CreationReductionTargetDescription", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonsforUse", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonsforUseSummary", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonsforCreation", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonsforCreationSummary", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "StatementNoOptionImplementedYN", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ReasonsNoOptionImplemented", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "OptionReductionCategory", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "ActivityTaken", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "DescriptionofOption", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedUseReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedCreationReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedContainedinProductReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedAirReleasesReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedAirReleasesReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedWaterReleasesReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedWaterReleasesReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedLandReleasesReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedLandReleasesReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedOnsiteDisposalsReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedOnsiteDisposalsReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedOffsiteDisposalsReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedOffsiteDisposalsReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedOffsiteRecyclingReduction", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "EstimatedOffsiteRecyclingReductionPercent", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "RationaleWhyOptionImplemented", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "DescriptionofAnyAdditionalActionsTaken", "TEXT", "", "", "10000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(PlanSummaryFeatureClass, "VersionofthePlan", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
#arcpy.AddField_management(PlanSummaryFeatureClass, "ID", "LONG", "", "", "", "", "NULLABLE", "REQUIRED", "")

wb = xlrd.open_workbook('input\\Data\\TRA - Plan Summary - 2011 - 20130815 - Final.xls')
sh = wb.sheet_by_name(u'Data')
planSummaryDict = {}
try:
	with arcpy.da.InsertCursor(PlanSummary, ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingPeriod", "OrganizationName", "FacilityName", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactEMail", "HighestRankingEmployee", "SubstanceName", "SubstanceCAS", "IntenttoReduceUseYN", "StatementofIntenttoReduceUseText", "ReasonforNoIntenttoReduceUseText", "IntenttoReduceCreationYN", "StatementofIntenttoReduceCreationText", "ReasonforNoIntenttoReduceCreationText", "PlanObjectives", "UseReductionQuantityTargetValue", "UseReductionQuantityTargetUnit", "UseReductionTimelineTargetYears", "UseReductionTargetDescription", "CreationReductionQuantityTargetValue", "CreationReductionQuantityTargetUnit", "CreationReductionTimelineTargetYears", "CreationReductionTargetDescription", "ReasonsforUse", "ReasonsforUseSummary", "ReasonsforCreation", "ReasonsforCreationSummary", "StatementNoOptionImplementedYN", "ReasonsNoOptionImplemented", "OptionReductionCategory", "ActivityTaken", "DescriptionofOption", "EstimatedUseReductionPercent", "EstimatedCreationReductionPercent", "EstimatedContainedinProductReductionPercent", "EstimatedAirReleasesReduction", "EstimatedAirReleasesReductionPercent", "EstimatedWaterReleasesReduction", "EstimatedWaterReleasesReductionPercent", "EstimatedLandReleasesReduction", "EstimatedLandReleasesReductionPercent", "EstimatedOnsiteDisposalsReduction", "EstimatedOnsiteDisposalsReductionPercent", "EstimatedOffsiteDisposalsReduction", "EstimatedOffsiteDisposalsReductionPercent", "EstimatedOffsiteRecyclingReduction", "EstimatedOffsiteRecyclingReductionPercent", "RationaleWhyOptionImplemented", "DescriptionofAnyAdditionalActionsTaken", "VersionofthePlan")) as cur:
		cntr = 1
		for rownum in range(1, sh.nrows):
			#print cntr
			row = sh.row_values(rownum)
			#ID = IDDict[row[0]]
			if row[0] in planSummaryDict:
				planSummaryDict[row[0]] = planSummaryDict[row[0]] + 1
			else:
				planSummaryDict[row[0]] = 1
			rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36], row[37], row[38], row[39], row[40], row[41], row[42], row[43], row[44], row[45], row[46], row[47], row[48], row[49], row[50], row[51], row[52], row[53], row[54], row[55], row[56], row[57], row[58], row[59], row[60], row[61], row[62]]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print "Error: " + e.message
arcpy.Project_management(PlanSummaryFeatureClass, arcpy.env.workspace + "\\PlanSummary", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(PlanSummaryFeatureClass, "FeatureClass")
print "Finish Creating Plan Summary feature class."

print "Create Exit Records feature class"
ExitRecords = "ExitRecords_NAD83"
ExitRecordsFeatureClass = arcpy.env.workspace + "\\"  + ExitRecords
arcpy.CreateFeatureclass_management(arcpy.env.workspace, ExitRecords, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(ExitRecordsFeatureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")

arcpy.AddField_management(ExitRecordsFeatureClass, "UniqueFacilityID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "NPRIID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "ReportingYear", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "OrganizationName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "FacilityName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "NAICS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "NumberofEmployees", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "StreetAddressPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "MunicipalityCityPhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "ProvincePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "PostalCodePhysicalAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "PublicContactFullName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "PublicContactPosition", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "PublicContactTelephone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "PublicContactEMail", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "HighestRankingEmployee", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "SubstanceName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "SubstanceCAS", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "DateofSubmission", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "Reason", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(ExitRecordsFeatureClass, "DescriptionofCircumstances", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", "")
#arcpy.AddField_management(ExitRecordsFeatureClass, "ID", "LONG", "", "", "", "", "NULLABLE", "REQUIRED", "")


wb = xlrd.open_workbook('input\\Data\\TRA - Exit Records - 2011 - 20130815 - Final.xls')
sh = wb.sheet_by_name(u'Data')
exitRecordDict = {}
try:
	with arcpy.da.InsertCursor(ExitRecords, ("SHAPE@XY", "UniqueFacilityID", "NPRIID", "ReportingYear", "OrganizationName", "FacilityName", "NAICS", "NumberofEmployees", "StreetAddressPhysicalAddress", "MunicipalityCityPhysicalAddress", "ProvincePhysicalAddress", "PostalCodePhysicalAddress", "UTMZone", "UTMEasting", "UTMNorthing", "PublicContactFullName", "PublicContactPosition", "PublicContactTelephone", "PublicContactEMail", "HighestRankingEmployee", "SubstanceName", "SubstanceCAS", "DateofSubmission", "Reason", "DescriptionofCircumstances")) as cur:
		cntr = 1
		for rownum in range(1, sh.nrows):
			#print cntr
			row = sh.row_values(rownum)
			#ID = IDDict[row[0]]
			if row[0] in exitRecordDict:
				exitRecordDict[row[0]] = exitRecordDict[row[0]] + 1
			else:
				exitRecordDict[row[0]] = 1			
			rowValue = [(0, 0), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23]]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print "Error: " + e.message
arcpy.Project_management(ExitRecordsFeatureClass, arcpy.env.workspace + "\\ExitRecords", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(ExitRecordsFeatureClass, "FeatureClass")
print "Finish Creating Exit Records feature class."

print "Create Facility feature class"
Facility = "Facility"
FacilityFeatureClass = arcpy.env.workspace + "\\"  + Facility
arcpy.CreateFeatureclass_management(arcpy.env.workspace, Facility, "POINT", "", "DISABLED", "DISABLED", "", "", "0", "0", "0")
# Process: Define Projection
arcpy.DefineProjection_management(FacilityFeatureClass, "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
# Process: Add Fields	
arcpy.AddField_management(FacilityFeatureClass, "UniqueID", "TEXT", "", "", "", "", "NON_NULLABLE", "REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NPRIId", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "MOEId", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Organisation", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Facility", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NAICS", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Year", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "StreetAddress", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "City", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "PostalCode", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "UTMZone", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "UTMEasting", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "UTMNorthing", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Latitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Longitude", "DOUBLE", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NumberofSubstances", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "SourceDataset", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "SourceXMLID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
#arcpy.AddField_management(FacilityFeatureClass, "ID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
#arcpy.AddField_management(FacilityFeatureClass, "FacilityName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
#arcpy.AddField_management(FacilityFeatureClass, "Address", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
#arcpy.AddField_management(FacilityFeatureClass, "OrganizationName", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NPRI_ID", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Sector", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "SectorDesc", "TEXT", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NUMsubst", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "Substance_List", "TEXT", "", "", "2000", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NUMPlanSummary", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")
arcpy.AddField_management(FacilityFeatureClass, "NUMExitRecord", "LONG", "", "", "", "", "NULLABLE", "NON_REQUIRED", "")

# import the data from Excel
wb = xlrd.open_workbook('input\\Data\\TRA - Facility Table - 2010 and 2011 (M) - 20130815 - Final.xls')
sh = wb.sheet_by_name(u'Main')
try:
#	with arcpy.da.InsertCursor(Facility, ("SHAPE@XY", "UniqueID", "NPRIId", "MOEId", "Organisation", "Facility", "NAICS", "Year", "StreetAddress", "City", "PostalCode", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "NumberofSubstances", "SourceDataset", "SourceXMLID", "ID", "FacilityName", "Address", "OrganizationName", "NPRI_ID", "Sector", "SectorDesc", "NUMsubst", "Substance_List", "NUMPlanSummary", "NUMExitRecord")) as cur:
	with arcpy.da.InsertCursor(Facility, ("SHAPE@XY", "UniqueID", "NPRIId", "MOEId", "Organisation", "Facility", "NAICS", "Year", "StreetAddress", "City", "PostalCode", "UTMZone", "UTMEasting", "UTMNorthing", "Latitude", "Longitude", "NumberofSubstances", "SourceDataset", "SourceXMLID", "NPRI_ID", "Sector", "SectorDesc", "NUMsubst", "Substance_List", "NUMPlanSummary", "NUMExitRecord")) as cur:
		cntr = 1
		for rownum in range(1, sh.nrows):
			row = sh.row_values(rownum)
			#print row
			#ID = row[0]
			#FacilityName = row[4]
			#Address = row[7] + " / " + row[8]
			#OrganizationName = row[3]
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
			NUMExitRecord = 0
			if row[0] in exitRecordDict:
				NUMExitRecord = exitRecordDict[row[0]]

			if len(str(row[14]).strip()) == 0:
				row[14] = 0.0
			if len(str(row[13]).strip()) == 0:
				row[13] = 0.0
			if len(str(row[5]).strip()) == 0:
				row[5] = None
			if row[14] > 0:
				row[14] = -row[14]
			#rowValue = [(row[14], row[13]), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], ID, FacilityName, Address, OrganizationName, NPRI_ID, Sector, SectorDesc, NUMsubst, Substance_List, NUMPlanSummary, NUMExitRecord]
			rowValue = [(row[14], row[13]), row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], NPRI_ID, Sector, SectorDesc, NUMsubst, Substance_List, NUMPlanSummary, NUMExitRecord]
			cur.insertRow(rowValue)
			cntr = cntr + 1
except Exception as e:
	print "Error: " + e.message
arcpy.Project_management(FacilityFeatureClass, arcpy.env.workspace + "\\Facilities", "PROJCS['WGS_1984_Web_Mercator_Auxiliary_Sphere',GEOGCS['GCS_WGS_1984',DATUM['D_WGS_1984',SPHEROID['WGS_1984',6378137.0,298.257223563]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]],PROJECTION['Mercator_Auxiliary_Sphere'],PARAMETER['False_Easting',0.0],PARAMETER['False_Northing',0.0],PARAMETER['Central_Meridian',0.0],PARAMETER['Standard_Parallel_1',0.0],PARAMETER['Auxiliary_Sphere_Type',0.0],UNIT['Meter',1.0]]", "NAD_1983_To_WGS_1984_5", "GEOGCS['GCS_North_American_1983',DATUM['D_North_American_1983',SPHEROID['GRS_1980',6378137.0,298.257222101]],PRIMEM['Greenwich',0.0],UNIT['Degree',0.0174532925199433]]")
arcpy.Delete_management(FacilityFeatureClass, "FeatureClass")
print "Finish Creating Facility feature class."

elapsed_time = time.time() - start_time
print elapsed_time