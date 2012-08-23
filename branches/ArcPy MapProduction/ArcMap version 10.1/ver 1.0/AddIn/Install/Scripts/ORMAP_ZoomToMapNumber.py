### ---------------------------------------------------------------------------
### Ormap_ZoomToMap.py
### Created by: Shad Campbell
### Date: 3/11/2011
### Updated by:
### Description: This script zooms and configures a map based upon the MapNumber
### and PageSize arguemnents passed into it.  It is used by itself and also used
### when printing/exporting.
### ---------------------------------------------------------------------------
##
import arcpy, arcpy.mapping as MAP
import pythonaddins
import os, sys

sys.path.append(os.path.dirname(__file__).replace('Scripts', 'Config')) #path to config files
import ORMAP_LayersConfig as OrmapLayers


#IMPORT PARAMETERS - SHOULD ONLY BE A SINGLE STRING ITEMS
MapNumber = arcpy.GetParameterAsText(0)
##PageSize = arcpy.GetParameterAsText(1)

#LOAD THE PAGE INFORMATION
import ORMAP_MapConfig as PageConfig
##import ORMAP_18x20MapConfig
##import ORMAP_18x24MapConfig
##if PageSize=='18x20':
##    PageConfig = ORMAP_18x20MapConfig
##else:
##    PageConfig = ORMAP_18x24MapConfig

#REFERENCE MAP DOCUMENT
MXD = MAP.MapDocument("CURRENT")

#REFERENCE EACH DATAFRAME - only mainDF is required.
if len(MAP.ListDataFrames(MXD, "MainDF"))>0:
    mainDF = MAP.ListDataFrames(MXD, "MainDF")[0]
else:
    sys.exit("Unable to find a DataFrame named mainDF.  Please make sure default dataframe is named 'mainDF'.")
locatorDF = None
if len(MAP.ListDataFrames(MXD, "LocatorDF"))>0:
    locatorDF = MAP.ListDataFrames(MXD, "LocatorDF")[0]
sectDF = None
if len(MAP.ListDataFrames(MXD, "SectionsDF"))>0:
    sectDF = MAP.ListDataFrames(MXD, "SectionsDF")[0]
qSectDF = None
if len(MAP.ListDataFrames(MXD, "QSectionsDF"))>0:
    qSectDF = MAP.ListDataFrames(MXD, "QSectionsDF")[0]


#REFERENCE PAGELAYOUT TABLE
pageLayoutRow = None
if len(MAP.ListTableViews(MXD, OrmapLayers.PAGELAYOUT_TABLE, mainDF))>0:
    pageLayoutTable = MAP.ListTableViews(MXD, OrmapLayers.PAGELAYOUT_TABLE, mainDF)[0]
    pageLayoutCursor = arcpy.SearchCursor(pageLayoutTable.name, "MapNumber = '" + MapNumber + "'")
    pageLayoutRow = pageLayoutCursor.next()
else:
    sys.exit("Unable to find PageLayoutTable table.  Please check your TOC and config file.")

#REFERENCE CANCELLED NUMBERS TABLE
cancelledRow = None
if len(MAP.ListTableViews(MXD, OrmapLayers.CANCELLEDNUMBERS_TABLE, mainDF))>0:
    cancelledTable = MAP.ListTableViews(MXD, OrmapLayers.CANCELLEDNUMBERS_TABLE, mainDF)[0]
    cancelledCursor = arcpy.SearchCursor(cancelledTable.name, "MapNumber = '" + MapNumber + "'", "", "", PageConfig.CancelledSortField)
    cancelledRow = cancelledCursor.next()
else:
    sys.exit("Unable to find Cancelled Numbers table.  Please check your TOC and config file.")


#REFERENCE CUSTOM DEFINITION QUERIES TABLE
defQueryRow = None
if len(MAP.ListTableViews(MXD, OrmapLayers.CUSTOMDEFINITIONQUERIES_TABLE, mainDF))>0:
    defQueryTable = MAP.ListTableViews(MXD, OrmapLayers.CUSTOMDEFINITIONQUERIES_TABLE, mainDF)[0]
    defQueryCursor = arcpy.SearchCursor(defQueryTable.name, "'MapNumber' = '" + MapNumber + "'") #--CUSTOM
    defQueryRow = defQueryCursor.next()


#REFERENCE MAPINDEX LAYER
mapIndexRow = None
if len(MAP.ListLayers(MXD, OrmapLayers.MAPINDEX_LAYER, mainDF))>0:
    MapIndex = MAP.ListLayers(MXD, OrmapLayers.MAPINDEX_LAYER, mainDF)[0]
    MapIndex.definitionQuery = ""
    mapIndexCursor = arcpy.SearchCursor(MapIndex.name, "MapNumber = '" + MapNumber + "'")
    mapIndexRow = mapIndexCursor.next()
    if mapIndexRow==None:
        sys.exit("Unable to find specified MapNumber in your MapIndex feature class.")
else:
    sys.exit("Unable to find MapIndex layer.  Please check your TOC and config file.")



arcpy.AddMessage("")
arcpy.AddMessage("Processing: " + MapNumber)
arcpy.AddMessage("")

#COLLECT MAP INDEX POLYGON INFORMATION AND GET FEATURE EXTENT
geom = mapIndexRow.shape
featureExtent = geom.extent

if pageLayoutRow==None:
    #SET PAGE LAYOUT LOCATIONS FROM CONFIG
    arcpy.AddMessage("Reading Page Layout items from Configuration File")
    DataFrameMinX = PageConfig.DataFrameMinX
    DataFrameMinY = PageConfig.DataFrameMinY
    DataFrameMaxX = PageConfig.DataFrameMaxX
    DataFrameMaxY = PageConfig.DataFrameMaxY
    mapExtent = featureExtent
    MapAngle = PageConfig.MapAngle
    TitleX = PageConfig.TitleX
    TitleY = PageConfig.TitleY
    DisclaimerX = PageConfig.DisclaimerX
    DisclaimerY = PageConfig.DisclaimerY
    CancelNumX = PageConfig.CancelNumX
    CancelNumY = PageConfig.CancelNumY
    DateX = PageConfig.DateX
    DateY = PageConfig.DateY
    URCornerNumX = PageConfig.URCornerNumX
    URCornerNumY = PageConfig.URCornerNumY
    LRCornerNumX = PageConfig.LRCornerNumX
    LRCornerNumY = PageConfig.LRCornerNumY
    ScaleBarX = PageConfig.ScaleBarX
    ScaleBarY = PageConfig.ScaleBarY
    NorthX = PageConfig.NorthX
    NorthY = PageConfig.NorthY

else:
    #SET PAGE LAYOUT LOCATIONS FROM PAGELAYOUT TABLE
    arcpy.AddMessage("Reading Page Layout items from PAGELAYOUTTABLE")
    MapAngle = pageLayoutRow.MapAngle

    DataFrameMinX = pageLayoutRow.DataFrameMinX
    DataFrameMinY = pageLayoutRow.DataFrameMinY
    DataFrameMaxX = pageLayoutRow.DataFrameMaxX
    DataFrameMaxY = pageLayoutRow.DataFrameMaxY

    MapPositionX = pageLayoutRow.MapPositionX
    MapPositionY = pageLayoutRow.MapPositionY

    mapExtent = arcpy.Extent(MapPositionX-1, MapPositionY-1, MapPositionX+1, MapPositionY+1)

    MapAngle = pageLayoutRow.MapAngle
    TitleX = pageLayoutRow.TitleX
    TitleY = pageLayoutRow.TitleY
    DisclaimerX = pageLayoutRow.DisclaimerX
    DisclaimerY = pageLayoutRow.DisclaimerY
    CancelNumX = pageLayoutRow.CancelNumX
    CancelNumY = pageLayoutRow.CancelNumY
    DateX = pageLayoutRow.DateX
    DateY = pageLayoutRow.DateY
    URCornerNumX = pageLayoutRow.URCornerNumX
    URCornerNumY = pageLayoutRow.URCornerNumY
    LRCornerNumX = pageLayoutRow.LRCornerNumX
    LRCornerNumY = pageLayoutRow.LRCornerNumY
    ScaleBarX = pageLayoutRow.ScaleBarX
    ScaleBarY = pageLayoutRow.ScaleBarY
    NorthX = pageLayoutRow.NorthX
    NorthY = pageLayoutRow.NorthY

#MISC Relative Map Distances
CountyNameDist = PageConfig.CountyNameDist
MapScaleDist = PageConfig.MapScaleDist

#GET OTHER TABLE ATTRIBUTES
MapScale = str(mapIndexRow.MapScale)
MapNumber = mapIndexRow.MapNumber
ORMapNum = mapIndexRow.ORMapNum
CityName = mapIndexRow.CityName

#SET QUERY DEFINITIONS FOR EACH LAYER.  SEARCH FOR AN ITEM IN THE DEF QUERY TABLE FIRST... OTHERWISE SET TO CONFIG TABLE VALUES.
if defQueryRow != None:

    #-- Items stored in long string input defQueryString field.  Pull them out and into thier own array.
    defQueryList = defQueryRow.defQueryString.split(";")
    defQueryLayers = []
    defQueryValues = []

    for defQueryItem in defQueryList:
        lyr,qry = defQueryItem.split(":")
        defQueryLayers.append(lyr)
        defQueryValues.append(qry)

    for lyr in MAP.ListLayers(MXD, "", mainDF):
        if lyr.supports("DATASETNAME"):
            if lyr.name in defQueryLayers:
                lyr.definitionQuery = defQueryValues[defQueryLayers.index(lyr.name)]

else:
    for lyr in MAP.ListLayers(MXD, "", mainDF):
        if lyr.supports("DATASETNAME"):
            if lyr.name == OrmapLayers.LOTSANNO_LAYER:
                lyr.definitionQuery = OrmapLayers.LOTSANNO_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.PLATSANNO_LAYER:
                lyr.definitionQuery = OrmapLayers.PLATSANNO_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.TAXCODEANNO_LAYER:
                lyr.definitionQuery = OrmapLayers.TAXCODEANNO_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.TAXNUMANNO_LAYER:
                lyr.definitionQuery = OrmapLayers.TAXNUMANNO_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ACRESANNO_LAYER:
                lyr.definitionQuery = OrmapLayers.ACRESANNO_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO10_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO10_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO20_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO20_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO30_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO30_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO40_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO40_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO50_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO50_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO60_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO60_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO100_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO100_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO200_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO200_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO400_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO400_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO800_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO800_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.ANNO2000_LAYER:
                lyr.definitionQuery = OrmapLayers.ANNO2000_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.CORNER_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.CORNER_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.TAXCODELINES_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.TAXCODELINES_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.TAXLOTLINES_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.TAXLOTLINES_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.REFLINES_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.REFLINES_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.CARTOLINES_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.CARTOLINES_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.WATERLINES_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.WATERLINES_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.WATER_ABOVE_LAYER:
                lyr.definitionQuery = OrmapLayers.WATER_ABOVE_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.MAPINDEXSEEMAP_LAYER:
                lyr.definitionQuery = OrmapLayers.MAPINDEXSEEMAP_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.MAPINDEX_LAYER:
                lyr.definitionQuery = OrmapLayers.MAPINDEX_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.CORNER_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.CORNER_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.TAXCODELINES_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.TAXCODELINES_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.TAXLOTLINES_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.TAXLOTLINES_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.REFLINES_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.REFLINES_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.CARTOLINES_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.CARTOLINES_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.WATERLINES_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.WATERLINES_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.WATER_BELOW_LAYER:
                lyr.definitionQuery = OrmapLayers.WATER_BELOW_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            #EXTRA LAYERS
            if lyr.name == OrmapLayers.EXTRA1_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA1_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA2_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA2_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA3_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA3_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA4_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA4_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA5_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA5_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA6_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA6_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA7_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA7_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA8_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA8_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA9_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA9_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA10_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA10_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA11_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA11_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA12_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA12_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA13_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA13_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA14_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA14_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA15_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA15_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA16_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA16_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA17_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA17_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA18_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA18_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA19_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA19_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)
            if lyr.name == OrmapLayers.EXTRA20_LAYER:
                lyr.definitionQuery = OrmapLayers.EXTRA20_QD.replace("*MapNumber*", MapNumber).replace("*MapScale*", MapScale)


#PARSE ORMAP MAPNUMBER TO DEVELOP MAP TITLE
sTown1 = ORMapNum[2]
sTown2 = ORMapNum[3]
sTownPart = ORMapNum[5:7]
sTownDir = ORMapNum[7]
sRange1 = ORMapNum[8]
sRange2 = ORMapNum[9]
sRangePart = ORMapNum[11:13]
sRangeDir = ORMapNum[13]
sSection1 = ORMapNum[14]
sSection2 = ORMapNum[15]
sQtr = ORMapNum[16]
sQtrQtr = ORMapNum[17]
sAnomaly = ORMapNum[18:20]
sMapType = ORMapNum[20]
sMapNum1 = ORMapNum[21]
sMapNum2 = ORMapNum[22]
sMapNum3 = ORMapNum[23]


#BUILD TOWNSHIP TEXT TO EXCLUDE LEADING ZEROS
sTownship = ""
if sTown1 <> "0":
    sTownship = sTown1
sTownship = sTownship + sTown2

#BUILD PARTIAL TOWNSHIP TEXT
sTP = ""
if sTownPart == "25":
    sTP = " 1/4"
if sTownPart == "50":
    sTP = " 1/2"
if sTownPart == "75":
    sTP = " 3/4"

#BUILD RANGE TEXT TO EXCLUDE LEADING ZEROS
sRange = ""
if sRange1 <> "0":
    sRange = sRange1
sRange = sRange + sRange2

#BUILD PARTIAL RANGE TEXT
sRP = ""
if sRangePart == "25":
    sRP = " 1/4"
if sRangePart == "50":
    sRP = " 1/2"
if sRangePart == "75":
    sRP = " 3/4"

#BUILD SECTION TEXT TO EXCLUDE LEADING ZEROS
sSection = ""
if sSection1 <> "0":
    sSection = sSection1
sSection = sSection + sSection2

#BUILD QTR/QTR TEXT
sSectionText = ""
if sQtr == "A" and sQtrQtr == "0":
    sSectionText = "N.E.1/4"
elif sQtr == "A" and sQtrQtr == "A":
    sSectionText = "N.E.1/4 N.E.1/4"
elif sQtr == "A" and sQtrQtr == "B":
    sSectionText = "N.W.1/4 N.E.1/4"
elif sQtr == "A" and sQtrQtr == "C":
    sSectionText = "S.W.1/4 N.E.1/4"
elif sQtr == "A" and sQtrQtr == "D":
    sSectionText = "S.E.1/4 N.E.1/4"

if sQtr == "B" and sQtrQtr == "0":
    sSectionText = "N.W.1/4"
elif sQtr == "B" and sQtrQtr == "A":
    sSectionText = "N.E.1/4 N.W.1/4"
elif sQtr == "B" and sQtrQtr == "B":
    sSectionText = "N.W.1/4 N.W.1/4"
elif sQtr == "B" and sQtrQtr == "C":
    sSectionText = "S.W.1/4 N.W.1/4"
elif sQtr == "B" and sQtrQtr == "D":
    sSectionText = "S.E.1/4 N.W.1/4"

if sQtr == "C" and sQtrQtr == "0":
    sSectionText = "S.W.1/4"
elif sQtr == "C" and sQtrQtr == "A":
    sSectionText = "N.E.1/4 S.W.1/4"
elif sQtr == "C" and sQtrQtr == "B":
    sSectionText = "N.W.1/4 S.W.1/4"
elif sQtr == "C" and sQtrQtr == "C":
    sSectionText = "S.W.1/4 S.W.1/4"
elif sQtr == "C" and sQtrQtr == "D":
    sSectionText = "S.E.1/4 S.W.1/4"

if sQtr == "D" and sQtrQtr == "0":
    sSectionText = "S.E.1/4"
elif sQtr == "D" and sQtrQtr == "A":
    sSectionText = "N.E.1/4 S.E.1/4"
elif sQtr == "D" and sQtrQtr == "B":
    sSectionText = "N.W.1/4 S.E.1/4"
elif sQtr == "D" and sQtrQtr == "C":
    sSectionText = "S.W.1/4 S.E.1/4"
elif sQtr == "D" and sQtrQtr == "D":
    sSectionText = "S.E.1/4 S.E.1/4"

#BUILD MAP SUFFIX TYPE AND MAP NUMBER TEXT
sMN = ""
if sMapNum1 <> "0":
    sMN += sMapNum1
if sMapNum2 <> "0":
    sMN += sMapNum2
if sMapNum3 <> "0":
    sMN += sMapNum3

#GENERATE TEXT FOR SHORT TITLES (UR & LR)
shortMapTitle = sTown1 + sTown2 + " " + sRange1 + sRange2 + " " + sSection1 + sSection2
if sQtr <> "0":
    shortMapTitle += sQtr
if sQtrQtr <> "0":
    shortMapTitle += sQtrQtr

#ADD CITY NAME IF EXISTS TO SHORT TITLES
if CityName != None:
    shortMapTitle += "\n" + CityName.replace(",","\n")

#CREATE TEXT FOR LONG MAP TITLE BASED ON SCALE FORMATS PROVIDED BY DOR.
sLongMapTitle = ""
sMapScale = "1\" = " + str(int(MapScale)/12) + "'"
if MapScale == "24000":
    sLongMapTitle = "T." + str(sTP) + str(sTownship) + str(sTownDir) + ". R." + str(sRange) + str(sRP) + str(sRangeDir) + ". W.M."
elif MapScale == "4800":
    sLongMapTitle = "SECTION " + str(sSection) + " T." + str(sTownship) + str(sTP)  + str(sTownDir) + ". R." + str(sRange) + str(sRP) + str(sRangeDir) + ". W.M."
elif MapScale == "2400":
    sLongMapTitle = str(sSectionText) + " SEC." + str(sSection) + " T." + str(sTownship) + str(sTP) + str(sTownDir) + ". R." + str(sRange) + str(sRP)  + str(sRangeDir) + ". W.M."
elif MapScale == "1200":
    sLongMapTitle = str(sSectionText) + " SEC." + str(sSection) + " T." + str(sTownship) + str(sTP)  + str(sTownDir) + ". R." + str(sRange) + str(sRP)  + str(sRangeDir) + ". W.M."
else:
    sLongMapTitle = str(sSectionText) + " SEC." + str(sSection) + " T." + str(sTownship) + str(sTP)  + str(sTownDir) + ". R." + str(sRange) + str(sRP)  + str(sRangeDir) + ". W.M."
    if str(sSectionText)=="":
        sLongMapTitle = "SECTION " + str(sSection) + " T." + str(sTownship) + str(sTP) + str(sTownDir) + ". R." + str(sRange) + str(sRP)  + str(sRangeDir) + ". W.M."
    if str(sSection)=="":
        sLongMapTitle = "T." + str(sTP) + str(sTownship) + str(sTownDir) + ". R." + str(sRange) + str(sRangeDir) + ". W.M."


#MODIFY TITLE FOR NON-STANDARD MAPS
if sMapType == "S":
    sLongMapTitle = "SUPPLEMENTAL MAP NO. " + str(sMN) + "\n" + sLongMapTitle
if sMapType == "D":
    sLongMapTitle = "DETAIL MAP NO. " + str(sMN) + "\n" + sLongMapTitle
if sMapType == "T":
    sLongMapTitle = "SHEET NO. " + str(sMN) + "\n" + sLongMapTitle


#REPOSITION AND MODIFY PAGE ELEMENTS
for elm in MAP.ListLayoutElements(MXD):
    #TEXT ELEMENTS
    if elm.name=="MapNumber":
        elm.text = MapNumber
    if elm.name == "MainMapTitle":
        elm.text = sLongMapTitle
        elm.elementPositionX = TitleX
        elm.elementPositionY = TitleY
    if elm.name == "CountyName":
        elm.text = PageConfig.CountyName
        elm.elementPositionX = TitleX
        elm.elementPositionY = TitleY - CountyNameDist
    if elm.name == "MainMapScale":
        elm.text = sMapScale
        elm.elementPositionX = TitleX
        elm.elementPositionY = TitleY - MapScaleDist

    if elm.name == "UpperLeftMapNum":
        elm.text = shortMapTitle
    if elm.name == "UpperRightMapNum":
        elm.text = shortMapTitle
        elm.elementPositionX = URCornerNumX
        elm.elementPositionY = URCornerNumY
    if elm.name == "LowerLeftMapNum":
        elm.text = shortMapTitle
    if elm.name == "LowerRightMapNum":
        elm.text = shortMapTitle
        elm.elementPositionX = LRCornerNumX
        elm.elementPositionY = LRCornerNumY

    if elm.name == "smallMapTitle":
        elm.text = sLongMapTitle
    if elm.name == "smallMapScale":
        elm.text = sMapScale
    if elm.name == "PlotDate":
        now = datetime.datetime.now()
        elm.text = str("%d/%d/%d"%(now.month, now.day, now.year))
        elm.elementPositionX = DateX
        elm.elementPositionY = DateY
    if elm.name == "Disclaimer" or elm.name == "DisclaimerBox":
        elm.elementPositionX = DisclaimerX
        elm.elementPositionY = DisclaimerY

    #PAGE ELEMENTS
    if elm.name == "MainDF":
        elm.elementHeight = DataFrameMaxY - DataFrameMinY
        elm.elementPositionX = DataFrameMinX
        elm.elementPositionY = DataFrameMinY
        elm.elementWidth = DataFrameMaxX - DataFrameMinX

    if elm.name == "NorthArrow":
        elm.elementPositionX = NorthX
        elm.elementPositionY = NorthY
    if elm.name == "ScaleBar":
        elm.elementPositionX = ScaleBarX
        elm.elementPositionY = ScaleBarY

    if elm.name == "CanMapNumber":
        elm.text = " " #-- Important that this element has some text in it (event just a single space) so ArcMap does not "lose" it.
        cancelledElm2 = None
        if len(MAP.ListLayoutElements(MXD, "TEXT_ELEMENT", "CanMapNumber2"))>0:
            cancelledElm2 = MAP.ListLayoutElements(MXD, "TEXT_ELEMENT", "CanMapNumber2")[0]
            cancelledElm2.text = " " #-- Important that this element has some text in it (event just a single space) so ArcMap does not "lose" it.
        cancelledElm3 = None
        if len(MAP.ListLayoutElements(MXD, "TEXT_ELEMENT", "CanMapNumber3"))>0:
            cancelledElm3 = MAP.ListLayoutElements(MXD, "TEXT_ELEMENT", "CanMapNumber3")[0]
            cancelledElm3.text = " " #-- Important that this element has some text in it (event just a single space) so ArcMap does not "lose" it.

        n = 0
        maxRows = PageConfig.MaxCancelledRows

        if cancelledRow==None:
            arcpy.AddMessage(" ")
            arcpy.AddMessage("------- WARNING -------")
            arcpy.AddMessage("THERE WERE NO CANCELED NUMBERS FOR THIS MAPNUMBER.  IF THERE ARE SUPPOSED TO BE THEN MAKE SURE YOU DO NOT HAVE ANY RECORDS CURRENTLY SELECTED OR HAVE A DEFINITION QUERY SET ON YOUR CANCELLED NUMBERS TABLE.")
            arcpy.AddMessage(" ")

        while cancelledRow:
            #-- If there is not a second text box for Cancelled Numbers force into the first text box.
            if n >= maxRows and cancelledElm2 == None:
                n = 0
            #-- If there is not a third text box for Cancelled Numbers force into the second text box.
            if n >= (maxRows*2) and cancelledElm2 != None and cancelledElm3 == None:
                n = maxRows

            if n < maxRows:
                elm.text += cancelledRow.Taxlot + "\n"
            elif n >= maxRows and n < (maxRows*2):
                cancelledElm2.text += cancelledRow.Taxlot + "\n"
            else:
                cancelledElm3.text += cancelledRow.Taxlot + "\n"

            n += 1
            cancelledRow = cancelledCursor.next()

        elm.text = PageConfig.CancelledNumberPrefix + "\n" + elm.text.strip() if elm.text != " " else " " #-- Important that this element has some text in it (event just a single space) so ArcMap does not "lose" it.
        elm.elementPositionX = CancelNumX
        elm.elementPositionY = CancelNumY
        if cancelledElm2 != None:
            cancelledElm2.text = cancelledElm2.text.strip() if cancelledElm2.text != " " else " " #-- Important that this element has some text in it (event just a single space) so ArcMap does not "lose" it.
            cancelledElm2.elementPositionX = elm.elementPositionX + cancelledElm2.elementWidth + .05
            cancelledElm2.elementPositionY = elm.elementPositionY
            if cancelledElm3 != None:
                cancelledElm3.text = cancelledElm3.text.strip() if cancelledElm3.text != " " else " " #-- Important that this element has some text in it (event just a single space) so ArcMap does not "lose" it.
                cancelledElm3.elementPositionX = cancelledElm2.elementPositionX + cancelledElm3.elementWidth + .05
                cancelledElm3.elementPositionY = elm.elementPositionY



#MODIFY MAIN DATAFRAME PROPERTIES
mainDF.extent = mapExtent
mainDF.scale = MapScale
mainDF.rotation = MapAngle

#MODIFY LOCATOR DATAFRAME
if locatorDF != None:
    mapIndexLayer = MAP.ListLayers(MXD, "MapIndex", locatorDF)[0]
    locatorWhere = "MapNumber = '" + mapNumber + "'"
    arcpy.management.SelectLayerByAttribute(mapIndexLayer, "NEW_SELECTION", locatorWhere)

#MODIFY SECTIONS DATAFRAME
if sectDF != None:
    sectionsLayer = MAP.ListLayers(MXD, "Sections_Select", sectDF)[0]
    sectionsLayer.definitionQuery = "[SectionNum] = " + str(sSection)

#MODIFY QUARTER SECTIONS DATAFRAME
if qSectDF != None:
    qSectionsLayer = MAP.ListLayers(MXD, "QtrSections_Select", qSectDF)[0]
    qSectionsLayer.definitionQuery = ""

    if sQtr == "A" and sQtrQtr == "0":
        qSectionsLayer.definitionQuery = "[QSectName] = 'A' or [QSectName]= 'AA' or [QSectName]= 'AB' or [QSectName]= 'AC' or [QSectName]= 'AD'"
    elif sQtr == "A" and sQtrQtr == "A":
        qSectionsLayer.definitionQuery = "[QSectName] = 'AA'"
    elif sQtr == "A" and sQtrQtr == "B":
        qSectionsLayer.definitionQuery = "[QSectName] = 'AB'"
    elif sQtr == "A" and sQtrQtr == "C":
        qSectionsLayer.definitionQuery = "[QSectName] = 'AC'"
    elif sQtr == "A" and sQtrQtr == "D":
        qSectionsLayer.definitionQuery = "[QSectName] = 'AD'"

    if sQtr == "B" and sQtrQtr == "0":
        qSectionsLayer.definitionQuery = "[QSectName] = 'B' or [QSectName]= 'BA' or [QSectName]= 'BB' or [QSectName]= 'BC' or [QSectName]= 'BD'"
    elif sQtr == "B" and sQtrQtr == "A":
        qSectionsLayer.definitionQuery = "[QSectName] = 'BA'"
    elif sQtr == "B" and sQtrQtr == "B":
        qSectionsLayer.definitionQuery = "[QSectName] = 'BB'"
    elif sQtr == "B" and sQtrQtr == "C":
        qSectionsLayer.definitionQuery = "[QSectName] = 'BC'"
    elif sQtr == "B" and sQtrQtr == "D":
        qSectionsLayer.definitionQuery = "[QSectName] = 'BD'"

    if sQtr == "C" and sQtrQtr == "0":
        qSectionsLayer.definitionQuery = "[QSectName] = 'C' or [QSectName]= 'CA' or [QSectName]= 'CB' or [QSectName]= 'CC' or [QSectName]= 'CD'"
    elif sQtr == "C" and sQtrQtr == "A":
        qSectionsLayer.definitionQuery = "[QSectName] = 'CA'"
    elif sQtr == "C" and sQtrQtr == "B":
        qSectionsLayer.definitionQuery = "[QSectName] = 'CB'"
    elif sQtr == "C" and sQtrQtr == "C":
        qSectionsLayer.definitionQuery = "[QSectName] = 'CC'"
    elif sQtr == "C" and sQtrQtr == "D":
        qSectionsLayer.definitionQuery = "[QSectName] = 'CD'"

    if sQtr == "D" and sQtrQtr == "0":
        qSectionsLayer.definitionQuery = "[QSectName] = 'D' or [QSectName]= 'DA' or [QSectName]= 'DB' or [QSectName]= 'DC' or [QSectName]= 'DD'"
    elif sQtr == "D" and sQtrQtr == "A":
        qSectionsLayer.definitionQuery = "[QSectName] = 'DA'"
    elif sQtr == "D" and sQtrQtr == "B":
        qSectionsLayer.definitionQuery = "[QSectName] = 'DB'"
    elif sQtr == "D" and sQtrQtr == "C":
        qSectionsLayer.definitionQuery = "[QSectName] = 'DC'"
    elif sQtr == "D" and sQtrQtr == "D":
        qSectionsLayer.definitionQuery = "[QSectName] = 'DD'"


#REFRESH THE CURRENT MAP DISPLAY/LAYOUT
del mapIndexCursor, mapIndexRow

arcpy.RefreshActiveView()
