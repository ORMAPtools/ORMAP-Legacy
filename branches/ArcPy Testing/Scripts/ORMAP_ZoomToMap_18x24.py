import arcpy, arcgisscripting, datetime, string, sys
import arcpy.mapping as MAP

gp = arcgisscripting.create()

#IMPORT PARAMETERS - SHOULD ONLY BE A SINGLE STRING ITEM
MapNumber = arcpy.GetParameterAsText(0)
#MapNumber = "7.4.1A"

#REFERENCE MAP DOCUMENT
MXD = MAP.MapDocument("CURRENT")
#MXD = MAP.MapDocument(r"C:\Active\ArcPY\ClientProjects\ORMAP_Mapping\MXD\MapProduction18x24_UsingPython.mxd")

#REFERENCE EACH DATAFRAME
mainDF = MAP.ListDataFrames(MXD, "MainDF")[0]
locatorDF = MAP.ListDataFrames(MXD, "LocatorDF")[0]
sectDF = MAP.ListDataFrames(MXD, "SectionsDF")[0]
qSectDF = MAP.ListDataFrames(MXD, "QSectionsDF")[0]


#REFERENCE MAPINXEX LAYER
for lyr in MAP.ListLayers(MXD, "MapIndex", mainDF):
    if lyr.name == "MapIndex":
        mapIndexCursor = arcpy.SearchCursor(lyr.dataSource, "[MapNumber] = '" + MapNumber + "'")

#REFERENCE PAGELAYOUT TABLE
pageLayoutTable = MAP.ListTableViews(MXD, "PageLayoutElements", mainDF)[0]
pageLayoutCursor = arcpy.SearchCursor(pageLayoutTable.dataSource, "[MapNumber] = '" + MapNumber + "'")

mapIndexRow = mapIndexCursor.next()
while mapIndexRow:
    
    #SET DEFAULT PAGE LAYOUT LOCATIONS
    DataFrameMinX = 0.25
    DataFrameMinY = 0.25
    DataFrameMaxX = 19.75
    DataFrameMaxY = 17.75
    MapPositionX = 1
    MapPositionY = 1
    MapAngle = 0
    TitleX = 14
    TitleY = 17.5
    DisClaimerX = 2.25
    DisClaimerY = 17.25
    CancelNumX = 17.75
    CancelNumY = 15
    DateX = 19
    DateY = 7
    URCornerNumX = 19.5
    URCornerNumY = 17.8
    LRCornerNumX = 1
    LRCornerNumY = 1
    ScaleBarX = 5.5
    ScaleBarY = 0.5
    NorthX = 0.5
    NorthY = 0.4

    #MISC Relative Map Distances
    CountyNameDist = 0.4
    MapScaleDist = 0.8
    
    #COLLECT MAP INDEX POLYGON INFORMATION
    #GET FEATURE EXTENT
    geom = mapIndexRow.shape
    featureExtent = geom.extent

    #GET OTHER TABLE ATTRIBUTES
    MapScale = mapIndexRow.MapScale
    mapNumber = mapIndexRow.MapNumber
    ORMapNum = mapIndexRow.ORMapNum
    CityName = mapIndexRow.CityName

    arcpy.AddMessage("")
    arcpy.AddMessage("Processing: " + MapNumber)
    arcpy.AddMessage("")
    
    #READ NON-DEFAULT INFORMATION FROM PAGELAYOUT TABLE
    pageLayoutRow = pageLayoutCursor.next()
    while pageLayoutRow:

        MapAngle = pageLayoutRow.MapAngle
        DataFrameMinX = pageLayoutRow.DataFrameMinX
        DataFrameMinY = pageLayoutRow.DataFrameMinY
        DataFrameMaxX = pageLayoutRow.DataFrameMaxX
        DataFrameMaxY = pageLayoutRow.DataFrameMaxY
        MapPositionX = pageLayoutRow.MapPositionX
        MapPositionY = pageLayoutRow.MapPositionY
        MapAngle = pageLayoutRow.MapAngle
        TitleX = pageLayoutRow.TitleX
        TitleY = pageLayoutRow.TitleY
        DisClaimerX = pageLayoutRow.DisClaimerX
        DisClaimerY = pageLayoutRow.DisClaimerY
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
        
        pageLayoutRow = pageLayoutCursor.next()
    
    #SET QUERY DEFINITIONS FOR EACH LAYER
    for lyr in MAP.ListLayers(MXD, "", mainDF):
        if lyr.name == "LotsAnno":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "PlatsAnno":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "TaxCodeAnno":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "TaxlotNumberAnno":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "TaxlotAcreageAnno":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0010scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0020scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0030scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0040scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0050scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0100scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0200scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0400scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno0800scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Anno2000scale":
            lyr.definitionQuery = "[MapNumber] = '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Corner - Above":
            lyr.definitionQuery = ""
        if lyr.name == "TaxCodeLines - Above":
            lyr.definitionQuery = ""
        if lyr.name == "TaxLotLines - Above":
            lyr.definitionQuery = "[LineType] = 8 or [LineType] = 14"
        if lyr.name == "ReferenceLines - Above":
            lyr.definitionQuery = ""
        if lyr.name == "CartographicLines - Above":
            lyr.definitionQuery = ""
        if lyr.name == "WaterLines - Above":
            lyr.definitionQuery = ""
        if lyr.name == "Water":
            lyr.definitionQuery = ""
        if lyr.name == "MapIndex - SeeMaps":
            lyr.definitionQuery = ""  ## NEED TO IMPLEMENT WITH SPATIAL QUERY
        if lyr.name == "MapIndex - Mask":
            lyr.definitionQuery = "[MapNumber] <> '" + mapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
        if lyr.name == "Corner - Below":
            lyr.definitionQuery = ""
        if lyr.name == "TaxCodeLines - Below":
            lyr.definitionQuery = "[CurrentLine] = 'Y'"
        if lyr.name == "TaxlotLines - Below":
            lyr.definitionQuery = ""
        if lyr.name == "ReferenceLines - Below":
            lyr.definitionQuery = ""
        if lyr.name == "CartographicLines - Below":
            lyr.definitionQuery = ""
        if lyr.name == "WaterLines - Below":
            lyr.definitionQuery = ""
        if lyr.name == "Water - Below":
            lyr.definitionQuery = ""

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
        sMN = sMN & sMapNum1
    if sMapNum2 <> "0":
        sMN = sMN & sMapNum2
    if sMapNum3 <> "0":
        sMN = sMN & sMapNum3
     
    #GENERATE TEXT FOR SHORT TITLES (UR & LR)
    shortMapTitle = sTown1 + sTown2 + " " + sRange1 + sRange2 + " " + sSection1 + sSection2
    if sQtr <> "0":
        shortMapTitle = shortMapTitle + " " + sQtr
    if sQtrQtr <> "0":
        shortMapTitle = shortMapTitle + " " + sQtrQtr
     
    #GENERATE TEXT FOR LONG MAP TITLE
    sLongMapTitle = ""
     
    #CREATE MAP TITLE BASED ON SCALE FORMATS PROVIDED BY DOR.
    if MapScale == 24000:
        sLongMapTitle = "T." + str(sTP) + str(sTownship) + str(sTownDir) + ". R." + str(sRange) + str(sRangeDir) + ". W.M."
        sMapScale = "1\" = 2000'"
    elif MapScale == 4800:
        sLongMapTitle = "SECTION " + str(sSection) + " T." + str(sTP) + str(sTownship) + str(sTownDir) + ". R." + str(sRange) + str(sRangeDir) + ". W.M."
        sMapScale = "1\" = 400'"
    elif MapScale == 2400:
        sLongMapTitle = str(sSectionText) + " SEC." + str(sSection) + " T." + str(sTP) + str(sTownship) + str(sTownDir) + ". R." + str(sRange) + str(sRangeDir) + ". W.M."
        sMapScale = "1\" = 200'"
    elif MapScale == 1200:
        sLongMapTitle = str(sSectionText) + " SEC." + str(sSection) + " T." + str(sTP) + str(sTownship) + str(sTownDir) + ". R." + str(sRange) + str(sRangeDir) + ". W.M."
        sMapScale = "1\" = 100'"
    else:
        sLongMapTitle = "MapTitle Format not defined for scales < 100, 800, 1000"
    
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
        if elm.name == "MainMapTitle":
            elm.text = sLongMapTitle
            elm.elementPositionX = TitleX
            elm.elementPositionY = TitleY
        if elm.name == "CountyName":
            elm.text = "Polk County"
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
        if elm.name == "LowerLeftMapNum":
            elm.text = shortMapTitle
        if elm.name == "LowerRightMapNum":
            elm.text = shortMapTitle
        if elm.name == "CanMapNumber":
            elm.text = shortMapTitle
        if elm.name == "smallMapTitle":
            elm.text = sLongMapTitle
        if elm.name == "smallMapScale":
            elm.text = sMapScale
        if elm.name == "PlotDate":
            now = datetime.datetime.now()
            elm.text = str(now.date())
        if elm.name == "MapNumber":
            elm.text = mapNumber
            
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
    
    #MODIFY MAIN DATAFRAME PROPERTIES
    mainExtent = arcpy.Extent(featureExtent.XMin, featureExtent.YMin, featureExtent.XMax, featureExtent.YMax)
    mainDF.extent = mainExtent
    mainDF.scale = MapScale
    mainDF.rotation = MapAngle

    #MODIFY LOCATOR DATAFRAME
    mapIndexLayer = MAP.ListLayers(MXD, "MapIndex", locatorDF)[0]
    locatorWhere = "[MapNumber] = '" + mapNumber + "'"
    arcpy.management.SelectLayerByAttribute(mapIndexLayer, "NEW_SELECTION", locatorWhere) 
    
    #MODIFY SECTIONS DATAFRAME
    sectionsLayer = MAP.ListLayers(MXD, "Sections_Select", sectDF)[0]
    sectionsLayer.definitionQuery = "[SectionNum] = " + str(sSection)

    #MODIFY QUARTER SECTIONS DATAFRAME
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
    
    mapIndexRow = mapIndexCursor.next()     #RETURN TO TOP OF MAIN LOOP

#REFRESH THE CURRENT MAP DISPLAY/LAYOUT
arcpy.RefreshActiveView()
