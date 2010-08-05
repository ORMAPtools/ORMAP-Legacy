import arcpy, datetime, string, sys
import arcpy.mapping as MAP

myInput = arcpy.GetParameterAsText(0)
myMapNumberList = myInput.split(";")
#myMapNumberList = ["7.4.1"]

arcpy.AddMessage(" ")

mainLoopCount = 0
for map in myMapNumberList:
    myMapNumber = myMapNumberList[mainLoopCount]
    arcpy.AddMessage("Processing: " + myMapNumber)

    #SET DEFAULT PAGE LAYOUT LOCATIONS
    DataFrameMinX = 0.25
    DataFrameMinY = 0.25
    DataFrameMaxX = 17.75
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

    #REFERENCE MAP DOCUMENT
    myMXD = MAP.MapDocument("CURRENT")
    #myMXD = MAP.MapDocument(r"C:\Active\ArcPY\ClientProjects\ORMAP_Mapping\MXD\MapProduction18x24_UsingPython.mxd")

    #REFERENCE DATAFRAME
    mainDF = MAP.ListDataFrames(myMXD, "MainDF")[0]
    locatorDF = MAP.ListDataFrames(myMXD, "LocatorDF")[0]
    sectDF = MAP.ListDataFrames(myMXD, "SectionsDF")[0]
    qSectDF = MAP.ListDataFrames(myMXD, "QSectionsDF")[0]

    #REFERENCE MAPINXEX LAYER
    for lyr in MAP.ListLayers(myMXD, "MapIndex", mainDF):
        if lyr.name == "MapIndex":
            mapIndexCursor = arcpy.SearchCursor(lyr.dataSource, "[MapNumber] = '" + myMapNumber + "'")

    #REFERENCE PAGELAYOUT TABLE
    pageLayoutTable = MAP.ListTableViews(myMXD)[0]
    pageLayoutCursor = arcpy.SearchCursor(pageLayoutTable.dataSource, "[MapNumber] = '" + myMapNumber + "'")
    
    mapIndexRow = mapIndexCursor.next()         #MAIN LOOP - THROUGH EACH MAPINDEX POLYGON     
    while mapIndexRow:                              
        #COLLECT MAP INDEX POLYGON INFORMATION
        #GET FEATURE EXTENT
        geom = mapIndexRow.shape
        featureExtent = geom.extent

        #GET OTHER TABLE ATTRIBUTES
        MapScale = mapIndexRow.MapScale
        MapNumber = mapIndexRow.MapNumber
        ORMapNum = mapIndexRow.ORMapNum
        CityName = mapIndexRow.CityName
        
        #READ INFORMATION FROM PAGELAYOUT TABLE
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
        myLayers = MAP.ListLayers(myMXD, "", mainDF)
        for myLayer in myLayers:
            if myLayer.name == "LotsAnno":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "PlatsAnno":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "TaxCodeAnno":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "TaxlotNumberAnno":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "TaxlotAcreageAnno":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0010scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0020scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0030scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0040scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0050scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0100scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0200scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0400scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno0800scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Anno2000scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Corner - Above":
                myLayer.definitionQuery = ""
            if myLayer.name == "TaxCodeLines - Above":
                myLayer.definitionQuery = ""
            if myLayer.name == "TaxLotLines - Above":
                myLayer.definitionQuery = "[LineType] = 8 or [LineType] = 14"
            if myLayer.name == "ReferenceLines - Above":
                myLayer.definitionQuery = ""
            if myLayer.name == "CartographicLines - Above":
                myLayer.definitionQuery = ""
            if myLayer.name == "WaterLines - Above":
                myLayer.definitionQuery = ""
            if myLayer.name == "Water":
                myLayer.definitionQuery = ""
            if myLayer.name == "MapIndex - SeeMaps":
                myLayer.definitionQuery = ""  ## NEED TO IMPLEMENT
            if myLayer.name == "MapIndex - Mask":
                myLayer.definitionQuery = "[MapNumber] <> '" + myMapNumber + "' OR [MapNumber] is NULL OR [MapNumber] = ''"
            if myLayer.name == "Corner - Below":
                myLayer.definitionQuery = ""
            if myLayer.name == "TaxCodeLines - Below":
                myLayer.definitionQuery = "[CurrentLine] = 'Y'"
            if myLayer.name == "TaxlotLines - Below":
                myLayer.definitionQuery = ""
            if myLayer.name == "ReferenceLines - Below":
                myLayer.definitionQuery = ""
            if myLayer.name == "CartographicLines - Below":
                myLayer.definitionQuery = ""
            if myLayer.name == "WaterLines - Below":
                myLayer.definitionQuery = ""
            if myLayer.name == "Water - Below":
                myLayer.definitionQuery = ""

        #PARSE ORMAPNUMBER TO DEVELOP MAP TITLES
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
        for elm in MAP.ListLayoutElements(myMXD):
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
                elm.text = myMapNumber
            
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
        aLayer = MAP.ListLayers(myMXD, "MapIndex", locatorDF)[0]
        locatorWhere = "[MapNumber] = '" + myMapNumber + "'"
        arcpy.management.SelectLayerByAttribute(aLayer, "NEW_SELECTION", locatorWhere) 

        #MODIFY SECTIONS DATAFRAME
        bLayer = MAP.ListLayers(myMXD, "Sections_Select", sectDF)[0]
        bLayer.definitionQuery = "[SectionNum] = " + str(sSection)

        #MODIFY QUARTER SECTIONS DATAFRAME
        cLayer = MAP.ListLayers(myMXD, "QtrSections_Select", qSectDF)[0]
        cLayer.definitionQuery = ""
        
        if sQtr == "A" and sQtrQtr == "0":
            cLayer.definitionQuery = "[QSectName] = 'A' or [QSectName]= 'AA' or [QSectName]= 'AB' or [QSectName]= 'AC' or [QSectName]= 'AD'"
        elif sQtr == "A" and sQtrQtr == "A":
            cLayer.definitionQuery = "[QSectName] = 'AA'"
        elif sQtr == "A" and sQtrQtr == "B":
            cLayer.definitionQuery = "[QSectName] = 'AB'"
        elif sQtr == "A" and sQtrQtr == "C":
            cLayer.definitionQuery = "[QSectName] = 'AC'"
        elif sQtr == "A" and sQtrQtr == "D":
            cLayer.definitionQuery = "[QSectName] = 'AD'"

        if sQtr == "B" and sQtrQtr == "0":
            cLayer.definitionQuery = "[QSectName] = 'B' or [QSectName]= 'BA' or [QSectName]= 'BB' or [QSectName]= 'BC' or [QSectName]= 'BD'"
        elif sQtr == "B" and sQtrQtr == "A":
            cLayer.definitionQuery = "[QSectName] = 'BA'"
        elif sQtr == "B" and sQtrQtr == "B":
            cLayer.definitionQuery = "[QSectName] = 'BB'"
        elif sQtr == "B" and sQtrQtr == "C":
            cLayer.definitionQuery = "[QSectName] = 'BC'"
        elif sQtr == "B" and sQtrQtr == "D":
            cLayer.definitionQuery = "[QSectName] = 'BD'"

        if sQtr == "C" and sQtrQtr == "0":
            cLayer.definitionQuery = "[QSectName] = 'C' or [QSectName]= 'CA' or [QSectName]= 'CB' or [QSectName]= 'CC' or [QSectName]= 'CD'"
        elif sQtr == "C" and sQtrQtr == "A":
            cLayer.definitionQuery = "[QSectName] = 'CA'"
        elif sQtr == "C" and sQtrQtr == "B":
            cLayer.definitionQuery = "[QSectName] = 'CB'"
        elif sQtr == "C" and sQtrQtr == "C":
            cLayer.definitionQuery = "[QSectName] = 'CC'"
        elif sQtr == "C" and sQtrQtr == "D":
            cLayer.definitionQuery = "[QSectName] = 'CD'"

        if sQtr == "D" and sQtrQtr == "0":
            cLayer.definitionQuery = "[QSectName] = 'D' or [QSectName]= 'DA' or [QSectName]= 'DB' or [QSectName]= 'DC' or [QSectName]= 'DD'"
        elif sQtr == "D" and sQtrQtr == "A":
            cLayer.definitionQuery = "[QSectName] = 'DA'"
        elif sQtr == "D" and sQtrQtr == "B":
            cLayer.definitionQuery = "[QSectName] = 'DB'"
        elif sQtr == "D" and sQtrQtr == "C":
            cLayer.definitionQuery = "[QSectName] = 'DC'"
        elif sQtr == "D" and sQtrQtr == "D":
            cLayer.definitionQuery = "[QSectName] = 'DD'"

        mapIndexRow = mapIndexCursor.next()         #RETURN TO TOP TO PROCESS NEXT MAPINDEX POLYGON

    arcpy.gp.refreshgraphics()

    #EXPORT TO PDF - OUTPUT IS RELATIVE TO SCRIPT LOCATION
    scriptPath = sys.path[0]
    pdfOutputPath = scriptPath[:-7] + "MXD/Output/" + shortMapTitle + ".pdf"
    MAP.ExportToPDF(myMXD, pdfOutputPath)
    mainLoopCount = mainLoopCount + 1

print "SCRIPT COMPLETED SUCCESSFULLY"
