###  Abbreviated ORMAP Script simply to show sample arcpy.mapping syntax - not expected to work as is.
###  Jeff Barrette
###  April 16, 2009

import arcpy, datetime, os, sys, shutil, string
import arcpy.mapping as MAP

#RELATIVE PATH SETUP
scriptPath = sys.path[0]
outputPath = scriptPath[:-7] + "MXD/Output/"

#REPLACE EXITING PDF WITH TEMPLATE COPY
shutil.copyfile(outputPath + "MapSeriesBook_blank.pdf", outputPath + "MapSeriesBook.pdf")
pdf = arcpy.mapping.PDFDocumentOpen(outputPath + "MapSeriesBook.pdf")

#READ INPUT ARGUMENTS
myInput = arcpy.GetParameterAsText(0)
myMapNumberList = myInput.split(";")

arcpy.AddMessage(" ")

#MAIN LOOP - FOR EACH PAGE IN LIST
for eachMap in myMapNumberList:
    arcpy.AddMessage("Processing: " + eachMap)

    #SET DEFAULT PAGE ELEMENT LOCATIONS
    DataFrameMinX = 0.25
    DataFrameMinY = 0.25
    DataFrameMaxX = 17.75
    DataFrameMaxY = 17.75
    #ETC - a whole bunch cut our here

    #REFERENCE MAP DOCUMENT
    #myMXD = MAP.MapDocument("CURRENT")
    myMXD = MAP.MapDocument("F:/Active/ArcPY/ClientProjects/ORMAP_Mapping/MXD/MapProduction18x24_UsingPython.mxd")

    #REFERENCE DATAFRAMES
    locatorDF = MAP.ListDataFrames(myMXD)[2]
    sectDF = MAP.ListDataFrames(myMXD)[1]
    qSectDF = MAP.ListDataFrames(myMXD)[0]
    mainDF = MAP.ListDataFrames(myMXD)[3]

    #REFERENCE MAPINXEX LAYER
    for myMapNumber in MAP.ListLayers(myMXD, "", mainDF):
        if myMapNumber.name == "MapIndex":
            mapIndexCursor = arcpy.SearchCursor(lyr.dataSource, "[MapNumber] = '" + myMapNumber + "'")

    #REFERENCE PAGELAYOUT TABLE
    pageLayoutTable = MAP.ListTableViews(myMXD)[0]
    pageLayoutCursor = arcpy.SearchCursor(pageLayoutTable.dataSource, "[MapNumber] = '" + myMapNumber + "'")

    mapIndexRow = mapIndexCursor.next()         #LOOP - THROUGH EACH MAPINDEX POLYGON     
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
            #ETC - a whole bunch cut our here
            
            pageLayoutRow = pageLayoutCursor.next()

        #SET QUERY DEFINITIONS FOR EACH LAYER
        myLayers = MAP.ListLayers(myMXD, "", mainDF)
        for myLayer in myLayers:
            if myLayer.name == "Anno0100scale":
                myLayer.definitionQuery = "[MapNumber] = '" + myMapNumber
            if myLayer.name == "TaxLotLines - Above":
                myLayer.definitionQuery = "[LineType] = 8 or [LineType] = 14"
            #ETC - a whole bunch cut our here

        #DEVELOP MAP TITLE
        #PARSE ORMAPNUMBER
        sTown1 = ORMapNum[2]
        sTown2 = ORMapNum[3]
        #ETC - a whole bunch cut our here
        
        #BUILD TOWNSHIP TEXT TO EXCLUDE LEADING ZEROS
        sTownship = ""
        if sTown1 <> "0":
            sTownship = sTown1
        sTownship = sTownship + sTown2

        #ETC - a whole bunch cut our here


        #REPOSITION AND MODIFY PAGE ELEMENTS
        pageElements = MAP.ListElements(myMXD)
        for elm in pageElements:
            #TEXT ELEMENTS
            if elm.name == "MainMapTitle":
                elm.text = sLongMapTitle
                elm.setPosition(TitleX, TitleY)
            if elm.name == "CountyName":
                elm.text = "Polk County"
                elm.setPosition(TitleX, TitleY - CountyNameDist)
            if elm.name == "MainMapScale":
                elm.text = sMapScale
                elm.setPosition(TitleX, TitleY - MapScaleDist)
            #ETC - a whole bunch cut our here
                
            #PAGE ELEMENTS
            if elm.name == "NorthArrow":
                elm.setPosition(NorthX ,NorthY)
            if elm.name == "ScaleBar":
                elm.setPosition(ScaleBarX ,ScaleBarY)

        #MODIFY MAIN DATAFRAME PROPERTIES
        mainExtent = arcpy.Extent(featureExtent.xMin, featureExtent.yMin, featureExtent.xMax, featureExtent.yMax)
        mainDF.zoomToExtent(mainExtent)
        mainDF.scale = 45                     #ISSUE - doesn't really have effect on saved output
        mainDF.rotation = MapAngle            #ISSUE - doesn't really have effect on saved output

        #MODIFY LOCATOR DATAFRAME
        aLayer = MAP.ListLayers(myMXD, "", locatorDF)[0]
        locatorWhere = "[MapNumber] = '" + myMapNumber + "'"
        arcpy.management.SelectLayerByAttribute_management(aLayer, "NEW_SELECTION", locatorWhere) 

        #MODIFY SECTIONS DATAFRAME
        bLayer = MAP.ListLayers(myMXD, "", sectDF)[1]
        bLayer.definitionQuery = "[SectionNum] = " + str(sSection)

        #MODIFY QUARTER SECTIONS DATAFRAME
        cLayer = MAP.ListLayers(myMXD, "", qSectDF)[1]
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

        #ETC - a whole bunch cut our here
        
        mapIndexRow = mapIndexCursor.next()

    #arcpy.gp.refreshgraphics()

    #EXPORT TO PDF
    pdfOutputPath = scriptPath[:-7] + "MXD/Output/" + shortMapTitle + ".pdf"
    MAP.ExportToPDF(myMXD, pdfOutputPath)
    pdf.appendPages(pdfOutputPath)
    os.remove(scriptPath[:-7] + "MXD/Output/" + shortMapTitle + ".pdf")

indexPDF = scriptPath[:-7] + "MXD/Output/Index.pdf"
pdf.appendPages(indexPDF)
del pdf
del myMXD

arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
