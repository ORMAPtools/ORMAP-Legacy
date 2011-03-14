# ---------------------------------------------------------------------------
# Ormap_SavePageLayoutElements.py
# Created by: Shad Campbell
# Date: 3/11/2011
# Updated by: 
# Description: This script captures the current x/y locations of the map elements
# and stores them in the PageLayoutElements table.  Once stored a map uses these
# values rather than the ones stored in the configuration files.
# ---------------------------------------------------------------------------

import sys, arcpy, arcpy.mapping as MAP
import ORMAP_LayersConfig as OrmapLayers

#IMPORT PARAMETERS
MapNumber = arcpy.GetParameterAsText(0)

#REFERENCE MAP DOCUMENT
MXD = MAP.MapDocument("CURRENT")

#COLLECT DATAFRAME INFORMATION
if len(MAP.ListDataFrames(MXD, "MainDF"))>0:
    mainDF = MAP.ListDataFrames(MXD, "MainDF")[0]
else:
    sys.exit("Unable to find a DataFrame named mainDF.  Please make sure default dataframe is named 'mainDF'.")

#REFERENCE PAGELAYOUT TABLE
pageLayoutRow = None
if len(MAP.ListTableViews(MXD, OrmapLayers.PAGELAYOUT_TABLE, mainDF))>0:
    pageLayoutTable = MAP.ListTableViews(MXD, OrmapLayers.PAGELAYOUT_TABLE, mainDF)[0]
    pageLayoutCursor = arcpy.SearchCursor(pageLayoutTable.name, "MapNumber = '" + MapNumber + "'")
    pageLayoutRow = pageLayoutCursor.next()
else:
    sys.exit("Unable to find PageLayoutTable table.  Please check your TOC and config file.")


#DETERMINE WHETHER TO INSERT OR UPDATE PAGELAYOUT ROW.
if pageLayoutRow == None:
    theCursor = arcpy.InsertCursor(pageLayoutTable.name)
    theRow = theCursor.newRow()
else:
    theCursor = arcpy.UpdateCursor(pageLayoutTable.name, "MapNumber = '" + MapNumber + "'")
    theRow = theCursor.next()

#LOOP THROUGH THE PAGE ELEMENTS AND STORE THE VALUES IF POSSIBLE
for elm in MAP.ListLayoutElements(MXD):

    theRow.MapNumber = MapNumber

    if len(arcpy.ListFields(pageLayoutTable.name,"MapAngle")) > 0:
        theRow.MapAngle = mainDF.rotation
   
    if elm.name == "MainDF":
        if len(arcpy.ListFields(pageLayoutTable.name,"DataFrameMinX")) > 0:
            theRow.DataFrameMinX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"DataFrameMinY")) > 0:
            theRow.DataFrameMinY = elm.elementPositionY
        if len(arcpy.ListFields(pageLayoutTable.name,"DataFrameMaxX")) > 0:
            theRow.DataFrameMaxX = elm.elementPositionX + elm.elementWidth
        if len(arcpy.ListFields(pageLayoutTable.name,"DataFrameMaxY")) > 0:
            theRow.DataFrameMaxY = elm.elementPositionY + elm.elementHeight
        mapExtent = mainDF.extent
        if len(arcpy.ListFields(pageLayoutTable.name,"MapPositionX")) > 0:
            theRow.MapPositionX = (((mapExtent.XMax - mapExtent.XMin) / 2) + mapExtent.XMin)
        if len(arcpy.ListFields(pageLayoutTable.name,"MapPositionY")) > 0:
            theRow.MapPositionY = (((mapExtent.YMax - mapExtent.YMin) / 2) + mapExtent.YMin)

    if elm.name == "MainMapTitle":
        if len(arcpy.ListFields(pageLayoutTable.name,"TitleX")) > 0:
            theRow.titleX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"titleY")) > 0:
            theRow.titleY = elm.elementPositionY

    if elm.name == "NorthArrow":
        if len(arcpy.ListFields(pageLayoutTable.name,"NorthX")) > 0:
            theRow.NorthX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"NorthY")) > 0:
            theRow.NorthY = elm.elementPositionY

    if elm.name == "ScaleBar":
        if len(arcpy.ListFields(pageLayoutTable.name,"ScaleBarX")) > 0:
            theRow.ScaleBarX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"ScaleBarY")) > 0:
            theRow.ScaleBarY = elm.elementPositionY

    if elm.name == "UpperRightMapNum":
        if len(arcpy.ListFields(pageLayoutTable.name,"URCornerNumX")) > 0:
            theRow.URCornerNumX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"URCornerNumY")) > 0:
            theRow.URCornerNumY = elm.elementPositionY

##    if elm.name == "UpperLeftMapNum":
##        if len(arcpy.ListFields(pageLayoutTable.name,"ULCornerNumX")) > 0:
##            theRow.ULCornerNumX = elm.elementPositionX
##        if len(arcpy.ListFields(pageLayoutTable.name,"ULCornerNumY")) > 0:
##            theRow.ULCornerNumY = elm.elementPositionY

    if elm.name == "LowerRightMapNum":
        if len(arcpy.ListFields(pageLayoutTable.name,"LRCornerNumX")) > 0:
            theRow.LRCornerNumX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"LRCornerNumY")) > 0:
            theRow.LRCornerNumY = elm.elementPositionY

##    if elm.name == "LowerLeftMapNum":
##        if len(arcpy.ListFields(pageLayoutTable.name,"LLCornerNumX")) > 0:
##            theRow.LLCornerNumX = elm.elementPositionX
##        if len(arcpy.ListFields(pageLayoutTable.name,"LLCornerNumY")) > 0:
##            theRow.LLCornerNumY = elm.elementPositionY

    if elm.name == "PlotDate":
        if len(arcpy.ListFields(pageLayoutTable.name,"DateX")) > 0:
            theRow.DateX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"DateY")) > 0:
            theRow.DateY = elm.elementPositionY

    if elm.name == "Disclaimer":
        if len(arcpy.ListFields(pageLayoutTable.name,"DisclaimerX")) > 0:
            theRow.DisclaimerX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"DisclaimerY")) > 0:
            theRow.DisclaimerY = elm.elementPositionY

    if elm.name == "CanMapNumber":
        if len(arcpy.ListFields(pageLayoutTable.name,"CancelNumX")) > 0:
            theRow.CancelNumX = elm.elementPositionX
        if len(arcpy.ListFields(pageLayoutTable.name,"CancelNumY")) > 0:
            theRow.CancelNumY = elm.elementPositionY

#INSERT OR UPDATE ROW.
if pageLayoutRow == None:
    theCursor.insertRow(theRow)
else:
    theCursor.updateRow(theRow)
    

del pageLayoutTable, pageLayoutCursor, pageLayoutRow, theCursor, theRow


