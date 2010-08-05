import arcpy, datetime, string, sys
import arcpy.mapping as MAP

#myMapNumber = arcpy.GetParameterAsText(0)

#REFERENCE MAP DOCUMENT
MXD = MAP.MapDocument("CURRENT")
#MXD.save()

#COLLECT DATAFRAME INFORMATION
mainDF = MAP.ListDataFrames(MXD, "MainDF")[0]
mapAngle = mainDF.rotation

#COLLECT X,Y COORDINATE INFORMATION FOR PAGE LAYOUT ELEMENTS
for elm in MAP.ListLayoutElements(MXD):
    if elm.name == "MainDF":
        dfMinX = elm.elementPositionX
        dfMinY = elm.elementPositionY
        dfMaxX = elm.elementPositionX + elm.elementWidth
        dfMaxY = elm.elementPositionY + elm.elementHeight
    if elm.name == "MainMapTitle":
        titleX = elm.elementPositionX
        titleY = elm.elementPositionY
    if elm.name == "NorthArrow":
        northX = elm.elementPositionX
        northY = elm.elementPositionY
    if elm.name == "ScaleBar":
        scaleBarX = elm.elementPositionX
        scaleBarY = elm.elementPositionY
    if elm.name == "MapNumber":
        myMapNumber = elm.text

#REFERENCE PAGELAYOUT TABLE
pageLayoutTable = MAP.ListTableViews(MXD, "PageLayoutElements", mainDF)[0]

#READ INFORMATION FROM PAGELAYOUT TABLE
pageLayoutCursor = arcpy.SearchCursor(pageLayoutTable.dataSource, "[MapNumber] = '" + myMapNumber + "'")
pageLayoutRow = pageLayoutCursor.next()
if pageLayoutRow == None:               #INSERT A NEW ROW
    pageInsertCursor = arcpy.InsertCursor(pageLayoutTable.dataSource)
    pageInsertRow = pageInsertCursor.newRow()

    pageInsertRow.MapNumber = myMapNumber
    pageInsertRow.DataFrameMinX = dfMinX
    pageInsertRow.DataFrameMinY = dfMinY
    pageInsertRow.DataFrameMaxX = dfMaxX
    pageInsertRow.DataFrameMaxY = dfMaxY
##        pageInsertRow.MapPositionX = MapPositionX
##        pageInsertRow.MapPositionY = MapPositionY
    pageInsertRow.MapAngle = mapAngle
    pageInsertRow.TitleX = titleX
    pageInsertRow.TitleY = titleY

##        pageInsertRow.DisClaimerX = DisClaimerX
##        pageInsertRow.DisClaimerY = DisClaimerY
##        pageInsertRow.CancelNumX = CancelNumX
##        pageInsertRow.CancelNumY = CancelNumY
##        pageInsertRow.DateX = DateX
##        pageInsertRow.DateY = DateY
##        pageInsertRow.URCornerNumX = URCornerNumX
##        pageInsertRow.URCornerNumY = URCornerNumY
##        pageInsertRow.LRCornerNumX = LRCornerNumX
##        pageInsertRow.LRCornerNumY = LRCornerNumY
    pageInsertRow.ScaleBarX = scaleBarX
    pageInsertRow.ScaleBarY = scaleBarY
    pageInsertRow.NorthX = northX
    pageInsertRow.NorthY = northY

    pageInsertCursor.insertRow(pageInsertRow)       
    
else:                                   #UPDATE EXISTING ROW
    pageUpdateCursor = arcpy.UpdateCursor(pageLayoutTable.dataSource, "[MapNumber] = '" + myMapNumber + "'")
    pageUpdateRow = pageUpdateCursor.next()
    while pageUpdateRow:
        pageUpdateRow.DataFrameMinX = dfMinX
        pageUpdateRow.DataFrameMinY = dfMinY
        pageUpdateRow.DataFrameMaxX = dfMaxX
        pageUpdateRow.DataFrameMaxY = dfMaxY
##        pageUpdateRow.MapPositionX = MapPositionX
##        pageUpdateRow.MapPositionY = MapPositionY
        pageUpdateRow.MapAngle = mapAngle
        pageUpdateRow.TitleX = titleX
        pageUpdateRow.TitleY = titleY
##        pageUpdateRow.DisClaimerX = DisClaimerX
##        pageUpdateRow.DisClaimerY = DisClaimerY
##        pageUpdateRow.CancelNumX = CancelNumX
##        pageUpdateRow.CancelNumY = CancelNumY
##        pageUpdateRow.DateX = DateX
##        pageUpdateRow.DateY = DateY
##        pageUpdateRow.URCornerNumX = URCornerNumX
##        pageUpdateRow.URCornerNumY = URCornerNumY
##        pageUpdateRow.LRCornerNumX = LRCornerNumX
##        pageUpdateRow.LRCornerNumY = LRCornerNumY
        pageUpdateRow.ScaleBarX = scaleBarX
        pageUpdateRow.ScaleBarY = scaleBarY
        pageUpdateRow.NorthX = northX
        pageUpdateRow.NorthY = northY

        pageUpdateCursor.updateRow(pageUpdateRow)
        pageUpdateRow = pageUpdateCursor.next()      
   
    pageLayoutRow = pageLayoutCursor.next()
