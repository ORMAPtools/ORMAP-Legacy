# ---------------------------------------------------------------------------
# Ormap_Print.py
# Created by: Shad Campbell
# Date: 3/11/2011
# Updated by:
# Description: This script prints or exports a map or collection of maps.  It
# utilizes the ORMAP_ZoomToMap.py script for zooming/configuring maps in a collection.
# ---------------------------------------------------------------------------

import arcpy, tempfile, sys, arcpy.mapping as MAP
arcpy.ImportToolbox(sys.path[0] + "\ORMAP Map Production.tbx")

printCurrent = arcpy.GetParameterAsText(0)
printMultiple = arcpy.GetParameterAsText(1)
printFromFile = arcpy.GetParameterAsText(2)
printType = arcpy.GetParameterAsText(3)
res = arcpy.GetParameterAsText(4)
tifCompression = arcpy.GetParameterAsText(5)
printOutputFolder = arcpy.GetParameterAsText(6)
printer = arcpy.GetParameterAsText(7)

##import pythonaddins
##pythonaddins.MessageBox(arcpy.GetParameterAsText(5), '', 0)

MXD = MAP.MapDocument("CURRENT")

##mapNumberList = printMultiple.split(";")
##for thisMap in mapNumberList:
##    arcpy.AddMessage(thisMap)
##    arcpy.ZoomToMapNumber_ORMAPMapProduction(str(thisMap), pageSize)
##    arcpy.AddMessage("DONE")


def printMap(mapNumber):

    arcpy.AddMessage("Printing " + str(mapNumber))

    if printType=="Printer":
        #-- Send to printer
        MAP.PrintMap(MXD, printer)
    else:
        #-- If the output folder is not specified use the temp folder.
        if printOutputFolder!="":
            outputPath = printOutputFolder
        else:
            outputPath = tempfile.gettempdir()

        #-- Export
        if printType=="PDF":
            arcpy.AddMessage("Exporting to " + outputPath + "\\" + mapNumber + ".pdf")
            arcpy.mapping.ExportToPDF(MXD, outputPath + "\\" + mapNumber + ".pdf", "PAGE_LAYOUT", resolution=res)
        elif  printType=="TIFF":
            arcpy.AddMessage("Exporting to " + outputPath + "\\" + mapNumber + ".tif")
            arcpy.mapping.ExportToTIFF(MXD, outputPath + "\\" + mapNumber + ".tif", "PAGE_LAYOUT", resolution=res, tiff_compression=tifCompression)
        else:
            arcpy.AddMessage("Unrecognized Print format.  Unable to print map.")



if printCurrent=="true":

    #-- Retrieve the mapnumber by looking at the mapnumber text in the upper right of layout
    if len(MAP.ListLayoutElements(MXD, "TEXT_ELEMENT", "MapNumber"))>0:
        theMapNumberElm = MAP.ListLayoutElements(MXD, "TEXT_ELEMENT", "MapNumber")[0]
        theMapNumber = theMapNumberElm.text
    else:
        theMapNumber = "ORMAP_Map"

    printMap(theMapNumber)

else:

    if printFromFile!="":

        text_file = open(printFromFile, "r")
        mapNumberList = text_file.readlines()
        for thisMap in mapNumberList:
            if thisMap.replace("\n", "")!="":
                arcpy.ZoomToMapNumber_ORMAPMapProduction(thisMap.replace("\n", ""))
                printMap(thisMap.replace("\n", ""))
        text_file.close()

    else:

        mapNumberList = printMultiple.split(";")
        for thisMap in mapNumberList:
            arcpy.AddMessage(thisMap)
            arcpy.ZoomToMapNumber_ORMAPMapProduction(str(thisMap))
            arcpy.AddMessage("DONE")



