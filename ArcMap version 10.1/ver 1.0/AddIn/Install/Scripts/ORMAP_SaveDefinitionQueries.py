# ---------------------------------------------------------------------------
# Ormap_SaveDefinitionQueries.py
# Created by: Shad Campbell
# Date: 3/11/2011
# Updated by: 
# Description: This script captures the current definition queries for each layer
# in the TOC and stores them in the CustomDefinitionQueries table.  Once stored a
# map uses these values rather than the ones stored in the configuration files.
# ---------------------------------------------------------------------------

import sys, os, arcpy, arcpy.mapping as MAP

sys.path.append(os.path.dirname(__file__).replace('Scripts', 'Config')) #path to config files
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

#REFERENCE CUSTOM DEFINITION QUERIES TABLE
defQueryRow = None
if len(MAP.ListTableViews(MXD, OrmapLayers.CUSTOMDEFINITIONQUERIES_TABLE, mainDF))>0:
    defQueryTable = MAP.ListTableViews(MXD, OrmapLayers.CUSTOMDEFINITIONQUERIES_TABLE, mainDF)[0]
    defQueryCursor = arcpy.SearchCursor(defQueryTable.name, "MapNumber = '" + MapNumber + "'") #--CUSTOM
    defQueryRow = defQueryCursor.next()    
else:
    sys.exit("Unable to find the Custom Definition Queries table.  Please check your TOC and config file.")  


#GET QUERY DEFINITIONS FOR EACH LAYER INTO A STRING
defQueryString = ""
for lyr in MAP.ListLayers(MXD, "", mainDF):
    if lyr.supports("DATASETNAME") and lyr.definitionQuery!="":
        defQueryString += ";" if len(defQueryString) > 0 else ""
        defQueryString += lyr.name + ":" + lyr.definitionQuery

#DETERMINE WHETHER TO INSERT OR UPDATE PAGELAYOUT ROW.
if defQueryRow == None:
    theCursor = arcpy.InsertCursor(defQueryTable.name)
    theRow = theCursor.newRow()
else:
    theCursor = arcpy.UpdateCursor(defQueryTable.name, "MapNumber = '" + MapNumber + "'")
    theRow = theCursor.next()

theRow.MapNumber = MapNumber

if len(arcpy.ListFields(defQueryTable.name, "DefQueryString")) > 0:
    theRow.DefQueryString = defQueryString


#INSERT OR UPDATE ROW.
if defQueryRow == None:
    theCursor.insertRow(theRow)
else:
    theCursor.updateRow(theRow)

