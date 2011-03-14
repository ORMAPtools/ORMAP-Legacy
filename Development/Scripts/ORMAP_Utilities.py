# ---------------------------------------------------------------------------
# OrmapUtilities.py
# Created by: Shad Campbell
# Date: 3/11/2011
# Updated by: 
# Description: Utility class used in Tool Validation scripts. 
# ---------------------------------------------------------------------------

import arcpy, arcpy.mapping as MAP
import ORMAP_LayersConfig as OrmapLayers


def getMapNumbers():

    mapNumbers = []
    MXD = MAP.MapDocument("CURRENT")
    
    if len(MAP.ListDataFrames(MXD, "MainDF"))>0:
        mainDF = MAP.ListDataFrames(MXD, "MainDF")[0]
        if len(MAP.ListLayers(MXD, OrmapLayers.MAPINDEX_LAYER, mainDF))>0:
            MapIndex = MAP.ListLayers(MXD, OrmapLayers.MAPINDEX_LAYER, mainDF)[0]
            orgDefQuery = MapIndex.definitionQuery
            MapIndex.definitionQuery = ""
            mapIndexCursor = arcpy.SearchCursor(MapIndex.name, "", "", "MAPNUMBER")

            row = mapIndexCursor.next()

            # Create an empty list
            while row:
                if row.MAPNUMBER not in mapNumbers:
                    mapNumbers.append(row.MAPNUMBER)
                row = mapIndexCursor.next()

            mapNumbers.sort()
            MapIndex.definitionQuery = orgDefQuery

    return mapNumbers


    
    





