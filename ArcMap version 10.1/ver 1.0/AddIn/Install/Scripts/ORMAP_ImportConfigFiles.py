# ---------------------------------------------------------------------------
# ORMAP_ImportConfigFiles.py
# Created by: Shad Campbell
# Updated by: 
# Description: Imports the config files 
# ---------------------------------------------------------------------------

import arcpy
import os, shutil

layerConfigFile = arcpy.GetParameterAsText(0)
mapConfigFile = arcpy.GetParameterAsText(1)

configFolder = os.path.dirname(__file__).replace('Scripts', 'Config')

if layerConfigFile!=None:
    shutil.copyfile(layerConfigFile, configFolder + r'\\ORMAP_LayersConfig.py')

if mapConfigFile!=None:
    shutil.copyfile(mapConfigFile, configFolder + r'\\ORMAP_MapConfig.py')


    




