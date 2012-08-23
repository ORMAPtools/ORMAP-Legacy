import arcpy
import pythonaddins
import os

tbxfile = os.path.join(os.path.dirname(__file__), r'Scripts\ORMAP Map Production.tbx')
layerConfigFile = os.path.join(os.path.dirname(__file__), r'Config\ORMAP_LayersConfig.py')
mapConfigFile = os.path.join(os.path.dirname(__file__), r'Config\ORMAP_MapConfig.py')

def configFilesExist():
    if os.path.exists(layerConfigFile) and os.path.exists(mapConfigFile):
        isEnabled = True
    else:
        isEnabled = False
        pythonaddins.MessageBox('Please Import your Configuration File before using this tool', 'Missing Configuration Files', 0)
    return isEnabled

class PageLayoutElementReport(object):
    """Implementation for TheAddIn_addin.PageLayoutElementReport (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        pythonaddins.GPToolDialog(tbxfile, "PagelayoutElementReport")

class PrintMaps(object):
    """Implementation for TheAddIn_addin.PrintMaps (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        if configFilesExist():
            pythonaddins.GPToolDialog(tbxfile, "PrintMap")

class SaveDefinitionQueries(object):
    """Implementation for TheAddIn_addin.SaveDefinitionQueries (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        if configFilesExist():
            pythonaddins.GPToolDialog(tbxfile, "SaveDefinitionQueries")

class SavePageLayoutElements(object):
    """Implementation for TheAddIn_addin.SavePageLayoutElements (Button)"""
    def __init__(self):
        self.enabled = True

        self.checked = False
    def onClick(self):
        if configFilesExist():
            pythonaddins.GPToolDialog(tbxfile, "SavePageLayoutElements")

class ZoomToMapNumber(object):
    """Implementation for TheAddIn_addin.ZoomToMapNumber (Button)"""
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        if configFilesExist():
            pythonaddins.GPToolDialog(tbxfile, "ZoomToMapNumber")

class ImportConfigFiles(object):
    def __init__(self):
        self.enabled = True
        self.checked = False
    def onClick(self):
        pythonaddins.GPToolDialog(tbxfile, "ImportConfigFiles")

