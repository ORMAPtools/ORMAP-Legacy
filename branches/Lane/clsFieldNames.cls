VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFieldNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    Copyright (C) 2006  opet developers opet-developers@lists.sourceforge.net
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details located in AppSpecs.bas file.
'
'    You should have received a copy of the GNU General Public License along
'    with this program; if not, write to the Free Software Foundation, Inc.,
'    51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'
' Keyword expansion for source code control
' Tag for this file : $Name$
' SCC Revision number: $Revision$
' Date of last change: $Date$
'
' File name:            clsFieldNames
'
' Initial Author:       Type your name here
'
' Date Created:
'
' Description: MODULE FOR STANDARDIZING REFERENCES TO TABLES, FEATURE CLASSES, AND FIELD NAMES
'
'
'
' Entry points:
'       List the public variables and their purposes.
'       List the properties and routines that the module exposes to the rest of the program.
'
' Dependencies:
'       How does this file depend or relate to other files?
'
' Issues:
'       What are unsolved bugs, bottlenecks,
'       possible future enhancements, and
'       descriptions of other issues.
'
' Method:
'       ANY NEW ADDITIONS WILL REQUIRE ENTRIES IN THE INI FILE AND THREE LOCATIONS WITHIN THIS MODULE.
'
'       Describe any complex details that make sense on the file level.  This includes explanations
'       of complex algorithms, how different routines within the module interact, and a description
'       of a data structure used in the module.
'
' Updates:
'++  JWM 10/11/2006 added this header file

Option Explicit
'******************************
' Global/Public Definitions
'------------------------------
' Public API Declarations
'------------------------------

'------------------------------
' Public Enums and Constants
'------------------------------

'------------------------------
' Public variables
'------------------------------

'------------------------------
' Public Types
'------------------------------

'------------------------------
' Public loop variables
'------------------------------

'******************************
' Private Definitions
'------------------------------
' Private API declarations
'------------------------------

'------------------------------
' Private Variables
'------------------------------
'MAPINDEX FIELDS
Private m_sMICountyFN As String           'County Name
Private m_sMIMapNumberFN As String
Private m_sMIORMAPMapNumberFN As String
Private m_sMIMapScaleFN As String
Private m_sMIMapSuffTypeFN As String
Private m_sMIMapSuffNumFN As String
Private m_sMIPageFN As String
Private m_sMIReliabFN As String
'TAXCODE FIELDS

'TAXCODE LINES Fields

'TAXLOT FIELDS
Private m_sTLCountyFN As String
Private m_sTLTownFN As String
Private m_sTLTownPartFN As String
Private m_sTLTownDirFN As String
Private m_sTLRangeFN As String
Private m_sTLRangePartFN As String
Private m_sTLRangeDirFN As String
Private m_sTLSectNumberFN As String
Private m_sTLQtrFN As String
Private m_sTLQtrQtrFN As String
Private m_sTLMapNumberFN As String
Private m_sTLOrmapMapNumberFN As String
Private m_sTLOrmapTaxlotFN As String
Private m_sTLTaxlotFN As String
Private m_sTLMapTaxlotFN As String
'++ START JWM 11/03/2006
Private m_sTLMapTaxlotFormatString As String
'++ END JWM 11/03/2006
Private m_sTLSpecInterestFN As String
Private m_sTLSufNumFN As String
Private m_sTLSufTypeFN As String
Private m_sTLMapAcresFN As String
Private m_sTLAnomalyFN As String

'TAXLOT LINES FIELDS
Private m_sTLLinesLineTypeFN As String
'CARTOGRAPHIC LINES FIELDS
Private m_sCLinesLineTypeFN As String


'FIELDS PRESENT IN ALL TABLES
Private m_sAutoDateFN As String
Private m_sAutoMethodFN As String
Private m_sAutoWhoFN As String
'TABLE NAMES
Private m_sFCMapIndex As String
Private m_sFCTaxCode As String
Private m_sFCTaxlot As String
Private m_sFCTaxlotLines As String
Private m_sFCPlats As String
Private m_sFCCartoLines As String
Private m_sFCReferenceLines As String
Private m_sFCTaxCodeAnno As String
Private m_sFCLotsAnno As String
Private m_sFCTLAcrAnno As String
Private m_sFCTLNumAnno As String
Private m_sFCAnno10 As String
Private m_sFCAnno20 As String
Private m_sFCAnno30 As String
Private m_sFCAnno40 As String
Private m_sFCAnno50 As String
Private m_sFCAnno100 As String
Private m_sFCAnno200 As String
Private m_sFCAnno400 As String
Private m_sFCAnno800 As String
Private m_sFCAnno2000 As String
Private m_sTabCancelled As String
'DEFAULT VALUES
Private m_sDefCounty As String
Private m_sDefTownPart As String
Private m_sDefRangePart As String
Private m_sDefSuffType As String
Private m_sDefSuffNum As String
Private m_sDefTownDir As String
Private m_sDefRangeDir As String
Private m_sDefQtr As String
Private m_sDefQtrQtr As String
Private m_sDefAnomaly As String
'++ New coded added 10/21/05
'ANNO TEXT SIZES
Private m_dAnnoSizeTLAcr120 As Double
Private m_dAnnoSizeTLAcr240 As Double
Private m_dAnnoSizeTLAcr360 As Double
Private m_dAnnoSizeTLAcr480 As Double
Private m_dAnnoSizeTLAcr600 As Double
Private m_dAnnoSizeTLAcr1200 As Double
Private m_dAnnoSizeTLAcr2400 As Double
Private m_dAnnoSizeTLAcr4800 As Double
Private m_dAnnoSizeTLAcr9600 As Double
Private m_dAnnoSizeTLAcr24000 As Double
Private m_dAnnoSizeTLNum120 As Double
Private m_dAnnoSizeTLNum240 As Double
Private m_dAnnoSizeTLNum360 As Double
Private m_dAnnoSizeTLNum480 As Double
Private m_dAnnoSizeTLNum600 As Double
Private m_dAnnoSizeTLNum1200 As Double
Private m_dAnnoSizeTLNum2400 As Double
Private m_dAnnoSizeTLNum4800 As Double
Private m_dAnnoSizeTLNum9600 As Double
Private m_dAnnoSizeTLNum24000 As Double
'++ End new code
'------------------------------
'Private Constants and Enums
'------------------------------
Private Const INIFILE As String = "ormap.ini"
Private Const m_sSection As String = "GENERAL"
'------------------------------
' Private Types
'------------------------------

'------------------------------
' Private loop variables
'------------------------------

Private Sub Class_Initialize()

  If Not TestIniFile(app.Path & "\" & INIFILE) Then
    MsgBox "Unable to open Inifile at " & app.Path & "\" & INIFILE, vbExclamation
    Exit Sub
  End If
  
  'Read ini file and set references to field names
  
  'MAPINDEX FIELDS
  m_sMICountyFN = ReadIniFile(m_sSection, "MI_COUNTY")
  m_sMIMapNumberFN = ReadIniFile(m_sSection, "MI_MAPNUMBER")
  m_sMIORMAPMapNumberFN = ReadIniFile(m_sSection, "MI_ORMAP_MAPNUMBER")
  m_sMIMapScaleFN = ReadIniFile(m_sSection, "MI_MAPSCALE")
  m_sMIMapSuffTypeFN = ReadIniFile(m_sSection, "MI_MAPSUFFIXTYPE")
  m_sMIMapSuffNumFN = ReadIniFile(m_sSection, "MI_MAPSUFFIXNUM")
  m_sMIPageFN = ReadIniFile(m_sSection, "MI_PAGE")
  m_sMIReliabFN = ReadIniFile(m_sSection, "MI_RELIABILITYCODE")

  'TAXCODE FIELDS

 'TAXCODE LINES Fields
  
  'TAXLOT FIELDS
  m_sTLCountyFN = ReadIniFile(m_sSection, "TL_COUNTY")
  m_sTLTownFN = ReadIniFile(m_sSection, "TL_TOWN")
  m_sTLTownPartFN = ReadIniFile(m_sSection, "TL_TOWNPART")
  m_sTLTownDirFN = ReadIniFile(m_sSection, "TL_TOWNDIR")
  m_sTLRangeFN = ReadIniFile(m_sSection, "TL_RANGE")
  m_sTLRangePartFN = ReadIniFile(m_sSection, "TL_RANGEPART")
  m_sTLRangeDirFN = ReadIniFile(m_sSection, "TL_RANGEDIR")
  m_sTLSectNumberFN = ReadIniFile(m_sSection, "TL_SECTNUMBER")
  m_sTLQtrFN = ReadIniFile(m_sSection, "TL_QTR")
  m_sTLQtrQtrFN = ReadIniFile(m_sSection, "TL_QTRQTR")
  m_sTLMapNumberFN = ReadIniFile(m_sSection, "TL_MAPNUMBER")
  m_sTLOrmapMapNumberFN = ReadIniFile(m_sSection, "TL_ORMAP_MAPNUMBER")
  m_sTLOrmapTaxlotFN = ReadIniFile(m_sSection, "TL_ORMAP_TAXLOT")
'++ START JWM 11/03/2006 Need the format string for MAPTAXLOT value
  m_sTLMapTaxlotFormatString = ReadIniFile(m_sSection, "TL_MAP_TAXLOT_MASK")
'++ END JWM 11/03/2006
  m_sTLTaxlotFN = ReadIniFile(m_sSection, "TL_TAXLOT")
  m_sTLMapTaxlotFN = ReadIniFile(m_sSection, "TL_MAP_TAXLOT")
  m_sTLSpecInterestFN = ReadIniFile(m_sSection, "TL_SPEC_INTEREST")
  m_sTLSufNumFN = ReadIniFile(m_sSection, "TL_MAP_SUF_NUM")
  m_sTLSufTypeFN = ReadIniFile(m_sSection, "TL_MAP_SUF_TYPE")
  m_sTLMapAcresFN = ReadIniFile(m_sSection, "TL_MAPACRES")
  m_sTLAnomalyFN = ReadIniFile(m_sSection, "TL_ANOMALY")


  'TAXLOT LINES FIELDS
  m_sTLLinesLineTypeFN = ReadIniFile(m_sSection, "TLLINES_LINETYPE")
  
  'CARTOGRAPHIC LINES FIELDS
  m_sCLinesLineTypeFN = ReadIniFile(m_sSection, "CARTOLINES_LINETYPE")


  'TAXLOT POINTS FIELDS
  
  'FIELDS PRESENT IN ALL TABLES
  m_sAutoDateFN = ReadIniFile(m_sSection, "AUTO_DATE")
  m_sAutoMethodFN = ReadIniFile(m_sSection, "AUTO_METHOD")
  m_sAutoWhoFN = ReadIniFile(m_sSection, "AUTO_WHO")
  'TABLE NAMES
  m_sFCMapIndex = ReadIniFile(m_sSection, "FC_MAPINDEX")
  m_sFCTaxCode = ReadIniFile(m_sSection, "FC_TAXCODE")
  m_sFCTaxlot = ReadIniFile(m_sSection, "FC_TAXLOT")
  m_sFCTaxlotLines = ReadIniFile(m_sSection, "FC_TAXLOTLINES")
  m_sFCPlats = ReadIniFile(m_sSection, "FC_PLATS")
  m_sFCCartoLines = ReadIniFile(m_sSection, "FC_CARTOLINES")
  m_sFCReferenceLines = ReadIniFile(m_sSection, "FC_REFERENCELINES")
  m_sFCTaxCodeAnno = ReadIniFile(m_sSection, "FC_TAXCODEANNO")
  m_sFCLotsAnno = ReadIniFile(m_sSection, "FC_LOTSANNO")
  m_sFCTLAcrAnno = ReadIniFile(m_sSection, "FC_TAXLOTACREAGEANNO")
  m_sFCTLNumAnno = ReadIniFile(m_sSection, "FC_TAXLOTNUMBERANNO")
  m_sFCAnno10 = ReadIniFile(m_sSection, "FC_ANNO0010")
  m_sFCAnno20 = ReadIniFile(m_sSection, "FC_ANNO0020")
  m_sFCAnno30 = ReadIniFile(m_sSection, "FC_ANNO0030")
  m_sFCAnno40 = ReadIniFile(m_sSection, "FC_ANNO0040")
  m_sFCAnno50 = ReadIniFile(m_sSection, "FC_ANNO0050")
  m_sFCAnno100 = ReadIniFile(m_sSection, "FC_ANNO0100")
  m_sFCAnno200 = ReadIniFile(m_sSection, "FC_ANNO0200")
  m_sFCAnno400 = ReadIniFile(m_sSection, "FC_ANNO0400")
  m_sFCAnno800 = ReadIniFile(m_sSection, "FC_ANNO0800")
  m_sFCAnno2000 = ReadIniFile(m_sSection, "FC_ANNO2000")
  m_sTabCancelled = ReadIniFile(m_sSection, "TAB_Cancelled")
  
  'DEFAULT VALUES
  m_sDefCounty = ReadIniFile(m_sSection, "DEFAULT_COUNTY")
  m_sDefTownPart = ReadIniFile(m_sSection, "DEFAULT_TOWNPART")
  m_sDefRangePart = ReadIniFile(m_sSection, "DEFAULT_RANGEPART")
  m_sDefSuffType = ReadIniFile(m_sSection, "DEFAULT_SUFFTYPE")
  m_sDefSuffNum = ReadIniFile(m_sSection, "DEFAULT_SUFFNUM")
  m_sDefTownDir = ReadIniFile(m_sSection, "DEFAULT_TOWNDIR")
  m_sDefRangeDir = ReadIniFile(m_sSection, "DEFAULT_RANGEDIR")
  m_sDefQtr = ReadIniFile(m_sSection, "DEFAULT_QTR")
  m_sDefQtrQtr = ReadIniFile(m_sSection, "DEFAULT_QTRQTR")
  m_sDefAnomaly = ReadIniFile(m_sSection, "DEFAULT_ANOMALY")

  '++ New coded added 10/2/05
  'ANNO TEXT SIZES
  m_dAnnoSizeTLAcr120 = ReadIniFile(m_sSection, "ANNOSIZETLACR120")
  m_dAnnoSizeTLAcr240 = ReadIniFile(m_sSection, "ANNOSIZETLACR240")
  m_dAnnoSizeTLAcr360 = ReadIniFile(m_sSection, "ANNOSIZETLACR360")
  m_dAnnoSizeTLAcr480 = ReadIniFile(m_sSection, "ANNOSIZETLACR480")
  m_dAnnoSizeTLAcr600 = ReadIniFile(m_sSection, "ANNOSIZETLACR600")
  m_dAnnoSizeTLAcr1200 = ReadIniFile(m_sSection, "ANNOSIZETLACR1200")
  m_dAnnoSizeTLAcr2400 = ReadIniFile(m_sSection, "ANNOSIZETLACR2400")
  m_dAnnoSizeTLAcr4800 = ReadIniFile(m_sSection, "ANNOSIZETLACR4800")
  m_dAnnoSizeTLAcr9600 = ReadIniFile(m_sSection, "ANNOSIZETLACR9600")
  m_dAnnoSizeTLAcr24000 = ReadIniFile(m_sSection, "ANNOSIZETLACR24000")
  m_dAnnoSizeTLNum120 = ReadIniFile(m_sSection, "ANNOSIZETLNUM120")
  m_dAnnoSizeTLNum240 = ReadIniFile(m_sSection, "ANNOSIZETLNUM240")
  m_dAnnoSizeTLNum360 = ReadIniFile(m_sSection, "ANNOSIZETLNUM360")
  m_dAnnoSizeTLNum480 = ReadIniFile(m_sSection, "ANNOSIZETLNUM480")
  m_dAnnoSizeTLNum600 = ReadIniFile(m_sSection, "ANNOSIZETLNUM600")
  m_dAnnoSizeTLNum1200 = ReadIniFile(m_sSection, "ANNOSIZETLNUM1200")
  m_dAnnoSizeTLNum2400 = ReadIniFile(m_sSection, "ANNOSIZETLNUM2400")
  m_dAnnoSizeTLNum4800 = ReadIniFile(m_sSection, "ANNOSIZETLNUM4800")
  m_dAnnoSizeTLNum9600 = ReadIniFile(m_sSection, "ANNOSIZETLNUM9600")
  m_dAnnoSizeTLNum24000 = ReadIniFile(m_sSection, "ANNOSIZETLNUM24000")
  '++ End new code
End Sub

'MAPINDEX FIELDS
Public Property Get MICountyFN() As String

    MICountyFN = m_sMICountyFN

End Property
Public Property Get MIMapNumberFN() As String

    MIMapNumberFN = m_sMIMapNumberFN

End Property
Public Property Get MIORMAPMapNumberFN() As String

    MIORMAPMapNumberFN = m_sMIORMAPMapNumberFN

End Property
Public Property Get MIMapScaleFN() As String

    MIMapScaleFN = m_sMIMapScaleFN

End Property
Public Property Get MIMapSuffTypeFN() As String

    MIMapSuffTypeFN = m_sMIMapSuffTypeFN

End Property
Public Property Get MIMapSuffNumFN() As String

    MIMapSuffNumFN = m_sMIMapSuffNumFN

End Property
Public Property Get MIPageFN() As String

    MIPageFN = m_sMIPageFN

End Property
Public Property Get MIReliabFN() As String

    MIReliabFN = m_sMIReliabFN

End Property

'TAXCODE FIELDS
'TAXCODE LINES Fields
'TAXLOT FIELDS
Public Property Get TLCountyFN() As String

    TLCountyFN = m_sTLCountyFN

End Property
Public Property Get TLTownFN() As String

    TLTownFN = m_sTLTownFN

End Property
Public Property Get TLTownPartFN() As String

    TLTownPartFN = m_sTLTownPartFN

End Property
Public Property Get TLTownDirFN() As String

    TLTownDirFN = m_sTLTownDirFN

End Property
Public Property Get TLRangeFN() As String

    TLRangeFN = m_sTLRangeFN

End Property
Public Property Get TLRangePartFN() As String

    TLRangePartFN = m_sTLRangePartFN

End Property
Public Property Get TLRangeDirFN() As String

    TLRangeDirFN = m_sTLRangeDirFN

End Property
Public Property Get TLSectNumberFN() As String

    TLSectNumberFN = m_sTLSectNumberFN

End Property
Public Property Get TLQtrFN() As String

    TLQtrFN = m_sTLQtrFN

End Property
Public Property Get TLQtrQtrFN() As String

    TLQtrQtrFN = m_sTLQtrQtrFN

End Property
Public Property Get TLMapNumberFN() As String

    TLMapNumberFN = m_sTLMapNumberFN

End Property
Public Property Get TLOrmapMapNumberFN() As String

    TLOrmapMapNumberFN = m_sTLOrmapMapNumberFN

End Property
Public Property Get TLOrmapTaxlotFN() As String
    TLOrmapTaxlotFN = m_sTLOrmapTaxlotFN
End Property
Public Property Get TLTaxlotFN() As String
    TLTaxlotFN = m_sTLTaxlotFN
End Property

Public Property Get TLMapTaxlotFN() As String
    TLMapTaxlotFN = m_sTLMapTaxlotFN
End Property
Public Property Get TLSpecInterestFN() As String
    TLSpecInterestFN = m_sTLSpecInterestFN
End Property
Public Property Get TLSufNumFN() As String
    TLSufNumFN = m_sTLSufNumFN
End Property
Public Property Get TLSufTypeFN() As String
    TLSufTypeFN = m_sTLSufTypeFN
End Property
Public Property Get TLMapAcresFN() As String
    TLMapAcresFN = m_sTLMapAcresFN
End Property
Public Property Get TLAnomalyFN() As String
    TLAnomalyFN = m_sTLAnomalyFN
End Property
'TAXLOT LINES FIELDS
Public Property Get TLLinesLineTypeFN() As String
    TLLinesLineTypeFN = m_sTLLinesLineTypeFN
End Property
'CARTOGRAPHIC LINES FIELDS
Public Property Get CLinesLineTypeFN() As String
    CLinesLineTypeFN = m_sCLinesLineTypeFN
End Property
'TAXLOT POINTS FIELDS

'FIELDS PRESENT IN ALL TABLES
Public Property Get AutoDateFN() As String

    AutoDateFN = m_sAutoDateFN

End Property
Public Property Get AutoMethodFN() As String

    AutoMethodFN = m_sAutoMethodFN

End Property
Public Property Get AutoWhoFN() As String

    AutoWhoFN = m_sAutoWhoFN

End Property
'TABLE NAMES
Public Property Get FCMapIndex() As String

    FCMapIndex = m_sFCMapIndex

End Property
Public Property Get FCTaxCode() As String

    FCTaxCode = m_sFCTaxCode

End Property
Public Property Get FCTaxlot() As String

    FCTaxlot = m_sFCTaxlot

End Property
Public Property Get FCTaxlotLines() As String

    FCTaxlotLines = m_sFCTaxlotLines

End Property
Public Property Get FCPlats() As String

    FCPlats = m_sFCPlats

End Property
Public Property Get FCCartoLines() As String

    FCCartoLines = m_sFCCartoLines

End Property
Public Property Get FCReferenceLines() As String
    FCReferenceLines = m_sFCReferenceLines
End Property
Public Property Get FCTaxCodeAnno() As String

    FCTaxCodeAnno = m_sFCTaxCodeAnno

End Property
Public Property Get FCLotsAnno() As String

    FCLotsAnno = m_sFCLotsAnno

End Property
Public Property Get FCTLAcrAnno()
    FCTLAcrAnno = m_sFCTLAcrAnno
End Property
Public Property Get FCTLNumAnno()
    FCTLNumAnno = m_sFCTLNumAnno
End Property
Public Property Get FCAnno10() As String

    FCAnno10 = m_sFCAnno10

End Property
Public Property Get FCAnno20() As String

    FCAnno20 = m_sFCAnno20

End Property
Public Property Get FCAnno30() As String

    FCAnno30 = m_sFCAnno30

End Property
Public Property Get FCAnno40() As String

    FCAnno40 = m_sFCAnno40

End Property
Public Property Get FCAnno50() As String

    FCAnno50 = m_sFCAnno50

End Property
Public Property Get FCAnno100() As String

    FCAnno100 = m_sFCAnno100

End Property
Public Property Get FCAnno200() As String

    FCAnno200 = m_sFCAnno200

End Property
Public Property Get FCAnno400() As String

    FCAnno400 = m_sFCAnno400

End Property
Public Property Get FCAnno800() As String

    FCAnno800 = m_sFCAnno800

End Property
Public Property Get FCAnno2000() As String

    FCAnno2000 = m_sFCAnno2000
End Property

Public Property Get TabCancelled() As String
    TabCancelled = m_sTabCancelled
End Property
'DEFAULT VALUES
Public Property Get DefCounty() As String
    DefCounty = m_sDefCounty
End Property
Public Property Get DefTownPart() As String
    DefTownPart = m_sDefTownPart
End Property
Public Property Get DefRangePart() As String
    DefRangePart = m_sDefRangePart
End Property
Public Property Get DefSuffType() As String
    DefSuffType = m_sDefSuffType
End Property
Public Property Get DefSuffNum() As String
    DefSuffNum = m_sDefSuffNum
End Property
Public Property Get DefTownDir() As String
    DefTownDir = m_sDefTownDir
End Property
Public Property Get DefRangeDir() As String
    DefRangeDir = m_sDefRangeDir
End Property
Public Property Get DefQtr() As String
    DefQtr = m_sDefQtr
End Property
Public Property Get DefQtrQtr() As String
    DefQtrQtr = m_sDefQtrQtr
End Property
Public Property Get DefAnomaly() As String
    DefAnomaly = m_sDefAnomaly
End Property
'++ New coded added 10/2/05
'ANNO TEXT SIZES
Public Property Get AnnoSizeTLAcr120() As Double
    AnnoSizeTLAcr120 = m_dAnnoSizeTLAcr120
End Property

Public Property Get AnnoSizeTLAcr240() As Double
    AnnoSizeTLAcr240 = m_dAnnoSizeTLAcr240
End Property

Public Property Get AnnoSizeTLAcr360() As Double
    AnnoSizeTLAcr360 = m_dAnnoSizeTLAcr360
End Property

Public Property Get AnnoSizeTLAcr480() As Double
    AnnoSizeTLAcr480 = m_dAnnoSizeTLAcr480
End Property

Public Property Get AnnoSizeTLAcr600() As Double
    AnnoSizeTLAcr600 = m_dAnnoSizeTLAcr600
End Property

Public Property Get AnnoSizeTLAcr1200() As Double
    AnnoSizeTLAcr1200 = m_dAnnoSizeTLAcr1200
End Property

Public Property Get AnnoSizeTLAcr2400() As Double
    AnnoSizeTLAcr2400 = m_dAnnoSizeTLAcr2400
End Property

Public Property Get AnnoSizeTLAcr4800() As Double
    AnnoSizeTLAcr4800 = m_dAnnoSizeTLAcr4800
End Property

Public Property Get AnnoSizeTLAcr9600() As Double
    AnnoSizeTLAcr9600 = m_dAnnoSizeTLAcr9600
End Property

Public Property Get AnnoSizeTLAcr24000() As Double
    AnnoSizeTLAcr24000 = m_dAnnoSizeTLAcr24000
End Property

Public Property Get AnnoSizeTLNum120() As Double
    AnnoSizeTLNum120 = m_dAnnoSizeTLNum120
End Property

Public Property Get AnnoSizeTLNum240() As Double
    AnnoSizeTLNum240 = m_dAnnoSizeTLNum240
End Property

Public Property Get AnnoSizeTLNum360() As Double
    AnnoSizeTLNum360 = m_dAnnoSizeTLNum360
End Property

Public Property Get AnnoSizeTLNum480() As Double
    AnnoSizeTLNum480 = m_dAnnoSizeTLNum480
End Property

Public Property Get AnnoSizeTLNum600() As Double
    AnnoSizeTLNum600 = m_dAnnoSizeTLNum600
End Property

Public Property Get AnnoSizeTLNum1200() As Double
    AnnoSizeTLNum1200 = m_dAnnoSizeTLNum1200
End Property

Public Property Get AnnoSizeTLNum2400() As Double
    AnnoSizeTLNum2400 = m_dAnnoSizeTLNum2400
End Property

Public Property Get AnnoSizeTLNum4800() As Double
    AnnoSizeTLNum4800 = m_dAnnoSizeTLNum4800
End Property

Public Property Get AnnoSizeTLNum9600() As Double
    AnnoSizeTLNum9600 = m_dAnnoSizeTLNum9600
End Property

Public Property Get AnnoSizeTLNum24000() As Double
    AnnoSizeTLNum24000 = m_dAnnoSizeTLNum24000
End Property
'++ End new code

Public Property Get MapTaxlotFormatString() As String
'++  JWM 11/03/2006 Added this assessor function for the format string
MapTaxlotFormatString = m_sTLMapTaxlotFormatString
End Property


