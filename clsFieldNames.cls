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
'MODULE FOR STANDARDIZING REFERENCES TO TABLES, FEATURE CLASSES, AND FIELD NAMES
'ANY NEW ADDITIONS WILL REQUIRE ENTRIES IN THE INI FILE AND THREE LOCATIONS WITHIN
'THIS MODULE.

Option Explicit

Private Const INIFILE As String = "ormap.ini"
Private Const m_sSection As String = "GENERAL"
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




Private Sub Class_Initialize()

93:   If Not TestIniFile(app.Path & "\" & INIFILE) Then
94:     MsgBox "Unable to open Inifile at " & app.Path & "\" & INIFILE, vbExclamation
    Exit Sub
96:   End If
  
  'Read ini file and set references to field names
  
  'MAPINDEX FIELDS
101:   m_sMICountyFN = ReadIniFile(m_sSection, "MI_COUNTY")
102:   m_sMIMapNumberFN = ReadIniFile(m_sSection, "MI_MAPNUMBER")
103:   m_sMIORMAPMapNumberFN = ReadIniFile(m_sSection, "MI_ORMAP_MAPNUMBER")
104:   m_sMIMapScaleFN = ReadIniFile(m_sSection, "MI_MAPSCALE")
105:   m_sMIMapSuffTypeFN = ReadIniFile(m_sSection, "MI_MAPSUFFIXTYPE")
106:   m_sMIMapSuffNumFN = ReadIniFile(m_sSection, "MI_MAPSUFFIXNUM")
107:   m_sMIPageFN = ReadIniFile(m_sSection, "MI_PAGE")
108:   m_sMIReliabFN = ReadIniFile(m_sSection, "MI_RELIABILITYCODE")


  'TAXCODE FIELDS

 'TAXCODE LINES Fields
  
  'TAXLOT FIELDS
116:   m_sTLCountyFN = ReadIniFile(m_sSection, "TL_COUNTY")
117:   m_sTLTownFN = ReadIniFile(m_sSection, "TL_TOWN")
118:   m_sTLTownPartFN = ReadIniFile(m_sSection, "TL_TOWNPART")
119:   m_sTLTownDirFN = ReadIniFile(m_sSection, "TL_TOWNDIR")
120:   m_sTLRangeFN = ReadIniFile(m_sSection, "TL_RANGE")
121:   m_sTLRangePartFN = ReadIniFile(m_sSection, "TL_RANGEPART")
122:   m_sTLRangeDirFN = ReadIniFile(m_sSection, "TL_RANGEDIR")
123:   m_sTLSectNumberFN = ReadIniFile(m_sSection, "TL_SECTNUMBER")
124:   m_sTLQtrFN = ReadIniFile(m_sSection, "TL_QTR")
125:   m_sTLQtrQtrFN = ReadIniFile(m_sSection, "TL_QTRQTR")
126:   m_sTLMapNumberFN = ReadIniFile(m_sSection, "TL_MAPNUMBER")
127:   m_sTLOrmapMapNumberFN = ReadIniFile(m_sSection, "TL_ORMAP_MAPNUMBER")
128:   m_sTLOrmapTaxlotFN = ReadIniFile(m_sSection, "TL_ORMAP_TAXLOT")
129:   m_sTLTaxlotFN = ReadIniFile(m_sSection, "TL_TAXLOT")
130:   m_sTLMapTaxlotFN = ReadIniFile(m_sSection, "TL_MAP_TAXLOT")
131:   m_sTLSpecInterestFN = ReadIniFile(m_sSection, "TL_SPEC_INTEREST")
132:   m_sTLSufNumFN = ReadIniFile(m_sSection, "TL_MAP_SUF_NUM")
133:   m_sTLSufTypeFN = ReadIniFile(m_sSection, "TL_MAP_SUF_TYPE")
134:   m_sTLMapAcresFN = ReadIniFile(m_sSection, "TL_MAPACRES")
135:   m_sTLAnomalyFN = ReadIniFile(m_sSection, "TL_ANOMALY")


  'TAXLOT LINES FIELDS
139:   m_sTLLinesLineTypeFN = ReadIniFile(m_sSection, "TLLINES_LINETYPE")
  
  'CARTOGRAPHIC LINES FIELDS
142:   m_sCLinesLineTypeFN = ReadIniFile(m_sSection, "CARTOLINES_LINETYPE")


  'TAXLOT POINTS FIELDS
  
  'FIELDS PRESENT IN ALL TABLES
148:   m_sAutoDateFN = ReadIniFile(m_sSection, "AUTO_DATE")
149:   m_sAutoMethodFN = ReadIniFile(m_sSection, "AUTO_METHOD")
150:   m_sAutoWhoFN = ReadIniFile(m_sSection, "AUTO_WHO")
  'TABLE NAMES
152:   m_sFCMapIndex = ReadIniFile(m_sSection, "FC_MAPINDEX")
153:   m_sFCTaxCode = ReadIniFile(m_sSection, "FC_TAXCODE")
154:   m_sFCTaxlot = ReadIniFile(m_sSection, "FC_TAXLOT")
155:   m_sFCTaxlotLines = ReadIniFile(m_sSection, "FC_TAXLOTLINES")
156:   m_sFCPlats = ReadIniFile(m_sSection, "FC_PLATS")
157:   m_sFCCartoLines = ReadIniFile(m_sSection, "FC_CARTOLINES")
158:   m_sFCReferenceLines = ReadIniFile(m_sSection, "FC_REFERENCELINES")
159:   m_sFCTaxCodeAnno = ReadIniFile(m_sSection, "FC_TAXCODEANNO")
160:   m_sFCLotsAnno = ReadIniFile(m_sSection, "FC_LOTSANNO")
161:   m_sFCTLAcrAnno = ReadIniFile(m_sSection, "FC_TAXLOTACREAGEANNO")
162:   m_sFCTLNumAnno = ReadIniFile(m_sSection, "FC_TAXLOTNUMBERANNO")
163:   m_sFCAnno10 = ReadIniFile(m_sSection, "FC_ANNO0010")
164:   m_sFCAnno20 = ReadIniFile(m_sSection, "FC_ANNO0020")
165:   m_sFCAnno30 = ReadIniFile(m_sSection, "FC_ANNO0030")
166:   m_sFCAnno40 = ReadIniFile(m_sSection, "FC_ANNO0040")
167:   m_sFCAnno50 = ReadIniFile(m_sSection, "FC_ANNO0050")
168:   m_sFCAnno100 = ReadIniFile(m_sSection, "FC_ANNO0100")
169:   m_sFCAnno200 = ReadIniFile(m_sSection, "FC_ANNO0200")
170:   m_sFCAnno400 = ReadIniFile(m_sSection, "FC_ANNO0400")
171:   m_sFCAnno800 = ReadIniFile(m_sSection, "FC_ANNO0800")
172:   m_sFCAnno2000 = ReadIniFile(m_sSection, "FC_ANNO2000")
173:   m_sTabCancelled = ReadIniFile(m_sSection, "TAB_Cancelled")
  
  'DEFAULT VALUES
176:   m_sDefCounty = ReadIniFile(m_sSection, "DEFAULT_COUNTY")
177:   m_sDefTownPart = ReadIniFile(m_sSection, "DEFAULT_TOWNPART")
178:   m_sDefRangePart = ReadIniFile(m_sSection, "DEFAULT_RANGEPART")
179:   m_sDefSuffType = ReadIniFile(m_sSection, "DEFAULT_SUFFTYPE")
180:   m_sDefSuffNum = ReadIniFile(m_sSection, "DEFAULT_SUFFNUM")
181:   m_sDefTownDir = ReadIniFile(m_sSection, "DEFAULT_TOWNDIR")
182:   m_sDefRangeDir = ReadIniFile(m_sSection, "DEFAULT_RANGEDIR")
183:   m_sDefQtr = ReadIniFile(m_sSection, "DEFAULT_QTR")
184:   m_sDefQtrQtr = ReadIniFile(m_sSection, "DEFAULT_QTRQTR")

End Sub

'MAPINDEX FIELDS
Public Property Get MICountyFN() As String

191:     MICountyFN = m_sMICountyFN

End Property
Public Property Get MIMapNumberFN() As String

196:     MIMapNumberFN = m_sMIMapNumberFN

End Property
Public Property Get MIORMAPMapNumberFN() As String

201:     MIORMAPMapNumberFN = m_sMIORMAPMapNumberFN

End Property
Public Property Get MIMapScaleFN() As String

206:     MIMapScaleFN = m_sMIMapScaleFN

End Property
Public Property Get MIMapSuffTypeFN() As String

211:     MIMapSuffTypeFN = m_sMIMapSuffTypeFN

End Property
Public Property Get MIMapSuffNumFN() As String

216:     MIMapSuffNumFN = m_sMIMapSuffNumFN

End Property
Public Property Get MIPageFN() As String

221:     MIPageFN = m_sMIPageFN

End Property
Public Property Get MIReliabFN() As String

226:     MIReliabFN = m_sMIReliabFN

End Property

'TAXCODE FIELDS


'TAXCODE LINES Fields


'TAXLOT FIELDS
Public Property Get TLCountyFN() As String

239:     TLCountyFN = m_sTLCountyFN

End Property
Public Property Get TLTownFN() As String

244:     TLTownFN = m_sTLTownFN

End Property
Public Property Get TLTownPartFN() As String

249:     TLTownPartFN = m_sTLTownPartFN

End Property
Public Property Get TLTownDirFN() As String

254:     TLTownDirFN = m_sTLTownDirFN

End Property
Public Property Get TLRangeFN() As String

259:     TLRangeFN = m_sTLRangeFN

End Property
Public Property Get TLRangePartFN() As String

264:     TLRangePartFN = m_sTLRangePartFN

End Property
Public Property Get TLRangeDirFN() As String

269:     TLRangeDirFN = m_sTLRangeDirFN

End Property
Public Property Get TLSectNumberFN() As String

274:     TLSectNumberFN = m_sTLSectNumberFN

End Property
Public Property Get TLQtrFN() As String

279:     TLQtrFN = m_sTLQtrFN

End Property
Public Property Get TLQtrQtrFN() As String

284:     TLQtrQtrFN = m_sTLQtrQtrFN

End Property
Public Property Get TLMapNumberFN() As String

289:     TLMapNumberFN = m_sTLMapNumberFN

End Property
Public Property Get TLOrmapMapNumberFN() As String

294:     TLOrmapMapNumberFN = m_sTLOrmapMapNumberFN

End Property
Public Property Get TLOrmapTaxlotFN() As String
298:     TLOrmapTaxlotFN = m_sTLOrmapTaxlotFN
End Property
Public Property Get TLTaxlotFN() As String
301:     TLTaxlotFN = m_sTLTaxlotFN
End Property

Public Property Get TLMapTaxlotFN() As String
305:     TLMapTaxlotFN = m_sTLMapTaxlotFN
End Property
Public Property Get TLSpecInterestFN() As String
308:     TLSpecInterestFN = m_sTLSpecInterestFN
End Property
Public Property Get TLSufNumFN() As String
311:     TLSufNumFN = m_sTLSufNumFN
End Property
Public Property Get TLSufTypeFN() As String
314:     TLSufTypeFN = m_sTLSufTypeFN
End Property
Public Property Get TLMapAcresFN() As String
317:     TLMapAcresFN = m_sTLMapAcresFN
End Property
Public Property Get TLAnomalyFN() As String
320:     TLAnomalyFN = m_sTLAnomalyFN
End Property
'TAXLOT LINES FIELDS
Public Property Get TLLinesLineTypeFN() As String
324:     TLLinesLineTypeFN = m_sTLLinesLineTypeFN
End Property
'CARTOGRAPHIC LINES FIELDS
Public Property Get CLinesLineTypeFN() As String
328:     CLinesLineTypeFN = m_sCLinesLineTypeFN
End Property
'TAXLOT POINTS FIELDS

'FIELDS PRESENT IN ALL TABLES
Public Property Get AutoDateFN() As String

335:     AutoDateFN = m_sAutoDateFN

End Property
Public Property Get AutoMethodFN() As String

340:     AutoMethodFN = m_sAutoMethodFN

End Property
Public Property Get AutoWhoFN() As String

345:     AutoWhoFN = m_sAutoWhoFN

End Property
'TABLE NAMES
Public Property Get FCMapIndex() As String

351:     FCMapIndex = m_sFCMapIndex

End Property
Public Property Get FCTaxCode() As String

356:     FCTaxCode = m_sFCTaxCode

End Property
Public Property Get FCTaxlot() As String

361:     FCTaxlot = m_sFCTaxlot

End Property
Public Property Get FCTaxlotLines() As String

366:     FCTaxlotLines = m_sFCTaxlotLines

End Property
Public Property Get FCPlats() As String

371:     FCPlats = m_sFCPlats

End Property
Public Property Get FCCartoLines() As String

376:     FCCartoLines = m_sFCCartoLines

End Property
Public Property Get FCReferenceLines() As String
380:     FCReferenceLines = m_sFCReferenceLines
End Property
Public Property Get FCTaxCodeAnno() As String

384:     FCTaxCodeAnno = m_sFCTaxCodeAnno

End Property
Public Property Get FCLotsAnno() As String

389:     FCLotsAnno = m_sFCLotsAnno

End Property
Public Property Get FCTLAcrAnno()
393:     FCTLAcrAnno = m_sFCTLAcrAnno
End Property
Public Property Get FCTLNumAnno()
396:     FCTLNumAnno = m_sFCTLNumAnno
End Property
Public Property Get FCAnno10() As String

400:     FCAnno10 = m_sFCAnno10

End Property
Public Property Get FCAnno20() As String

405:     FCAnno20 = m_sFCAnno20

End Property
Public Property Get FCAnno30() As String

410:     FCAnno30 = m_sFCAnno30

End Property
Public Property Get FCAnno40() As String

415:     FCAnno40 = m_sFCAnno40

End Property
Public Property Get FCAnno50() As String

420:     FCAnno50 = m_sFCAnno50

End Property
Public Property Get FCAnno100() As String

425:     FCAnno100 = m_sFCAnno100

End Property
Public Property Get FCAnno200() As String

430:     FCAnno200 = m_sFCAnno200

End Property
Public Property Get FCAnno400() As String

435:     FCAnno400 = m_sFCAnno400

End Property
Public Property Get FCAnno800() As String

440:     FCAnno800 = m_sFCAnno800

End Property
Public Property Get FCAnno2000() As String

445:     FCAnno2000 = m_sFCAnno2000
End Property

Public Property Get TabCancelled() As String
449:     TabCancelled = m_sTabCancelled
End Property
'DEFAULT VALUES
Public Property Get DefCounty() As String
453:     DefCounty = m_sDefCounty
End Property
Public Property Get DefTownPart() As String
456:     DefTownPart = m_sDefTownPart
End Property
Public Property Get DefRangePart() As String
459:     DefRangePart = m_sDefRangePart
End Property
Public Property Get DefSuffType() As String
462:     DefSuffType = m_sDefSuffType
End Property
Public Property Get DefSuffNum() As String
465:     DefSuffNum = m_sDefSuffNum
End Property
Public Property Get DefTownDir() As String
468:     DefTownDir = m_sDefTownDir
End Property
Public Property Get DefRangeDir() As String
471:     DefRangeDir = m_sDefRangeDir
End Property
Public Property Get DefQtr() As String
474:     DefQtr = m_sDefQtr
End Property
Public Property Get DefQtrQtr() As String
477:     DefQtrQtr = m_sDefQtrQtr
End Property
