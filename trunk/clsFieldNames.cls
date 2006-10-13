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

183:   If Not TestIniFile(app.Path & "\" & INIFILE) Then
184:     MsgBox "Unable to open Inifile at " & app.Path & "\" & INIFILE, vbExclamation
    Exit Sub
186:   End If
  
  'Read ini file and set references to field names
  
  'MAPINDEX FIELDS
191:   m_sMICountyFN = ReadIniFile(m_sSection, "MI_COUNTY")
192:   m_sMIMapNumberFN = ReadIniFile(m_sSection, "MI_MAPNUMBER")
193:   m_sMIORMAPMapNumberFN = ReadIniFile(m_sSection, "MI_ORMAP_MAPNUMBER")
194:   m_sMIMapScaleFN = ReadIniFile(m_sSection, "MI_MAPSCALE")
195:   m_sMIMapSuffTypeFN = ReadIniFile(m_sSection, "MI_MAPSUFFIXTYPE")
196:   m_sMIMapSuffNumFN = ReadIniFile(m_sSection, "MI_MAPSUFFIXNUM")
197:   m_sMIPageFN = ReadIniFile(m_sSection, "MI_PAGE")
198:   m_sMIReliabFN = ReadIniFile(m_sSection, "MI_RELIABILITYCODE")


  'TAXCODE FIELDS

 'TAXCODE LINES Fields
  
  'TAXLOT FIELDS
206:   m_sTLCountyFN = ReadIniFile(m_sSection, "TL_COUNTY")
207:   m_sTLTownFN = ReadIniFile(m_sSection, "TL_TOWN")
208:   m_sTLTownPartFN = ReadIniFile(m_sSection, "TL_TOWNPART")
209:   m_sTLTownDirFN = ReadIniFile(m_sSection, "TL_TOWNDIR")
210:   m_sTLRangeFN = ReadIniFile(m_sSection, "TL_RANGE")
211:   m_sTLRangePartFN = ReadIniFile(m_sSection, "TL_RANGEPART")
212:   m_sTLRangeDirFN = ReadIniFile(m_sSection, "TL_RANGEDIR")
213:   m_sTLSectNumberFN = ReadIniFile(m_sSection, "TL_SECTNUMBER")
214:   m_sTLQtrFN = ReadIniFile(m_sSection, "TL_QTR")
215:   m_sTLQtrQtrFN = ReadIniFile(m_sSection, "TL_QTRQTR")
216:   m_sTLMapNumberFN = ReadIniFile(m_sSection, "TL_MAPNUMBER")
217:   m_sTLOrmapMapNumberFN = ReadIniFile(m_sSection, "TL_ORMAP_MAPNUMBER")
218:   m_sTLOrmapTaxlotFN = ReadIniFile(m_sSection, "TL_ORMAP_TAXLOT")
219:   m_sTLTaxlotFN = ReadIniFile(m_sSection, "TL_TAXLOT")
220:   m_sTLMapTaxlotFN = ReadIniFile(m_sSection, "TL_MAP_TAXLOT")
221:   m_sTLSpecInterestFN = ReadIniFile(m_sSection, "TL_SPEC_INTEREST")
222:   m_sTLSufNumFN = ReadIniFile(m_sSection, "TL_MAP_SUF_NUM")
223:   m_sTLSufTypeFN = ReadIniFile(m_sSection, "TL_MAP_SUF_TYPE")
224:   m_sTLMapAcresFN = ReadIniFile(m_sSection, "TL_MAPACRES")
225:   m_sTLAnomalyFN = ReadIniFile(m_sSection, "TL_ANOMALY")


  'TAXLOT LINES FIELDS
229:   m_sTLLinesLineTypeFN = ReadIniFile(m_sSection, "TLLINES_LINETYPE")
  
  'CARTOGRAPHIC LINES FIELDS
232:   m_sCLinesLineTypeFN = ReadIniFile(m_sSection, "CARTOLINES_LINETYPE")


  'TAXLOT POINTS FIELDS
  
  'FIELDS PRESENT IN ALL TABLES
238:   m_sAutoDateFN = ReadIniFile(m_sSection, "AUTO_DATE")
239:   m_sAutoMethodFN = ReadIniFile(m_sSection, "AUTO_METHOD")
240:   m_sAutoWhoFN = ReadIniFile(m_sSection, "AUTO_WHO")
  'TABLE NAMES
242:   m_sFCMapIndex = ReadIniFile(m_sSection, "FC_MAPINDEX")
243:   m_sFCTaxCode = ReadIniFile(m_sSection, "FC_TAXCODE")
244:   m_sFCTaxlot = ReadIniFile(m_sSection, "FC_TAXLOT")
245:   m_sFCTaxlotLines = ReadIniFile(m_sSection, "FC_TAXLOTLINES")
246:   m_sFCPlats = ReadIniFile(m_sSection, "FC_PLATS")
247:   m_sFCCartoLines = ReadIniFile(m_sSection, "FC_CARTOLINES")
248:   m_sFCReferenceLines = ReadIniFile(m_sSection, "FC_REFERENCELINES")
249:   m_sFCTaxCodeAnno = ReadIniFile(m_sSection, "FC_TAXCODEANNO")
250:   m_sFCLotsAnno = ReadIniFile(m_sSection, "FC_LOTSANNO")
251:   m_sFCTLAcrAnno = ReadIniFile(m_sSection, "FC_TAXLOTACREAGEANNO")
252:   m_sFCTLNumAnno = ReadIniFile(m_sSection, "FC_TAXLOTNUMBERANNO")
253:   m_sFCAnno10 = ReadIniFile(m_sSection, "FC_ANNO0010")
254:   m_sFCAnno20 = ReadIniFile(m_sSection, "FC_ANNO0020")
255:   m_sFCAnno30 = ReadIniFile(m_sSection, "FC_ANNO0030")
256:   m_sFCAnno40 = ReadIniFile(m_sSection, "FC_ANNO0040")
257:   m_sFCAnno50 = ReadIniFile(m_sSection, "FC_ANNO0050")
258:   m_sFCAnno100 = ReadIniFile(m_sSection, "FC_ANNO0100")
259:   m_sFCAnno200 = ReadIniFile(m_sSection, "FC_ANNO0200")
260:   m_sFCAnno400 = ReadIniFile(m_sSection, "FC_ANNO0400")
261:   m_sFCAnno800 = ReadIniFile(m_sSection, "FC_ANNO0800")
262:   m_sFCAnno2000 = ReadIniFile(m_sSection, "FC_ANNO2000")
263:   m_sTabCancelled = ReadIniFile(m_sSection, "TAB_Cancelled")
  
  'DEFAULT VALUES
266:   m_sDefCounty = ReadIniFile(m_sSection, "DEFAULT_COUNTY")
267:   m_sDefTownPart = ReadIniFile(m_sSection, "DEFAULT_TOWNPART")
268:   m_sDefRangePart = ReadIniFile(m_sSection, "DEFAULT_RANGEPART")
269:   m_sDefSuffType = ReadIniFile(m_sSection, "DEFAULT_SUFFTYPE")
270:   m_sDefSuffNum = ReadIniFile(m_sSection, "DEFAULT_SUFFNUM")
271:   m_sDefTownDir = ReadIniFile(m_sSection, "DEFAULT_TOWNDIR")
272:   m_sDefRangeDir = ReadIniFile(m_sSection, "DEFAULT_RANGEDIR")
273:   m_sDefQtr = ReadIniFile(m_sSection, "DEFAULT_QTR")
274:   m_sDefQtrQtr = ReadIniFile(m_sSection, "DEFAULT_QTRQTR")
275:   m_sDefAnomaly = ReadIniFile(m_sSection, "DEFAULT_ANOMALY")

  '++ New coded added 10/2/05
  'ANNO TEXT SIZES
279:   m_dAnnoSizeTLAcr120 = ReadIniFile(m_sSection, "ANNOSIZETLACR120")
280:   m_dAnnoSizeTLAcr240 = ReadIniFile(m_sSection, "ANNOSIZETLACR240")
281:   m_dAnnoSizeTLAcr360 = ReadIniFile(m_sSection, "ANNOSIZETLACR360")
282:   m_dAnnoSizeTLAcr480 = ReadIniFile(m_sSection, "ANNOSIZETLACR480")
283:   m_dAnnoSizeTLAcr600 = ReadIniFile(m_sSection, "ANNOSIZETLACR600")
284:   m_dAnnoSizeTLAcr1200 = ReadIniFile(m_sSection, "ANNOSIZETLACR1200")
285:   m_dAnnoSizeTLAcr2400 = ReadIniFile(m_sSection, "ANNOSIZETLACR2400")
286:   m_dAnnoSizeTLAcr4800 = ReadIniFile(m_sSection, "ANNOSIZETLACR4800")
287:   m_dAnnoSizeTLAcr9600 = ReadIniFile(m_sSection, "ANNOSIZETLACR9600")
288:   m_dAnnoSizeTLAcr24000 = ReadIniFile(m_sSection, "ANNOSIZETLACR24000")
289:   m_dAnnoSizeTLNum120 = ReadIniFile(m_sSection, "ANNOSIZETLNUM120")
290:   m_dAnnoSizeTLNum240 = ReadIniFile(m_sSection, "ANNOSIZETLNUM240")
291:   m_dAnnoSizeTLNum360 = ReadIniFile(m_sSection, "ANNOSIZETLNUM360")
292:   m_dAnnoSizeTLNum480 = ReadIniFile(m_sSection, "ANNOSIZETLNUM480")
293:   m_dAnnoSizeTLNum600 = ReadIniFile(m_sSection, "ANNOSIZETLNUM600")
294:   m_dAnnoSizeTLNum1200 = ReadIniFile(m_sSection, "ANNOSIZETLNUM1200")
295:   m_dAnnoSizeTLNum2400 = ReadIniFile(m_sSection, "ANNOSIZETLNUM2400")
296:   m_dAnnoSizeTLNum4800 = ReadIniFile(m_sSection, "ANNOSIZETLNUM4800")
297:   m_dAnnoSizeTLNum9600 = ReadIniFile(m_sSection, "ANNOSIZETLNUM9600")
298:   m_dAnnoSizeTLNum24000 = ReadIniFile(m_sSection, "ANNOSIZETLNUM24000")
  '++ End new code
End Sub

'MAPINDEX FIELDS
Public Property Get MICountyFN() As String

305:     MICountyFN = m_sMICountyFN

End Property
Public Property Get MIMapNumberFN() As String

310:     MIMapNumberFN = m_sMIMapNumberFN

End Property
Public Property Get MIORMAPMapNumberFN() As String

315:     MIORMAPMapNumberFN = m_sMIORMAPMapNumberFN

End Property
Public Property Get MIMapScaleFN() As String

320:     MIMapScaleFN = m_sMIMapScaleFN

End Property
Public Property Get MIMapSuffTypeFN() As String

325:     MIMapSuffTypeFN = m_sMIMapSuffTypeFN

End Property
Public Property Get MIMapSuffNumFN() As String

330:     MIMapSuffNumFN = m_sMIMapSuffNumFN

End Property
Public Property Get MIPageFN() As String

335:     MIPageFN = m_sMIPageFN

End Property
Public Property Get MIReliabFN() As String

340:     MIReliabFN = m_sMIReliabFN

End Property

'TAXCODE FIELDS


'TAXCODE LINES Fields


'TAXLOT FIELDS
Public Property Get TLCountyFN() As String

353:     TLCountyFN = m_sTLCountyFN

End Property
Public Property Get TLTownFN() As String

358:     TLTownFN = m_sTLTownFN

End Property
Public Property Get TLTownPartFN() As String

363:     TLTownPartFN = m_sTLTownPartFN

End Property
Public Property Get TLTownDirFN() As String

368:     TLTownDirFN = m_sTLTownDirFN

End Property
Public Property Get TLRangeFN() As String

373:     TLRangeFN = m_sTLRangeFN

End Property
Public Property Get TLRangePartFN() As String

378:     TLRangePartFN = m_sTLRangePartFN

End Property
Public Property Get TLRangeDirFN() As String

383:     TLRangeDirFN = m_sTLRangeDirFN

End Property
Public Property Get TLSectNumberFN() As String

388:     TLSectNumberFN = m_sTLSectNumberFN

End Property
Public Property Get TLQtrFN() As String

393:     TLQtrFN = m_sTLQtrFN

End Property
Public Property Get TLQtrQtrFN() As String

398:     TLQtrQtrFN = m_sTLQtrQtrFN

End Property
Public Property Get TLMapNumberFN() As String

403:     TLMapNumberFN = m_sTLMapNumberFN

End Property
Public Property Get TLOrmapMapNumberFN() As String

408:     TLOrmapMapNumberFN = m_sTLOrmapMapNumberFN

End Property
Public Property Get TLOrmapTaxlotFN() As String
412:     TLOrmapTaxlotFN = m_sTLOrmapTaxlotFN
End Property
Public Property Get TLTaxlotFN() As String
415:     TLTaxlotFN = m_sTLTaxlotFN
End Property

Public Property Get TLMapTaxlotFN() As String
419:     TLMapTaxlotFN = m_sTLMapTaxlotFN
End Property
Public Property Get TLSpecInterestFN() As String
422:     TLSpecInterestFN = m_sTLSpecInterestFN
End Property
Public Property Get TLSufNumFN() As String
425:     TLSufNumFN = m_sTLSufNumFN
End Property
Public Property Get TLSufTypeFN() As String
428:     TLSufTypeFN = m_sTLSufTypeFN
End Property
Public Property Get TLMapAcresFN() As String
431:     TLMapAcresFN = m_sTLMapAcresFN
End Property
Public Property Get TLAnomalyFN() As String
434:     TLAnomalyFN = m_sTLAnomalyFN
End Property
'TAXLOT LINES FIELDS
Public Property Get TLLinesLineTypeFN() As String
438:     TLLinesLineTypeFN = m_sTLLinesLineTypeFN
End Property
'CARTOGRAPHIC LINES FIELDS
Public Property Get CLinesLineTypeFN() As String
442:     CLinesLineTypeFN = m_sCLinesLineTypeFN
End Property
'TAXLOT POINTS FIELDS

'FIELDS PRESENT IN ALL TABLES
Public Property Get AutoDateFN() As String

449:     AutoDateFN = m_sAutoDateFN

End Property
Public Property Get AutoMethodFN() As String

454:     AutoMethodFN = m_sAutoMethodFN

End Property
Public Property Get AutoWhoFN() As String

459:     AutoWhoFN = m_sAutoWhoFN

End Property
'TABLE NAMES
Public Property Get FCMapIndex() As String

465:     FCMapIndex = m_sFCMapIndex

End Property
Public Property Get FCTaxCode() As String

470:     FCTaxCode = m_sFCTaxCode

End Property
Public Property Get FCTaxlot() As String

475:     FCTaxlot = m_sFCTaxlot

End Property
Public Property Get FCTaxlotLines() As String

480:     FCTaxlotLines = m_sFCTaxlotLines

End Property
Public Property Get FCPlats() As String

485:     FCPlats = m_sFCPlats

End Property
Public Property Get FCCartoLines() As String

490:     FCCartoLines = m_sFCCartoLines

End Property
Public Property Get FCReferenceLines() As String
494:     FCReferenceLines = m_sFCReferenceLines
End Property
Public Property Get FCTaxCodeAnno() As String

498:     FCTaxCodeAnno = m_sFCTaxCodeAnno

End Property
Public Property Get FCLotsAnno() As String

503:     FCLotsAnno = m_sFCLotsAnno

End Property
Public Property Get FCTLAcrAnno()
507:     FCTLAcrAnno = m_sFCTLAcrAnno
End Property
Public Property Get FCTLNumAnno()
510:     FCTLNumAnno = m_sFCTLNumAnno
End Property
Public Property Get FCAnno10() As String

514:     FCAnno10 = m_sFCAnno10

End Property
Public Property Get FCAnno20() As String

519:     FCAnno20 = m_sFCAnno20

End Property
Public Property Get FCAnno30() As String

524:     FCAnno30 = m_sFCAnno30

End Property
Public Property Get FCAnno40() As String

529:     FCAnno40 = m_sFCAnno40

End Property
Public Property Get FCAnno50() As String

534:     FCAnno50 = m_sFCAnno50

End Property
Public Property Get FCAnno100() As String

539:     FCAnno100 = m_sFCAnno100

End Property
Public Property Get FCAnno200() As String

544:     FCAnno200 = m_sFCAnno200

End Property
Public Property Get FCAnno400() As String

549:     FCAnno400 = m_sFCAnno400

End Property
Public Property Get FCAnno800() As String

554:     FCAnno800 = m_sFCAnno800

End Property
Public Property Get FCAnno2000() As String

559:     FCAnno2000 = m_sFCAnno2000
End Property

Public Property Get TabCancelled() As String
563:     TabCancelled = m_sTabCancelled
End Property
'DEFAULT VALUES
Public Property Get DefCounty() As String
567:     DefCounty = m_sDefCounty
End Property
Public Property Get DefTownPart() As String
570:     DefTownPart = m_sDefTownPart
End Property
Public Property Get DefRangePart() As String
573:     DefRangePart = m_sDefRangePart
End Property
Public Property Get DefSuffType() As String
576:     DefSuffType = m_sDefSuffType
End Property
Public Property Get DefSuffNum() As String
579:     DefSuffNum = m_sDefSuffNum
End Property
Public Property Get DefTownDir() As String
582:     DefTownDir = m_sDefTownDir
End Property
Public Property Get DefRangeDir() As String
585:     DefRangeDir = m_sDefRangeDir
End Property
Public Property Get DefQtr() As String
588:     DefQtr = m_sDefQtr
End Property
Public Property Get DefQtrQtr() As String
591:     DefQtrQtr = m_sDefQtrQtr
End Property
Public Property Get DefAnomaly() As String
594:     DefAnomaly = m_sDefAnomaly
End Property
'++ New coded added 10/2/05
'ANNO TEXT SIZES
Public Property Get AnnoSizeTLAcr120() As Double
599:     AnnoSizeTLAcr120 = m_dAnnoSizeTLAcr120
End Property

Public Property Get AnnoSizeTLAcr240() As Double
603:     AnnoSizeTLAcr240 = m_dAnnoSizeTLAcr240
End Property

Public Property Get AnnoSizeTLAcr360() As Double
607:     AnnoSizeTLAcr360 = m_dAnnoSizeTLAcr360
End Property

Public Property Get AnnoSizeTLAcr480() As Double
611:     AnnoSizeTLAcr480 = m_dAnnoSizeTLAcr480
End Property

Public Property Get AnnoSizeTLAcr600() As Double
615:     AnnoSizeTLAcr600 = m_dAnnoSizeTLAcr600
End Property

Public Property Get AnnoSizeTLAcr1200() As Double
619:     AnnoSizeTLAcr1200 = m_dAnnoSizeTLAcr1200
End Property

Public Property Get AnnoSizeTLAcr2400() As Double
623:     AnnoSizeTLAcr2400 = m_dAnnoSizeTLAcr2400
End Property

Public Property Get AnnoSizeTLAcr4800() As Double
627:     AnnoSizeTLAcr4800 = m_dAnnoSizeTLAcr4800
End Property

Public Property Get AnnoSizeTLAcr9600() As Double
631:     AnnoSizeTLAcr9600 = m_dAnnoSizeTLAcr9600
End Property

Public Property Get AnnoSizeTLAcr24000() As Double
635:     AnnoSizeTLAcr24000 = m_dAnnoSizeTLAcr24000
End Property

Public Property Get AnnoSizeTLNum120() As Double
639:     AnnoSizeTLNum120 = m_dAnnoSizeTLNum120
End Property

Public Property Get AnnoSizeTLNum240() As Double
643:     AnnoSizeTLNum240 = m_dAnnoSizeTLNum240
End Property

Public Property Get AnnoSizeTLNum360() As Double
647:     AnnoSizeTLNum360 = m_dAnnoSizeTLNum360
End Property

Public Property Get AnnoSizeTLNum480() As Double
651:     AnnoSizeTLNum480 = m_dAnnoSizeTLNum480
End Property

Public Property Get AnnoSizeTLNum600() As Double
655:     AnnoSizeTLNum600 = m_dAnnoSizeTLNum600
End Property

Public Property Get AnnoSizeTLNum1200() As Double
659:     AnnoSizeTLNum1200 = m_dAnnoSizeTLNum1200
End Property

Public Property Get AnnoSizeTLNum2400() As Double
663:     AnnoSizeTLNum2400 = m_dAnnoSizeTLNum2400
End Property

Public Property Get AnnoSizeTLNum4800() As Double
667:     AnnoSizeTLNum4800 = m_dAnnoSizeTLNum4800
End Property

Public Property Get AnnoSizeTLNum9600() As Double
671:     AnnoSizeTLNum9600 = m_dAnnoSizeTLNum9600
End Property

Public Property Get AnnoSizeTLNum24000() As Double
675:     AnnoSizeTLNum24000 = m_dAnnoSizeTLNum24000
End Property
'++ End new code
