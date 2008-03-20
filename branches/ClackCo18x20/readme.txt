'Enhancement
'Author: Jonathan McDowell, GISP  Clackamas County GIS

'Summery:  This enhancement allows Clackamas County to Add DLC information and book number to the Map layout.
'          This enhancement requires two fields added to the MapIndex name DLCName and book
'          The name of the textelement assigned to this enhancement is tbSecondTitle
'          Secondtitle is a text Field.  The length is 75 characters.
'          Book is a short integer field.
'          The string format is  DLC 1, DLC 2, DLC 3, DLC.....
'
'          Fields were also added to the the pagelayoutelements for the SecondTitle
'          The fields are sectitlex and sectitley
'          Since the book is fixed to the lower righthand corner of the map, fields were not added to the pagelayoutelements table.
'
'          I added an element that allows the book number to be placed in the lower right corner.
'          This enhancement requires a field named book to be added to the mapindex feature class.
'