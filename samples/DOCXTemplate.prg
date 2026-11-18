// Filling a template: a document with placeholders like {{customer}} is read,
// the placeholders are replaced, a table row is repeated for every item and new
// content is inserted in place of a marker paragraph.
// Step 1 creates template.docx (normally it is made in Word), step 2 fills it.
// Build: hbmk2 docxtemplate.hbp   -> template.docx, invoice.docx
#include "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main()
LOCAL oDoc, aItems := {}, aRow, nTotal := 0, nList

hb_cdpSelect( "UTF8EX" )
SET DATE FORMAT "dd.mm.yyyy"

MakeTemplate( "template.docx" )

oDoc := WordDocument():New()
IF ! oDoc:Read( "template.docx" )
   ? oDoc:cError
   RETURN
ENDIF

// one value or a hash of values; dates and numbers are converted to text
oDoc:Replace( { "{{customer}}" => "Petar Petrović", ;
                "{{address}}"  => "Kralja Petra I 1" + Chr( 10 ) + "36000 Kraljevo", ;   // Chr(10) = new line
                "{{number}}"   => "2026-118", ;
                "{{date}}"     => Date() } )

// the row with {{item}} is repeated for every element; elements are hashes,
// or arrays with the placeholder names in the third parameter
FOR EACH aRow IN { { "Olovka", 120, 35.5 }, { "Sveska A4", 45, 189.99 }, { "Gumica", 300, 12 } }
   AAdd( aItems, { aRow[ 1 ], aRow[ 2 ], Transform( aRow[ 3 ], "@E 999,999.99" ), Transform( aRow[ 2 ] * aRow[ 3 ], "@E 999,999.99" ) } )
   nTotal += aRow[ 2 ] * aRow[ 3 ]
NEXT
oDoc:RepeatRow( "{{item}}", aItems, { "{{item}}", "{{qty}}", "{{price}}", "{{amount}}" } )
oDoc:Replace( "{{total}}", Transform( nTotal, "@E 999,999.99" ) )

// new content goes in place of the paragraph with {{notes}}
oDoc:InsertAt( "{{notes}}" )
oDoc:Heading( "Napomene", 2 )
nList := oDoc:NewList( "decimal" )
oDoc:ListItem( "Rok plaćanja je 15 dana.", nList )
oDoc:ListItem( "Reklamacije u roku od 8 dana.", nList )

oDoc:cName := "invoice.docx"                  // save under another name
IF oDoc:Save()
   ? "Created", oDoc:cFilePath + oDoc:cName
ELSE
   ? oDoc:cError
ENDIF

RETURN

// a template made with the class itself, so the sample is self-contained
STATIC PROCEDURE MakeTemplate( cFile )
LOCAL oDoc, oTable, nBold, nHead, nRight

oDoc := WordDocument():New( cFile )
nBold  := oDoc:NewFont( , , .T. )
nHead  := oDoc:NewStyle( oDoc:NewFont( , , .T., , , , "FFFFFF" ), 2 )
nRight := oDoc:NewStyle( nBold, 3 )

oDoc:AddHeader( "S.O.K. doo Kraljevo", "", "Račun {{number}}" )
oDoc:AddFooter( "", "Strana &P od &N", "" )

oDoc:Title( "Račun br. {{number}}" )
oDoc:Paragraph( "Datum: " ):Text( "{{date}}", nBold )
oDoc:Paragraph( "Kupac:" )
oDoc:Paragraph( "{{customer}}", oDoc:NewStyle( nBold, , , 0 ) )
oDoc:Paragraph( "{{address}}" )

oTable := oDoc:Table( { 7, 2.5, 3, 3.5 } )
oTable:nHeaderRows := 1
oTable:AddRow( { "Artikal", "Količina", "Cena", "Iznos" }, nHead, "305496" )
oTable:aColStyles := { , oDoc:NewStyle( , 3 ), oDoc:NewStyle( , 3 ), oDoc:NewStyle( , 3 ) }   // right aligned
oTable:AddRow( { "{{item}}", "{{qty}}", "{{price}}", "{{amount}}" } )
oTable:AddRow( { "Ukupno", , , "{{total}}" }, nRight )
oTable:Merge( 3, 1, 3, 3 )

oDoc:Paragraph()
oDoc:Paragraph( "{{notes}}" )
oDoc:Paragraph( "Hvala na saradnji." )
oDoc:Save()

RETURN
