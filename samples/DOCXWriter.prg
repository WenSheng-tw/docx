// Writing a .docx: title, table of contents, headings, formatted text, lists,
// tables, pictures, a hyperlink, header and footer.
// Build: hbmk2 docxwriter.hbp   -> sample.docx
#include "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main()
LOCAL oDoc, oTable, oP, nList, nNum, i, nRow, nTotal := 0
LOCAL nBold, nItalic, nRed, nWhite, nHead, nCenter, nJustify, nCode, nTotalStyle
LOCAL cLogo := hb_DirBase() + "logo.png"
LOCAL aData := { { "Olovka", 120, 35.5 }, { "Sveska A4", 45, 189.99 }, { "Gumica", 300, 12 }, ;
                 { "Fascikla", 80, 74.2 }, { "Marker", 60, 99.9 } }

hb_cdpSelect( "UTF8EX" )
SET DATE FORMAT "dd.mm.yyyy"

oDoc := WordDocument():New( "sample.docx" )
oDoc:cTitle    := "DOCX Class sample"
oDoc:cAuthor   := "S.O.K. doo Kraljevo"
oDoc:cLanguage := "sr-Latn-RS"           // spell checking language

// NewFont( cFont, nSize, lBold, lItalic, lUnderline, lStrike, cColor, cBackColor )
nBold   := oDoc:NewFont( , , .T. )
nItalic := oDoc:NewFont( , , , .T. )
nRed    := oDoc:NewFont( , , .T., , , , "C00000" )
nWhite  := oDoc:NewFont( , , .T., , , , "FFFFFF" )

// NewStyle( nFont, nAlign, nSpaceBefore, nSpaceAfter, nLineSpacing, nIndent, nFirstLine, cShading, lKeepNext, lPageBreakBefore )
nHead       := oDoc:NewStyle( nWhite, 2 )                     // table header: white bold, centered
nTotalStyle := oDoc:NewStyle( nBold, 3 )                      // bold, right
nCenter     := oDoc:NewStyle( , 2 )
nJustify    := oDoc:NewStyle( , 4, , 6, 1.15, , 1 )           // justified, first line indented 1 cm
nCode       := oDoc:NewStyle( oDoc:NewFont( "Consolas", 9, , , , , "404040" ), 1, 6, 6, 1, 0.5, , "F2F2F2" )

// header and footer: left, center, right; &P page, &N pages, &D date, &F file name
oDoc:AddHeader( "S.O.K. doo Kraljevo", "", "&D" )
oDoc:AddFooter( "&F", "", "Strana &P od &N" )

oDoc:Title( "DOCX Class" )
oDoc:Paragraph( "Word documents from Harbour, without Word, OLE or DLLs", nCenter )
oDoc:TableOfContents( "Sadržaj" )
oDoc:PageBreak()

oDoc:Heading( "Text", 1 )
oP := oDoc:Paragraph( "A paragraph is built from parts with their own font: ", nJustify )
oP:Text( "bold" , nBold ):Text( ", " ):Text( "italic", nItalic ):Text( " and " ):Text( "red", nRed )
oP:Text( ". Text is written as UTF-8, so Latin (čćžšđ ČĆŽŠĐ) and Cyrillic (Србија) both work; " + ;
         "text in another codepage, e.g. from a DBF, is converted automatically." )
oDoc:Paragraph( "First line" + Chr( 10 ) + "second line of the same paragraph" + Chr( 10 ) + "Name:" + Chr( 9 ) + "a tab" )
oDoc:Paragraph( "Link: " ):Link( "sokdoo.com", "https://sokdoo.com" )
oDoc:Paragraph( 'oDoc:Paragraph( "Hello world" )', nCode )

oDoc:Heading( "Lists", 1 )
nList := oDoc:NewList( "bullet" )
oDoc:ListItem( "First item", nList )
oDoc:ListItem( "Second item", nList )
oDoc:ListItem( "Nested item", nList, 2 )
oDoc:ListItem( "Third item", nList )
nNum := oDoc:NewList( "decimal" )
oDoc:ListItem( "Step one", nNum )
oDoc:ListItem( "Step two", nNum )
oDoc:ListItem( "Detail a", nNum, 2 )
oDoc:ListItem( "Detail b", nNum, 2 )
oDoc:ListItem( "Step three", nNum )

oDoc:Heading( "Tables", 1 )
oDoc:Heading( "Report with totals", 2 )
oTable := oDoc:Table( { 7, 2.5, 3, 3.5 } )          // column widths in cm
oTable:nHeaderRows := 1                              // repeated on every page
oTable:aPictures   := { , "@E 999,999", "@E 999,999.99", "@E 999,999.99" }
oTable:AddRow( { "Artikal", "Količina", "Cena", "Iznos" }, nHead, "305496" )
FOR i := 1 TO Len( aData )
   nRow := oTable:AddRow( { aData[ i, 1 ], aData[ i, 2 ], aData[ i, 3 ], aData[ i, 2 ] * aData[ i, 3 ] } )
   IF i % 2 == 0
      oTable:Shade( nRow, , "DEEAF6" )               // whole row
   ENDIF
   nTotal += aData[ i, 2 ] * aData[ i, 3 ]
NEXT
nRow := oTable:AddRow()
oTable:Merge( nRow, 1, nRow, 3 )
oTable:Cell( nRow, 1, "Ukupno", nTotalStyle )
oTable:Cell( nRow, 4, nTotal, nTotalStyle )

oDoc:Heading( "Merged cells, a picture in a cell", 2 )
oTable := oDoc:Table( 3 )                            // 3 equal columns over the text width
oTable:nBorder := 4                                  // dashed
oTable:cBorderColor := "808080"
oTable:Merge( 1, 1, 3, 1 )
oTable:Cell( 1, 1 ):Picture( cLogo, 3, , 2 )
oTable:VAlign( 1, 1, 2 )
oTable:Cell( 1, 2, "B1" )
oTable:Cell( 1, 3, "C1" )
oTable:Merge( 2, 2, 2, 3 )
oTable:Cell( 2, 2, "B2:C2 merged", nCenter )
oTable:Shade( 2, 2, "FFF2CC" )
oTable:Cell( 3, 2, 12.5 )
oTable:Cell( 3, 3, Date() )

oDoc:Heading( "Pictures", 1 )
oDoc:Picture( cLogo, 5, , 2 )                        // 5 cm wide, centered
oDoc:Paragraph( "A small picture " ):Picture( cLogo, , 0.5 ):Text( " inside a line of text." )

IF oDoc:Save()
   ? "Created", oDoc:cFilePath + oDoc:cName
ELSE
   ? oDoc:cError
ENDIF

RETURN
