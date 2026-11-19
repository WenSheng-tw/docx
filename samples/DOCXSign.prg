// A PDF signed with a digital certificate (A1 / A3 of ICP-Brasil, any X.509 certificate):
// the signature is visible at the row and column given, the same as @ nRow, nCol on a printer
// (6 rows and 10 columns per inch, from the top left corner of the page).
//
//   docxsign                                  a contract -> contract.pdf -> contract_signed.pdf
//   docxsign file.pdf                         signs file.pdf -> file_signed.pdf
//   docxsign file.pdf row col                 at that row and column (row < 0: rows above the bottom)
//   docxsign file.pdf row col cert.pfx pass   with a .pfx certificate (Linux: always)
//   docxsign file.pdf row col cert.pfx        the password is asked (not shown, not in the command line)
//
// Windows without a .pfx: the window of Windows to choose the certificate (A1 installed,
// A3 token or card - Windows asks for the PIN).
// Build: hbmk2 docxsign.hbp
#include "hbclass.ch"
#include "inkey.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main( cPdf, cRow, cCol, cPfx, cPass )
LOCAL oDoc, hOpt, cErr := "", cOut, nJustify, nTitle, oTable

hb_cdpSelect( "UTF8EX" )
SET DATE FORMAT "dd/mm/yyyy"

IF Empty( cPdf )
   oDoc := WordDocument():New( "contract.docx" )
   oDoc:cFont := "Arial"
   nTitle := oDoc:NewStyle( oDoc:NewFont( , 16, .T. ), 2 )   // centered, bold 16
   nJustify := oDoc:NewStyle( , 4, , 6, 1.15 )                // justified
   oDoc:Paragraph( "Service agreement", nTitle )
   oDoc:Paragraph()
   oDoc:Paragraph( "Between Malc Informática Ltda. (the provider) and Clínica Demonstração Ltda. " + ;
                   "(the client), on " + DToC( Date() ) + ". The provider keeps the occupational " + ;
                   "health system of the client running and updated for twelve months.", nJustify )
   oDoc:Paragraph()
   oDoc:Paragraph( "This document is signed digitally; the signature below can be checked in any " + ;
                   "PDF reader (Adobe Reader: the Signatures panel).", nJustify )
   oDoc:Paragraph()
   oTable := oDoc:Table( { 8, 1, 8 } )
   oTable:nBorder := 0
   oTable:AddRow( { "Marília, " + DToC( Date() ), "", "" } )
   IF ! oDoc:SavePDF( "contract.pdf" )
      ? oDoc:cError
      RETURN
   ENDIF
   ? "Created contract.pdf"
   cPdf := "contract.pdf"
ENDIF

/* where and what: rows and columns as on a printer; the text with {name} {date} {reason} {location} */
hOpt := { "row"      => IIF( Empty( cRow ), -10, Val( cRow ) ), ;
          "col"      => IIF( Empty( cCol ), 10, Val( cCol ) ), ;
          "width"    => 90, ;                             // mm
          "height"   => 20, ;
          "page"     => 0, ;                              // the last page
          "reason"   => "Agreement of the parties", ;
          "location" => "Marília - SP", ;
          "text"     => { "Assinado digitalmente por {name}", "Data: {date}", "Motivo: {reason}", "Local: {location}" } }
IF ! Empty( cPfx )
   hOpt[ "pfx" ] := cPfx
   hOpt[ "password" ] := IIF( cPass == NIL, AskPassword( cPfx ), cPass )
ENDIF

cOut := hb_FNameExtSet( cPdf, "" ) + "_signed.pdf"
IF PDF_Sign( cPdf, cOut, hOpt, @cErr )
   ? "Signed:", cOut
ELSE
   ? "Not signed:", cErr
ENDIF
RETURN

/* the password of the .pfx, typed without showing it */
STATIC FUNCTION AskPassword( cPfx )

LOCAL cPass := "", nKey

? "Password of " + cPfx + ": "
DO WHILE ( nKey := Inkey( 0 ) ) != K_ENTER
   DO CASE
   CASE nKey == K_ESC
      cPass := ""
      EXIT
   CASE nKey == K_BS
      IF ! Empty( cPass )
         cPass := Left( cPass, Len( cPass ) - 1 )
         ?? Chr( 8 ) + " " + Chr( 8 )
      ENDIF
   CASE Len( hb_keyChar( nKey ) ) > 0
      cPass += hb_keyChar( nKey )
      ?? "*"
   ENDCASE
ENDDO
?
RETURN cPass
