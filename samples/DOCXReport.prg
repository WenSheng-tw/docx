// A business report as programs used to make it with Word through OLE - without Word:
// cover page without header / footer, header with the logo and the company in a table,
// "Page X of Y" in the footer (numbering can start at any page), table of contents,
// justified text, lists, a table with a repeated header row, merged and shaded cells,
// a signature table; then the same document as PDF (Word or LibreOffice).
// Build: hbmk2 docxreport.hbp   -> report.docx, report.pdf
//   docxreport sign               -> also report_signed.pdf: signed with a certificate
//                                    (Windows: chosen in the window of Windows; see DOCXSign.prg)
#include "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main( cSign )
LOCAL oDoc, oHead, oTable, oCell, aRow, nRow, nGroup, cGroup := ""
LOCAL nTitle, nBig, nBold, nWhite, nSmall
LOCAL nCenter, nRight, nJustify, nHeadCell, nBoldCell, nSign
LOCAL cLogo := hb_DirBase() + "logo.png", cErr := ""
LOCAL aExams := { ;
   { "Administration", "Ana Souza",      "Clinical exam",     SToD( "20261103" ), "Fit" }, ;
   { "Administration", "Bruno Lima",     "Clinical exam",     SToD( "20261103" ), "Fit" }, ;
   { "Production",     "Carlos Pereira", "Audiometry",        SToD( "20261105" ), "Fit with restrictions" }, ;
   { "Production",     "Daniela Costa",  "Spirometry",        SToD( "20261105" ), "Fit" }, ;
   { "Production",     "Eduardo Alves",  "Clinical exam",     SToD( "20261106" ), "Fit" }, ;
   { "Warehouse",      "Fernanda Rocha", "Visual acuity",     SToD( "20261110" ), "Fit" }, ;
   { "Warehouse",      "Gustavo Ramos",  "Clinical exam",     SToD( "20261110" ), "Unfit" } }

hb_cdpSelect( "UTF8EX" )
SET DATE FORMAT "dd/mm/yyyy"

oDoc := WordDocument():New( "report.docx" )
oDoc:cTitle      := "Occupational health report"
oDoc:cAuthor     := "Malc Informatica"
oDoc:cFont       := "Arial"
oDoc:cLanguage   := "pt-BR"
oDoc:nSpaceAfter := 4
oDoc:nLeftMargin := 1.8                  // cm (Word OLE PageSetup takes points: 50 pt = 1.76 cm)
oDoc:nRightMargin := 1.2
oDoc:nTopMargin  := 1.2
oDoc:nBottomMargin := 1.2
oDoc:lTitlePage  := .T.                  // the cover page has no header and footer
// oDoc:nPageStart := 15                 // numbering from page 15: "Page 15 of 18" ... "Page 18 of 18"

// fonts and paragraph styles (NewFont / NewStyle return a number used later)
nTitle  := oDoc:NewFont( , 24, .T., , , , "1F3864" )
nBig    := oDoc:NewFont( , 13, .T. )
nBold   := oDoc:NewFont( , , .T. )
nWhite  := oDoc:NewFont( , 10, .T., , , , "FFFFFF" )
nSmall  := oDoc:NewFont( , 9 )
nCenter   := oDoc:NewStyle( , 2 )
nRight    := oDoc:NewStyle( , 3 )
nJustify  := oDoc:NewStyle( , 4, , 6, 1.15 )
nHeadCell := oDoc:NewStyle( nWhite, 2, 0, 0, 1 )
nBoldCell := oDoc:NewStyle( nBold, 1, 0, 0, 1 )
nSign     := oDoc:NewStyle( nSmall, 2, 0, 0, 1 )

// header: logo on the left, the company on the right (Malc_Word:SetCabecalhoPadrao)
oHead := oDoc:Header():Table( { 2.2, 15.8 } )
oHead:nBorder := 0
oHead:AddRow()
oHead:VAlign( 1, NIL, 2 )
oHead:Cell( 1, 1 ):Picture( cLogo, 1.8 )
oCell := oHead:Cell( 1, 2 )
oCell:Paragraph( "MALC INFORMATICA LTDA", oDoc:NewStyle( nBig, 1, 0, 0, 1 ) )
oCell:Paragraph( "Rua das Flores, 100 - Phone: (14) 3333-0000", oDoc:NewStyle( nSmall, 1, 0, 0, 1 ) )
oCell:Paragraph( "Marilia/SP", oDoc:NewStyle( nSmall, 1, 0, 0, 1 ) )

// footer: "Page X of Y" on the right (Malc_Word:SetNumeracao); &D date, &F file name
oDoc:AddFooter( "&F", "", "Page &P of &" + "N", nSmall )

// ---- cover page (Malc_Word:CapaDocto)
oDoc:Paragraph()
oDoc:Picture( cLogo, 4, , 2 )
oDoc:Paragraph()
oDoc:Paragraph( "OCCUPATIONAL HEALTH REPORT", oDoc:NewStyle( nTitle, 2, 48, 12 ) )
oDoc:Paragraph( "Periodic examinations - November 2026", oDoc:NewStyle( nBig, 2, 0, 48 ) )
oTable := oDoc:Table( { 6, 6 } )
oTable:nAlign := 2
oTable:nHeaderRows := 1
oTable:AddRow( { "RESPONSIBLE TECHNICIAN", "OCCUPATIONAL PHYSICIAN" }, nHeadCell, "1F3864" )
oTable:AddRow( { "João da Silva" + Chr( 10 ) + "Safety engineer", "Dr. Maria Oliveira" + Chr( 10 ) + "CRM 123456" }, nSign )
oTable:RowHeight( 2, 1.5 )
oTable:VAlign( 2, NIL, 2 )
oDoc:PageBreak()

// ---- table of contents (Malc_Word:AddSumario); Word fills it when the file is opened
oDoc:TableOfContents( "Contents" )
oDoc:PageBreak()

// ---- text (Malc_Word:AddTitulo, Texto)
oDoc:Heading( "1. Introduction", 1 )
oDoc:Paragraph( "This report lists the periodic examinations of the employees made in November 2026, " + ;
                "with the result of every examination. It was made by a Harbour program with DOCXClass, " + ;
                "without Word: the class writes the XML of the document and packs it into a .docx file.", nJustify )
oDoc:Paragraph( "Results: ", nJustify ):Text( "Fit", nBold ):Text( ", " ):Text( "Fit with restrictions", nBold ):Text( " or " ):Text( "Unfit", nBold ):Text( "." )

// ---- list (Malc_Word:AddBullet)
oDoc:Heading( "2. Examinations", 1 )
oDoc:ListItem( "Clinical exam - every employee" )
oDoc:ListItem( "Audiometry - employees exposed to noise" )
oDoc:ListItem( "Spirometry - employees exposed to dust" )
oDoc:ListItem( "Visual acuity - drivers and forklift operators" )

// ---- table (Malc_Word:TabelaCriar, TabelaTexto, TabelaMesclar): a row per department,
// merged over all columns and shaded, the header row repeated on every page
oDoc:Heading( "3. Results", 1 )
oTable := oDoc:Table( { 5.5, 4, 2.6, 4.2 } )
oTable:nHeaderRows := 1
oTable:aPictures := { , , "@D", }
oTable:AddRow( { "Employee", "Examination", "Date", "Result" }, nHeadCell, "305496" )
nGroup := 0
FOR EACH aRow IN aExams
   IF !( aRow[ 1 ] == cGroup )
      cGroup := aRow[ 1 ]
      nRow := oTable:AddRow( { cGroup }, nBoldCell, "D9E2F3" )
      oTable:Merge( nRow, 1, nRow, 4 )
      nGroup++
   ENDIF
   nRow := oTable:AddRow( { aRow[ 2 ], aRow[ 3 ], aRow[ 4 ], aRow[ 5 ] } )
   IF aRow[ 5 ] == "Unfit"
      oTable:Shade( nRow, 4, "F8CBAD" )
   ENDIF
NEXT
nRow := oTable:AddRow()
oTable:Merge( nRow, 1, nRow, 3 )
oTable:Cell( nRow, 1, "Employees examined", oDoc:NewStyle( nBold, 3 ) )
oTable:Cell( nRow, 4, Len( aExams ), oDoc:NewStyle( nBold, 2 ) )
oTable:Shade( nRow, NIL, "F2F2F2" )
oDoc:Paragraph( hb_ntos( nGroup ) + " departments, " + hb_ntos( Len( aExams ) ) + " employees.", nRight )

// ---- signatures (Malc_Word:Protocolo)
oDoc:Heading( "4. Signatures", 1 )
oDoc:Paragraph( "Marilia, " + DToC( Date() ) + ".", nJustify )
oDoc:Paragraph()
oTable := oDoc:Table( { 8, 1, 8 } )
oTable:nBorder := 0
oTable:AddRow( { Replicate( "_", 34 ), "", Replicate( "_", 34 ) }, nCenter )
oTable:AddRow( { "João da Silva" + Chr( 10 ) + "Safety engineer", "", "Dr. Maria Oliveira" + Chr( 10 ) + "Occupational physician" }, nSign )

IF ! oDoc:Save()
   ? "Error:", oDoc:cError
   RETURN
ENDIF
? "Created", oDoc:cName

// PDF: Word if it is installed, else LibreOffice (Malc_Word:Finalizar with lGerarPDF)
IF oDoc:SavePDF( "report.pdf" )
   ? "Created report.pdf"
   // a certificate chosen: the PDF signed, the signature on the last page, 6 rows above the bottom
   IF Lower( hb_defaultValue( cSign, "" ) ) == "sign"
      IF PDF_Sign( "report.pdf", "report_signed.pdf", ;
                   { "page" => 0, "row" => -6, "col" => 45, "width" => 85, "height" => 18, ;
                     "reason" => "Occupational health report", "location" => "Marília - SP", ;
                     "text" => { "Assinado digitalmente por {name}", "Data: {date}", "Motivo: {reason}" } }, @cErr )
         ? "Created report_signed.pdf"
      ELSE
         ? "Not signed:", cErr
      ENDIF
   ENDIF
ELSE
   ? oDoc:cError
ENDIF
RETURN
