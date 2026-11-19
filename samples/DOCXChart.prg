// Charts in a .docx: column, bar, line, area, pie and doughnut charts, with
// the data in the chart and in an embedded workbook (Word: right click - Edit Data).
// Build: hbmk2 docxchart.hbp   -> charts.docx
#include "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main()
LOCAL oDoc, oChart, oTable, i, nHead, nWhite
LOCAL aMonths := { "Jan", "Feb", "Mar", "Apr", "May", "Jun" }
LOCAL aSales  := { { "Beograd",  { 120, 150,  90, 170, 210, 190 } }, ;
                   { "Novi Sad", {  80,  95, 110, 105, 130, 160 } }, ;
                   { "Niš",      {  60,  70,  65,  90,  85, 120 } } }

hb_cdpSelect( "UTF8EX" )

oDoc := WordDocument():New( "charts.docx" )
oDoc:cTitle    := "DOCX Class charts"
oDoc:cLanguage := "sr-Latn-RS"
nWhite := oDoc:NewFont( , , .T., , , , "FFFFFF" )
nHead  := oDoc:NewStyle( nWhite, 2 )

oDoc:Title( "Charts" )
oDoc:Paragraph( "Every chart keeps its data in an embedded workbook: right click - Edit Data in Word." )

// the data as a table
oDoc:Heading( "Sales by month", 1 )
oTable := oDoc:Table( 4 )
oTable:AddRow( { "Month", aSales[ 1 ][ 1 ], aSales[ 2 ][ 1 ], aSales[ 3 ][ 1 ] }, nHead, "4472C4" )
FOR i := 1 TO Len( aMonths )
   oTable:AddRow( { aMonths[ i ], aSales[ 1 ][ 2 ][ i ], aSales[ 2 ][ 2 ][ i ], aSales[ 3 ][ 2 ][ i ] } )
NEXT
oDoc:Paragraph()

// Chart( cTitle, cType, aCategories, aSeries, nWidth, nHeight, nAlign )
// cType: "bar", "hbar", "line", "area", "pie", "doughnut"
oChart := oDoc:Chart( "Sales 2026", "bar", aMonths, aSales, , , 2 )
oChart:cValTitle := "thousand RSD"

oDoc:Heading( "Line chart with values", 1 )
oChart := oDoc:Chart( "Beograd and Niš", "line", aMonths, { aSales[ 1 ], aSales[ 3 ] }, , , 2 )
oChart:lValues := .T.

oDoc:Heading( "Horizontal bars, own colors", 1 )
oChart := oDoc:Chart( "First half of the year", "hbar", { "Beograd", "Novi Sad", "Niš" }, ;
                      { { "Total", { Sum( aSales[ 1 ][ 2 ] ), Sum( aSales[ 2 ][ 2 ] ), Sum( aSales[ 3 ][ 2 ] ) }, "70AD47" } }, 14, 6, 2 )
oChart:lLegend := .F.
oChart:lValues := .T.
oChart:cNumFormat := "#,##0"

oDoc:PageBreak()
oDoc:Heading( "Pie and doughnut", 1 )
oDoc:Chart( "Share of sales", "pie", { "Beograd", "Novi Sad", "Niš" }, ;
            { Sum( aSales[ 1 ][ 2 ] ), Sum( aSales[ 2 ][ 2 ] ), Sum( aSales[ 3 ][ 2 ] ) }, 12, 7, 2 )
oChart := oDoc:Chart( "Costs", "doughnut", { "Salaries", "Rent", "Material", "Other" }, { 55, 20, 15, 10 }, 12, 7, 2 )
oChart:aColors := { "264478", "4472C4", "8FAADC", "D9D9D9" }

oDoc:Heading( "Area chart", 1 )
oChart := oDoc:Chart( "Novi Sad", "area", aMonths, { aSales[ 2 ] }, , , 2 )
oChart:cCatTitle := "month"

// a chart in a table cell
oDoc:Heading( "Charts in a table", 1 )
oTable := oDoc:Table( 2 )
oTable:Cell( 1, 1 ):Chart( "Beograd", "bar", aMonths, { aSales[ 1 ] }, 7.5, 5 ):lLegend := .F.
oTable:Cell( 1, 2 ):Chart( "Novi Sad", "line", aMonths, { aSales[ 2 ] }, 7.5, 5 ):lLegend := .F.

IF oDoc:Save()
   ? "Created", oDoc:cFilePath + oDoc:cName
ELSE
   ? "Error:", oDoc:cError
ENDIF
RETURN

STATIC FUNCTION Sum( a )
LOCAL n := 0
AEval( a, {| x | n += x } )
RETURN n
