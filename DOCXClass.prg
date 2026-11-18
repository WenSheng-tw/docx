/*
 *
 * DOCX writer/reader Class - Pure Harbour
 *
 * Copyright 2026 Srđan Dragojlović <digikv@yahoo.com>
 * Company  S.O.K. doo Kraljevo
 * www - https://sokdoo.com
 *
 * Dual licensed: GNU GPL version 3 or later (COPYING), or a commercial
 * license from the author - see LICENSE.
*/

#require "hbmzip"
#include "hbclass.ch"

REQUEST HB_CODEPAGE_UTF8EX

// 1 cm = 567 twip (1/20 pt), 1 cm = 360000 EMU
#define CM2TW( n )   Round( ( n ) * 1440 / 2.54, 0 )
#define CM2EMU( n )  Round( ( n ) * 360000, 0 )

#define XML_HEAD   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + Chr( 13 ) + Chr( 10 )
#define NS_W       "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
#define NS_R       "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
#define NS_WP      "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
#define NS_A       "http://schemas.openxmlformats.org/drawingml/2006/main"
#define NS_PIC     "http://schemas.openxmlformats.org/drawingml/2006/picture"
#define NS_PKG     "http://schemas.openxmlformats.org/package/2006/relationships"
#define NS_CT      "http://schemas.openxmlformats.org/package/2006/content-types"
#define REL_OFFICE "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
#define CT_WML     "application/vnd.openxmlformats-officedocument.wordprocessingml."

STATIC s_aAlign  := { "left", "center", "right", "both" }
STATIC s_aVAlign := { "top", "center", "bottom" }
STATIC s_aBorder := { "single", "double", "dotted", "dashed", "thick" }
STATIC s_aListTypes := { "bullet", "decimal", "lowerLetter", "upperLetter", "lowerRoman", "upperRoman" }

// ---------------------------------------------------------------------------
// Sadrzaj: niz pasusa i tabela. Isti kod sluzi za telo dokumenta (WordDocument
// ga nasledjuje), zaglavlje, podnozje i celije tabele.
// XML se pravi tek u Save(), jer slike i linkovi traze relacije dela (part)
// u kome se nalaze, a sirina teksta zavisi od podesavanja stranice.
// ---------------------------------------------------------------------------
CLASS DocxContent
   DATA oDoc
   DATA aItems
   METHOD New( oDoc )
   METHOD Paragraph( xText, nStyle )
   METHOD Heading( xText, nLevel, nStyle )
   METHOD Title( xText, nStyle )
   METHOD ListItem( xText, nList, nLevel, nStyle )
   METHOD Table( xCols )
   METHOD Picture( cFile, nWidth, nHeight, nAlign )
   METHOD PageBreak()
   METHOD TableOfContents( cTitle, nLevels )
   METHOD XML( hCtx, lCell )
ENDCLASS

METHOD New( oDoc ) CLASS DocxContent
::oDoc := oDoc
::aItems := {}
RETURN Self

METHOD Paragraph( xText, nStyle ) CLASS DocxContent
LOCAL oP := DocxParagraph():New( ::oDoc, nStyle )
IF ! HB_ISNIL( xText )
   oP:Text( xText )
ENDIF
AAdd( ::aItems, oP )
RETURN oP

// naslov 1-6 (ugradjeni stilovi "heading N": navigacija i sadrzaj u Wordu)
METHOD Heading( xText, nLevel, nStyle ) CLASS DocxContent
LOCAL oP := ::Paragraph( xText, nStyle )
nLevel := IIF( HB_ISNUMERIC( nLevel ), Max( 1, Min( 6, Int( nLevel ) ) ), 1 )
oP:cStyleName := "heading " + hb_NToS( nLevel )
RETURN oP

METHOD Title( xText, nStyle ) CLASS DocxContent
LOCAL oP := ::Paragraph( xText, nStyle )
oP:cStyleName := "Title"
RETURN oP

// stavka liste; nList iz oDoc:NewList(), bez njega poslednja napravljena lista
// (ili nova lista sa tackama), nLevel 1-9
METHOD ListItem( xText, nList, nLevel, nStyle ) CLASS DocxContent
LOCAL oP
IF ! HB_ISNUMERIC( nList ) .OR. nList < 1 .OR. nList > Len( ::oDoc:aLists )
   nList := IIF( Empty( ::oDoc:aLists ), ::oDoc:NewList(), Len( ::oDoc:aLists ) )
ENDIF
oP := ::Paragraph( xText, nStyle )
oP:cStyleName := "List Paragraph"
oP:nList  := nList
oP:nLevel := IIF( HB_ISNUMERIC( nLevel ), Max( 1, Min( 9, Int( nLevel ) ) ), 1 ) - 1
RETURN oP

// xCols: broj kolona (jednake sirine) ili niz sirina kolona u cm
METHOD Table( xCols ) CLASS DocxContent
LOCAL oT := DocxTable():New( ::oDoc, xCols )
AAdd( ::aItems, oT )
RETURN oT

// slika u svom pasusu; nWidth / nHeight u cm, nAlign 1 levo, 2 centar, 3 desno
METHOD Picture( cFile, nWidth, nHeight, nAlign ) CLASS DocxContent
LOCAL oP := ::Paragraph()
oP:nAlign := nAlign
oP:Picture( cFile, nWidth, nHeight )
RETURN oP

METHOD PageBreak() CLASS DocxContent
LOCAL oP := ::Paragraph()
AAdd( oP:aRuns, { "br", "page" } )
RETURN oP

// Sadrzaj (TOC polje). Word ga popunjava pri otvaranju (pita za azuriranje
// polja), a rucno: desni klik - Update Field.
METHOD TableOfContents( cTitle, nLevels ) CLASS DocxContent
LOCAL oP
IF HB_ISSTRING( cTitle ) .AND. ! Empty( cTitle )
   oP := ::Paragraph( cTitle )
   oP:cStyleName := "TOC Heading"
ENDIF
nLevels := IIF( HB_ISNUMERIC( nLevels ), Max( 1, Min( 9, Int( nLevels ) ) ), 3 )
oP := ::Paragraph()
AAdd( oP:aRuns, { "toc", 'TOC \o "1-' + hb_NToS( nLevels ) + '" \h \z \u', ;
                  "Right-click and choose Update Field to show the table of contents." } )
::oDoc:lUpdateFields := .T.
RETURN oP

// lCell: celija tabele (i zaglavlje/podnozje) mora da ima bar jedan pasus i da
// se zavrsava pasusom, inace Word prijavi ostecen fajl
METHOD XML( hCtx, lCell ) CLASS DocxContent
LOCAL aOut := {}, o
FOR EACH o IN ::aItems
   AAdd( aOut, o:XML( hCtx ) )
NEXT
IF HB_ISLOGICAL( lCell ) .AND. lCell .AND. ( Empty( ::aItems ) .OR. ATail( ::aItems ):ClassName() == "DOCXTABLE" )
   AAdd( aOut, "<w:p/>" )
ENDIF
RETURN _Join( aOut, "" )

// ---------------------------------------------------------------------------
// Pasus: niz delova teksta (run) sa svojim fontom, linkovi, polja, slike.
// Metode za dodavanje vracaju Self, pa se mogu nizati:
//    oDoc:Paragraph( "Strana " ):PageNumber():Text( " od " ):PageCount()
// ---------------------------------------------------------------------------
CLASS DocxParagraph
   DATA oDoc
   DATA nStyle INIT 0       // NewStyle()
   DATA cStyleName          // ugradjeni stil: "heading 1", "Title", "List Paragraph"...
   DATA nAlign              // 1 levo, 2 centar, 3 desno, 4 obostrano (menja poravnanje stila)
   DATA nList               // lista (NewList) i nivo za ListItem()
   DATA nLevel INIT 0
   DATA lHFTabs INIT .F.    // AddHeader()/AddFooter(): tab na sredini i desno
   DATA aRuns
   METHOD New( oDoc, nStyle )
   METHOD Text( xText, nFont )
   METHOD Link( xText, cURL, nFont )
   METHOD Field( cInstr, nFont, cShow )
   METHOD PageNumber( nFont ) INLINE ::Field( "PAGE", nFont, "1" )
   METHOD PageCount( nFont ) INLINE ::Field( "NUMPAGES", nFont, "1" )
   METHOD Picture( cFile, nWidth, nHeight )
   METHOD LineBreak()
   METHOD XML( hCtx )
ENDCLASS

METHOD New( oDoc, nStyle ) CLASS DocxParagraph
::oDoc := oDoc
IF HB_ISNUMERIC( nStyle ) .AND. nStyle >= 1 .AND. nStyle <= Len( oDoc:aStyles )
   ::nStyle := Int( nStyle )
ENDIF
::aRuns := {}
RETURN Self

// tekst; Chr(10) (i CRLF) = novi red u istom pasusu, Chr(9) = tab
METHOD Text( xText, nFont ) CLASS DocxParagraph
AAdd( ::aRuns, { "t", ::oDoc:ToUTF8( _Str( xText ) ), nFont } )
RETURN Self

METHOD Link( xText, cURL, nFont ) CLASS DocxParagraph
IF HB_ISSTRING( cURL ) .AND. ! Empty( cURL )
   AAdd( ::aRuns, { "link", ::oDoc:ToUTF8( _Str( xText ) ), ::oDoc:ToUTF8( AllTrim( cURL ) ), nFont } )
ELSE
   ::Text( xText, nFont )
ENDIF
RETURN Self

// polje Worda: "PAGE", "NUMPAGES", 'DATE \@ "dd.MM.yyyy"', "FILENAME"...
// cShow = tekst koji se vidi dok Word ne azurira polje
METHOD Field( cInstr, nFont, cShow ) CLASS DocxParagraph
IF HB_ISSTRING( cInstr ) .AND. ! Empty( cInstr )
   AAdd( ::aRuns, { "fld", ::oDoc:ToUTF8( cInstr ), nFont, IIF( HB_ISSTRING( cShow ), ::oDoc:ToUTF8( cShow ), "" ) } )
ENDIF
RETURN Self

// slika u redu teksta; nWidth / nHeight u cm (jedna zadata = druga proporcionalno,
// nijedna = prirodna velicina pri 96 dpi, najvise sirina teksta)
METHOD Picture( cFile, nWidth, nHeight ) CLASS DocxParagraph
IF HB_ISSTRING( cFile ) .AND. hb_FileExists( cFile )
   AAdd( ::aRuns, { "pic", cFile, nWidth, nHeight } )
ELSE
   ::oDoc:cError := "Slika ne postoji: " + hb_CStr( cFile )
ENDIF
RETURN Self

METHOD LineBreak() CLASS DocxParagraph
AAdd( ::aRuns, { "br", "" } )
RETURN Self

METHOD XML( hCtx ) CLASS DocxParagraph
LOCAL hSt, cP := "", cDefR := "", cSpacing := "", a, aOut := {}, nAlign, nTW
IF ::nStyle > 0
   hSt := ::oDoc:aStyles[ ::nStyle ]
   cDefR := _FontRPr( ::oDoc, hSt[ "font" ] )
   cSpacing := hSt[ "spacing" ]
ENDIF
// celije tabele, zaglavlje i podnozje: bez razmaka posle pasusa (kao Wordov "Table Grid")
IF Empty( cSpacing ) .AND. Empty( ::cStyleName ) .AND. hCtx[ "compact" ]
   cSpacing := '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
ENDIF
// redosled elemenata u pPr propisuje sema (Word odbija drugaciji)
IF ! Empty( ::cStyleName )
   cP += '<w:pStyle w:val="' + _StyleId( hCtx, ::cStyleName ) + '"/>'
ENDIF
IF hSt != NIL
   cP += hSt[ "keep" ]
ENDIF
IF HB_ISNUMERIC( ::nList )
   cP += '<w:numPr><w:ilvl w:val="' + hb_NToS( ::nLevel ) + '"/><w:numId w:val="' + hb_NToS( hCtx[ "numBase" ] + ::nList ) + '"/></w:numPr>'
ENDIF
IF hSt != NIL
   cP += hSt[ "shd" ]
ENDIF
IF ::lHFTabs
   nTW := CM2TW( hCtx[ "textWidth" ] )
   cP += '<w:tabs><w:tab w:val="center" w:pos="' + hb_NToS( Int( nTW / 2 ) ) + '"/><w:tab w:val="right" w:pos="' + hb_NToS( nTW ) + '"/></w:tabs>'
ENDIF
cP += cSpacing
IF hSt != NIL
   cP += hSt[ "ind" ]
ENDIF
nAlign := IIF( HB_ISNUMERIC( ::nAlign ) .AND. ::nAlign >= 1 .AND. ::nAlign <= 4, Int( ::nAlign ), IIF( hSt != NIL, hSt[ "align" ], 0 ) )
IF nAlign > 0
   cP += '<w:jc w:val="' + s_aAlign[ nAlign ] + '"/>'
ENDIF
IF ! Empty( cDefR )
   cP += "<w:rPr>" + cDefR + "</w:rPr>"
ENDIF
AAdd( aOut, "<w:p>" + IIF( Empty( cP ), "", "<w:pPr>" + cP + "</w:pPr>" ) )
FOR EACH a IN ::aRuns
   DO CASE
   CASE a[ 1 ] == "t"
      AAdd( aOut, _RunXML( a[ 2 ], _RunRPr( ::oDoc, a[ 3 ], cDefR ) ) )
   CASE a[ 1 ] == "link"
      AAdd( aOut, '<w:hyperlink xmlns:r="' + NS_R + '" r:id="' + _AddRel( hCtx, "hyperlink", a[ 3 ], .T. ) + '" w:history="1">' + ;
                  _RunXML( a[ 2 ], '<w:rStyle w:val="' + _StyleId( hCtx, "Hyperlink" ) + '"/>' + _RunRPr( ::oDoc, a[ 4 ], cDefR ) ) + ;
                  "</w:hyperlink>" )
   CASE a[ 1 ] == "fld"
      AAdd( aOut, '<w:fldSimple w:instr=" ' + XmlEsc( a[ 2 ] ) + ' ">' + _RunXML( a[ 4 ], _RunRPr( ::oDoc, a[ 3 ], cDefR ) ) + "</w:fldSimple>" )
   CASE a[ 1 ] == "toc"
      AAdd( aOut, '<w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>' + ;
                  '<w:r><w:instrText xml:space="preserve"> ' + XmlEsc( a[ 2 ] ) + ' </w:instrText></w:r>' + ;
                  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + _RunXML( a[ 3 ], "" ) + ;
                  '<w:r><w:fldChar w:fldCharType="end"/></w:r>' )
   CASE a[ 1 ] == "pic"
      AAdd( aOut, _PictureXML( hCtx, a[ 2 ], a[ 3 ], a[ 4 ] ) )
   CASE a[ 1 ] == "br"
      AAdd( aOut, IIF( a[ 2 ] == "page", '<w:r><w:br w:type="page"/></w:r>', "<w:r><w:br/></w:r>" ) )
   ENDCASE
NEXT
AAdd( aOut, "</w:p>" )
RETURN _Join( aOut, "" )

// ---------------------------------------------------------------------------
// Tabela. Celija je DocxContent, pa moze da sadrzi vise pasusa, slike i
// ugnezdene tabele: oTable:Cell( 2, 1 ):Picture( "logo.png", 3 )
// ---------------------------------------------------------------------------
CLASS DocxTable
   DATA oDoc
   DATA aWidths             // sirine kolona u cm (NIL = deli ostatak sirine teksta)
   DATA aCells              // aCells[ nRow, nCol ] := DocxContent ili NIL
   DATA aHeights
   DATA hSpan PROTECTED     // "r,c" -> { nColSpan (0 = pokrivena celija), cVMerge }
   DATA hShade PROTECTED    // "r,c" / "r,0" -> boja pozadine
   DATA hVAlign PROTECTED   // "r,c" / "r,0" / "0,0" -> vertikalno poravnanje
   DATA nHeaderRows INIT 0  // prvih n redova se ponavlja na vrhu svake strane
   DATA nBorder INIT 1      // 0 bez okvira, 1 single, 2 double, 3 dotted, 4 dashed, 5 thick
   DATA nBorderSize INIT 0.5   // debljina linije u pt
   DATA cBorderColor INIT "000000"
   DATA nAlign INIT 1       // tabela na strani: 1 levo, 2 centar, 3 desno
   DATA nCellMargin INIT 0.19  // levi i desni razmak u celiji, cm
   DATA aColStyles          // stil (NewStyle) po koloni, za celije bez svog stila
   DATA aPictures           // Transform() slika po koloni za brojeve i datume
   METHOD New( oDoc, xCols )
   METHOD AddRow( aValues, nStyle, cShading )
   METHOD Cell( nRow, nCol, xValue, nStyle )
   METHOD Merge( nRow1, nCol1, nRow2, nCol2 )
   METHOD Shade( nRow, nCol, cColor )
   METHOD VAlign( nRow, nCol, nAlign )
   METHOD RowHeight( nRow, nHeight )
   METHOD RowCount() INLINE Len( ::aCells )
   METHOD ColCount() INLINE Len( ::aWidths )
   METHOD XML( hCtx )
   METHOD EnsureRow( nRow ) PROTECTED
ENDCLASS

METHOD New( oDoc, xCols ) CLASS DocxTable
::oDoc := oDoc
IF HB_ISARRAY( xCols ) .AND. ! Empty( xCols )
   ::aWidths := AClone( xCols )
ELSE
   ::aWidths := Array( IIF( HB_ISNUMERIC( xCols ) .AND. xCols >= 1, Int( xCols ), 1 ) )
ENDIF
::aCells     := {}
::aHeights   := {}
::hSpan      := { => }
::hShade     := { => }
::hVAlign    := { => }
::aColStyles := Array( Len( ::aWidths ) )
::aPictures  := Array( Len( ::aWidths ) )
RETURN Self

METHOD EnsureRow( nRow ) CLASS DocxTable
DO WHILE Len( ::aCells ) < nRow
   AAdd( ::aCells, Array( Len( ::aWidths ) ) )
   AAdd( ::aHeights, 0 )
ENDDO
RETURN NIL

// novi red sa vrednostima (NIL = prazna celija); vraca broj reda
METHOD AddRow( aValues, nStyle, cShading ) CLASS DocxTable
LOCAL nRow := Len( ::aCells ) + 1, i
::EnsureRow( nRow )
IF HB_ISARRAY( aValues )
   FOR i := 1 TO Min( Len( aValues ), Len( ::aWidths ) )
      IF ! HB_ISNIL( aValues[ i ] )
         ::Cell( nRow, i, aValues[ i ], nStyle )
      ENDIF
   NEXT
ENDIF
IF HB_ISSTRING( cShading ) .AND. ! Empty( cShading )
   ::Shade( nRow, , cShading )
ENDIF
RETURN nRow

// Sa xValue: upisuje vrednost (zamenjuje raniji sadrzaj celije). Broj bez stila
// se poravnava desno, kao u Excelu. Vraca sadrzaj celije (DocxContent) za
// dodavanje pasusa, slika... NIL ako je adresa van tabele.
METHOD Cell( nRow, nCol, xValue, nStyle ) CLASS DocxTable
LOCAL oC, oP, cPic
IF ! HB_ISNUMERIC( nRow ) .OR. ! HB_ISNUMERIC( nCol ) .OR. nRow < 1 .OR. nCol < 1 .OR. nCol > Len( ::aWidths )
   RETURN NIL
ENDIF
::EnsureRow( nRow )
IF HB_ISNIL( xValue ) .AND. ::aCells[ nRow, nCol ] != NIL
   RETURN ::aCells[ nRow, nCol ]
ENDIF
oC := DocxContent():New( ::oDoc )
::aCells[ nRow, nCol ] := oC
IF ! HB_ISNIL( xValue )
   IF ! HB_ISNUMERIC( nStyle ) .OR. nStyle < 1
      nStyle := _AGet( ::aColStyles, nCol )
   ENDIF
   cPic := _AGet( ::aPictures, nCol )
   IF HB_ISSTRING( cPic ) .AND. ! Empty( cPic ) .AND. ! HB_ISSTRING( xValue )
      xValue := AllTrim( Transform( xValue, cPic ) )
      oP := oC:Paragraph( xValue, nStyle )
      IF ! HB_ISNUMERIC( nStyle )
         oP:nAlign := 3
      ENDIF
   ELSE
      oP := oC:Paragraph( xValue, nStyle )
      IF HB_ISNUMERIC( xValue ) .AND. ! HB_ISNUMERIC( nStyle )
         oP:nAlign := 3
      ENDIF
   ENDIF
ENDIF
RETURN oC

// spaja pravougaonik celija; sadrzaj ima gornja leva celija
METHOD Merge( nRow1, nCol1, nRow2, nCol2 ) CLASS DocxTable
LOCAL r, c
IF ! HB_ISNUMERIC( nRow1 ) .OR. ! HB_ISNUMERIC( nCol1 )
   RETURN Self
ENDIF
nRow2 := IIF( HB_ISNUMERIC( nRow2 ), nRow2, nRow1 )
nCol2 := IIF( HB_ISNUMERIC( nCol2 ), nCol2, nCol1 )
IF nRow1 < 1 .OR. nCol1 < 1 .OR. nRow2 < nRow1 .OR. nCol2 < nCol1 .OR. nCol2 > Len( ::aWidths )
   RETURN Self
ENDIF
::EnsureRow( nRow2 )
FOR r := nRow1 TO nRow2
   ::hSpan[ _RC( r, nCol1 ) ] := { nCol2 - nCol1 + 1, IIF( nRow2 > nRow1, IIF( r == nRow1, "restart", "continue" ), "" ) }
   FOR c := nCol1 + 1 TO nCol2
      ::hSpan[ _RC( r, c ) ] := { 0, "" }
   NEXT
NEXT
RETURN Self

// boja pozadine celije, ili celog reda kad nCol nije zadat
METHOD Shade( nRow, nCol, cColor ) CLASS DocxTable
IF HB_ISNUMERIC( nRow ) .AND. HB_ISSTRING( cColor )
   ::hShade[ _RC( nRow, IIF( HB_ISNUMERIC( nCol ), nCol, 0 ) ) ] := _RGB( cColor )
ENDIF
RETURN Self

// vertikalno poravnanje 1 gore, 2 sredina, 3 dole: celija, ceo red (nCol NIL)
// ili cela tabela (nRow NIL)
METHOD VAlign( nRow, nCol, nAlign ) CLASS DocxTable
IF HB_ISNUMERIC( nAlign )
   ::hVAlign[ _RC( IIF( HB_ISNUMERIC( nRow ), nRow, 0 ), IIF( HB_ISNUMERIC( nCol ), nCol, 0 ) ) ] := nAlign
ENDIF
RETURN Self

// najmanja visina reda u cm
METHOD RowHeight( nRow, nHeight ) CLASS DocxTable
IF HB_ISNUMERIC( nRow ) .AND. nRow >= 1 .AND. HB_ISNUMERIC( nHeight )
   ::EnsureRow( nRow )
   ::aHeights[ nRow ] := nHeight
ENDIF
RETURN Self

METHOD XML( hCtx ) CLASS DocxTable
LOCAL aOut := {}, nCols := Len( ::aWidths ), aTw := Array( nCols ), nTotal := 0, nFree := 0, nRest
LOCAL i, k, r, x, nSpan, cV, cShd, nVA, nW, oC, cB, cTr, lCompact, nTextWidth := hCtx[ "textWidth" ]
nRest := nTextWidth
FOR i := 1 TO nCols
   IF HB_ISNUMERIC( ::aWidths[ i ] ) .AND. ::aWidths[ i ] > 0
      nRest -= ::aWidths[ i ]
   ELSE
      nFree++
   ENDIF
NEXT
FOR i := 1 TO nCols
   aTw[ i ] := CM2TW( IIF( HB_ISNUMERIC( ::aWidths[ i ] ) .AND. ::aWidths[ i ] > 0, ::aWidths[ i ], Max( 1, nRest / Max( 1, nFree ) ) ) )
   nTotal += aTw[ i ]
NEXT
cB := ' w:val="' + IIF( ::nBorder >= 1 .AND. ::nBorder <= Len( s_aBorder ), s_aBorder[ ::nBorder ], "none" ) + ;
      '" w:sz="' + hb_NToS( Max( 2, Round( ::nBorderSize * 8, 0 ) ) ) + '" w:space="0" w:color="' + _RGB( ::cBorderColor ) + '"/>'
AAdd( aOut, '<w:tbl><w:tblPr><w:tblW w:w="' + hb_NToS( nTotal ) + '" w:type="dxa"/>' + ;
            IIF( ::nAlign == 2, '<w:jc w:val="center"/>', IIF( ::nAlign == 3, '<w:jc w:val="right"/>', "" ) ) + ;
            "<w:tblBorders><w:top" + cB + "<w:left" + cB + "<w:bottom" + cB + "<w:right" + cB + ;
            "<w:insideH" + cB + "<w:insideV" + cB + "</w:tblBorders>" + ;
            '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="' + hb_NToS( CM2TW( ::nCellMargin ) ) + '" w:type="dxa"/>' + ;
            '<w:right w:w="' + hb_NToS( CM2TW( ::nCellMargin ) ) + '" w:type="dxa"/></w:tblCellMar>' + ;
            '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' + ;
            "</w:tblPr><w:tblGrid>" )
FOR i := 1 TO nCols
   AAdd( aOut, '<w:gridCol w:w="' + hb_NToS( aTw[ i ] ) + '"/>' )
NEXT
AAdd( aOut, "</w:tblGrid>" )
lCompact := hCtx[ "compact" ]
hCtx[ "compact" ] := .T.
FOR r := 1 TO Len( ::aCells )
   cTr := IIF( ::aHeights[ r ] > 0, '<w:trHeight w:val="' + hb_NToS( CM2TW( ::aHeights[ r ] ) ) + '" w:hRule="atLeast"/>', "" ) + ;
          IIF( r <= ::nHeaderRows, "<w:tblHeader/>", "" )
   AAdd( aOut, "<w:tr>" + IIF( Empty( cTr ), "", "<w:trPr>" + cTr + "</w:trPr>" ) )
   FOR i := 1 TO nCols
      nSpan := 1
      cV := ""
      IF ( x := hb_HGetDef( ::hSpan, _RC( r, i ), NIL ) ) != NIL
         IF x[ 1 ] == 0
            LOOP
         ENDIF
         nSpan := Min( x[ 1 ], nCols - i + 1 )
         cV := x[ 2 ]
      ENDIF
      nW := 0
      FOR k := i TO i + nSpan - 1
         nW += aTw[ k ]
      NEXT
      cShd := hb_HGetDef( ::hShade, _RC( r, i ), hb_HGetDef( ::hShade, _RC( r, 0 ), "" ) )
      nVA  := hb_HGetDef( ::hVAlign, _RC( r, i ), hb_HGetDef( ::hVAlign, _RC( r, 0 ), hb_HGetDef( ::hVAlign, "0,0", 0 ) ) )
      oC := ::aCells[ r, i ]
      // u celiji je "sirina teksta" sirina celije (ugnezdene tabele, slike)
      hCtx[ "textWidth" ] := Max( 0.5, nW / ( 1440 / 2.54 ) - 2 * ::nCellMargin )
      AAdd( aOut, '<w:tc><w:tcPr><w:tcW w:w="' + hb_NToS( nW ) + '" w:type="dxa"/>' + ;
                  IIF( nSpan > 1, '<w:gridSpan w:val="' + hb_NToS( nSpan ) + '"/>', "" ) + ;
                  IIF( cV == "restart", '<w:vMerge w:val="restart"/>', IIF( cV == "continue", "<w:vMerge/>", "" ) ) + ;
                  IIF( Empty( cShd ), "", '<w:shd w:val="clear" w:color="auto" w:fill="' + cShd + '"/>' ) + ;
                  IIF( nVA >= 1 .AND. nVA <= 3, '<w:vAlign w:val="' + s_aVAlign[ nVA ] + '"/>', "" ) + "</w:tcPr>" + ;
                  IIF( cV == "continue" .OR. oC == NIL, "<w:p/>", oC:XML( hCtx, .T. ) ) + "</w:tc>" )
      hCtx[ "textWidth" ] := nTextWidth
   NEXT
   AAdd( aOut, "</w:tr>" )
NEXT
hCtx[ "compact" ] := lCompact
AAdd( aOut, "</w:tbl>" )
RETURN _Join( aOut, "" )

// ---------------------------------------------------------------------------
// Dokument
// ---------------------------------------------------------------------------
CLASS WordDocument INHERIT DocxContent
   DATA cName
   DATA cFilePath
   DATA cError INIT ""
   DATA cCodePage                  // kodna strana teksta koji nije UTF-8 (podraz. hb_cdpSelect())
   DATA cTitle INIT ""             // svojstva dokumenta (File - Info)
   DATA cSubject INIT ""
   DATA cAuthor INIT ""
   DATA cCompany INIT ""
   DATA Application INIT "DOCX Class"
   DATA cFont INIT "Calibri"       // osnovni font novog dokumenta
   DATA nFontSize INIT 11
   DATA cLanguage INIT "en-US"     // jezik za proveru pravopisa, npr. "sr-Latn-RS"
   DATA nSpaceAfter INIT 8         // razmak posle pasusa u pt
   DATA nLineSpacing INIT 1.08     // prored
   DATA nPaperSize INIT 9          // 9 A4, 1 Letter, 5 Legal, 8 A3, 11 A5
   DATA lLandscape INIT .F.
   DATA nLeftMargin INIT 2.5       // margine u cm
   DATA nRightMargin INIT 2.5
   DATA nTopMargin INIT 2.5
   DATA nBottomMargin INIT 2.5
   DATA nHeaderMargin INIT 1.25
   DATA nFooterMargin INIT 1.25
   DATA lUpdateFields INIT .F.     // Word azurira polja (sadrzaj) pri otvaranju
   DATA cInsertAt                  // Read(): novi sadrzaj ide na mesto pasusa sa ovim tekstom
   DATA oHeader
   DATA oFooter
   DATA aFonts
   DATA aStyles
   DATA aLists
   DATA aStyleKeys PROTECTED
   DATA hParts PROTECTED           // delovi ucitanog fajla: ime -> sadrzaj
   DATA cMainPart INIT "word/document.xml" PROTECTED
   DATA lRead INIT .F. PROTECTED
   METHOD New( cName )
   METHOD NewFont( cFont, nSize, lBold, lItalic, lUnderline, lStrike, cColor, cBackColor )
   METHOD NewStyle( nFont, nAlign, nSpaceBefore, nSpaceAfter, nLineSpacing, nIndent, nFirstLine, cShading, lKeepNext, lPageBreakBefore )
   METHOD NewList( cType )
   METHOD Header()
   METHOD Footer()
   METHOD AddHeader( cLeft, cCenter, cRight, nFont )
   METHOD AddFooter( cLeft, cCenter, cRight, nFont )
   METHOD InsertAt( cMarker )
   METHOD Read( cName )
   METHOD Paragraphs( lStyles )
   METHOD Tables()
   METHOD Text()
   METHOD Replace( xFind, xReplace )
   METHOD RepeatRow( cMarker, aRows, aKeys )
   METHOD Save( cName )
   METHOD ToUTF8( c )
   METHOD SetName( cName ) PROTECTED
   METHOD BaseParts() PROTECTED
   METHOD Body() PROTECTED
   METHOD TextWidth( hOut ) PROTECTED
ENDCLASS

METHOD New( cName ) CLASS WordDocument
::oDoc       := Self
::aItems     := {}
::aFonts     := {}
::aStyles    := {}
::aStyleKeys := {}
::aLists     := {}
::hParts     := { => }
::cCodePage  := hb_cdpSelect()
::SetName( cName )
RETURN Self

// putanja i ime posebno, kao u XLSX klasi: oDoc:cName := "novo.docx" snima u isti folder
METHOD SetName( cName ) CLASS WordDocument
IF ! HB_ISSTRING( cName )
   cName := ""
ENDIF
::cFilePath := hb_FNameDir( cName )
IF Empty( ::cFilePath )
   ::cFilePath := hb_cwd()
ENDIF
::cName := hb_FNameNameExt( cName )
RETURN NIL

// Tekst koji nije ispravan UTF-8 smatra se tekstom u kodnoj strani ::cCodePage
// (podrazumevano kodna strana aplikacije u trenutku New()) i konvertuje se.
METHOD ToUTF8( c ) CLASS WordDocument
IF HB_ISSTRING( c ) .AND. HB_ISSTRING( ::cCodePage ) .AND. ! Empty( ::cCodePage ) .AND. ;
   !( "UTF8" $ Upper( ::cCodePage ) ) .AND. ! hb_StrIsUTF8( c )
   c := hb_StrToUTF8( c, ::cCodePage )
ENDIF
RETURN c

// Font za deo teksta (oPara:Text( c, nFont )) ili za stil (NewStyle).
// NIL = nasledjuje se; lBold := .F. iskljucuje podebljanje (npr. u naslovu).
// Boje RGB "FF0000" (prihvata i ARGB "FFFF0000" kao u XLSX klasi).
METHOD NewFont( cFont, nSize, lBold, lItalic, lUnderline, lStrike, cColor, cBackColor ) CLASS WordDocument
LOCAL c := "", n
IF HB_ISSTRING( cFont ) .AND. ! Empty( cFont )
   cFont := XmlEsc( ::ToUTF8( cFont ) )
   c += '<w:rFonts w:ascii="' + cFont + '" w:hAnsi="' + cFont + '" w:eastAsia="' + cFont + '" w:cs="' + cFont + '"/>'
ENDIF
c += _OnOff( "b", lBold ) + _OnOff( "bCs", lBold ) + _OnOff( "i", lItalic ) + _OnOff( "iCs", lItalic ) + _OnOff( "strike", lStrike )
IF HB_ISSTRING( cColor ) .AND. ! Empty( cColor )
   c += '<w:color w:val="' + _RGB( cColor ) + '"/>'
ENDIF
IF HB_ISNUMERIC( nSize ) .AND. nSize > 0
   c += '<w:sz w:val="' + hb_NToS( Round( nSize * 2, 0 ) ) + '"/><w:szCs w:val="' + hb_NToS( Round( nSize * 2, 0 ) ) + '"/>'
ENDIF
IF HB_ISLOGICAL( lUnderline )
   c += '<w:u w:val="' + IIF( lUnderline, "single", "none" ) + '"/>'
ENDIF
IF HB_ISSTRING( cBackColor ) .AND. ! Empty( cBackColor )
   c += '<w:shd w:val="clear" w:color="auto" w:fill="' + _RGB( cBackColor ) + '"/>'
ENDIF
IF ( n := AScan( ::aFonts, {| x | x == c } ) ) == 0
   AAdd( ::aFonts, c )
   n := Len( ::aFonts )
ENDIF
RETURN n

// Stil pasusa: font, poravnanje (1 levo, 2 centar, 3 desno, 4 obostrano),
// razmak pre/posle u pt, prored (1, 1.15, 1.5, 2), uvlacenje levo i prvog reda
// u cm (negativno = viseci red), boja pozadine, drzi sa sledecim, nova strana.
METHOD NewStyle( nFont, nAlign, nSpaceBefore, nSpaceAfter, nLineSpacing, nIndent, nFirstLine, cShading, lKeepNext, lPageBreakBefore ) CLASS WordDocument
LOCAL h := { => }, c, n
h[ "font" ]  := IIF( HB_ISNUMERIC( nFont ) .AND. nFont >= 1 .AND. nFont <= Len( ::aFonts ), Int( nFont ), 0 )
h[ "align" ] := IIF( HB_ISNUMERIC( nAlign ) .AND. nAlign >= 1 .AND. nAlign <= 4, Int( nAlign ), 0 )
h[ "keep" ]  := IIF( HB_ISLOGICAL( lKeepNext ) .AND. lKeepNext, "<w:keepNext/>", "" ) + ;
                IIF( HB_ISLOGICAL( lPageBreakBefore ) .AND. lPageBreakBefore, "<w:pageBreakBefore/>", "" )
h[ "shd" ]   := IIF( HB_ISSTRING( cShading ) .AND. ! Empty( cShading ), '<w:shd w:val="clear" w:color="auto" w:fill="' + _RGB( cShading ) + '"/>', "" )
c := ""
IF HB_ISNUMERIC( nSpaceBefore )
   c += ' w:before="' + hb_NToS( Round( nSpaceBefore * 20, 0 ) ) + '"'
ENDIF
IF HB_ISNUMERIC( nSpaceAfter )
   c += ' w:after="' + hb_NToS( Round( nSpaceAfter * 20, 0 ) ) + '"'
ENDIF
IF HB_ISNUMERIC( nLineSpacing ) .AND. nLineSpacing > 0
   c += ' w:line="' + hb_NToS( Round( nLineSpacing * 240, 0 ) ) + '" w:lineRule="auto"'
ENDIF
h[ "spacing" ] := IIF( Empty( c ), "", "<w:spacing" + c + "/>" )
c := ""
IF HB_ISNUMERIC( nIndent )
   c += ' w:left="' + hb_NToS( CM2TW( nIndent ) ) + '"'
ENDIF
IF HB_ISNUMERIC( nFirstLine ) .AND. nFirstLine != 0
   c += IIF( nFirstLine > 0, ' w:firstLine="' + hb_NToS( CM2TW( nFirstLine ) ) + '"', ' w:hanging="' + hb_NToS( CM2TW( -nFirstLine ) ) + '"' )
ENDIF
h[ "ind" ] := IIF( Empty( c ), "", "<w:ind" + c + "/>" )
c := hb_NToS( h[ "font" ] ) + "|" + hb_NToS( h[ "align" ] ) + "|" + h[ "keep" ] + h[ "shd" ] + h[ "spacing" ] + h[ "ind" ]
IF ( n := AScan( ::aStyleKeys, {| x | x == c } ) ) == 0
   AAdd( ::aStyleKeys, c )
   AAdd( ::aStyles, h )
   n := Len( ::aStyles )
ENDIF
RETURN n

// nova lista (svaka pocinje od 1): "bullet", "decimal", "lowerLetter",
// "upperLetter", "lowerRoman", "upperRoman"
METHOD NewList( cType ) CLASS WordDocument
LOCAL n := 0
IF HB_ISSTRING( cType )
   n := AScan( s_aListTypes, {| x | Lower( x ) == Lower( AllTrim( cType ) ) } )
ENDIF
AAdd( ::aLists, s_aListTypes[ Max( 1, n ) ] )
RETURN Len( ::aLists )

METHOD Header() CLASS WordDocument
IF ::oHeader == NIL
   ::oHeader := DocxContent():New( Self )
ENDIF
RETURN ::oHeader

METHOD Footer() CLASS WordDocument
IF ::oFooter == NIL
   ::oFooter := DocxContent():New( Self )
ENDIF
RETURN ::oFooter

// Kao u Excelu: tekst levo, na sredini i desno, sa kodovima
// &P strana, &N broj strana, &D datum, &T vreme, &F ime fajla, && znak &
METHOD AddHeader( cLeft, cCenter, cRight, nFont ) CLASS WordDocument
RETURN _HFParagraph( ::Header():Paragraph(), cLeft, cCenter, cRight, nFont )

METHOD AddFooter( cLeft, cCenter, cRight, nFont ) CLASS WordDocument
RETURN _HFParagraph( ::Footer():Paragraph(), cLeft, cCenter, cRight, nFont )

// posle Read(): sadrzaj dodat metodama Paragraph(), Table()... upisuje se na
// mesto pasusa koji sadrzi cMarker (inace na kraj dokumenta)
METHOD InsertAt( cMarker ) CLASS WordDocument
::cInsertAt := IIF( HB_ISSTRING( cMarker ) .AND. ! Empty( cMarker ), ::ToUTF8( cMarker ), NIL )
RETURN Self

// Ucitava postojeci .docx. Vraca .T. ako je uspelo, inace .F. i opis greske u ::cError.
// Posle toga: Paragraphs(), Tables(), Text() za citanje, Replace() i RepeatRow()
// za popunjavanje sablona, Save() snima (i sav novi sadrzaj) pod ::cName.
METHOD Read( cName ) CLASS WordDocument
LOCAL hParts, cTok, cA, cTag, nKind, h
::cError := ""
IF ! HB_ISSTRING( cName ) .OR. ! hb_FileExists( cName )
   ::cError := "Fajl ne postoji: " + hb_CStr( cName )
   RETURN .F.
ENDIF
hParts := _ZipRead( cName )
IF ! hb_HHasKey( hParts, "[Content_Types].xml" )
   ::cError := "Nije ispravan .docx fajl: " + cName
   RETURN .F.
ENDIF
::cMainPart := "word/document.xml"
FOR EACH cTok IN hb_ATokens( hb_HGetDef( hParts, "_rels/.rels", "" ), "<" )
   cA := _XTok( cTok, @cTag, @nKind )
   IF cTag == "Relationship" .AND. nKind != 2
      h := _XAttrs( cA )
      IF Right( _HAttr( h, "Type" ), 15 ) == "/officeDocument"
         ::cMainPart := _PartPath( "", _HAttr( h, "Target" ) )
      ENDIF
   ENDIF
NEXT
IF ! hb_HHasKey( hParts, ::cMainPart )
   ::cError := "Nije ispravan .docx fajl: " + cName
   RETURN .F.
ENDIF
::hParts := hParts
::lRead := .T.
::SetName( cName )
RETURN .T.

METHOD Body() CLASS WordDocument
LOCAL cStyles
IF ! ::lRead
   RETURN {}
ENDIF
cStyles := _RelTarget( ::hParts, ::cMainPart, "/styles" )
RETURN _ParseBody( ::hParts[ ::cMainPart ], _StyleMap( hb_HGetDef( ::hParts, cStyles, "" ) )[ 2 ] )

// pasusi ucitanog dokumenta van tabela; lStyles: { { cTekst, cStil }, ... }
// (stil je ime iz Worda, npr. "heading 1", "Title", "" za obican pasus)
METHOD Paragraphs( lStyles ) CLASS WordDocument
LOCAL aOut := {}, a
FOR EACH a IN ::Body()
   IF a[ 1 ] == "p"
      AAdd( aOut, IIF( HB_ISLOGICAL( lStyles ) .AND. lStyles, { a[ 2 ], a[ 3 ] }, a[ 2 ] ) )
   ENDIF
NEXT
RETURN aOut

// tabele ucitanog dokumenta: niz tabela, tabela je niz redova, red niz teksta celija
METHOD Tables() CLASS WordDocument
LOCAL aOut := {}, a
FOR EACH a IN ::Body()
   IF a[ 1 ] == "t"
      AAdd( aOut, a[ 2 ] )
   ENDIF
NEXT
RETURN aOut

// ceo tekst; pasusi u redovima, celije tabele odvojene tabom
METHOD Text() CLASS WordDocument
LOCAL aOut := {}, a, aRow
FOR EACH a IN ::Body()
   IF a[ 1 ] == "p"
      AAdd( aOut, a[ 2 ] )
   ELSE
      FOR EACH aRow IN a[ 2 ]
         AAdd( aOut, _Join( aRow, Chr( 9 ) ) )
      NEXT
   ENDIF
NEXT
RETURN _Join( aOut, hb_eol() )

// Zamena teksta u ucitanom dokumentu (telo, zaglavlja, podnozja, fusnote).
// Replace( "{{ime}}", "Petar" ) ili Replace( { "{{ime}}" => "Petar", "{{iznos}}" => 1250 } ).
// Pronalazi tekst i kada ga je Word podelio na vise delova (run) sa razlicitim
// atributima; zamena dobija format prvog znaka. Vraca broj zamena.
METHOD Replace( xFind, xReplace ) CLASS WordDocument
LOCAL aPairs, cPart, nCount := 0, cDir
IF ! ::lRead
   RETURN 0
ENDIF
aPairs := _Pairs( Self, IIF( HB_ISHASH( xFind ), xFind, { xFind => xReplace } ) )
cDir := _Dir( ::cMainPart )
FOR EACH cPart IN hb_HKeys( ::hParts )
   IF cPart == ::cMainPart .OR. ;
      ( hb_BLeft( cPart, Len( cDir ) ) == cDir .AND. Right( cPart, 4 ) == ".xml" .AND. ;
        ( "/header" $ cPart .OR. "/footer" $ cPart .OR. "/footnotes" $ cPart .OR. "/endnotes" $ cPart ) .AND. ;
        !( "/_rels/" $ cPart ) )
      ::hParts[ cPart ] := _ReplaceXml( ::hParts[ cPart ], aPairs, @nCount )
   ENDIF
NEXT
RETURN nCount

// Red tabele u kome je cMarker umnozava se za svaki element aRows; u svakoj kopiji
// oznake se zamene vrednostima. Element je hash { "{{naziv}}" => "Olovka", ... },
// ili niz vrednosti kada je zadat aKeys (imena oznaka po redu).
// Prazan aRows brise red. Vraca .F. ako red nije nadjen.
METHOD RepeatRow( cMarker, aRows, aKeys ) CLASS WordDocument
LOCAL cXml, aEv := {}, aStack := {}, p, c, x, nS, nE, nBest := 0, nBS := 0, nBE := 0, cRow, aOut := {}
IF ! ::lRead .OR. ! HB_ISSTRING( cMarker ) .OR. Empty( cMarker ) .OR. ! HB_ISARRAY( aRows )
   RETURN .F.
ENDIF
cMarker := ::ToUTF8( cMarker )
cXml := ::hParts[ ::cMainPart ]
p := 1
DO WHILE ( p := hb_BAt( "<w:tr", cXml, p ) ) > 0
   c := hb_BSubStr( cXml, p + 5, 1 )
   IF c == ">" .OR. c == " "
      AAdd( aEv, { p, 1 } )
   ENDIF
   p++
ENDDO
p := 1
DO WHILE ( p := hb_BAt( "</w:tr>", cXml, p ) ) > 0
   AAdd( aEv, { p, 2 } )
   p++
ENDDO
ASort( aEv, , , {| a, b | a[ 1 ] < b[ 1 ] } )
// najuzi red (ugnezdene tabele) koji sadrzi oznaku
FOR EACH x IN aEv
   IF x[ 2 ] == 1
      AAdd( aStack, x[ 1 ] )
   ELSEIF ! Empty( aStack )
      nS := ATail( aStack )
      ASize( aStack, Len( aStack ) - 1 )
      nE := x[ 1 ] + 7
      IF ( nBest == 0 .OR. nE - nS < nBest ) .AND. cMarker $ _PlainText( hb_BSubStr( cXml, nS, nE - nS ) )
         nBest := nE - nS
         nBS := nS
         nBE := nE
      ENDIF
   ENDIF
NEXT
IF nBest == 0
   RETURN .F.
ENDIF
cRow := hb_BSubStr( cXml, nBS, nBest )
FOR EACH x IN aRows
   AAdd( aOut, _ReplaceXml( cRow, _Pairs( Self, x, aKeys ) ) )
NEXT
::hParts[ ::cMainPart ] := hb_BLeft( cXml, nBS - 1 ) + _Join( aOut, "" ) + hb_BSubStr( cXml, nBE )
RETURN .T.

// sirina teksta (strana bez margina) u cm
METHOD TextWidth( hOut ) CLASS WordDocument
LOCAL aP, cXml, nS, nE, cSect, n, hSz, hMar, nW
IF ! ::lRead
   aP := _PaperSize( ::nPaperSize )
   RETURN Max( 2, ( IIF( ::lLandscape, aP[ 2 ], aP[ 1 ] ) / ( 1440 / 2.54 ) ) - ::nLeftMargin - ::nRightMargin )
ENDIF
cXml := hOut[ ::cMainPart ]
nS := _LastSectPr( cXml )
nE := IIF( nS > 0, hb_BAt( "</w:sectPr>", cXml, nS ), 0 )
IF nE > 0
   cSect := hb_BSubStr( cXml, nS, nE - nS )
   n := hb_BAt( "<w:pgSz ", cSect )
   hSz := IIF( n > 0, _XAttrs( hb_BSubStr( cSect, n + 8, hb_BAt( ">", cSect, n ) - n - 8 ) ), { => } )
   n := hb_BAt( "<w:pgMar ", cSect )
   hMar := IIF( n > 0, _XAttrs( hb_BSubStr( cSect, n + 9, hb_BAt( ">", cSect, n ) - n - 9 ) ), { => } )
   nW := Val( _HAttr( hSz, "w" ) ) - Val( _HAttr( hMar, "left" ) ) - Val( _HAttr( hMar, "right" ) )
   IF nW > 567
      RETURN nW / ( 1440 / 2.54 )
   ENDIF
ENDIF
RETURN 16

METHOD Save( cName ) CLASS WordDocument
LOCAL hOut, hCtx, cFile, cMain, cDir, cStyles, cNum, cSet, aSM, cBody, cPart, cRels, cNormal, cAbs, cNums, nAbs
LOCAL i, k, a, aRels, c
::cError := ""
IF HB_ISSTRING( cName ) .AND. ! Empty( cName )
   ::SetName( cName )
ENDIF
IF Empty( ::cName )
   ::cError := "Nije zadato ime fajla: WordDocument():New( cFajl )"
   RETURN .F.
ENDIF
cFile := ::cFilePath + ::cName
// radi se na kopiji, pa Save() moze da se pozove vise puta
hOut  := IIF( ::lRead, hb_HClone( ::hParts ), ::BaseParts() )
cMain := ::cMainPart
cDir  := _Dir( cMain )
hCtx  := { "doc" => Self, "out" => hOut, "part" => cMain, "dir" => cDir, "rels" => { => }, "relN" => 0, ;
           "media" => { => }, "docPr" => 0, "numBase" => 0, "compact" => .F., "newStyles" => { => } }
hCtx[ "textWidth" ] := ::TextWidth( hOut )

// stilovi: ugradjeni stil se trazi po imenu ("heading 1"), jer lokalizovani
// Word daje drugi styleId; nedostajuci se dodaju u styles.xml
cStyles := _RelTarget( hOut, cMain, "/styles" )
IF Empty( cStyles )
   cStyles := cDir + "styles.xml"
   hOut[ cStyles ] := _BaseStyles( Self )
   _AddRel( hCtx, "styles", "styles.xml" )
   _CTOverride( hOut, cStyles, CT_WML + "styles+xml" )
ENDIF
aSM := _StyleMap( hOut[ cStyles ] )
hCtx[ "styles" ]   := aSM[ 1 ]
hCtx[ "styleIds" ] := aSM[ 3 ]

// jedinstveni id-jevi slika u celom dokumentu
FOR EACH cPart IN hb_HKeys( hOut )
   IF hb_BLeft( cPart, Len( cDir ) ) == cDir .AND. Right( cPart, 4 ) == ".xml"
      hCtx[ "docPr" ] := Max( hCtx[ "docPr" ], _MaxNum( hOut[ cPart ], 'docPr id="' ) )
   ENDIF
NEXT

// liste: numbering.xml (novi ili postojeci, sa id-jevima posle postojecih)
IF ! Empty( ::aLists )
   cNum := _RelTarget( hOut, cMain, "/numbering" )
   IF Empty( cNum )
      cNum := _FreePart( hOut, cDir + "numbering", ".xml" )
      hOut[ cNum ] := XML_HEAD + '<w:numbering xmlns:w="' + NS_W + '"></w:numbering>'
      _AddRel( hCtx, "numbering", hb_BSubStr( cNum, Len( cDir ) + 1 ) )
      _CTOverride( hOut, cNum, CT_WML + "numbering+xml" )
   ENDIF
   hCtx[ "numBase" ] := _MaxNum( hOut[ cNum ], 'w:numId="' )
   nAbs := _MaxNum( hOut[ cNum ], 'w:abstractNumId="' ) + 1
   cAbs := ""
   cNums := ""
   FOR i := 1 TO Len( ::aLists )
      cAbs  += _AbstractNum( nAbs + i - 1, ::aLists[ i ] )
      cNums += '<w:num w:numId="' + hb_NToS( hCtx[ "numBase" ] + i ) + '"><w:abstractNumId w:val="' + hb_NToS( nAbs + i - 1 ) + '"/></w:num>'
   NEXT
   hOut[ cNum ] := _InsertNumbering( hOut[ cNum ], cAbs, cNums )
ENDIF

// telo
cBody := ::XML( hCtx )
IF ! Empty( ::aItems ) .AND. ATail( ::aItems ):ClassName() == "DOCXTABLE"
   cBody += "<w:p/>"
ENDIF
IF ! Empty( cBody )
   hOut[ cMain ] := _InsertBody( hOut[ cMain ], cBody, ::cInsertAt )
ENDIF

// zaglavlje i podnozje
hCtx[ "compact" ] := .T.
FOR EACH a IN { { ::oHeader, "header", "hdr", "headerReference" }, { ::oFooter, "footer", "ftr", "footerReference" } }
   IF a[ 1 ] != NIL
      cPart := _FreePart( hOut, cDir + a[ 2 ], ".xml" )
      hCtx[ "part" ] := cPart
      hOut[ cPart ] := XML_HEAD + "<w:" + a[ 3 ] + ' xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '">' + ;
                       a[ 1 ]:XML( hCtx, .T. ) + "</w:" + a[ 3 ] + ">"
      hCtx[ "part" ] := cMain
      hOut[ cMain ] := _SetHFRef( hOut[ cMain ], a[ 4 ], _AddRel( hCtx, a[ 2 ], hb_BSubStr( cPart, Len( cDir ) + 1 ) ) )
      _CTOverride( hOut, cPart, CT_WML + a[ 2 ] + "+xml" )
   ENDIF
NEXT
hCtx[ "compact" ] := .F.

// stilovi koje dokument nema
IF ! Empty( hCtx[ "newStyles" ] )
   cNormal := hb_HGetDef( hCtx[ "styles" ], "normal", "Normal" )
   c := ""
   FOR EACH k IN hb_HKeys( hCtx[ "newStyles" ] )
      c += _StyleDef( k, hCtx[ "newStyles" ][ k ], cNormal )
   NEXT
   hOut[ cStyles ] := _InsertBefore( hOut[ cStyles ], "</w:styles>", c )
ENDIF

// relacije svih delova
FOR EACH cPart IN hb_HKeys( hCtx[ "rels" ] )
   c := ""
   FOR EACH aRels IN hb_HValues( hCtx[ "rels" ][ cPart ] )
      c += '<Relationship Id="' + aRels[ 1 ] + '" Type="' + REL_OFFICE + aRels[ 2 ] + '" Target="' + XmlEsc( aRels[ 3 ] ) + '"' + ;
           IIF( aRels[ 4 ], ' TargetMode="External"', "" ) + "/>"
   NEXT
   cRels := _RelsPath( cPart )
   IF hb_HHasKey( hOut, cRels )
      hOut[ cRels ] := _InsertBefore( hOut[ cRels ], "</Relationships>", c )
   ELSE
      hOut[ cRels ] := XML_HEAD + '<Relationships xmlns="' + NS_PKG + '">' + c + "</Relationships>"
   ENDIF
NEXT

// azuriranje polja (sadrzaj) pri otvaranju
IF ::lUpdateFields
   cSet := _RelTarget( hOut, cMain, "/settings" )
   IF ! Empty( cSet ) .AND. hb_HHasKey( hOut, cSet ) .AND. !( "<w:updateFields" $ hOut[ cSet ] )
      hOut[ cSet ] := _InsertSetting( hOut[ cSet ], '<w:updateFields w:val="true"/>' )
   ENDIF
ENDIF

// fajl otvoren u Wordu ne moze da se obrise
IF hb_FileExists( cFile ) .AND. FErase( cFile ) != 0
   ::cError := "Ne mogu da prepisem fajl " + cFile + ";Verovatno je otvoren u Wordu - zatvorite ga i pokusajte ponovo."
   RETURN .F.
ENDIF
IF ! _ZipWrite( cFile, hOut )
   ::cError := "Greska pri upisu fajla " + cFile
   FErase( cFile )
   RETURN .F.
ENDIF
RETURN .T.

// delovi praznog dokumenta
METHOD BaseParts() CLASS WordDocument
LOCAL h := { => }, aP := _PaperSize( ::nPaperSize ), nW, nH, cDate := _W3CDate()
nW := aP[ 1 ]
nH := aP[ 2 ]
IF ::lLandscape
   nW := aP[ 2 ]
   nH := aP[ 1 ]
ENDIF
h[ "[Content_Types].xml" ] := XML_HEAD + '<Types xmlns="' + NS_CT + '">' + ;
   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + ;
   '<Default Extension="xml" ContentType="application/xml"/>' + ;
   '<Override PartName="/word/document.xml" ContentType="' + CT_WML + 'document.main+xml"/>' + ;
   '<Override PartName="/word/styles.xml" ContentType="' + CT_WML + 'styles+xml"/>' + ;
   '<Override PartName="/word/settings.xml" ContentType="' + CT_WML + 'settings+xml"/>' + ;
   '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' + ;
   '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' + ;
   "</Types>"
h[ "_rels/.rels" ] := XML_HEAD + '<Relationships xmlns="' + NS_PKG + '">' + ;
   '<Relationship Id="rId1" Type="' + REL_OFFICE + 'officeDocument" Target="word/document.xml"/>' + ;
   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' + ;
   '<Relationship Id="rId3" Type="' + REL_OFFICE + 'extended-properties" Target="docProps/app.xml"/>' + ;
   "</Relationships>"
h[ "docProps/core.xml" ] := XML_HEAD + ;
   '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ' + ;
   'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ' + ;
   'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' + ;
   "<dc:title>" + XmlEsc( ::ToUTF8( ::cTitle ) ) + "</dc:title>" + ;
   "<dc:subject>" + XmlEsc( ::ToUTF8( ::cSubject ) ) + "</dc:subject>" + ;
   "<dc:creator>" + XmlEsc( ::ToUTF8( ::cAuthor ) ) + "</dc:creator>" + ;
   "<cp:lastModifiedBy>" + XmlEsc( ::ToUTF8( ::cAuthor ) ) + "</cp:lastModifiedBy>" + ;
   '<dcterms:created xsi:type="dcterms:W3CDTF">' + cDate + "</dcterms:created>" + ;
   '<dcterms:modified xsi:type="dcterms:W3CDTF">' + cDate + "</dcterms:modified>" + ;
   "</cp:coreProperties>"
h[ "docProps/app.xml" ] := XML_HEAD + ;
   '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" ' + ;
   'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' + ;
   "<Application>" + XmlEsc( ::ToUTF8( ::Application ) ) + "</Application>" + ;
   "<Company>" + XmlEsc( ::ToUTF8( ::cCompany ) ) + "</Company>" + ;
   "</Properties>"
h[ "word/_rels/document.xml.rels" ] := XML_HEAD + '<Relationships xmlns="' + NS_PKG + '">' + ;
   '<Relationship Id="rId1" Type="' + REL_OFFICE + 'styles" Target="styles.xml"/>' + ;
   '<Relationship Id="rId2" Type="' + REL_OFFICE + 'settings" Target="settings.xml"/>' + ;
   "</Relationships>"
h[ "word/document.xml" ] := XML_HEAD + '<w:document xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '" ' + ;
   'xmlns:wp="' + NS_WP + '" xmlns:a="' + NS_A + '" xmlns:pic="' + NS_PIC + '"><w:body>' + ;
   '<w:sectPr><w:pgSz w:w="' + hb_NToS( nW ) + '" w:h="' + hb_NToS( nH ) + '"' + IIF( ::lLandscape, ' w:orient="landscape"', "" ) + "/>" + ;
   '<w:pgMar w:top="' + hb_NToS( CM2TW( ::nTopMargin ) ) + '" w:right="' + hb_NToS( CM2TW( ::nRightMargin ) ) + ;
   '" w:bottom="' + hb_NToS( CM2TW( ::nBottomMargin ) ) + '" w:left="' + hb_NToS( CM2TW( ::nLeftMargin ) ) + ;
   '" w:header="' + hb_NToS( CM2TW( ::nHeaderMargin ) ) + '" w:footer="' + hb_NToS( CM2TW( ::nFooterMargin ) ) + '" w:gutter="0"/>' + ;
   '<w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>'
h[ "word/styles.xml" ] := _BaseStyles( Self )
h[ "word/settings.xml" ] := XML_HEAD + '<w:settings xmlns:w="' + NS_W + '">' + ;
   '<w:zoom w:percent="100"/><w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="doNotCompress"/>' + ;
   '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>' + ;
   "</w:settings>"
RETURN h

// ---------------------------------------------------------------------------
// Pomocne funkcije - pisanje
// ---------------------------------------------------------------------------

STATIC FUNCTION _BaseStyles( oDoc )
LOCAL cFont := XmlEsc( oDoc:ToUTF8( oDoc:cFont ) ), nSz := Round( oDoc:nFontSize * 2, 0 )
RETURN XML_HEAD + '<w:styles xmlns:w="' + NS_W + '"><w:docDefaults><w:rPrDefault><w:rPr>' + ;
   '<w:rFonts w:ascii="' + cFont + '" w:eastAsia="' + cFont + '" w:hAnsi="' + cFont + '" w:cs="' + cFont + '"/>' + ;
   '<w:sz w:val="' + hb_NToS( nSz ) + '"/><w:szCs w:val="' + hb_NToS( nSz ) + '"/>' + ;
   '<w:lang w:val="' + XmlEsc( oDoc:cLanguage ) + '"/></w:rPr></w:rPrDefault>' + ;
   '<w:pPrDefault><w:pPr><w:spacing w:after="' + hb_NToS( Round( oDoc:nSpaceAfter * 20, 0 ) ) + ;
   '" w:line="' + hb_NToS( Round( oDoc:nLineSpacing * 240, 0 ) ) + '" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' + ;
   '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' + ;
   '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/>' + ;
   '<w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' + ;
   '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/>' + ;
   '<w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/>' + ;
   '<w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' + ;
   '<w:style w:type="numbering" w:default="1" w:styleId="NoList"><w:name w:val="No List"/><w:uiPriority w:val="99"/>' + ;
   '<w:semiHidden/><w:unhideWhenUsed/></w:style></w:styles>'

// definicija ugradjenog stila koji dokument nema
STATIC FUNCTION _StyleDef( cKey, cId, cNormal )
LOCAL n, aSz := { 32, 28, 24, 22, 22, 22 }, aBefore := { 360, 240, 200, 160, 160, 160 }
LOCAL aColor := { "1F3864", "2F5496", "1F3763", "2F5496", "2F5496", "1F3763" }
LOCAL cHead := '<w:style w:type="paragraph" w:styleId="' + cId + '">'
DO CASE
CASE Left( cKey, 8 ) == "heading "
   n := Max( 1, Min( 6, Val( SubStr( cKey, 9 ) ) ) )
   RETURN cHead + '<w:name w:val="heading ' + hb_NToS( n ) + '"/><w:basedOn w:val="' + cNormal + '"/><w:next w:val="' + cNormal + '"/>' + ;
      '<w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/>' + ;
      '<w:spacing w:before="' + hb_NToS( aBefore[ n ] ) + '" w:after="' + IIF( n <= 2, "120", "80" ) + '"/>' + ;
      '<w:outlineLvl w:val="' + hb_NToS( n - 1 ) + '"/></w:pPr><w:rPr><w:b/><w:bCs/>' + IIF( n >= 4, "<w:i/><w:iCs/>", "" ) + ;
      '<w:color w:val="' + aColor[ n ] + '"/><w:sz w:val="' + hb_NToS( aSz[ n ] ) + '"/><w:szCs w:val="' + hb_NToS( aSz[ n ] ) + '"/></w:rPr></w:style>'
CASE cKey == "title"
   RETURN cHead + '<w:name w:val="Title"/><w:basedOn w:val="' + cNormal + '"/><w:next w:val="' + cNormal + '"/>' + ;
      '<w:uiPriority w:val="10"/><w:qFormat/><w:pPr><w:spacing w:after="240" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>' + ;
      '<w:rPr><w:color w:val="1F3864"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr></w:style>'
CASE cKey == "list paragraph"
   RETURN cHead + '<w:name w:val="List Paragraph"/><w:basedOn w:val="' + cNormal + '"/><w:uiPriority w:val="34"/><w:qFormat/>' + ;
      '<w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>'
CASE cKey == "toc heading"
   RETURN cHead + '<w:name w:val="TOC Heading"/><w:basedOn w:val="' + cNormal + '"/><w:next w:val="' + cNormal + '"/>' + ;
      '<w:uiPriority w:val="39"/><w:unhideWhenUsed/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="120"/>' + ;
      '<w:outlineLvl w:val="9"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr></w:style>'
CASE cKey == "hyperlink"
   RETURN '<w:style w:type="character" w:styleId="' + cId + '"><w:name w:val="Hyperlink"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/>' + ;
      '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
ENDCASE
RETURN cHead + '<w:name w:val="' + XmlEsc( cKey ) + '"/><w:basedOn w:val="' + cNormal + '"/></w:style>'

// styleId ugradjenog stila po imenu; ako ga dokument nema, belezi se za dodavanje
STATIC FUNCTION _StyleId( hCtx, cName )
LOCAL cKey := Lower( cName ), cId
IF hb_HHasKey( hCtx[ "styles" ], cKey )
   RETURN hCtx[ "styles" ][ cKey ]
ENDIF
cId := StrTran( cName, " ", "" )
cId := Upper( Left( cId, 1 ) ) + SubStr( cId, 2 )
DO WHILE hb_HHasKey( hCtx[ "styleIds" ], Lower( cId ) )
   cId := "Docx" + cId
ENDDO
hCtx[ "styleIds" ][ Lower( cId ) ] := cId
hCtx[ "styles" ][ cKey ] := cId
hCtx[ "newStyles" ][ cKey ] := cId
RETURN cId

STATIC FUNCTION _AbstractNum( nId, cType )
LOCAL c, i, cFmt, cText, cFont
LOCAL aBul := { hb_UTF8Chr( 0xF0B7 ), "o", hb_UTF8Chr( 0xF0A7 ) }, aBulF := { "Symbol", "Courier New", "Wingdings" }
LOCAL aNum := { "decimal", "lowerLetter", "lowerRoman" }
c := '<w:abstractNum w:abstractNumId="' + hb_NToS( nId ) + '"><w:multiLevelType w:val="hybridMultilevel"/>'
FOR i := 0 TO 8
   IF cType == "bullet"
      cFmt  := "bullet"
      cText := aBul[ i % 3 + 1 ]
      cFont := aBulF[ i % 3 + 1 ]
   ELSE
      cFmt  := IIF( i == 0, cType, aNum[ i % 3 + 1 ] )
      cText := "%" + hb_NToS( i + 1 ) + "."
      cFont := ""
   ENDIF
   c += '<w:lvl w:ilvl="' + hb_NToS( i ) + '"><w:start w:val="1"/><w:numFmt w:val="' + cFmt + '"/>' + ;
        '<w:lvlText w:val="' + cText + '"/><w:lvlJc w:val="left"/>' + ;
        '<w:pPr><w:ind w:left="' + hb_NToS( 720 * ( i + 1 ) ) + '" w:hanging="360"/></w:pPr>' + ;
        IIF( Empty( cFont ), "", '<w:rPr><w:rFonts w:ascii="' + cFont + '" w:hAnsi="' + cFont + '" w:hint="default"/></w:rPr>' ) + ;
        "</w:lvl>"
NEXT
RETURN c + "</w:abstractNum>"

// abstractNum mora pre svih <w:num>
STATIC FUNCTION _InsertNumbering( cXml, cAbs, cNums )
LOCAL p, q, n
q := hb_BAt( "<w:numIdMacAtCleanup", cXml )
IF q == 0
   q := hb_BRAt( "</w:numbering>", cXml )
ENDIF
cXml := hb_BLeft( cXml, q - 1 ) + cNums + hb_BSubStr( cXml, q )
p := hb_BAt( "<w:num ", cXml )
n := hb_BAt( "<w:num>", cXml )
IF n > 0 .AND. ( p == 0 .OR. n < p )
   p := n
ENDIF
RETURN hb_BLeft( cXml, p - 1 ) + cAbs + hb_BSubStr( cXml, p )

STATIC FUNCTION _HFParagraph( oP, cLeft, cCenter, cRight, nFont )
oP:lHFTabs := .T.
_HFCodes( oP, cLeft, nFont )
oP:Text( Chr( 9 ), nFont )
_HFCodes( oP, cCenter, nFont )
oP:Text( Chr( 9 ), nFont )
_HFCodes( oP, cRight, nFont )
RETURN oP

STATIC PROCEDURE _HFCodes( oP, cText, nFont )
LOCAL n := 1, cBuf := "", ch
IF ! HB_ISSTRING( cText )
   RETURN
ENDIF
DO WHILE n <= Len( cText )
   ch := SubStr( cText, n, 1 )
   IF ch == "&" .AND. n < Len( cText ) .AND. Upper( SubStr( cText, n + 1, 1 ) ) $ "PNDTF&"
      ch := Upper( SubStr( cText, n + 1, 1 ) )
      IF ch == "&"
         cBuf += "&"
      ELSE
         IF Len( cBuf ) > 0
            oP:Text( cBuf, nFont )
            cBuf := ""
         ENDIF
         DO CASE
         CASE ch == "P" ; oP:PageNumber( nFont )
         CASE ch == "N" ; oP:PageCount( nFont )
         CASE ch == "D" ; oP:Field( 'DATE \@ "' + StrTran( Lower( Set( _SET_DATEFORMAT ) ), "m", "M" ) + '"', nFont, DToC( Date() ) )
         CASE ch == "T" ; oP:Field( 'TIME \@ "HH:mm"', nFont, Left( Time(), 5 ) )
         CASE ch == "F" ; oP:Field( "FILENAME", nFont, oP:oDoc:cName )
         ENDCASE
      ENDIF
      n += 2
      LOOP
   ENDIF
   cBuf += ch
   n++
ENDDO
IF Len( cBuf ) > 0
   oP:Text( cBuf, nFont )
ENDIF
RETURN

// deo teksta; Chr(10) -> <w:br/>, Chr(9) -> <w:tab/>
STATIC FUNCTION _RunXML( cText, cRPr )
LOCAL c := "<w:r>" + IIF( Empty( cRPr ), "", "<w:rPr>" + cRPr + "</w:rPr>" ), a, i, cPart
cText := StrTran( StrTran( cText, Chr( 13 ) + Chr( 10 ), Chr( 10 ) ), Chr( 13 ), Chr( 10 ) )
a := hb_ATokens( cText, Chr( 10 ) )
FOR i := 1 TO Len( a )
   IF i > 1
      c += "<w:br/>"
   ENDIF
   FOR EACH cPart IN hb_ATokens( a[ i ], Chr( 9 ) )
      IF cPart:__enumIndex() > 1
         c += "<w:tab/>"
      ENDIF
      IF Len( cPart ) > 0
         c += '<w:t xml:space="preserve">' + XmlEsc( cPart ) + "</w:t>"
      ENDIF
   NEXT
NEXT
RETURN c + "</w:r>"

STATIC FUNCTION _RunRPr( oDoc, nFont, cDef )
IF HB_ISNUMERIC( nFont ) .AND. nFont >= 1 .AND. nFont <= Len( oDoc:aFonts )
   RETURN oDoc:aFonts[ nFont ]
ENDIF
RETURN cDef

STATIC FUNCTION _FontRPr( oDoc, nFont )
RETURN IIF( nFont >= 1 .AND. nFont <= Len( oDoc:aFonts ), oDoc:aFonts[ nFont ], "" )

STATIC FUNCTION _OnOff( cTag, l )
IF HB_ISLOGICAL( l )
   RETURN IIF( l, "<w:" + cTag + "/>", "<w:" + cTag + ' w:val="0"/>' )
ENDIF
RETURN ""

// "#FF0000", "ff0000", "FFFF0000" (ARGB) -> "FF0000"
STATIC FUNCTION _RGB( c )
c := Upper( StrTran( AllTrim( c ), "#", "" ) )
RETURN IIF( Len( c ) == 8, Right( c, 6 ), c )

STATIC FUNCTION _PictureXML( hCtx, cFile, nW, nH )
LOCAL cKey := Lower( cFile ), cPart, cData, aPx, nId, cRId, cExt, nMax
IF hb_HHasKey( hCtx[ "media" ], cKey )
   cPart := hCtx[ "media" ][ cKey ][ 1 ]
   aPx   := hCtx[ "media" ][ cKey ][ 2 ]
ELSE
   cData := hb_MemoRead( cFile )
   IF Empty( cData )
      RETURN ""
   ENDIF
   cExt := Lower( SubStr( hb_FNameExt( cFile ), 2 ) )
   aPx := _ImageSize( cData )
   cPart := _FreePart( hCtx[ "out" ], hCtx[ "dir" ] + "media/docx_image", "." + cExt )
   hCtx[ "out" ][ cPart ] := cData
   _CTDefault( hCtx[ "out" ], cExt, _Mime( cExt ) )
   hCtx[ "media" ][ cKey ] := { cPart, aPx }
ENDIF
cRId := _AddRel( hCtx, "image", hb_BSubStr( cPart, Len( hCtx[ "dir" ] ) + 1 ) )
nW := IIF( HB_ISNUMERIC( nW ) .AND. nW > 0, nW, NIL )
nH := IIF( HB_ISNUMERIC( nH ) .AND. nH > 0, nH, NIL )
DO CASE
CASE nW == NIL .AND. nH == NIL
   IF aPx[ 1 ] > 0 .AND. aPx[ 2 ] > 0
      nW := aPx[ 1 ] * 2.54 / 96
      nH := aPx[ 2 ] * 2.54 / 96
      nMax := hCtx[ "textWidth" ]
      IF nW > nMax
         nH := nH * nMax / nW
         nW := nMax
      ENDIF
   ELSE
      nW := nH := 5
   ENDIF
CASE nH == NIL
   nH := IIF( aPx[ 1 ] > 0, nW * aPx[ 2 ] / aPx[ 1 ], nW )
CASE nW == NIL
   nW := IIF( aPx[ 2 ] > 0, nH * aPx[ 1 ] / aPx[ 2 ], nH )
ENDCASE
nId := ++hCtx[ "docPr" ]
nW := hb_NToS( CM2EMU( nW ) )
nH := hb_NToS( CM2EMU( nH ) )
RETURN '<w:r><w:drawing><wp:inline xmlns:wp="' + NS_WP + '" distT="0" distB="0" distL="0" distR="0">' + ;
   '<wp:extent cx="' + nW + '" cy="' + nH + '"/><wp:effectExtent l="0" t="0" r="0" b="0"/>' + ;
   '<wp:docPr id="' + hb_NToS( nId ) + '" name="Picture ' + hb_NToS( nId ) + '"/>' + ;
   '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="' + NS_A + '" noChangeAspect="1"/></wp:cNvGraphicFramePr>' + ;
   '<a:graphic xmlns:a="' + NS_A + '"><a:graphicData uri="' + NS_PIC + '"><pic:pic xmlns:pic="' + NS_PIC + '">' + ;
   '<pic:nvPicPr><pic:cNvPr id="0" name="' + XmlEsc( hCtx[ "doc" ]:ToUTF8( hb_FNameNameExt( cFile ) ) ) + '"/><pic:cNvPicPr/></pic:nvPicPr>' + ;
   '<pic:blipFill><a:blip xmlns:r="' + NS_R + '" r:embed="' + cRId + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' + ;
   '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + nW + '" cy="' + nH + '"/></a:xfrm>' + ;
   '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

// velicina slike u pikselima iz zaglavlja PNG, JPEG, GIF, BMP; { 0, 0 } ako nije poznata
STATIC FUNCTION _ImageSize( c )
LOCAL n, nLen := hb_BLen( c ), nMark, nW, nH
DO CASE
CASE hb_BLeft( c, 4 ) == hb_BChar( 137 ) + "PNG"
   RETURN { _BE( c, 17, 4 ), _BE( c, 21, 4 ) }
CASE hb_BLeft( c, 3 ) == "GIF"
   RETURN { _LE( c, 7, 2 ), _LE( c, 9, 2 ) }
CASE hb_BLeft( c, 2 ) == "BM"
   nW := _LE( c, 19, 4 )
   nH := _LE( c, 23, 4 )
   RETURN { nW, IIF( nH > 0x7FFFFFFF, 0x100000000 - nH, nH ) }
CASE hb_BLeft( c, 2 ) == hb_BChar( 255 ) + hb_BChar( 216 )
   n := 3
   DO WHILE n + 8 <= nLen
      IF hb_BPeek( c, n ) != 255
         n++
         LOOP
      ENDIF
      nMark := hb_BPeek( c, n + 1 )
      DO CASE
      CASE nMark == 255
         n++
      CASE nMark >= 0xC0 .AND. nMark <= 0xCF .AND. nMark != 0xC4 .AND. nMark != 0xC8 .AND. nMark != 0xCC
         RETURN { _BE( c, n + 7, 2 ), _BE( c, n + 5, 2 ) }
      CASE nMark == 0x01 .OR. ( nMark >= 0xD0 .AND. nMark <= 0xD8 )
         n += 2
      OTHERWISE
         n += 2 + _BE( c, n + 2, 2 )
      ENDCASE
   ENDDO
ENDCASE
RETURN { 0, 0 }

STATIC FUNCTION _BE( c, nPos, nBytes )
LOCAL n := 0, i
FOR i := 0 TO nBytes - 1
   n := n * 256 + hb_BPeek( c, nPos + i )
NEXT
RETURN n

STATIC FUNCTION _LE( c, nPos, nBytes )
LOCAL n := 0, i
FOR i := nBytes - 1 TO 0 STEP -1
   n := n * 256 + hb_BPeek( c, nPos + i )
NEXT
RETURN n

STATIC FUNCTION _Mime( cExt )
DO CASE
CASE cExt == "jpeg" .OR. cExt == "jpg" ; RETURN "image/jpeg"
CASE cExt == "png"  ; RETURN "image/png"
CASE cExt == "gif"  ; RETURN "image/gif"
CASE cExt == "bmp"  ; RETURN "image/bmp"
CASE cExt == "emf"  ; RETURN "image/x-emf"
CASE cExt == "wmf"  ; RETURN "image/x-wmf"
CASE cExt == "tif" .OR. cExt == "tiff" ; RETURN "image/tiff"
ENDCASE
RETURN "application/octet-stream"

// { sirina, visina } u twip
STATIC FUNCTION _PaperSize( n )
DO CASE
CASE n == 1  ; RETURN { 12240, 15840 }   // Letter
CASE n == 5  ; RETURN { 12240, 20160 }   // Legal
CASE n == 8  ; RETURN { 16838, 23811 }   // A3
CASE n == 11 ; RETURN { 8391, 11906 }    // A5
ENDCASE
RETURN { 11906, 16838 }                   // A4

STATIC FUNCTION _W3CDate()
LOCAL cDate := DToS( Date() )
RETURN Left( cDate, 4 ) + "-" + SubStr( cDate, 5, 2 ) + "-" + Right( cDate, 2 ) + "T" + Left( Time(), 8 ) + "Z"

// ---------------------------------------------------------------------------
// Pomocne funkcije - paket (zip delovi, relacije, content types)
// ---------------------------------------------------------------------------

// ceo zip u memoriju: ime dela -> sadrzaj (bez privremenih foldera)
STATIC FUNCTION _ZipRead( cFile )
LOCAL hParts := { => }, hUnzip, nErr, cName, nSize, cData, cBuf, nRead
hUnzip := hb_unzipOpen( cFile )
IF Empty( hUnzip )
   RETURN hParts
ENDIF
nErr := hb_unzipFileFirst( hUnzip )
DO WHILE nErr == 0
   cName := ""
   nSize := 0
   hb_unzipFileInfo( hUnzip, @cName, , , , , , @nSize )
   cName := StrTran( cName, "\", "/" )
   IF !( Right( cName, 1 ) == "/" ) .AND. hb_unzipFileOpen( hUnzip ) == 0
      cData := ""
      cBuf := Space( 65536 )
      DO WHILE ( nRead := hb_unzipFileRead( hUnzip, @cBuf ) ) > 0
         cData += hb_BLeft( cBuf, nRead )
      ENDDO
      hb_unzipFileClose( hUnzip )
      hParts[ cName ] := cData
   ENDIF
   nErr := hb_unzipFileNext( hUnzip )
ENDDO
hb_unzipClose( hUnzip )
RETURN hParts

STATIC FUNCTION _ZipWrite( cFile, hParts )
LOCAL hZip, cPart, lOk := .T., aKeys, n, tNow := hb_DateTime()
hZip := hb_zipOpen( cFile )
IF Empty( hZip )
   RETURN .F.
ENDIF
aKeys := hb_HKeys( hParts )
IF ( n := AScan( aKeys, {| x | x == "[Content_Types].xml" } ) ) > 1
   ADel( aKeys, n )
   AIns( aKeys, 1 )
   aKeys[ 1 ] := "[Content_Types].xml"
ENDIF
FOR EACH cPart IN aKeys
   IF hb_zipFileCreate( hZip, cPart, tNow ) == 0
      IF hb_BLen( hParts[ cPart ] ) > 0
         lOk := ( hb_zipFileWrite( hZip, hParts[ cPart ] ) == 0 ) .AND. lOk
      ENDIF
      lOk := ( hb_zipFileClose( hZip ) == 0 ) .AND. lOk
   ELSE
      lOk := .F.
   ENDIF
NEXT
lOk := ( hb_zipClose( hZip ) == 0 ) .AND. lOk
RETURN lOk

// relacija iz tekuceg dela (hCtx["part"]); isti cilj = isti id
STATIC FUNCTION _AddRel( hCtx, cType, cTarget, lExt )
LOCAL cPart := hCtx[ "part" ], hR := hCtx[ "rels" ], cKey, cId, cRels
lExt := HB_ISLOGICAL( lExt ) .AND. lExt
IF ! hb_HHasKey( hR, cPart )
   hR[ cPart ] := { => }
ENDIF
cKey := cType + "|" + cTarget
IF hb_HHasKey( hR[ cPart ], cKey )
   RETURN hR[ cPart ][ cKey ][ 1 ]
ENDIF
cRels := hb_HGetDef( hCtx[ "out" ], _RelsPath( cPart ), "" )
DO WHILE .T.
   hCtx[ "relN" ]++
   cId := "rIdDx" + hb_NToS( hCtx[ "relN" ] )
   IF !( '"' + cId + '"' $ cRels )
      EXIT
   ENDIF
ENDDO
hR[ cPart ][ cKey ] := { cId, cType, cTarget, lExt }
RETURN cId

// putanja dela na koji pokazuje relacija tipa ".../cTypeSuffix", "" ako je nema
STATIC FUNCTION _RelTarget( hParts, cPart, cTypeSuffix )
LOCAL cTok, cA, cTag, nKind, h
FOR EACH cTok IN hb_ATokens( hb_HGetDef( hParts, _RelsPath( cPart ), "" ), "<" )
   cA := _XTok( cTok, @cTag, @nKind )
   IF cTag == "Relationship" .AND. nKind != 2
      h := _XAttrs( cA )
      IF Right( _HAttr( h, "Type" ), Len( cTypeSuffix ) ) == cTypeSuffix .AND. !( _HAttr( h, "TargetMode" ) == "External" )
         RETURN _PartPath( _Dir( cPart ), _HAttr( h, "Target" ) )
      ENDIF
   ENDIF
NEXT
RETURN ""

STATIC FUNCTION _PartPath( cBase, cTarget )
LOCAL c, n
IF Left( cTarget, 1 ) == "/"
   RETURN SubStr( cTarget, 2 )
ENDIF
c := cBase + cTarget
DO WHILE ( n := At( "/../", c ) ) > 0
   c := hb_BLeft( c, hb_BRAt( "/", hb_BLeft( c, n - 1 ) ) ) + hb_BSubStr( c, n + 4 )
ENDDO
RETURN StrTran( c, "./", "" )

STATIC FUNCTION _RelsPath( cPart )
LOCAL n := hb_BRAt( "/", cPart )
RETURN hb_BLeft( cPart, n ) + "_rels/" + hb_BSubStr( cPart, n + 1 ) + ".rels"

STATIC FUNCTION _Dir( cPart )
RETURN hb_BLeft( cPart, hb_BRAt( "/", cPart ) )

STATIC FUNCTION _FreePart( hParts, cPrefix, cExt )
LOCAL n := 1
DO WHILE hb_HHasKey( hParts, cPrefix + hb_NToS( n ) + cExt )
   n++
ENDDO
RETURN cPrefix + hb_NToS( n ) + cExt

STATIC PROCEDURE _CTDefault( hOut, cExt, cMime )
LOCAL c := hOut[ "[Content_Types].xml" ]
IF !( 'extension="' + Lower( cExt ) + '"' $ Lower( c ) )
   hOut[ "[Content_Types].xml" ] := _InsertBefore( c, "</Types>", '<Default Extension="' + cExt + '" ContentType="' + cMime + '"/>' )
ENDIF
RETURN

STATIC PROCEDURE _CTOverride( hOut, cPart, cType )
LOCAL c := hOut[ "[Content_Types].xml" ]
IF !( 'PartName="/' + cPart + '"' $ c )
   hOut[ "[Content_Types].xml" ] := _InsertBefore( c, "</Types>", '<Override PartName="/' + cPart + '" ContentType="' + cType + '"/>' )
ENDIF
RETURN

STATIC FUNCTION _InsertBefore( c, cTag, cIns )
LOCAL n := hb_BRAt( cTag, c )
IF n == 0
   RETURN c + cIns
ENDIF
RETURN hb_BLeft( c, n - 1 ) + cIns + hb_BSubStr( c, n )

// updateFields mora na svoje mesto u settings.xml (pre compat, rsids...)
STATIC FUNCTION _InsertSetting( cXml, cIns )
LOCAL aAfter := { "<w:hdrShapeDefaults", "<w:footnotePr", "<w:endnotePr", "<w:compat", "<w:docVars", "<w:rsids", ;
                  "<m:mathPr", "<w:attachedSchema", "<w:themeFontLang", "<w:clrSchemeMapping", "<w:doNotIncludeSubdocsInStats", ;
                  "<w:doNotAutoCompressPictures", "<w:forceUpgrade", "<w:captions", "<w:readModeInkLockDown", "<w:smartTagType", ;
                  "<sl:schemaLibrary", "<w:shapeDefaults", "<w:doNotEmbedSmartTags", "<w:decimalSymbol", "<w:listSeparator", "</w:settings>" }
LOCAL c, n, nMin := 0
FOR EACH c IN aAfter
   IF ( n := hb_BAt( c, cXml ) ) > 0 .AND. ( nMin == 0 .OR. n < nMin )
      nMin := n
   ENDIF
NEXT
IF nMin == 0
   RETURN cXml
ENDIF
RETURN hb_BLeft( cXml, nMin - 1 ) + cIns + hb_BSubStr( cXml, nMin )

// pozicija poslednjeg <w:sectPr> (sekcija tela dokumenta), 0 ako ga nema
STATIC FUNCTION _LastSectPr( cXml )
RETURN Max( hb_BRAt( "<w:sectPr>", cXml ), hb_BRAt( "<w:sectPr ", cXml ) )

// novi sadrzaj tela: na mesto pasusa sa oznakom cMarker, ili pre <w:sectPr> tela
STATIC FUNCTION _InsertBody( cXml, cBody, cMarker )
LOCAL aP, n
IF HB_ISSTRING( cMarker ) .AND. ! Empty( aP := _FindParagraph( cXml, cMarker ) )
   RETURN hb_BLeft( cXml, aP[ 1 ] - 1 ) + cBody + hb_BSubStr( cXml, aP[ 2 ] )
ENDIF
n := _LastSectPr( cXml )
IF n == 0
   n := hb_BRAt( "</w:body>", cXml )
ENDIF
IF n == 0
   RETURN cXml
ENDIF
RETURN hb_BLeft( cXml, n - 1 ) + cBody + hb_BSubStr( cXml, n )

// { pocetak, kraj+1 } pasusa koji sadrzi tekst cMarker, {} ako ga nema
STATIC FUNCTION _FindParagraph( cXml, cMarker )
LOCAL aScan := _TScan( cXml ), aNodes := aScan[ 2 ], aPos := aScan[ 3 ], i := 1, nSeg, cFull, nFirst, nS, nE, cLeft
DO WHILE i <= Len( aNodes )
   nSeg := aNodes[ i, 3 ]
   nFirst := i
   cFull := ""
   DO WHILE i <= Len( aNodes ) .AND. aNodes[ i, 3 ] == nSeg
      cFull += aNodes[ i, 2 ]
      i++
   ENDDO
   IF cMarker $ cFull
      cLeft := hb_BLeft( cXml, aPos[ aNodes[ nFirst, 1 ] ] )
      nS := Max( hb_BRAt( "<w:p>", cLeft ), hb_BRAt( "<w:p ", cLeft ) )
      nE := hb_BAt( "</w:p>", cXml, aPos[ aNodes[ nFirst, 1 ] ] )
      IF nS > 0 .AND. nE > 0
         RETURN { nS, nE + 6 }
      ENDIF
   ENDIF
ENDDO
RETURN {}

// referenca na zaglavlje/podnozje u sekciji tela (zamenjuje postojecu "default")
STATIC FUNCTION _SetHFRef( cXml, cTag, cRId )
LOCAL nS := _LastSectPr( cXml ), nGt, nE, cSect, cRef, p, q
IF nS == 0
   RETURN cXml
ENDIF
nGt := hb_BAt( ">", cXml, nS )
cRef := "<w:" + cTag + ' xmlns:r="' + NS_R + '" w:type="default" r:id="' + cRId + '"/>'
IF hb_BSubStr( cXml, nGt - 1, 1 ) == "/"
   RETURN hb_BLeft( cXml, nGt - 2 ) + ">" + cRef + "</w:sectPr>" + hb_BSubStr( cXml, nGt + 1 )
ENDIF
nE := hb_BAt( "</w:sectPr>", cXml, nGt )
IF nE == 0
   RETURN cXml
ENDIF
cSect := hb_BSubStr( cXml, nGt + 1, nE - nGt - 1 )
p := 1
DO WHILE ( p := hb_BAt( "<w:" + cTag, cSect, p ) ) > 0
   q := hb_BAt( ">", cSect, p )
   IF 'w:type="default"' $ hb_BSubStr( cSect, p, q - p + 1 )
      cSect := hb_BLeft( cSect, p - 1 ) + hb_BSubStr( cSect, q + 1 )
   ELSE
      p := q
   ENDIF
ENDDO
RETURN hb_BLeft( cXml, nGt ) + cRef + cSect + hb_BSubStr( cXml, nE )

// najveci broj posle cPrefix (npr. 'w:numId="'), 0 ako ga nema
STATIC FUNCTION _MaxNum( cXml, cPrefix )
LOCAL n := 0, p := 1
DO WHILE ( p := hb_BAt( cPrefix, cXml, p ) ) > 0
   p += hb_BLen( cPrefix )
   n := Max( n, Val( hb_BSubStr( cXml, p, 12 ) ) )
ENDDO
RETURN n

// ---------------------------------------------------------------------------
// Pomocne funkcije - citanje i zamena teksta
// ---------------------------------------------------------------------------

// Tokeni XML-a (po "<"), cvorovi teksta <w:t> i granice pasusa.
// Vraca { aTok, aNodes, aPos }: aNodes[i] := { nTok, cTekst, nSegment },
// aPos[nTok] = bajt pozicija znaka "<" ispred tokena.
// Sve na nivou bajtova (hb_B*), nezavisno od kodne strane aplikacije.
STATIC FUNCTION _TScan( cXml )
LOCAL aTok := hb_ATokens( cXml, "<" ), aNodes := {}, aPos := Array( Len( aTok ) ), nSeg := 1, i, c, c4, n, nPos
nPos := hb_BLen( aTok[ 1 ] ) + 1
aPos[ 1 ] := 0
FOR i := 2 TO Len( aTok )
   aPos[ i ] := nPos
   c := aTok[ i ]
   c4 := hb_BLeft( c, 4 )
   IF c4 == "w:t>" .OR. c4 == "w:t "
      n := hb_BAt( ">", c )
      IF n > 0 .AND. !( hb_BSubStr( c, n - 1, 1 ) == "/" )
         AAdd( aNodes, { i, _XEnt( hb_BSubStr( c, n + 1 ) ), nSeg } )
      ENDIF
   ELSEIF c4 == "w:p>" .OR. c4 == "w:p " .OR. c4 == "w:p/" .OR. hb_BLeft( c, 5 ) == "/w:p>"
      nSeg++
   ENDIF
   nPos += hb_BLen( c ) + 1
NEXT
RETURN { aTok, aNodes, aPos }

STATIC FUNCTION _PlainText( cXml )
LOCAL c := "", a
FOR EACH a IN _TScan( cXml )[ 2 ]
   c += a[ 2 ]
NEXT
RETURN c

// zamena parova { cTrazi, cZameni } u tekstu pasusa, i kada je tekst podeljen
// na vise <w:t>; menjaju se samo pogodjeni cvorovi
STATIC FUNCTION _ReplaceXml( cXml, aPairs, nCount )
LOCAL aScan, aTok, aNodes, aIdx, aTxt, aDirty, cFull, aPair, i, k, nSeg, p, nFrom, nLen, nOff, nL, nI, nJ, nOffI, nOffJ, lChg := .F.
IF ! HB_ISNUMERIC( nCount )
   nCount := 0
ENDIF
IF Empty( aPairs ) .OR. !( "<w:t" $ cXml )
   RETURN cXml
ENDIF
aScan  := _TScan( cXml )
aTok   := aScan[ 1 ]
aNodes := aScan[ 2 ]
i := 1
DO WHILE i <= Len( aNodes )
   nSeg := aNodes[ i, 3 ]
   aIdx := {}
   aTxt := {}
   DO WHILE i <= Len( aNodes ) .AND. aNodes[ i, 3 ] == nSeg
      AAdd( aIdx, aNodes[ i, 1 ] )
      AAdd( aTxt, aNodes[ i, 2 ] )
      i++
   ENDDO
   aDirty := AFill( Array( Len( aIdx ) ), .F. )
   FOR EACH aPair IN aPairs
      nFrom := 1
      nLen := hb_BLen( aPair[ 1 ] )
      DO WHILE .T.
         cFull := _Join( aTxt, "" )
         IF ( p := hb_BAt( aPair[ 1 ], cFull, nFrom ) ) == 0
            EXIT
         ENDIF
         // cvorovi u kojima su prvi (nI) i poslednji (nJ) znak pogotka
         nI := nJ := 0
         nOff := 1
         FOR k := 1 TO Len( aTxt )
            nL := hb_BLen( aTxt[ k ] )
            IF nI == 0 .AND. p < nOff + nL
               nI := k
               nOffI := nOff
            ENDIF
            IF nI > 0 .AND. p + nLen - 1 < nOff + nL
               nJ := k
               nOffJ := nOff
               EXIT
            ENDIF
            nOff += nL
         NEXT
         IF nI == nJ
            aTxt[ nI ] := hb_BLeft( aTxt[ nI ], p - nOffI ) + aPair[ 2 ] + hb_BSubStr( aTxt[ nI ], p - nOffI + nLen + 1 )
         ELSE
            aTxt[ nI ] := hb_BLeft( aTxt[ nI ], p - nOffI ) + aPair[ 2 ]
            FOR k := nI + 1 TO nJ - 1
               aTxt[ k ] := ""
            NEXT
            aTxt[ nJ ] := hb_BSubStr( aTxt[ nJ ], p + nLen - nOffJ + 1 )
         ENDIF
         FOR k := nI TO nJ
            aDirty[ k ] := .T.
         NEXT
         nCount++
         nFrom := p + hb_BLen( aPair[ 2 ] )
      ENDDO
   NEXT
   FOR k := 1 TO Len( aIdx )
      IF aDirty[ k ]
         aTok[ aIdx[ k ] ] := 'w:t xml:space="preserve">' + _TEsc( aTxt[ k ] )
         lChg := .T.
      ENDIF
   NEXT
ENDDO
RETURN IIF( lChg, _Join( aTok, "<" ), cXml )

// tekst za <w:t>; novi red i tab zatvaraju <w:t> i ubacuju <w:br/> / <w:tab/>
STATIC FUNCTION _TEsc( c )
c := XmlEsc( StrTran( StrTran( c, Chr( 13 ) + Chr( 10 ), Chr( 10 ) ), Chr( 13 ), Chr( 10 ) ) )
c := StrTran( c, Chr( 10 ), '</w:t><w:br/><w:t xml:space="preserve">' )
RETURN StrTran( c, Chr( 9 ), '</w:t><w:tab/><w:t xml:space="preserve">' )

// hash ili niz (sa aKeys) -> { { cTrazi, cZameni }, ... } u UTF-8
STATIC FUNCTION _Pairs( oDoc, x, aKeys )
LOCAL aPairs := {}, k, i
IF HB_ISHASH( x )
   FOR EACH k IN hb_HKeys( x )
      IF HB_ISSTRING( k ) .AND. ! Empty( k )
         AAdd( aPairs, { oDoc:ToUTF8( k ), oDoc:ToUTF8( _Str( x[ k ] ) ) } )
      ENDIF
   NEXT
ELSEIF HB_ISARRAY( x ) .AND. HB_ISARRAY( aKeys )
   FOR i := 1 TO Min( Len( x ), Len( aKeys ) )
      IF HB_ISSTRING( aKeys[ i ] ) .AND. ! Empty( aKeys[ i ] )
         AAdd( aPairs, { oDoc:ToUTF8( aKeys[ i ] ), oDoc:ToUTF8( _Str( x[ i ] ) ) } )
      ENDIF
   NEXT
ENDIF
RETURN aPairs

// stilovi iz styles.xml: { ime (malim slovima) -> id, id -> ime, id (malim) -> id }
STATIC FUNCTION _StyleMap( cXml )
LOCAL hN := { => }, hI := { => }, hL := { => }, cId := "", cTok, cA, cTag, nKind, c
FOR EACH cTok IN hb_ATokens( cXml, "<" )
   cA := _XTok( cTok, @cTag, @nKind )
   IF cTag == "style"
      IF nKind == 2
         cId := ""
      ELSE
         cId := _HAttr( _XAttrs( cA ), "styleId" )
         hL[ Lower( cId ) ] := cId
      ENDIF
   ELSEIF cTag == "name" .AND. nKind != 2 .AND. ! Empty( cId )
      c := _HAttr( _XAttrs( cA ), "val" )
      hI[ cId ] := c
      IF ! hb_HHasKey( hN, Lower( c ) )
         hN[ Lower( c ) ] := cId
      ENDIF
   ENDIF
NEXT
RETURN { hN, hI, hL }

// telo dokumenta: { { "p", cTekst, cStil }, { "t", aRedovi }, ... }
STATIC FUNCTION _ParseBody( cXml, hNames )
LOCAL aItems := {}, aStack := {}, nTbl := 0, nSkip := 0, lTabs := .F.
LOCAL aTbl, aRow, cCell := "", lFirst := .T., nSpan := 1, cStyle := "", cTok, cA, cTag, nKind, cText, c, i
FOR EACH cTok IN hb_ATokens( cXml, "<" )
   cA := _XTok( cTok, @cTag, @nKind, @cText )
   // mc:Fallback ponavlja sadrzaj iz mc:Choice (stari format okvira za tekst)
   IF nSkip > 0
      IF cTag == "Fallback"
         nSkip += IIF( nKind == 1, 1, IIF( nKind == 2, -1, 0 ) )
      ENDIF
      LOOP
   ENDIF
   DO CASE
   CASE cTag == "Fallback"
      IF nKind == 1
         nSkip := 1
      ENDIF
   CASE cTag == "tbl"
      IF nKind == 1
         IF ++nTbl == 1
            aTbl := {}
         ENDIF
      ELSEIF nKind == 2
         IF nTbl == 1
            AAdd( aItems, { "t", aTbl } )
         ENDIF
         nTbl := Max( 0, nTbl - 1 )
      ENDIF
   CASE cTag == "tr"
      IF nTbl == 1
         IF nKind == 1
            aRow := {}
         ELSEIF nKind == 2
            AAdd( aTbl, aRow )
         ENDIF
      ENDIF
   CASE cTag == "tc"
      IF nTbl == 1
         IF nKind == 1
            cCell := ""
            lFirst := .T.
            nSpan := 1
         ELSEIF nKind == 2
            DO WHILE Right( cCell, 1 ) == Chr( 10 )     // prazan pasus posle ugnezdene tabele
               cCell := hb_StrShrink( cCell, 1 )
            ENDDO
            AAdd( aRow, cCell )
            FOR i := 2 TO nSpan
               AAdd( aRow, "" )
            NEXT
         ENDIF
      ENDIF
   CASE cTag == "gridSpan"
      IF nTbl == 1 .AND. nKind != 2
         nSpan := Max( 1, Val( _HAttr( _XAttrs( cA ), "val" ) ) )
      ENDIF
   CASE cTag == "p"
      IF nKind == 1 .OR. nKind == 3
         AAdd( aStack, "" )
         IF Len( aStack ) == 1
            cStyle := ""
         ENDIF
      ENDIF
      IF ( nKind == 2 .OR. nKind == 3 ) .AND. ! Empty( aStack )
         c := ATail( aStack )
         ASize( aStack, Len( aStack ) - 1 )
         IF ! Empty( aStack )                 // pasus u okviru za tekst
            aStack[ Len( aStack ) ] += IIF( Empty( c ), "", c + " " )
         ELSEIF nTbl == 0
            AAdd( aItems, { "p", c, cStyle } )
         ELSE
            cCell += IIF( lFirst, "", Chr( 10 ) ) + c
            lFirst := .F.
         ENDIF
      ENDIF
   CASE cTag == "pStyle"
      IF Len( aStack ) == 1 .AND. nTbl == 0 .AND. nKind != 2
         c := _HAttr( _XAttrs( cA ), "val" )
         cStyle := hb_HGetDef( hNames, c, c )
      ENDIF
   CASE cTag == "tabs"
      lTabs := ( nKind == 1 )
   CASE Empty( aStack )
   CASE cTag == "t" .AND. nKind == 1
      aStack[ Len( aStack ) ] += _XEnt( cText )
   CASE cTag == "tab" .AND. nKind != 2 .AND. ! lTabs
      aStack[ Len( aStack ) ] += Chr( 9 )
   CASE ( cTag == "br" .OR. cTag == "cr" ) .AND. nKind != 2
      aStack[ Len( aStack ) ] += Chr( 10 )
   CASE cTag == "noBreakHyphen" .AND. nKind != 2
      aStack[ Len( aStack ) ] += "-"
   ENDCASE
NEXT
RETURN aItems

// ---------------------------------------------------------------------------
// Pomocne funkcije - opste (iste kao u XLSXClass.prg)
// ---------------------------------------------------------------------------

// Escape za XML tekst/atribut (izbacuje i nedozvoljene kontrolne znakove).
// '&' mora prvi, inace bi se dupli-escape-ovali &lt; &gt; &quot;.
STATIC FUNCTION XmlEsc( c )
c := XmlClean( c )
c := StrTran( c, "&", "&amp;" )
c := StrTran( c, "<", "&lt;" )
c := StrTran( c, ">", "&gt;" )
c := StrTran( c, '"', "&quot;" )
RETURN c

// XML 1.0 ne dozvoljava znakove 0-31 osim TAB, LF i CR (ostaci chr(0)/chr(26) u DBF poljima)
STATIC FUNCTION XmlClean( c )
STATIC s_cCtrl
LOCAL i
IF s_cCtrl == NIL
   s_cCtrl := ""
   FOR i := 0 TO 31
      IF i != 9 .AND. i != 10 .AND. i != 13
         s_cCtrl += Chr( i )
      ENDIF
   NEXT
ENDIF
RETURN hb_StrReplace( c, s_cCtrl, "" )

// vrednost kao tekst: broj, datum, logicka...
STATIC FUNCTION _Str( x )
DO CASE
CASE HB_ISSTRING( x )
   RETURN x
CASE HB_ISNIL( x )
   RETURN ""
CASE HB_ISNUMERIC( x )
   RETURN hb_NToS( x )
CASE HB_ISDATE( x )
   RETURN IIF( Empty( x ), "", DToC( x ) )
CASE HB_ISDATETIME( x )
   RETURN IIF( Empty( x ), "", hb_TToC( x ) )
ENDCASE
RETURN hb_CStr( x )

STATIC FUNCTION _AGet( a, n )
RETURN IIF( HB_ISARRAY( a ) .AND. n >= 1 .AND. n <= Len( a ), a[ n ], NIL )

STATIC FUNCTION _RC( r, c )
RETURN hb_NToS( r ) + "," + hb_NToS( c )

// spaja niz stringova; deljenjem na pola, jer bi += u petlji za velike
// dokumente kopirao ceo tekst pri svakom dodavanju
STATIC FUNCTION _Join( a, cSep, nFrom, nTo )
LOCAL c, i, nMid
IF nFrom == NIL
   nFrom := 1
   nTo := Len( a )
ENDIF
IF nTo < nFrom
   RETURN ""
ENDIF
IF nTo - nFrom < 64
   c := a[ nFrom ]
   FOR i := nFrom + 1 TO nTo
      c += cSep + a[ i ]
   NEXT
   RETURN c
ENDIF
nMid := Int( ( nFrom + nTo ) / 2 )
RETURN _Join( a, cSep, nFrom, nMid ) + cSep + _Join( a, cSep, nMid + 1, nTo )

// Jedan deo XML-a: vraca tekst atributa, a u cName ime elementa (bez prefiksa
// "w:"), u nKind 1 <a>, 2 </a>, 3 <a/>, 0 ostalo (<?xml?>, komentar), u cText
// tekst posle elementa (do sledeceg "<", nedekodiran).
STATIC FUNCTION _XTok( cTok, cName, nKind, cText )
LOCAL n := At( ">", cTok ), cHead, ch
cName := ""
nKind := 0
cText := ""
IF n == 0
   RETURN ""
ENDIF
// ">" sme da stoji u vrednosti atributa: trazi se prvi van navodnika
IF '"' $ cTok
   DO WHILE n > 0 .AND. ( Len( Left( cTok, n ) ) - Len( StrTran( Left( cTok, n ), '"', "" ) ) ) % 2 == 1
      n := hb_At( ">", cTok, n + 1 )
   ENDDO
   IF n == 0
      RETURN ""
   ENDIF
ENDIF
cHead := Left( cTok, n - 1 )
cText := SubStr( cTok, n + 1 )
ch := Left( cHead, 1 )
IF ch == "?" .OR. ch == "!"
   RETURN ""
ENDIF
IF ch == "/"
   nKind := 2
   cHead := SubStr( cHead, 2 )
ELSEIF Right( cHead, 1 ) == "/"
   nKind := 3
   cHead := Left( cHead, Len( cHead ) - 1 )
ELSE
   nKind := 1
ENDIF
cHead := hb_StrReplace( cHead, Chr( 9 ) + Chr( 10 ) + Chr( 13 ), "   " )
n := At( " ", cHead )
cName := IIF( n == 0, cHead, Left( cHead, n - 1 ) )
IF ( ch := At( ":", cName ) ) > 0
   cName := SubStr( cName, ch + 1 )
ENDIF
RETURN IIF( n == 0, "", SubStr( cHead, n + 1 ) )

// 'a="1" b="x &amp; y"' -> { "a" => "1", "b" => "x & y" }
STATIC FUNCTION _XAttrs( cA )
LOCAL h := { => }, n, m, q, cKey
DO WHILE ( n := At( "=", cA ) ) > 0
   cKey := AllTrim( Left( cA, n - 1 ) )
   cA := LTrim( SubStr( cA, n + 1 ) )
   q := Left( cA, 1 )
   IF !( q == '"' .OR. q == "'" ) .OR. ( m := hb_At( q, cA, 2 ) ) == 0
      EXIT
   ENDIF
   h[ cKey ] := _XEnt( SubStr( cA, 2, m - 2 ) )
   cA := SubStr( cA, m + 1 )
ENDDO
RETURN h

// atribut iz hash-a ili ""; "val" nalazi i "w:val"
STATIC FUNCTION _HAttr( h, cName )
LOCAL k
IF hb_HHasKey( h, cName )
   RETURN h[ cName ]
ENDIF
FOR EACH k IN hb_HKeys( h )
   IF Right( k, Len( cName ) + 1 ) == ":" + cName
      RETURN h[ k ]
   ENDIF
NEXT
RETURN ""

// XML entiteti: &amp; &lt; &gt; &quot; &apos; &#NN; &#xHH;
STATIC FUNCTION _XEnt( c )
LOCAL a, cOut, i, n, cE
IF ! "&" $ c
   RETURN c
ENDIF
a := hb_ATokens( c, "&" )
cOut := a[ 1 ]
FOR i := 2 TO Len( a )
   n := At( ";", a[ i ] )
   cE := IIF( n > 1 .AND. n <= 10, Left( a[ i ], n - 1 ), "" )
   DO CASE
   CASE cE == "amp"  ; cOut += "&"
   CASE cE == "lt"   ; cOut += "<"
   CASE cE == "gt"   ; cOut += ">"
   CASE cE == "quot" ; cOut += '"'
   CASE cE == "apos" ; cOut += "'"
   CASE Lower( Left( cE, 2 ) ) == "#x" ; cOut += hb_UTF8Chr( hb_HexToNum( SubStr( cE, 3 ) ) )
   CASE Left( cE, 1 ) == "#" .AND. Len( cE ) > 1 ; cOut += hb_UTF8Chr( Val( SubStr( cE, 2 ) ) )
   OTHERWISE
      cOut += "&" + a[ i ]
      LOOP
   ENDCASE
   cOut += SubStr( a[ i ], n + 1 )
NEXT
RETURN cOut
