// Reading a .docx: paragraphs with their styles, tables, all text.
// Build: hbmk2 docxreader.hbp   (run docxwriter and docxtemplate first)
#include "hbclass.ch"
REQUEST HB_GT_WIN_DEFAULT
REQUEST HB_CODEPAGE_UTF8EX

PROCEDURE Main( cFile )
LOCAL oDoc, aPara, aTable, aRow

hb_cdpSelect( "UTF8EX" )
hb_SetTermCP( "UTF8EX" )

oDoc := WordDocument():New()
IF ! oDoc:Read( IIF( Empty( cFile ), "invoice.docx", cFile ) )
   ? oDoc:cError
   RETURN
ENDIF

? "Paragraphs:"
FOR EACH aPara IN oDoc:Paragraphs( .T. )                 // { cText, cStyle }
   IF ! Empty( aPara[ 1 ] )
      ? " ", PadR( aPara[ 2 ], 16 ), StrTran( aPara[ 1 ], Chr( 10 ), " / " )
   ENDIF
NEXT

FOR EACH aTable IN oDoc:Tables()
   ?
   ? "Table", hb_ntos( aTable:__enumIndex() ), "-", hb_ntos( Len( aTable ) ), "rows"
   FOR EACH aRow IN aTable
      ? " ", hb_ValToExp( aRow )
   NEXT
NEXT

RETURN
