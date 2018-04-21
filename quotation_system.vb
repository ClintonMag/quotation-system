REM  *****  BASIC  *****

Option Explicit

Const QROW As Integer = 6
Const QCOL As Integer = 5
'Number of digits in a quotation number
Const QNUM_LENGTH = 7
Dim Doc As Object
Dim AllSheets As Object
Dim QNum As Long

Function Cells(sheet As Object, x As Integer, y As Integer) As Object
	REM Simplify calls to retrieve cell position by coordinates
	
	Cells = sheet.getCellByPosition(x,y)
End Function

Sub Main
	REM Read quotation/invoice number from a 'QINUM' sheet
	REM Update it for the current quotation/invoice
	REM Place new quotation/invoice number in the correct cell
	REM Write new quotation/invoice, date, and client name to the
	REM 'QINUM.csv' file
	
	InitializeGlobals
	GenerateQuotationNumber
	NextQuoteNumber

End Sub
	

Sub InitializeGlobals
	REM Initializes the global variables
	
	'The current spreadsheet document
	Doc = ThisComponent
	
	'All sheets contained by the current spreadsheet document
	AllSheets = Doc.Sheets
	
	'Sheet used for making quotations
	QSheet = AllSheets.getByName("Quotation")
	
	'Cell containing the quotation number
	QCel = Cells(QSheet,QROW,QCOL)
	
End Sub

Sub GenerateQuotationNumber
	REM Read the last quotation number

	QNum = Val(Right(qCel.String, QNUM_LENGTH))

End Sub

Sub NextQuoteNumber
	REM Calculate the next quotion number
	
	QCel = Cells(QSheet,QROW,QCOL)
	'Extract integer portion of quotation number, convert to int.
	qCel.Value = qNum + 50
	
End Sub