REM  *****  BASIC  *****

REM QuotationSystem

REM Read quotation/invoice number from a 'QINUM' sheet
REM Update it for the current quotation/invoice
REM Place new quotation/invoice number in the correct cell
REM Write new quotation/invoice, date, and client name to the
REM 'QINUM.csv' file

Option Explicit

'Row and Column position of the quotation/invoice number in respective sheet
'Row 1 in sheet = row 1 in this code
'Col A in sheet = col 1 in this code
Const QIROW As Integer = 7
Const QICOL As Integer = 6

'Number of digits in a quotation/invoice number
Const QINUM_LENGTH = 7

'The current spreadsheet document
Dim Doc As Object

'All sheets contained by the current spreadsheet document
Dim AllSheets As Object

'The applicable working sheet, either 'Quotation' or 'Invoice'
Dim QISheet As Object

'Cell containing the quotation or invoice number
Dim QICell as Object

'Quotation/Invoice Number
Dim QINum As Long

'Store the full quotation number as a string
Dim QIString As String

'New Quotation/Invoice Number
Dim NewQINum as Long

'Indicate whether we are working on a quotation or an invoice
'0 for quotation, 1 for invoice
Dim QuoteOrInv as Integer

'Index of the sheet named 'Quotation' in the ThisComponent.Sheets list
Dim QSheetIndex as Integer


Function Cells(sheet As Object, x As Integer, y As Integer) As Object
	REM Simplify calls to retrieve cell position by coordinates
	
	Cells = sheet.getCellByPosition(y-1, x-1)
End Function
	
Sub MainQuotation
	REM Entry point if working on a quotation
	
	QuoteOrInv = 0
	MainProcess

End Sub

Sub MainInvoice
	REM Entry point if working on an invoice
	
	QuoteOrInv = 1
	MainProcess

End Sub


Sub MainProcess
	REM Start the main invoice/quotation nunmber update procedure once it's
	REM which is to be updated
	
	InitializeGlobals
	GenerateQINumber

End Sub

Sub InitializeGlobals
	REM Initializes the global variables
	
	Doc = ThisComponent
	
	AllSheets = Doc.Sheets
	
	QSheetIndex = AllSheets.getByName("Quotation").RangeAddress.Sheet
	
	' If working on quotation, QSheetIndex + QoutOrInv will be the "Quotation"
	' sheet. If working on invoice,vQSheetIndex + QoutOrInv will be the
	' "Invoice" sheet
	QISheet = AllSheets.getByIndex(QSheetIndex + QuoteOrInv)
	
	QICell = Cells(QISheet, QIROW, QICOL)
	QIString = Cells(AllSheets.getByName("QINUM"),1,1).String
	
End Sub

Sub GenerateQINumber
	REM Read the last quotation/invoice number, increase it by 1 and place it
	REM in the correct cell
	
	'String format to be used by the Format function to specify number of zeros
	'for left-padding
	Dim QINumFormatString As String
	Dim i as Integer
	
	QINumFormatString = "0"
	
	'Compute format string. E.g. if quotation number is 7 digits, format string
	'will be 7 zeros: 0000000
	For i = 1 to QINUM_LENGTH-1
		QINumFormatString = QINumFormatString + "0"
	Next i

	QINum = Val(Right(QIString, QINUM_LENGTH))
	NewQINum = QINum + 50
	
	If QuoteOrInv = 0 Then
		QICell.String = "Q" + Format(NewQINum, QINumFormatString)
	Else
		QICell.String = "I" + Str(NewQINum, QINumFormatString)
	End If

End Sub

Sub Test

	Doc = ThisComponent
	AllSheets = Doc.Sheets
	
	Dim MySheet As Object
	
	MySheet = AllSheets.getByName("QINUM")
	
	MySheet.getCellByPosition(0,0).String = "Yes " + "And No"

End Sub
