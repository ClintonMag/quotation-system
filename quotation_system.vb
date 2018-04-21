REM  *****  BASIC  *****

REM QuotationSystem

REM Read quotation/invoice number from the 'QINUM' sheet
REM Update it for the current quotation/invoice
REM Place new quotation/invoice number in the correct cell
REM Write new quotation/invoice, date, and client name to the
REM 'QINUM.csv' file

Option Explicit

' ***** GLOBAL VARIABLES ***** '

'Row and Column position of the quotation/invoice number in the sheets where
'they will be used in the quotation/invoice.
'Row 1 in sheet = row 1 in this code
'Col A in sheet = col 1 in this code
Const QIROW As Integer = 7
Const QICOL As Integer = 6

'Rows and Column of last used quotation and invoice number. invoice number
'must be one row below quotation number
Const QROW As Integer = 1
Const QCOL As Integer = 1
Const IROW As Integer = 2

'Sheet in which quotations are made
Const QUOTATION_SHEET As String = "Quotation"

'Sheet in which last used quotation and invoice numbers are stored
Const QI_NUMBER_SOURCE_SHEET As String = "QINUM"

'Number of digits in a quotation/invoice number
Const QINUM_LENGTH = 7

'The current spreadsheet document
Dim Doc As Object

'All sheets contained by the current spreadsheet document
Dim AllSheets As Object

'The applicable working sheet, either 'Quotation' or 'Invoice'
Dim QISheet As Object

'Store the full quotation number as a string
Dim QIString As String

'Index of the sheet named 'Quotation' in the ThisComponent.Sheets list
Dim QSheetIndex as Integer


' ***** FUNCTIONS ***** '

Function Cells(sheet As Object, x As Integer, y As Integer) As Object
	REM Simplify calls to retrieve cell position by coordinates
	
	Cells = sheet.getCellByPosition(y-1, x-1)
End Function


' ***** SUBROUTINES ***** '
	
Sub MainQuotation
	REM Entry point if working on a quotation
	
	'Indicate whether we are working on a quotation or an invoice
	'0 for quotation, 1 for invoice
	Dim QuoteOrInv as Integer
	
	QuoteOrInv = 0
	MainProcess(QuoteOrInv)

End Sub

Sub MainInvoice
	REM Entry point if working on an invoice

	'Indicate whether we are working on a quotation or an invoice
	'0 for quotation, 1 for invoice
	Dim QuoteOrInv as Integer
	
	QuoteOrInv = 1
	MainProcess(QuoteOrInv)

End Sub


Sub MainProcess(QuoteOrInv)
	REM Start the main invoice/quotation nunmber update procedure once it's
	REM which is to be updated
	
	InitializeGlobals(QuoteOrInv)
	GenerateQINumber(QuoteOrInv)
	WriteQINumber(QuoteOrInv)

End Sub

Sub InitializeGlobals(QuoteOrInv As Integer)
	REM Initializes the global variables
	
	Doc = ThisComponent
	
	AllSheets = Doc.Sheets
	
	QSheetIndex = AllSheets.getByName(QUOTATION_SHEET).RangeAddress.Sheet
	
	' If working on quotation, QSheetIndex + QoutOrInv will be the "Quotation"
	' sheet. If working on invoice,vQSheetIndex + QoutOrInv will be the
	' "Invoice" sheet
	QISheet = AllSheets.getByIndex(QSheetIndex + QuoteOrInv)
	
	QIString = Cells(AllSheets.getByName(QI_NUMBER_SOURCE_SHEET), _
					 QROW + QuoteOrInv, QCOL).String
	
End Sub

Sub GenerateQINumber(QuoteOrInv As Integer)
	REM Read the last quotation/invoice number, increase it by 1 and place it
	REM in the correct cell
	
	Dim i as Integer
	
	'String format to be used by the Format function to specify number of zeros
	'for left-padding
	Dim QINumFormatString As String
	
	'Quotation/Invoice Number
	Dim QINum As Long
	
	'New Quotation/Invoice Number
	Dim NewQINum as Long
	
	'Cell where last used quotation/invoice number is stored
	Dim QISourceCell As Object
	
	'Cell to receive the new quotation or invoice number in the QUOTATION_SHEET
	'or the QI_NUMBER_SOURCE_SHEET
	Dim QICell as Object

	'Extract quotation/invoice integer portion, increment by 1.
	QINum = Val(Right(QIString, QINUM_LENGTH))
	NewQINum = QINum + 1
	
	QISourceCell = Cells(AllSheets.getByName(QI_NUMBER_SOURCE_SHEET), _
						 QROW + QuoteOrInv, QCOL)
	
	QICell = Cells(QISheet, QIROW, QICOL)
	
	QINumFormatString = "0"
	'Compute format string. E.g. if quotation number is 7 digits, format string
	'will be 7 zeros: 0000000
	For i = 1 to QINUM_LENGTH-1
		QINumFormatString = QINumFormatString + "0"
	Next i
	
	'Place new value of quotation/invoice number in QICell
	'Place new value of quotation/invoice number in QI_NUMBER_SOURCE_SHEET
	If QuoteOrInv = 0 Then
		QICell.String = "Q" + Format(NewQINum, QINumFormatString)
		QISourceCell.String = "Q" + Format(NewQINum, QINumFormatString)
	Else
		QICell.String = "I" + Format(NewQINum, QINumFormatString)
		QISourceCell.String = "I" + Format(NewQINum, QINumFormatString)
	End If

End Sub

Sub WriteQINumber(QuoteOrInv As Integer)

	

End Sub


Sub Test

	Doc = ThisComponent
	AllSheets = Doc.Sheets
	
	Dim MySheet As Object
	
	MySheet = AllSheets.getByName("QINUM")
	
	MySheet.getCellByPosition(0,0).String = "Yes " + "And No"

End Sub
