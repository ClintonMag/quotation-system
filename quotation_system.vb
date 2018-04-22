REM  *****  BASIC  *****


REM QuotationSystem

REM Read quotation/invoice number from the 'QINUM' sheet
REM Update it for the current quotation/invoice
REM Place new quotation/invoice number in the correct cell
REM Write new quotation/invoice, date, and client name to the
REM 'QINUM.csv' file

' ** IMPORTANT ** '
' The sheet 'Invoice' must immediately follow the sheet 'Quotation'
' The cell containing the last invoice number must be immediately below
' the cell containing the last quotation number.

Option Explicit

' ***** GLOBAL VARIABLES ***** '

'Represent the row and column of any cell, where the cell is represented by a
'1D array with row at index 0 and column at index 1.
'Every (row,col) array name ends in XY
Const ROW As Integer = 0
Const COL As Integer = 1

'Row and Column position of the quotation/invoice number in the sheets where
'they will be used in the quotation/invoice.
'Row 1 in sheet = row 1 in this code
'Col A in sheet = col 1 in this code
Dim QIRecipientXY(2) As Integer
Const QI_REC_ROW As Integer = 7
Const QI_REC_COL As Integer = 6

'Rows and Column of last used quotation and invoice number. invoice number
'must be one row below quotation number
Dim QISourceXY(2) As Integer
Const Q_SRC_ROW As Integer = 1
Const Q_SRC_COL As Integer = 1

'Location in Quotation/Invoice Sheet of the client name
Dim ClientXY(2) As Integer
Const CLIENT_ROW As Integer = 6
Const CLIENT_COL As Integer = 1

'Sheet in which quotations are made
Const QUOTATION_SHEET As String = "Quotation"

'Sheet in which last used quotation and invoice numbers are stored
'This sheet is hidden and protected
Const QI_NUMBER_SOURCE_SHEET As String = "QINUM"

'Number of digits in a quotation/invoice number
Const QINUM_LENGTH = 7

'The name of the ods document for which this macro is meant
Const PARENT_FILE As String = "Stock_And_Quotes.ods"

'The csv file where list of quotation numbers used will be stored
'It's a csv for now, database to be used in future
Const CSVFILE = "QINUM.csv"

'The current spreadsheet document
Dim Doc As Object

'All sheets contained by the current spreadsheet document
Dim AllSheets As Object

'The applicable working sheet, either 'Quotation' or 'Invoice'
Dim QISheet As Object

'Store the full quotation number as a string
Dim QIString As String

'New quotation number as a string
Dim NewQIString As String

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
	
	'Saving row-col values of 3 cells in their respective arrays
	
	QIRecipientXY(ROW) = QI_REC_ROW
	QIRecipientXY(COL) = QI_REC_COL
	
	QISourceXY(ROW) = Q_SRC_ROW
	QISourceXY(COL) = Q_SRC_COL
	
	ClientXY(ROW) = CLIENT_ROW
	ClientXY(COL) = CLIENT_COL
	
	QIString = Cells(AllSheets.getByName(QI_NUMBER_SOURCE_SHEET), _
					 QISourceXY(ROW) + QuoteOrInv, QISourceXY(COL)).String
	
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
						 QISourceXY(ROW) + QuoteOrInv, QISourceXY(COL))
	
	QICell = Cells(QISheet, QIRecipientXY(ROW), QIRecipientXY(COL))
	
	QINumFormatString = "0"
	'Compute format string. E.g. if quotation number is 7 digits, format string
	'will be 7 zeros: 0000000
	For i = 1 to QINUM_LENGTH-1
		QINumFormatString = QINumFormatString & "0"
	Next i
	
	If QuoteOrInv = 0 Then
		NewQIString = "Q"
	Else
		NewQIString = "I"
	End If
	
	NewQIString = NewQIString & Format(NewQINum, QINumFormatString)
	'Place new value of quotation/invoice number in QICell
	QICell.String = NewQIString
	'Place new value of quotation/invoice number in QI_NUMBER_SOURCE_SHEET
	QISourceCell.String = NewQIString

End Sub

Sub WriteQINumber(QuoteOrInv As Integer)
	REM Write an entry into the QINUM.csv file for each new quotation
	REM or invoice written.
	REM Format of one row of QINUM.csv:
	REM Quotation/Invoice Number,QuoteOrInv Boolean value, Date, Name of client

	'The string to be written to the QINUM.csv file
	Dim DataEntry As String
	'Client Name
	Dim Client As String
	'The path of 'Stock_And_Quotes.ods', the document this macro is for.
	Dim DocPath As String
	'Index of first appearance of parent file name in it's directory string
	Dim Index As Integer
	'File path of the csv file in url format
	Dim CSVFileURL As String
	'The file handle through which communication with the CSVFILE is done
	Dim FileNo As Integer
	'Directory to the folder in which PARENT_FILE is contained
	Dim ParentDir As String
	
	
	Client = Cells(QISheet, ClientXY(ROW), ClientXY(COL)).String
	' The keyword 'Date' used in this context refers to today's date
	DataEntry = NewQIString & "," & QuoteOrInv & "," & Date & "," & Client
	
	DocPath = Doc.getURL()
	Index = InStr(DocPath, PARENT_FILE)
	ParentDir = Mid(DocPath, 1, Index-1)
	CSVFileURL = ParentDir & CSVFILE
	
	'The DataEntry string will be written to the file specified by CSVFileURL
	
	'FreeFile function is used to get a free file handle	
	
	FileNo = FreeFile
	'Indentation used as file is free for use between Open and Close statements
	Open CSVFileURL For Append As #FileNo
		
		Print #FileNo, DataEntry
	
	Close #FileNo

End Sub


Sub Test
	'A subroutine to experiment with code to study it's behaviour.
	'Frequently replaced with new test code.

	Doc = ThisComponent
	AllSheets = Doc.Sheets
	
	Dim MySheet As Object
	
	MySheet = AllSheets.getByName("Quotation")
	

End Sub
