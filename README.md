[3/19, 08:49] Gokul Ct: Sub EstimateCashFlow()

Dim ws As Worksheet

Dim lastRow As Long

Dim lastColumn As Long

Dim startColumn As Long

Dim totalInflows As Double

Dim totalOutflows As Double

Dim netCashFlow As Double

Dim col As Integer

' Set the worksheet

Set ws = ThisWorkbook.Sheets("Sheet1")

' Find the last row in column A

lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

 

' Find the last column in row 1

lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

' Start column (assuming first column contains descriptions or dates)

startColumn = 2

' Initialize totals

totalInflows = 0

totalOutflows = 0
[3/19, 08:49] Gokul Ct: Loop through columns

For col = startColumn To lastColumn

If ws.Cells(1, col).Value Like "Inflow*" Then

totalInflows = totalInflows + WorksheetFunction.Sum(ws.Columns(col))

ElseIf ws.Cells(1, col).Value Like "Outflow*" Then

totalOutflows = totalOutflows + WorksheetFunction.Sum(ws.Columns(col))

End If

Next col

' Calculate net cash flow

netCashFlow = totalInflows - totalOutflows

' Display message box with results

MsgBox "Total Inflows: " & totalInflows & vbCrLf & _

"Total Outflows: " & totalOutflows & vbCrLf & _

"Net Cash Flow: " & netCashFlow, vbInformation, "Cash Flow Estimation"

End Sub
