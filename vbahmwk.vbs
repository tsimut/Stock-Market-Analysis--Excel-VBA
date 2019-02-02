Sub WorksheetLoop():

Dim WS_Count As Integer
Dim I As Integer
WS_Count = ActiveWorkbook.Worksheets.Count
For I = 1 To WS_Count

Dim Ticker As String
Dim Stock_Total As Double
Dim K As Double
Dim Lastrow As Long
Lastrow = ActiveWorkbook.Worksheets(I).Cells(Rows.Count, 1).End(xlUp).Row

Stock_Total = 0
Summary_row = 2
For K = 2 To Lastrow
If ActiveWorkbook.Worksheets(I).Cells(K + 1, 1).Value <> ActiveWorkbook.Worksheets(I).Cells(K, 1).Value Then
Ticker = ActiveWorkbook.Worksheets(I).Cells(K, 1).Value
Stock_Total = Stock_Total + ActiveWorkbook.Worksheets(I).Cells(K, 7).Value
ActiveWorkbook.Worksheets(I).Range("I" & Summary_row) = Ticker
ActiveWorkbook.Worksheets(I).Range("J" & Summary_row) = Stock_Total
Summary_row = Summary_row + 1
Stock_Total = 0
Else
Stock_Total = Stock_Total + ActiveWorkbook.Worksheets(I).Cells(K, 7).Value
End If
Next K

Next I


End Sub


