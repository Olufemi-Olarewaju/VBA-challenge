Sub Stock()

For Each ws In Worksheets

Dim WorksheetName As String
WorksheetName = ws.Name
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim RowNum As Integer
RowNum = 2
Dim Opening_Price As Double
Dim Closing_Price As Double
Dim Yearly_Change As Double
Dim Total As Double
Total = 0

Opening_Price = ws.Cells(2, 3)
Total = 0
ws.Range("I1").EntireColumn.Insert
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Stock Volume"

For i = 2 To lastrow

    If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    ws.Range("I" & RowNum) = ws.Cells(i, 1)
    Closing_Price = ws.Cells(i, 6)
    Yearly_Change = Closing_Price - Opening_Price
    Total = Total + ws.Cells(i, 7)
    ws.Range("K" & RowNum) = Yearly_Change / Opening_Price
    ws.Range("K" & RowNum).NumberFormat = "0.00%"
    ws.Range("J" & RowNum) = Yearly_Change
    ws.Range("L" & RowNum) = Total + ws.Cells(i, 7)
    RowNum = RowNum + 1
    Opening_Price = ws.Cells(i + 1, 3)


    Total = 0
    
    Else
    
    Total = Total + ws.Cells(i, 7)
    
    End If
    
    Next i
    
    For x = 2 To lastrow
    If ws.Cells(x, 11) > 0 Then
    ws.Cells(x, 11).Interior.ColorIndex = 4
    Else
    ws.Cells(x, 11).Interior.ColorIndex = 3
    End If
    Next x

    
    ws.Cells(1, 16) = "Ticker"
    ws.Cells(1, 17) = "Value"
    ws.Cells(2, 15) = "Greatest % Increase"
    ws.Cells(3, 15) = "Greatest % Decrease"
    ws.Cells(4, 15) = "Greatest Total Volume"
    
    For j = 2 To lastrow
    
    If ws.Cells(j, 11) = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) Then
    Dim Max_Increase As Double
    Dim Max_Stock As String
    Max_Increase = ws.Cells(j, 11)
    Max_Stock = ws.Cells(j, 9)
    ws.Range("P2") = Max_Stock
    ws.Range("Q2") = Max_Increase
    ws.Range("Q2").NumberFormat = "0.00%"
    End If
    
    If ws.Cells(j, 11) = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) Then
    Dim Min_Increase As Double
    Dim Min_Stock As String
    Min_Increase = ws.Cells(j, 11)
    Min_Stock = ws.Cells(j, 9)
    ws.Range("P3") = Min_Stock
    ws.Range("Q3") = Min_Increase
    ws.Range("Q3").NumberFormat = "0.00%"
    End If
    
    If ws.Cells(j, 12) = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) Then
    Dim Max_Total As Double
    Dim Max_Total_Stock As String
    Max_Total = ws.Cells(j, 12)
    Max_Total_Stock = ws.Cells(j, 9)
    ws.Range("P4") = Max_Total_Stock
    ws.Range("Q4") = Max_Total
    
    End If
    
    Next j
    
        


'Exit For
Next ws

End Sub



