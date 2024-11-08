Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()

Dim ws As Worksheet
For Each ws In Worksheets


'headers
ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Quarterly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Stock Volume"
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"
ws.Cells(2, "O").Value = "Greatest % Increase"
ws.Cells(3, "O").Value = "Greatest % Decrease"
ws.Cells(4, "O").Value = "Greatest Total Volume"

Dim ticker As String
Dim lastrow As Long
Dim i As Long
Dim j As Long

'to find last row number
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).row

Dim Open_Price As Double
Dim Close_Price As Double
Dim Quarterly_Change As Double
Dim Percent_Change As Double

Dim row As Long
row = 2
j = 2

For i = 2 To lastrow
    If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
        ticker = ws.Cells(i, "A").Value
        ws.Cells(row, "I").Value = ticker

        Open_Price = ws.Cells(j, "C").Value
        
        Close_Price = ws.Cells(i, "F").Value
        
        Quarterly_Change = Close_Price - Open_Price
        ws.Cells(row, "J").Value = Quarterly_Change
        ws.Cells(row, "J").Style = "currency"
        
        
        'conditional formating for Quarterly Change
            If Quarterly_Change < 0 Then
                ws.Cells(row, "J").Interior.ColorIndex = 3
            ElseIf Quarterly_Change > 0 Then
                ws.Cells(row, "J").Interior.ColorIndex = 4
            End If
        
        'conditional formating for Percent Change, and stop /0 error
            If Open_Price <> 0 Then
                Percent_Change = ((Close_Price - Open_Price) / Open_Price)
                ws.Cells(row, "K").Value = Percent_Change
                ws.Cells(row, "K").NumberFormat = "0.00%"
            ElseIf Open_Price = 0 Then
                ws.Cells(row, "K").Value = 0
                ws.Cells(row, "K").NumberFormat = "0.00%"
            End If
        
        'total stock volume
        ws.Cells(row, "L").Value = WorksheetFunction.Sum(Range(ws.Cells(j, "G"), ws.Cells(i, "G")))
        
        row = row + 1
        j = i + 1
    End If
Next i

Dim lastrow2 As Long
Dim Increase_Percent As Double
Dim Decrease_Percent As Double
Dim Total_Volume As Double

Increase_Percent = ws.Cells(2, "K").Value
Decrease_Percent = ws.Cells(2, "K").Value
Total_Volume = ws.Cells(2, "L").Value

lastrow2 = ws.Cells(Rows.Count, "I").End(xlUp).row

For i = 2 To lastrow2
    If ws.Cells(i, "K").Value > Increase_Percent Then
        Increase_Percent = ws.Cells(i, "K").Value
        ws.Cells(2, "P").Value = ws.Cells(i, "I").Value
        ws.Cells(2, "Q").Value = ws.Cells(i, "K").Value
        ws.Cells(2, "Q").NumberFormat = "0.00%"
    Else
        Increase_Percent = Increase_Percent
    
    If ws.Cells(i, "K").Value < Decrease_Percent Then
        Decrease_Percent = ws.Cells(i, "K").Value
        ws.Cells(3, "P").Value = ws.Cells(i, "I").Value
        ws.Cells(3, "Q").Value = ws.Cells(i, "K").Value
        ws.Cells(3, "Q").NumberFormat = "0.00%"
    Else
        Decrease_Percent = Decrease_Percent
        
    If ws.Cells(i, "L").Value > Total_Volume Then
        Total_Volume = ws.Cells(i, "L").Value
        ws.Cells(4, "P").Value = ws.Cells(i, "I").Value
        ws.Cells(4, "Q").Value = ws.Cells(i, "L").Value
    Else
        Total_Volume = Total_Volume
        
        
    End If
    End If
    End If
    
Next i
    
Next ws

End Sub
