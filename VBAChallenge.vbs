Attribute VB_Name = "Module2"
Sub SummaryStats():
'Create the new column headers and row labels
Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("I1:Q1").Columns.AutoFit
    ws.Range("O4").Columns.AutoFit
    
'Create variables for New Column Data
Dim Ticker As String
Dim StockVol As Variant
StockVol = 0
Dim SummaryRow As Integer
SummaryRow = 2
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double

lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1

For i = 2 To lastrow

'find the opening price
If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    OpenPrice = ws.Cells(i, 3).Value
End If

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    'find the ticker value
    Ticker = ws.Cells(i, 1).Value
    
    'find the closing price value
    ClosePrice = ws.Cells(i, 6).Value
    
    'calculate the yearly change
    YearlyChange = OpenPrice - ClosePrice
    
    'add to the stock volume
    StockVol = StockVol + ws.Cells(i, 7).Value
      
    'put the ticker value in the summary row
    ws.Range("I" & SummaryRow).Value = Ticker
    
    'put the stock volume in the summary row
    ws.Range("L" & SummaryRow).Value = StockVol
    
    ws.Range("J" & SummaryRow).Value = YearlyChange
    
    ws.Range("K" & SummaryRow).Value = FormatPercent(YearlyChange / OpenPrice)
    
    'increase the summary row for the next set
    SummaryRow = SummaryRow + 1
    
    'Reset stock volume for next ticker
    StockVol = 0

Else
    'running total of Stock Volume in the same ticker
    StockVol = StockVol + ws.Cells(i, 7).Value

End If
Next i

Next ws
End Sub


Sub NewDataActions():
Dim ws As Worksheet


For Each ws In Worksheets

'Conditional formatting
lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row + 1

For i = 2 To lastrow
'format yearly change cells

If ws.Cells(i, 10) > 0 Then
    ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)

Else
    ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    
End If

'Create Variables for greatest hits
Dim GreatestIncrease As Double
Dim GreatestDecrease As Doubleau
Dim GreatestVolume As Variant
Dim TickerIncrease As String
Dim TickerDecrease As String
Dim TickerVolume As String

TickerIncrease = ws.Cells(2, 16).Value
TickerDecrease = ws.Cells(3, 16).Value
TickerVolume = ws.Cells(4, 16).Value

lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row + 1

'find the greatest %increase
GreatestIncrease = Application.WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & lastrow))
ws.Cells(2, 17).Value = FormatPercent(GreatestIncrease)

'find the greatest %decrease
GreatestDecrease = Application.WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & lastrow))
ws.Cells(3, 17).Value = FormatPercent(GreatestDecrease)

'find the greatest total volume
GreatestVolume = Application.WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & lastrow))
ws.Cells(4, 17).Value = GreatestVolume
ws.Cells(4, 17).NumberFormat = "0"



'Find the tickers
If ws.Cells(i, 11).Value = GreatestIncrease Then
    TickerIncrease = ws.Cells(i, 9).Value
ElseIf ws.Cells(i, 11).Value = GreatestDecrease Then
     TickerDecrease = ws.Cells(i, 9).Value
ElseIf ws.Cells(i, 12).Value = GreatestVolume Then
     TickerVolume = ws.Cells(i, 9).Value
End If

Next i
    

Next ws

End Sub
