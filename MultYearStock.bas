Attribute VB_Name = "Module1"
Sub MultYearStock():

'using "ws" so that the below affects each sheet
For Each ws In Worksheets

'placing script in the appropriate cells
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

'setting column widths
ws.Range("J:J").ColumnWidth = 13
ws.Range("K:K").ColumnWidth = 14
ws.Range("L:L").ColumnWidth = 18
ws.Range("N:N").ColumnWidth = 20
ws.Range("P:P").ColumnWidth = 13

'formatting cells with color/bold lettering/centering/etc
ws.Range("N2:N4").Interior.ColorIndex = 24
ws.Range("N2:N4").Font.Bold = 1
ws.Range("O1:P1").Interior.ColorIndex = 24
ws.Range("O1:P1").Font.Bold = 1
ws.Range("I1:L1").Font.Bold = 1
ws.Range("I1:L1").HorizontalAlignment = xlCenter
ws.Range("O:O").HorizontalAlignment = xlCenter
ws.Range("P1").HorizontalAlignment = xlCenter
ws.Range("J:J").NumberFormat = "#,##0.00"

'defining the total volume amount for each individual stock and other variables
Dim TotalVolume As Double
Dim j As Integer
Dim OpenPrice As Double

'setting starting values for each variable
j = 2
TotalVolume = 0
OpenPrice = 2

'finding the last row in the set of data
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'setting a conditional for seperating each stock and combining their volumes
For i = 2 To lastrow

    'going down the list, if the stock symbol directly below the current stock
    'is different, then the current volume will be added to the running total
    'volume and that stock and the running total will be added to the list
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        ws.Range("I" & j).Value = ws.Cells(i, 1).Value
        ws.Range("L" & j).Value = TotalVolume
        
        'calculating yearly change. Open Price is set to start with stock "A"
        YearlyChange = ws.Cells(i, 6).Value - ws.Cells(OpenPrice, 3).Value
        ws.Range("J" & j).Value = YearlyChange
        
        'filling in color for positive and negative changes
        If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & j).Value < 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 3
        End If
        
        'calculating the percent change for the year
        PercentChange = (YearlyChange / ws.Cells(OpenPrice, 6).Value) * 100
        ws.Range("K" & j).Value = "%" & PercentChange
        
        'resetting yearly change for next i
        YearlyChange = 0
        
        'moving down one line for the stock symbol's data
        j = j + 1
        
        'resetting next stock symbol's total volume to 0
        TotalVolume = 0
        
    Else
        'the opposite of the condition above. If the next row has the same symbol
        'the total volume will contine to cumulate
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    End If
Next i

'finding the stocks with the greatest % increase, greatest % decrease and greatest vol
ws.Range("P2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K3500").Value) * 100
ws.Range("P3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K3500").Value) * 100
ws.Range("P4").Value = WorksheetFunction.Max(ws.Range("L2:L3500").Value)

'finding and defining the row numbers of the greatest % increase, greatest % decrease and greatest vol
MaxTicker = WorksheetFunction.Match(ws.Range("P2").Value, ws.Range("K2:K3500").Value, 0)
MinTicker = WorksheetFunction.Match(ws.Range("P3").Value, ws.Range("K2:K3500").Value, 0)
MaxVolTicker = WorksheetFunction.Match(ws.Range("P4").Value, ws.Range("L2:L3500").Value, 0)

'taking the row number and adding one bc of the header line and retrieving the ticker symbol
ws.Range("O2").Value = ws.Cells(MaxTicker + 1, 9)
ws.Range("O3").Value = ws.Cells(MinTicker + 1, 9)
ws.Range("O4").Value = ws.Cells(MaxVolTicker + 1, 9)

Next ws

End Sub
