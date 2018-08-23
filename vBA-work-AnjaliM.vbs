Sub Hard_Solution():

Dim LastRow As Double
Dim Lastcol As Integer
Dim TickerCol As Integer
Dim TotalCol As Integer, PercentageCol As Integer, PriceChangecol As Integer, TickerCol2 As Integer
Dim valueCol As Double
Dim sum As Double, i As Double, j As Integer
Dim closeprice As Double, openPrice As Double, priceChange As Double, percentage As Double
Dim loopcount  As Integer
Dim maxTotal As Double, minPercentage As Double, maxPercentage As Double
Dim indexMax As Double, indexMinPercentage As Double, indexmaxPercentage As Double
Dim maxVolume As Double
Dim maxVolumeIndex As Double

maxVolume = 0
maxVolumeIndex = 2
TickerCol = 10
PriceChangecol = 11
PercentageCol = 12
TotalCol = 13
TickerCol2 = 17
valueCol = 18

For Each ws In Worksheets
    'get the numbe rof Columns oer sheet
    Lastcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
   ' Get the Row count
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Fix the labels and cloumns Headers
    ws.Cells(1, TickerCol).Value = "Ticker"
    ws.Cells(1, PriceChangecol).Value = "Yearly Change"
    ws.Cells(1, PercentageCol).Value = "% change"
    ws.Cells(1, TotalCol).Value = "Total Stock Volume "
    ws.Cells(1, TickerCol2).Value = "Ticker"
    ws.Cells(1, valueCol).Value = "Value"
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease "
    ws.Cells(4, 16).Value = "Greatest %volume"
    
    sum = 0
    j = 2
    loopcount = 0
    'Main Sheet logic
    For i = 2 To LastRow
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            'loopcount = loopcount + 1
            closeprice = ws.Cells(i, 6).Value
            openPrice = ws.Cells(i - loopcount, 3).Value
            sum = sum + ws.Cells(i, 7).Value
            ws.Cells(j, TickerCol).Value = ws.Cells(i, 1).Value
            ws.Cells(j, TotalCol).Value = sum
            If openPrice = 0 Then
                 ws.Cells(j, PriceChangecol).Value = 0
                 ws.Cells(j, PercentageCol).Value = 0
            Else
                priceChange = closeprice - openPrice
                percentage = priceChange / openPrice
                ws.Cells(j, PriceChangecol).Value = priceChange
                ws.Cells(j, PercentageCol).Value = percentage
            End If
            j = j + 1
            sum = 0
            loopcount = 0
        Else
            sum = sum + ws.Cells(i, 7).Value
            loopcount = loopcount + 1
        End If
    Next i
    'Loopd to change the Color Red/Green
    For i = 2 To LastRow
        If ws.Cells(i, PriceChangecol).Value >= 0 Then
            ws.Cells(i, PriceChangecol).Interior.ColorIndex = 4
        Else
            ws.Cells(i, PriceChangecol).Interior.ColorIndex = 3
        End If
   
    Next i
    
    indexMax = 0
    maxTotal = 0
    minPercentage = 0
    maxPercentage = 0
    indexmaxPercentage = 0
    indexMinPercentage = 0
   

    ' Loop for last part to find Greatest Volume /% etc
    For k = 2 To j
        ' search for Greatest %volume
        If ws.Cells(k, TotalCol).Value > maxTotal Then
            maxTotal = ws.Cells(k, TotalCol).Value
            indexMax = k
        End If
        'Search for Greatest % Increase
        If ws.Cells(k, PercentageCol).Value > maxPercentage Then
            maxPercentage = ws.Cells(k, PercentageCol).Value
            indexmaxPercentage = k
        End If
        'Search for Greatest % Decrease"
        If ws.Cells(k, PercentageCol).Value < minPercentage Then
            minPercentage = ws.Cells(k, PercentageCol).Value
            indexMinPercentage = k
        End If

    Next k
    
    'Set values
    ws.Range("L2:L" & LastRow).NumberFormat = "0.00%"
    
    ws.Cells(2, TickerCol2).Value = ws.Cells(indexmaxPercentage, 10).Value
    ws.Cells(2, valueCol).Value = ws.Cells(indexmaxPercentage, PercentageCol).Value
    ws.Cells(2, valueCol).NumberFormat = "0.00%"
    
    ws.Cells(3, TickerCol2).Value = ws.Cells(indexMinPercentage, 10).Value
    ws.Cells(3, valueCol).Value = ws.Cells(indexMinPercentage, PercentageCol).Value
    ws.Cells(3, valueCol).NumberFormat = "0.00%"
    
    ws.Cells(4, TickerCol2).Value = ws.Cells(indexMax, 10).Value
    ws.Cells(4, valueCol).Value = ws.Cells(indexMax, TotalCol).Value
    
    ws.Columns("A:R").AutoFit

Next ws

End Sub