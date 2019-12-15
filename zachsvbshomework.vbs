Sub Stocks()
    ' loops through all of the sheets
Dim WS As Worksheet
Dim Tick As String  'Ticker
        Dim CPrice As Double 'Closing Price
        Dim YChange As Double
        Dim OPrice As Double 'Opening Price
        Dim PChange As Double 'Percent of change
        Dim Volume As Double
        Dim Row As Double
        Volume = 0
        Row = 2 
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
        ' Adds Heading for summary
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        
                 
        'Sets Initial Open Price
        OPrice = Cells(2, 3).Value
         ' Loop through all ticker symbols
        For i = 2 To LastRow
         ' Checks to see if still within the same ticker symbol, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Sets the Ticker name
                Tick = Cells(i, 1).Value
                Cells(Row, 9).Value = Tick
                ' Set Close Price
                CPrice = Cells(i, 6).Value
                ' Adds Yearly Change
                YChange = CPrice - OPrice
                Cells(Row, 10).Value = YChange
                ' Adds Percent Change
                If (OPrice = 0 And CPrice = 0) Then
                    PChange = 0
                ElseIf (OPrice = 0 And CPrice <> 0) Then
                    PChange = 1
                Else
                    PChange = YChange / OPrice
                    Cells(Row, 11).Value = PChange
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                ' Adds Total Volume
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
                Row = Row + 1
                ' resets the Open Price
                OPrice = Cells(i + 1, 3)
                ' resets the Volume Total
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        ' Sets the Cell Colors
        For lr = 2 To YCLastRow
            If (Cells(lr, 10).Value > 0 Or Cells(lr, 10).Value = 0) Then
                Cells(lr, 10).Interior.ColorIndex = 10
            ElseIf Cells(lr, 10).Value < 0 Then
                Cells(lr, 10).Interior.ColorIndex = 3
            End If
        Next lr
        ' Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Delta"
        ' Look through each of the rows to find the greatest value and its associate ticker
        For Yast = 2 To YCLastRow
            If Cells(Yast, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(Yast, 9).Value
                Cells(2, 17).Value = Cells(Yast, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(Yast, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(Yast, 9).Value
                Cells(3, 17).Value = Cells(Yast, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Yast, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(Yast, 9).Value
                Cells(4, 17).Value = Cells(Yast, 12).Value
            End If
        Next Yast    
    Next WS    
End Sub
