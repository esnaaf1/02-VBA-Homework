Sub StockDataAnalysis():
    
    Dim J As Integer
    Dim i As Long
    Dim TotalRows As Double
    Dim Ticker() As String
    Dim TotalStocksValue() As Double
    

    Dim YearPriceOpen() As Double
    Dim YearPriceClose() As Double
    
    Dim YearlyChange As Double
    
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    TotalRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    ReDim Ticker(TotalRows)
    ReDim TotalStocksValue(TotalRows)
    ReDim YearPriceOpen(TotalRows)
    ReDim YearPriceClose(TotalRows)
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    
    J = 0
    
    Ticker(J) = Cells(2, 1).Value
    TotalStocksValue(J) = Cells(2, 7).Value
    YearPriceOpen(J) = Cells(2, 3).Value
    
    
    For i = 3 To TotalRows
           
    
        ' if a new symbol then add it to the Ticker column and initialize the stock volume to the first row

        If (Ticker(J) <> Cells(i, 1).Value) Then
            J = J + 1
            Ticker(J) = Cells(i, 1).Value
            TotalStocksValue(J) = Cells(i, 7).Value
            YearPriceOpen(J) = Cells(i, 3).Value
        Else
            TotalStocksValue(J) = Cells(i, 7).Value + TotalStocksValue(J)
            YearPriceClose(J) = Cells(i, 6).Value
        End If

    Next i
    
    For L = 0 To J
    
       If (YearPriceOpen(L) <> 0 And YearPriceClose(L) <> 0) Then
       
            If (GreatestIncrease <= ((YearPriceClose(L) - YearPriceOpen(L)) / YearPriceOpen(L))) Then
                GreatestIncrease = (YearPriceClose(L) - YearPriceOpen(L)) / YearPriceOpen(L)
                GreatestIncreaseTicker = Ticker(L)
            End If
        
            If (GreatestDecrease > ((YearPriceClose(L) - YearPriceOpen(L)) / YearPriceOpen(L))) Then
                GreatestDecrease = (YearPriceClose(L) - YearPriceOpen(L)) / YearPriceOpen(L)
                GreatestDecreaseTicker = Ticker(L)
            End If
            If (GreatestVolume < TotalStocksValue(L)) Then
                GreatestVolume = TotalStocksValue(L)
                GreatestVolumeTicker = Ticker(L)
            End If
        End If
        
        Cells(L + 2, 9).Value = Ticker(L)
        YearlyChange = YearPriceClose(L) - YearPriceOpen(L)
        Cells(L + 2, 10).Value = YearlyChange
        Cells(L + 2, 10).NumberFormat = "0.000000000"
        
        ' Account for the opening prices at 0
        
        If (YearPriceOpen(L) = 0) Then
            Cells(L + 2, 11).Value = 0
        Else
            Cells(L + 2, 11).Value = YearlyChange / YearPriceOpen(L)
        End If
        Cells(L + 2, 11).NumberFormat = "0.00%"
        Cells(L + 2, 12).Value = TotalStocksValue(L)
        
        ' Highlight the yearly change cell
        
        If (Cells(L + 2, 10).Value <= 0) Then
            Cells(L + 2, 10).Interior.ColorIndex = 3
        Else
            Cells(L + 2, 10).Interior.ColorIndex = 4
        End If
         
    Next L
    
    Range("O2").Value = " Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P2").Value = GreatestIncreaseTicker
    Range("Q2").Value = GreatestIncrease
    Range("Q2").NumberFormat = "0.00%"
    
    
    Range("P3").Value = GreatestDecreaseTicker
    Range("Q3").Value = GreatestDecrease
    Range("Q3").NumberFormat = "0.00%"
    
    Range("P4").Value = GreatestVolumeTicker
    Range("Q4").Value = GreatestVolume
    Range("Q4").NumberFormat = "General"
    
    
End Sub



