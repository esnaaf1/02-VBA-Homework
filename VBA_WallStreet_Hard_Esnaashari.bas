Attribute VB_Name = "Module1"
Sub StockMarketAnalysis_Hard():
    
    ' hard version of the Stock Market Analysis
    ' Farshad Esnaashari
    
    ' Declare variables
    Dim i As Long
    Dim TotalRows As Double
    Dim Ticker() As String
    Dim TickerCount As Long
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
    
    ' find the total number of active rows
    TotalRows = Cells(Rows.Count, 1).End(xlUp).Row
    

    'Set the array variable dimentions
    ReDim Ticker(TotalRows)
    ReDim TotalStocksValue(TotalRows)
    ReDim YearPriceOpen(TotalRows)
    ReDim YearPriceClose(TotalRows)
    
    ' set the column headings
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    
    
    ' initialize the variables
    TickerCount = 0
    
    Ticker(TickerCount) = Cells(2, 1).Value
    TotalStocksValue(TickerCount) = Cells(2, 7).Value
    YearPriceOpen(TickerCount) = Cells(2, 3).Value
    
    ' populate the arrays variables

    For i = 3 To TotalRows
           
    
        ' if a new ticker symbol then increment the TickerCount add the ticker to the
        ' Ticker Array and initialize the stock volume to the first row

        If (Ticker(TickerCount) <> Cells(i, 1).Value) Then
            TickerCount = TickerCount + 1
            Ticker(TickerCount) = Cells(i, 1).Value
            TotalStocksValue(TickerCount) = Cells(i, 7).Value
            YearPriceOpen(TickerCount) = Cells(i, 3).Value
        Else
            TotalStocksValue(TickerCount) = Cells(i, 7).Value + TotalStocksValue(TickerCount)
            YearPriceClose(TickerCount) = Cells(i, 6).Value
        End If

    Next i
    
    ' Loop through the Array of Tickers
    For i = 0 To TickerCount
    
       If (YearPriceOpen(i) <> 0 And YearPriceClose(i) <> 0) Then
       
            If (GreatestIncrease <= ((YearPriceClose(i) - YearPriceOpen(i)) / YearPriceOpen(i))) Then
                GreatestIncrease = (YearPriceClose(i) - YearPriceOpen(i)) / YearPriceOpen(i)
                GreatestIncreaseTicker = Ticker(i)
            End If
        
            If (GreatestDecrease > ((YearPriceClose(i) - YearPriceOpen(i)) / YearPriceOpen(i))) Then
                GreatestDecrease = (YearPriceClose(i) - YearPriceOpen(i)) / YearPriceOpen(i)
                GreatestDecreaseTicker = Ticker(i)
            End If
            If (GreatestVolume < TotalStocksValue(i)) Then
                GreatestVolume = TotalStocksValue(i)
                GreatestVolumeTicker = Ticker(i)
            End If
        End If
        
        ' log the values to the excel sheet
        Cells(i + 2, 9).Value = Ticker(i)
        YearlyChange = YearPriceClose(i) - YearPriceOpen(i)
        Cells(i + 2, 10).Value = YearlyChange
        Cells(i + 2, 10).NumberFormat = "0.000000000"
        
        ' Account for the opening prices at 0
        
        If (YearPriceOpen(i) = 0) Then
            Cells(i + 2, 11).Value = 0
        Else
            Cells(i + 2, 11).Value = YearlyChange / YearPriceOpen(i)
        End If
        Cells(i + 2, 11).NumberFormat = "0.00%"
        Cells(i + 2, 12).Value = TotalStocksValue(i)
        
        ' Highlight the yearly change cell
        
        If (Cells(i + 2, 10).Value <= 0) Then
            Cells(i + 2, 10).Interior.ColorIndex = 3
        Else
            Cells(i + 2, 10).Interior.ColorIndex = 4
        End If
         
    Next i
    
    ' write out the greatest increase/descrease ticker and values to the columns
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
    
    'Auto fit the columns
    Range("O1:Q4").EntireColumn.AutoFit
    

End Sub




