Attribute VB_Name = "Module1"
Sub StockMarketAnalysis_Moderate():
    
    'Moderate version of the Stock Market Analysis
    'Farshad Esnashari
    
    
    
    ' Declare variables
    
    Dim i As Long
    Dim TickerCount As Long
    Dim TotalRows As Double
    Dim Ticker() As String
    Dim TotalStocksValue() As Double
    

    Dim YearPriceOpen() As Double
    Dim YearPriceClose() As Double
    
    Dim YearlyChange As Double
    
    
    
    ' Find the total number of active rows
    TotalRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set the dimention of the  arrays
    
    ReDim Ticker(TotalRows)
    ReDim TotalStocksValue(TotalRows)
    ReDim YearPriceOpen(TotalRows)
    ReDim YearPriceClose(TotalRows)
    
    'Initialize the variables
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    
    TickerCount = 0
    
    Ticker(TickerCount) = Cells(2, 1).Value
    TotalStocksValue(TickerCount) = Cells(2, 7).Value
    YearPriceOpen(TickerCount) = Cells(2, 3).Value
    
    
    For i = 3 To TotalRows
           
    
        ' if a new symbol then add it to the Ticker column and initialize the stock volume to the first row

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
    
    For i = 0 To TickerCount
    
        
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
    
    
End Sub




