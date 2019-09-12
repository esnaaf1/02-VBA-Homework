Attribute VB_Name = "Module1"
Sub StockMarketAnalysis_easy():
    
' Easy version of the WallStreet VBA
' Farshad Esnaashari


    'Declare variables
    
    Dim i As Long
    
    Dim tickerCount As Long
    
    Dim TotalRows As Double
    Dim Ticker() As String
    Dim TotalStocksValue() As Double
    
    
    'Determine the last active row on the sheet
    
    TotalRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    ReDim Ticker(TotalRows)
    ReDim TotalStocksValue(TotalRows)

    
    'Initialize the variable
    
    tickerCount = 0
    
    Ticker(tickerCount) = Cells(2, 1).Value
    TotalStocksValue(tickerCount) = Cells(2, 7).Value
    
    
    For i = 3 To TotalRows
           
    
        ' if a new symbol then add it to the Ticker column and initialize the stock volume to the first row

        If (Ticker(tickerCount) <> Cells(i, 1).Value) Then
            tickerCount = tickerCount + 1
            Ticker(tickerCount) = Cells(i, 1).Value
            TotalStocksValue(tickerCount) = Cells(i, 7).Value
        Else
            TotalStocksValue(tickerCount) = Cells(i, 7).Value + TotalStocksValue(tickerCount)
        End If

    Next i
    
    'Write the values to the cells
    ' Set the first row labels
    
    Range("I1") = "Ticker"
    Range("J1") = "Total Stock Volume"

    For i = 0 To tickerCount
        
        Cells(i + 2, 9).Value = Ticker(i)
        Cells(i + 2, 10).Value = TotalStocksValue(i)
    Next i
  
End Sub