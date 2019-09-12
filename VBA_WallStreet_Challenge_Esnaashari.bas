Attribute VB_Name = "Module2"
Sub StockMarketAnalysis_Challenge()

' Challenge version of the Stock Market Analysis
' Farshad Esnaashari

' Note: This VBA calls another sub to excute.  You must add the StockMarketAnalysis_Hard() VBA as a sub to the sheet before using this VBA code

    'Declare ws as a Worksheet objects

    Dim ws As Worksheet

    'Loop through the sheets
    
    For Each ws In Worksheets
        
        ' Active the sheet
        
        ws.Activate
        
        Call StockMarketAnalysis_Hard
        
    Next ws

End Sub
