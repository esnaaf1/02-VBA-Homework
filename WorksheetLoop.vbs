Sub WiorSheetLioop()

Dim M As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet
ws_num = ThisWorkbook.Worksheets.Count

For M = 1 To ws_num

    ThisWorkbook.Worksheets(M).Activate
    Call StockDataAnalysis
    
Next M

starting_ws.Activate 'activate the worksheet that was originally active

End Sub
