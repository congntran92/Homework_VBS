Sub Walstreet_Stock_Easy()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Long
Dim year_open As Double
Dim year_close As Double
Dim Summary_Table As Integer

On Error Resume Next
vol = 0
For Each ws In ThisWorkbook.Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Total Stock Volume"

Summary_Table = 2
    For I = 2 To ws.UsedRange.Rows.Count
            vol = ws.Cells(I, 7).Value + vol

         If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
            ticker = ws.Cells(I, 1).Value
            vol = ws.Cells(I, 7).Value + vol
            
            year_open = ws.Cells(I, 3).Value
            year_close = ws.Cells(I, 6).Value

            ws.Cells(Summary_Table, 9).Value = ticker
            ws.Cells(Summary_Table, 10).Value = vol
            Summary_Table = Summary_Table + 1

             vol = 0

        End If
    Next I
Next ws


End Sub
