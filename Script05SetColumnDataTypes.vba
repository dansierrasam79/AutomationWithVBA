Sub SetColumnDataTypes()
    Dim ws As Worksheet, targetRow As Long, r As Integer, targetCol As Long, rng As Range, cell As Range
    Set ws = ThisWorkbook.Sheets("DataCopy")
    
    'Find last row and column
    targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    targetCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Integer column
    For r = 2 To targetRow
        Cells(r, 1).NumberFormat = "@"
        Cells(r, 2).NumberFormat = "@"
        Cells(r, 3).NumberFormat = "@"
        Cells(r, 4).NumberFormat = "@"
        Cells(r, 5).NumberFormat = "@"
        Cells(r, 6).NumberFormat = "@"
        Cells(r, 7).NumberFormat = "@"
        Cells(r, 8).NumberFormat = "0"
        Cells(r, 9).NumberFormat = "dd-mm-yyyy"
        Cells(r, 10).NumberFormat = "$##,##0.00"
        Cells(r, 11).NumberFormat = "0%"
        Cells(r, 12).NumberFormat = "@"
        Cells(r, 13).NumberFormat = "@"
        Cells(r, 14).NumberFormat = "dd-mm-yyyy"
    Next r
End Sub
