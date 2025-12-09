Sub FormatWorksheet()
    Dim ws As Worksheet, rng As Range, hrng As Range, rrng As Range, lastRow As Long, lastCol As Long
    
    ' Work on specific worksheet
    Set ws = ThisWorkbook.Sheets("DataCopy")
    
    'Find the last row and column of the used range
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Get the used range
    Set rng = ws.UsedRange
    
    ' Apply borders to all sides
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = vbBlack
    End With
    
    MsgBox "Borders added to the used range!", vbInformation

    ' Center the data horizontally and vertically for the entire range
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    MsgBox "Data in the used range has been centered!", vbInformation
    
    ' Change font type, color, bold and background color for headers
    Set hrng = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    hrng.Font.Name = "Arial Black"
    'hrng.Font.Bold = True
    hrng.Font.Size = 14
    hrng.Font.Color = vbWhite
    hrng.Interior.Color = vbBlack
    MsgBox "Header range formatted successfully!", vbInformation
    
    ' Change font type, color and size for the rest of the range
    Set rrng = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    rrng.Font.Name = "Arial Narrow"
    rrng.Font.Size = 10
    rrng.Font.Color = vbBlack
    MsgBox "Remaining range values formatted successfully!", vbInformation
    
    ' AutoFit all columns
    ws.Cells.EntireColumn.AutoFit
    
    ' AutoFit all rows
    ws.Cells.EntireRow.AutoFit
    
    MsgBox "All rows and columns have been autofitted!", vbInformation
    
End Sub
