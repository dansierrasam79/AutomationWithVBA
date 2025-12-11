Sub CheckColumnDataTypes()
    Dim ws As Worksheet, targetRow As Long, targetCol As Long, r As Integer, cell As Range, rng As Range, errorCell As Integer
    Set ws = ThisWorkbook.Sheets("DataCopy")
    
    'Find last row and column
    targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    targetCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Delete contents in the next column and row after targetRow & targetCol
    ws.Rows(targetRow + 1).Clear
    ws.Columns(targetCol + 1).Clear
    
    ' Check if all cells are empty
    For Each cell In ws.UsedRange
        If cell.Value = "" Then
            Debug.Print cell.Value
            cell.Interior.Color = vbRed
            errorCell = errorCell + 1
        End If
    Next cell
    MsgBox errorCell & " Cells have issues and are marked in red"
    errorCell = 0
    
    'Checks if all cells in EEID is alphanumeric
    For r = 2 To targetRow
        If Not Cells(r, 1) Like WorksheetFunction.Rept("[A-Za-z0-9]", Len(Cells(r, 1))) Then
            cell.Interior.Color = vbRed
            errorCell = errorCell + 1
        End If
    Next r
    MsgBox errorCell & " Cells under EEID have issues and are marked in red"
    errorCell = 0
    
    'Integer = Age, Annual Salary
    ' Checking if cells in the "Age" column is of Integer type
    For r = 2 To targetRow
        If IsNumeric(Cells(r, 8)) Then
            If Not Cells(r, 8) = Int(Cells(r, 8)) Then
                Cells(r, 8).Interior.Color = vbRed
            End If
        End If
    Next r
    MsgBox errorCell & " Cells under Age have issues and are marked in red"
    errorCell = 0
    
    ' Checking if cells in the "Annual Salary" column is of Integer type
    For r = 2 To targetRow
        If IsNumeric(Cells(r, 10)) Then
            If Not Cells(r, 10) = Int(Cells(r, 10)) Then
                Cells(r, 10).Interior.Color = vbRed
            End If
        End If
    Next r
    MsgBox errorCell & " Cells under Annual Salary have issues and are marked in red"
    errorCell = 0
    
    'Date = Hire Date, Exit Date
    For r = 2 To targetRow
        If Not IsDate(Cells(r, 9)) Then
            Cells(r, 9).Interior.Color = vbRed
        End If
    Next r
    
    'Percentage = Bonus %
    ' Checking if cells in the Bonus % column is of Percentage type
    For r = 2 To targetRow
        If Not IsNumeric(Cells(r, 11)) And Not Cells(r, 11).NumberFormat Like "*%*" Then
            Cells(r, 11).Interior.Color = vbRed
        End If
    Next r
    MsgBox errorCell & " Cells under Bonus % have issues and are marked in red"
    errorCell = 0
    
    'Text = Full Name, Job Title, Department, Business Unit, Gender, Ethnicity, Country, City
    Set rng = Range(Cells(1, 2), Cells(targetRow, 2))
    
    ' Checks if Full Name has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 2)) Then
            Cells(r, 2).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if Job Title has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 3)) Then
            Cells(r, 3).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if Department has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 4)) Then
            Cells(r, 4).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if Business Unit has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 5)) Then
            Cells(r, 5).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if Gender has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 6)) Then
            Cells(r, 6).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if Ethnicity has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 7)) Then
            Cells(r, 7).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if Country has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 12)) Then
            Cells(r, 12).Interior.Color = vbRed
        End If
    Next cell
    
    ' Checks if City has any non-textual content
    For Each cell In rng
        If Not WorksheetFunction.IsText(Cells(r, 13)) Then
            Cells(r, 13).Interior.Color = vbRed
        End If
    Next cell
End Sub
