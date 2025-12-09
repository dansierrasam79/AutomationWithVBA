Sub CopyDataWithinWorkbook()
    Dim ws As Worksheet, sourceWS As Worksheet, destWS As Worksheet

    ' Set references to worksheets
    Set sourceWS = ThisWorkbook.Sheets("Master") ' Change to your source sheet name
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "DataCopy"
    MsgBox "New worksheet 'DataCopy' created!", vbInformation
    Set destWS = ThisWorkbook.Sheets("DataCopy")   ' Change to your destination sheet name
    
    ' Copy data (example: A1:N1001)
    sourceWS.Range("A1:N1001").Copy
    
    ' Paste into destination (starting at A1)
    destWS.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    MsgBox "Data copied successfully!", vbInformation

End Sub
