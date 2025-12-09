Sub ReadDataIntoWorksheet()
    Dim sourceWB As Workbook
    Dim destWB As Workbook
    Dim sourceWS As Worksheet
    Dim destWS As Worksheet
    Dim sourcePath As String
    Dim destPath As String
    
    ' Set file paths (update these paths!)
    sourcePath = "C:\Users\chakdaniel2\Documents\dbs\EmployeeData.xlsx"
    destPath = "C:\Users\chakdaniel2\Desktop\dmpVBA.xlsm"
    
    ' Open the source workbook
    Set sourceWB = Workbooks.Open(sourcePath)
    Set sourceWS = sourceWB.Sheets("Data") ' Change sheet name if needed
    
    ' Open the destination workbook
    Set destWB = Workbooks.Open(destPath)
    Set destWS = destWB.Sheets("Master") ' Change sheet name if needed
    
    ' Copy data (example: A1:D100)
    sourceWS.Range("A1:N1001").Copy
    
    ' Paste into destination (starting at A1)
    destWS.Range("A1").PasteSpecial Paste:=xlPasteValues
    
    ' Close source workbook
    sourceWB.Close SaveChanges:=False
    
    ' Clear clipboard
    Application.CutCopyMode = False
    
    MsgBox "Data copied successfully!", vbInformation
    
    ' Save and close destination workbook
    destWB.Save
    'destWB.Close
    
End Sub
