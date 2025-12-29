' Send completed Excel sheet via Outlook to the boss
Sub SendExcelSheetByOutlook()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim wb As Workbook
    Dim FilePath As String
    
    '--- Reference current workbook
    Set wb = ThisWorkbook
    
    '--- Save workbook to ensure attachment is up-to-date
    FilePath = wb.FullName
    wb.Save
    
    '--- Create Outlook application
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0 = MailItem
    
    '--- Build email
    With OutlookMail
        .To = "biggbossman@outlook.com"
        .CC = "smallbossman@gmail.com"
        .BCC = ""
        .Subject = "Excel Sheet with Processed Data, PivotTable and Charts"
        .Body = "Hi," & vbCrLf & vbCrLf & _
                "Please find attached the finished Excel sheet." & vbCrLf & vbCrLf & _
                "Regards," & vbCrLf & "Daniel"
        
        '--- Attach workbook
        .Attachments.Add FilePath
        
        '--- Display before sending (use .Send to auto-send)
        .Display
    End With
    
    MsgBox "Email created with Excel sheet attached!", vbInformation
End Sub
