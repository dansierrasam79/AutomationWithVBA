Sub CreatePivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim rngData As Range
    
    '--- Set references
    Set wsData = ThisWorkbook.Sheets("DataCopy")   ' Source sheet
    Set wsPivot = ThisWorkbook.Sheets("PivotSheet") ' Destination sheet
    
    '--- Define source range (dynamic last row/column)
    Set rngData = wsData.Range("A1").CurrentRegion
    
    '--- Create Pivot Cache
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rngData)
    
    '--- Add Pivot Table to destination sheet
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="EmployeesPivot")
    
    '--- Configure Pivot Table fields
    With pt
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("Country").Orientation = xlColumnField
        .PivotFields("Full Name").Orientation = xlDataField
    End With
    
    MsgBox "Pivot Table created successfully!", vbInformation
End Sub
