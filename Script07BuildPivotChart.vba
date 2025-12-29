Sub CreatePivotChart()
    Dim wsPivot As Worksheet
    Dim pt As PivotTable
    Dim chtObj As ChartObject
    Dim cht As Chart
    
    '--- Set references
    Set wsPivot = ThisWorkbook.Sheets("PivotSheet")   ' Sheet containing PivotTable
    Set pt = wsPivot.PivotTables("EmployeesPivot")   ' Name of your PivotTable
    
    '--- Add ChartObject to the sheet
    Set chtObj = wsPivot.ChartObjects.Add( _
        Left:=300, Top:=20, Width:=400, Height:=300)
    
    '--- Set chart reference
    Set cht = chtObj.Chart
    
    '--- Link chart to PivotTable
    cht.SetSourceData pt.TableRange1
    
    '--- Define chart type
    cht.ChartType = xlColumnClustered
    
    '--- Optional formatting
    With cht
        .HasTitle = True
        .ChartTitle.Text = "Employees by Country and Department"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Products"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Total Sales"
    End With
    
    MsgBox "Pivot Chart created successfully!", vbInformation
End Sub
