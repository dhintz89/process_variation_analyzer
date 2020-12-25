Sub CreateReport()

'Declare Variables
    Dim AppExcel As Object
    Dim xlBook As Object
    Dim uniqVals As Variant
    Dim resSheet As Object
    Dim dataSheet As Object
    Dim shapeSheet As Object
    Dim pn As Variant
    Dim lRow As Long
    Dim lcol As Long
    Dim validShapes(1 To 27) As String
    Dim i As Integer
    Dim shapeReport As String
    Const xlWorksheet As Long = -4167
    Const xlUp As Long = -4162
    Const xlToLeft As Long = -4159
    Const xlDown As Long = -4121
    Const xlDatabase As Long = 1
    Const xlCount As Long = -4112
    Const xlColumnStacked100 As Long = 53
    Const xlPercentOfRow As Long = 6
    Const xlRowField As Long = 1
    Const xlPageField As Long = 3
    Const xlLabelPositionInsideEnd As Long = 3
    Const xlColorIndexAutomatic As Long = -4105
    Const xlFormats As Long = -4122
    Const xlUnderlineStyleSingle As Long = 2
        
'Download ShapeReport from Visio
    shapeReport = "LocalVarianceRep"
    Visio.Application.Addons("VisRpt").Run ("/rptDefName=" + shapeReport)
    Set AppExcel = GetObject(, "Excel.Application")
    AppExcel.Visible = True
    Set xlBook = AppExcel.ActiveWorkbook
    Set dataSheet = xlBook.ActiveSheet
    Debug.Print (xlBook.Name)

'Create new sheet for Valid Shapes
    Set shapeSheet = xlBook.Worksheets.Add(Type:=xlWorksheet, After:=xlBook.Worksheets(1))
    With shapeSheet
        .Name = "ValidShapeTypeList"
    End With
    
'Define and paste valid shapes into Excel Doc (Need to update during SOP creation - don't forget range update in variable declaration)
    validShapes(1) = "PA_L0 Process"
    validShapes(2) = "PA_L1 Process"
    validShapes(3) = "PA_Decision"
    validShapes(4) = "Sales Associate"
    validShapes(5) = "Account Manager"
    validShapes(6) = "Sales Manager"
    validShapes(7) = "Sales Director"
    validShapes(8) = "Channel Specialist"
    validShapes(9) = "Customer Success Rep"
    validShapes(10) = "Cust Success Mgr"
    validShapes(11) = "Sales Dev Rep"
    validShapes(12) = "Inside Sales Rep"
    validShapes(13) = "Inside Sales Mgr"
    validShapes(14) = "QA Team"
    validShapes(15) = "Operations Rep"
    validShapes(16) = "Operations Manager"
    validShapes(17) = "Operations Director"
    validShapes(18) = "System Activity"
    validShapes(19) = "Client"
    validShapes(20) = "End User"
    validShapes(21) = "Reseller"
    validShapes(22) = "Distributor"
    validShapes(23) = "Data Specialist"
    validShapes(24) = "Data Engineer"
    validShapes(25) = "Business Intelligence"
    validShapes(26) = "Global Sales Ops"
    validShapes(27) = "Agency Billing"
    
    For i = 1 To UBound(validShapes)
        shapeSheet.Cells(i, 1).Value = validShapes(i)
    Next i
    
'Create new sheet for results
    Set resSheet = xlBook.Worksheets.Add(Type:=xlWorksheet, Before:=xlBook.Worksheets(1))
    With resSheet
        .Name = "Results"
    End With
    
    
'Insert Handover Column in column E
    dataSheet.Activate
    dataSheet.Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    lRow = dataSheet.Cells(dataSheet.Rows.Count, 3).End(xlUp).Row
    dataSheet.Range("E2").Value = "Handover"
    dataSheet.Range("E3", dataSheet.Cells(lRow, 5)).Value = "=IF(OR(E2=""Handover"",C2=""PA_Start/End.104"", H2=0),FALSE,IF(B3=LOOKUP(2,1/($H2:H$3=1),$B2:B$3),FALSE,TRUE))"
        
    
'Filter out shapes not in scope (wrong pages, connectors, headers, etc.)
'Changes to Visio Stencil will need to be addressed in ValidShapeTypeList
    'lRow = dataSheet.Cells(dataSheet.Rows.Count, 3).End(xlUp).Row
    dataSheet.Range("H2").Value = "Include?"
    dataSheet.Range("H3", dataSheet.Cells(lRow, 8)).Value = "=IF(ISERROR(FIND(""."",C3)),IF(OR(ISERROR(VLOOKUP(C3,ValidShapeTypeList!A:A,1,FALSE)),ISBLANK(D3)),0,1),IF(OR(ISERROR(VLOOKUP(LEFT(C3,FIND(""."",C3)-1),ValidShapeTypeList!A:A,1,FALSE)),ISBLANK(F3)),0,1))"


'Determine number of tables needed based on Unique Pages
    If IsEmpty(dataSheet.Range("J1")) Then
        dataSheet.Columns("J:J").Delete Shift:=xlToLeft
    End If
    xlBook.ActiveSheet.Range("F2", xlBook.ActiveSheet.Cells(lRow, 6)).Select
    AppExcel.Selection.Copy
    xlBook.ActiveSheet.Range("J1").Select
    xlBook.ActiveSheet.Paste
    AppExcel.CutCopyMode = False
    xlBook.ActiveSheet.Range("$J:$J").RemoveDuplicates Columns:=1, Header:=xlYes
    xlBook.ActiveSheet.Range("J2").Select
    If IsEmpty(AppExcel.Selection) Then AppExcel.Selection.Delete Shift:=xlUp  'empty cell is now showing up at the bottom, not the top
    uniqVals = xlBook.ActiveSheet.Range("J2", AppExcel.Selection.End(xlDown))
    xlBook.ActiveSheet.Cells(1, 10).Value = "Unique Pages"
    
'Create PivotTable for Each resultPage
    xlBook.Sheets("Results").Select
    AppExcel.ActiveSheet.Range("A27").Select
    For Each pn In uniqVals
        AppExcel.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
            "Sheet1!R2C1:R1048576C8", Version:=6).CreatePivotTable TableDestination:= _
            AppExcel.ActiveCell, TableName:=pn, DefaultVersion:=6
        With AppExcel.ActiveSheet.PivotTables(pn).PivotFields("PageName")
            .Orientation = xlPageField
            .Position = 1
        End With
        With AppExcel.ActiveSheet.PivotTables(pn).PivotFields("Include?")
            .Orientation = xlPageField
            .Position = 1
        End With
        AppExcel.ActiveSheet.PivotTables(pn).PivotFields("PageName").ClearAllFilters
        AppExcel.ActiveSheet.PivotTables(pn).PivotFields("PageName").CurrentPage = pn
        AppExcel.ActiveSheet.PivotTables(pn).PivotFields("Include?").ClearAllFilters
        AppExcel.ActiveSheet.PivotTables(pn).PivotFields("Include?").CurrentPage = "1"
        With AppExcel.ActiveSheet.PivotTables(pn).PivotFields("Variance")
            .Orientation = xlRowField
            .Position = 1
        End With
        AppExcel.ActiveSheet.PivotTables(pn).AddDataField AppExcel.ActiveSheet.PivotTables(pn) _
            .PivotFields("Displayed Text"), "Count of Displayed Text", _
            xlCount
        lcol = AppExcel.ActiveSheet.Cells(27, AppExcel.ActiveSheet.Columns.Count).End(xlToLeft).Column
        AppExcel.ActiveSheet.Cells(27, lcol).Select
        AppExcel.ActiveCell.Offset(0, 2).Select
    Next pn
    AppExcel.ActiveSheet.Range("A12").Select
    
'Build Concatenated Table
    AppExcel.ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R2C1:R1048576C8", Version:=6).CreatePivotTable TableDestination:= _
        "Results!R35C1", TableName:="CombinedTable", DefaultVersion:=6
    With AppExcel.ActiveSheet.PivotTables("CombinedTable").PivotFields("Include?")
        .Orientation = xlPageField
        .Position = 1
    End With
    AppExcel.ActiveSheet.PivotTables("CombinedTable").PivotFields("Include?").ClearAllFilters
    AppExcel.ActiveSheet.PivotTables("CombinedTable").PivotFields("Include?").CurrentPage = _
        "1"
    With AppExcel.ActiveSheet.PivotTables("CombinedTable").PivotFields("Variance")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With AppExcel.ActiveSheet.PivotTables("CombinedTable").PivotFields("PageName")
        .Orientation = xlRowField
        .Position = 1
    End With
    AppExcel.ActiveSheet.PivotTables("CombinedTable").AddDataField AppExcel.ActiveSheet.PivotTables( _
        "CombinedTable").PivotFields("Displayed Text"), "Count of Displayed Text", _
        xlCount
    With AppExcel.ActiveSheet.PivotTables("CombinedTable").PivotFields( _
        "Count of Displayed Text")
        .Calculation = xlPercentOfRow
        .NumberFormat = "0%"
    End With
    
'Build E2E Table
    AppExcel.ActiveWorkbook.Worksheets("Results").PivotTables("CombinedTable").PivotCache. _
        CreatePivotTable TableDestination:="Results!R35C6", TableName:= _
        "E2ETable", DefaultVersion:=6
    xlBook.Sheets("Results").Select
    AppExcel.ActiveSheet.Cells(27, 1).Select
    AppExcel.ActiveWorkbook.ShowPivotTableFieldList = True
    With AppExcel.ActiveSheet.PivotTables("E2ETable").PivotFields("Variance")
        .Orientation = xlRowField
        .Position = 1
    End With
    With AppExcel.ActiveSheet.PivotTables("E2ETable").PivotFields("Include?")
        .Orientation = xlPageField
        .Position = 1
    End With
    AppExcel.ActiveSheet.PivotTables("E2ETable").AddDataField AppExcel.ActiveSheet.PivotTables( _
        "E2ETable").PivotFields("Displayed Text"), "Count of Displayed Text", _
        xlCount
    AppExcel.ActiveSheet.PivotTables("E2ETable").PivotFields("Include?").ClearAllFilters
    AppExcel.ActiveSheet.PivotTables("E2ETable").PivotFields("Include?").CurrentPage = _
        "1"

'Create Bar Chart
    Dim Rng1
    Set Rng1 = AppExcel.ActiveSheet.Range("A2:G20")
    AppExcel.ActiveSheet.Range("A1").Select
    AppExcel.ActiveSheet.Shapes.AddChart2(297, xlColumnStacked100).Select
    AppExcel.ActiveChart.SetSourceData Source:=xlBook.ActiveSheet.Range("Results!$A$35:$D$36")
    With AppExcel.ActiveChart.Parent
        .Left = Rng1.Left
        .Top = Rng1.Top
    End With
    AppExcel.ActiveChart.ChartColor = 13
    AppExcel.ActiveChart.Axes(xlValue).Select
    AppExcel.ActiveChart.Axes(xlValue).MinimumScale = 0
    With AppExcel.ActiveSheet.Shapes("Chart 1").Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(146, 208, 80)
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Weight = 1.25
    End With
    AppExcel.ActiveSheet.ChartObjects("Chart 1").Activate
    AppExcel.ActiveChart.ShowReportFilterFieldButtons = False
    AppExcel.ActiveSheet.ChartObjects("Chart 1").Activate
    AppExcel.ActiveChart.ShowValueFieldButtons = False
    AppExcel.ActiveSheet.ChartObjects("Chart 1").Activate
    AppExcel.ActiveChart.ShowAxisFieldButtons = False
    AppExcel.ActiveChart.FullSeriesCollection(1).Select
    AppExcel.ActiveChart.FullSeriesCollection(1).ApplyDataLabels
    AppExcel.ActiveChart.PlotArea.Select
    If AppExcel.ActiveChart.FullSeriesCollection.Count > 1 Then
        AppExcel.ActiveChart.FullSeriesCollection(2).Select
        AppExcel.ActiveChart.FullSeriesCollection(2).ApplyDataLabels
    End If
    AppExcel.ActiveChart.ChartArea.Select
    AppExcel.CommandBars("Format Object").Visible = False
    AppExcel.ActiveChart.SetElement (msoElementChartTitleAboveChart)
    AppExcel.Selection.Caption = "Variance by Process Tab"
    AppExcel.Selection.Format.TextFrame2.TextRange.Font.UnderlineStyle = _
        msoUnderlineSingleLine
    AppExcel.ActiveSheet.Shapes("Chart 1").ScaleWidth 1.8, msoFalse, _
        msoScaleFromTopLeft
    AppExcel.ActiveSheet.Shapes("Chart 1").ScaleHeight 1.3, msoFalse, msoScaleFromTopLeft
    AppExcel.ActiveSheet.Shapes("Chart 1").IncrementLeft 25
    AppExcel.ActiveChart.FullSeriesCollection(1).Select
    AppExcel.ActiveChart.ChartGroups(1).GapWidth = 65
     
'Create E2E Pie Chart
    Dim Rng2
    Set Rng2 = AppExcel.ActiveSheet.Range("I2:N20")
    AppExcel.ActiveSheet.Range("I1").Select
    AppExcel.ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    AppExcel.ActiveChart.SetSourceData Source:=xlBook.ActiveSheet.Range("Results!$F$35:$G$35")
    With AppExcel.ActiveChart.Parent
        .Left = Rng2.Left
        .Top = Rng2.Top
    End With
    AppExcel.ActiveSheet.Shapes("Chart 2").IncrementLeft -25
    AppExcel.ActiveChart.FullSeriesCollection(1).Select
    AppExcel.ActiveChart.ChartArea.Select
    AppExcel.ActiveChart.ChartColor = 13
    With AppExcel.ActiveSheet.Shapes("Chart 2").Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(146, 208, 80)
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Weight = 1.25
    End With
    AppExcel.ActiveChart.SetElement (msoElementDataLabelBestFit)
    AppExcel.ActiveChart.ApplyDataLabels
    AppExcel.ActiveChart.FullSeriesCollection(1).DataLabels.Select
    AppExcel.Selection.ShowPercentage = True
    AppExcel.Selection.ShowValue = False
    AppExcel.Selection.Position = xlLabelPositionInsideEnd
    AppExcel.ActiveChart.ChartArea.Select
    AppExcel.Application.CommandBars("Format Object").Visible = False
    AppExcel.ActiveWorkbook.ShowPivotTableFieldList = False
    AppExcel.ActiveSheet.ChartObjects("Chart 2").Activate
    AppExcel.ActiveChart.HasTitle = True
    AppExcel.ActiveChart.ChartTitle.Select
    AppExcel.Application.CommandBars("Format Object").Visible = False
    AppExcel.ActiveChart.ChartTitle.Text = "Total Process Variance"
    AppExcel.Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Total Process Variance"
    With AppExcel.Selection.Format.TextFrame2.TextRange.Characters(1, 22).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With AppExcel.Selection.Format.TextFrame2.TextRange.Characters(1, 22).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoUnderlineSingleLine
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    AppExcel.ActiveChart.ChartArea.Select
    AppExcel.ActiveSheet.Shapes("Chart 2").ScaleWidth 1.3020833333, msoFalse, _
        msoScaleFromTopLeft
    AppExcel.ActiveSheet.Shapes("Chart 2").ScaleHeight 1.3020833333, msoFalse, _
        msoScaleFromTopLeft

    AppExcel.ActiveSheet.Range("A32").Select
    AppExcel.ActiveCell.FormulaR1C1 = "Variance By Tab"
    With AppExcel.Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Bold = True
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleSingle
        '.ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    AppExcel.ActiveSheet.Range("F32").Select
    AppExcel.ActiveCell.FormulaR1C1 = "Total Variance"
    With AppExcel.Selection.Font
        .Name = "Calibri"
        .Size = 16
        .Bold = True
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleSingle
        '.ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
'bring view back to top of page
    AppExcel.ActiveSheet.Shapes("Chart 1").Select
    AppExcel.ActiveSheet.Range("A1").Select
    
End Sub
