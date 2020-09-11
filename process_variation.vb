Sub process_variation()
'Tasks:
'1. Separate in-scope vs. out-scope pages (Scope: L0, L1, & L2)
'2. Creates Variance and PageName Attributes on all objects
'3. Analyzes and tags each shape with appropriate Variance and PageName value (1.5pt outline = variance)
'4. Runs Excel Macro to process variance into charts/tables in Excel
'5. Notifies of completion

    'Enable diagram services
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150

    Dim UndoScopeID2 As Long
    Dim vsoPages As Visio.Pages
    Dim selShape As Visio.Shape
    Dim intPropRow As Integer
    Dim val As String
    Set vsoPages = ActiveDocument.Pages
        
'1. Separate in-scope vs. out-scope pages
    For Each Item In vsoPages
        'determine in-scope pages
        If Left(Item.Name, 2) = "L0" Or Left(Item.Name, 2) = "L1" Or Left(Item.Name, 2) = "L2" Then
            Debug.Print Item.Name
            val = 2
            For Each selShape In Item.Shapes
            'Possible Future Enhancement: If Not selShape.OneD Then   'doesn't add fields/data to connectors, etc. report filter by exist?(Variance)
                UndoScopeID2 = Application.BeginUndoScope("Shape Data")
                'add Variance Attribute if not exist
'2. Create Variance and PageName Attributes
                If Not selShape.CellExists("Prop.Variance", 1) Then
                    intPropRow = selShape.AddRow(visSectionProp, visRowLast, visTagDefault)
                    selShape.Section(visSectionProp).Row(intPropRow).NameU = "Variance"
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsLabel).FormulaU = """Variance"""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsType).FormulaU = "3"
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsFormat).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsLangID).FormulaU = "1043"
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsCalendar).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsPrompt).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsValue).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsSortKey).FormulaU = ""
                End If
                'add PageName Attribute if not exist
                If Not selShape.CellExists("Prop.PageName", 1) Then
                    intPropRow = selShape.AddRow(visSectionProp, visRowLast, visTagDefault)
                    selShape.Section(visSectionProp).Row(intPropRow).NameU = "PageName"
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsLabel).FormulaU = """PageName"""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsType).FormulaU = "0"
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsFormat).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsLangID).FormulaU = "1043"
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsCalendar).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsPrompt).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsValue).FormulaU = ""
                    selShape.CellsSRC(visSectionProp, intPropRow, visCustPropsSortKey).FormulaU = ""
                End If
'3. Analyzes & Tags Variance/PageName value
                'populate Variance and PageName
                selShape.Cells("Prop.PageName").FormulaU = Chr(34) & Item.Name & Chr(34)
                If selShape.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "1.5 pt" Then
                    selShape.Cells("Prop.Variance").Formula = "TRUE"
                Else:
                    selShape.Cells("Prop.Variance").Formula = "FALSE"
                End If
            'End If  'part of OneD statement (future enhancement)
            Next selShape
        End If
    Next Item
    
'4. Run Excel Macro
    Call CreateReport

'5. Completion Message
    Msg = MsgBox("Process Variance Completed.", vbOKOnly, "Variation Process Complete")

    'Restore diagram services
    ActiveDocument.DiagramServicesEnabled = DiagramServices

End Sub
