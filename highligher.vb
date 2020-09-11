Sub Highlighter()
    
    Dim sel As Visio.Selection
    Dim shp As Visio.Shape
    
    Set sel = ActiveWindow.Selection
    Set shp = sel.PrimaryItem
    Set shps = shp.Shapes
    
    For Each lilshape In shps
        If lilshape.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "1.5 pt" Then
            lilshape.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.5 pt"
            lilshape.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = ""
            lilshape.CellsSRC(visSectionObject, visRowThemeProperties, visColorSchemeIndex).FormulaU = "33"
        Else
            lilshape.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(MSOTINT(THEMEVAL(""Light""),-50))"
            lilshape.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "1.5 pt"
        End If
    Next
    
    If shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "1.5 pt" Then
        shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "0.5 pt"
        shp.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = ""
        shp.CellsSRC(visSectionObject, visRowThemeProperties, visColorSchemeIndex).FormulaU = "33"
    Else
        shp.CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "1.5 pt"
        shp.CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(MSOTINT(THEMEVAL(""Light""),-50))"
    End If
    
    sel.DeselectAll
    
End Sub
