Attribute VB_Name = "CircleCells"
    '-----CIRCLE-----

Sub CircleBoxADD()
If TypeName(Selection) <> "Range" Then
        MsgBox "Please select 1 or more ranges before running the macro."
    Exit Sub
End If
Dim MyOval As Shape
    Dim cell As Range
    For Each cell In Selection
        t = cell.MergeArea.Top
        L = cell.MergeArea.Left
        h = cell.MergeArea.Height
        w = cell.MergeArea.Width
    

        Set MyOval = ActiveSheet.Shapes.AddShape(msoShapeOval, L + 2, t + 2, w - 4, h - 4)
        With MyOval
            .Name = "CircleMarckCell"
            .Fill.Visible = msoFalse
            .Line.Visible = msoTrue
            .Line.ForeColor.RGB = RGB(255, 0, 0)
            .Line.Transparency = 0
            .Line.Visible = msoTrue
            .Line.Weight = 0.5
        End With
    Next
End Sub

Sub CircleBoxREMOVE()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If shp.Type = 1 And shp.AutoShapeType = msoShapeOval _
        And shp.Name = "CircleMarckCell" Then
            shp.Delete
        End If
    Next
End Sub

'----END CIRCLE-----

