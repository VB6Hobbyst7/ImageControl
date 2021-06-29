Attribute VB_Name = "SelectShapes"
'-----SELECT-----

Sub SelectShapesWithinSelectedRange()
On Error Resume Next
        If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range first"
    Exit Sub
    End If
    Dim shp As Shape
    Dim R As Range
    Set R = Selection
    For Each shp In ActiveSheet.Shapes
        If Not Intersect(Range(shp.TopLeftCell, shp.BottomRightCell), R) Is Nothing Then
            shp.Select Replace:=False
        End If
    Next shp
End Sub

Sub SelectShapesByName()
On Error Resume Next
    Dim shp As Shape
    ActiveSheet.Range("A1").Select
    Dim str As String
    str = InputBox("contains in NAME?")

    For Each shp In ActiveSheet.Shapes
        If InStr(shp.Name, str) Then
            shp.Select Replace:=False
        End If
    Next shp
End Sub

Sub SelectShapesByText()
    
    Dim shp As Shape
    Dim str As String
    str = InputBox("contains in TEXT?")

    ActiveSheet.Range("A1").Select
    
    On Error GoTo nxt
    For Each shp In ActiveWorkbook.ActiveSheet.Shapes 'Selection.ShapeRange
        If shp.Type <> 13 Then
            With shp.TextFrame.Characters
                If InStr(1, .Text, str) Then
                    shp.Select Replace:=False
                End If
            End With
        End If
nxt:
        
    Next shp
End Sub

    '----END SELECT------
