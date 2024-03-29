Attribute VB_Name = "DrawArrowsBetweenCells"
Sub DrawArrows(FromRange As Range, ToRange As Range, Optional RGBcolor As Long, Optional LineType As String)
    '---------------------------------------------------------------------------------------------------
    '---Script: DrawArrows------------------------------------------------------------------------------
    '---Created by: Ryan Wells -------------------------------------------------------------------------
    '---Date: 10/2015-----------------------------------------------------------------------------------
    '---Description: This macro draws arrows or lines from the middle of one cell to the middle --------
    '----------------of another. Custom endpoints and shape colors are suppported ----------------------
    '---------------------------------------------------------------------------------------------------
    'https://wellsr.com/vba/2015/excel/draw-lines-or-arrows-between-cells-with-vba/
    Dim dleft1 As Double, dleft2 As Double
    Dim dtop1 As Double, dtop2 As Double
    Dim dheight1 As Double, dheight2 As Double
    Dim dwidth1 As Double, dwidth2 As Double
    dleft1 = FromRange.Left
    dleft2 = ToRange.Left
    dtop1 = FromRange.Top
    dtop2 = ToRange.Top
    dheight1 = FromRange.Height
    dheight2 = ToRange.Height
    dwidth1 = FromRange.Width
    dwidth2 = ToRange.Width
 
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, dleft1 + dwidth1 / 2, dtop1 + dheight1 / 2, dleft2 + dwidth2 / 2, dtop2 + dheight2 / 2).Select
    'format line
    With Selection.ShapeRange.Line
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadOpen
        .Weight = 1.75
        .Transparency = 0.5
        If UCase(LineType) = "DOUBLE" Then       'double arrows
            .BeginArrowheadStyle = msoArrowheadOpen
        ElseIf UCase(LineType) = "LINE" Then     'Line (no arows)
            .EndArrowheadStyle = msoArrowheadNone
        Else                                     'single arrow
            'defaults to an arrow with one head
        End If
        'color arrow
        If RGBcolor <> 0 Then
            .ForeColor.RGB = RGBcolor            'custom color
        Else
            .ForeColor.RGB = RGB(228, 108, 10)   'orange (DEFAULT)
        End If
    End With
 
End Sub

Sub DeleteArrows()
    For Each shp In ActiveSheet.Shapes
        If shp.Connector = msoTrue Then
            shp.Delete
        End If
    Next shp
End Sub

Sub HideArrows()
    For Each shp In ActiveSheet.Shapes
        If shp.Connector = msoTrue Then
            shp.Line.Transparency = 1
        End If
    Next shp
End Sub

