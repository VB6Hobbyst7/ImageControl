Attribute VB_Name = "AlignShapes"
'------ALIGN---------

Sub GridVertical()
    'Automatically space and align shapes to create vertical grid.

    Dim shp As Shape
    Dim lCnt As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim dSPACE As Variant
    Dim lRowCnt As Variant
    Dim dStart As Double
    Dim dMaxHeight As Double

    'Check if shapes are selected
    If TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If
  
    'Display an input box to ask user for the number of columns in the vertical grid.
    lRowCnt = Application.InputBox("Enter the number of columns for the vertical shape grid.", "Vertical Shape Grid", Type:=1)
  
    'Exit macro if user presses cancel
    If lRowCnt <= 0 Or lRowCnt = False Then
        Exit Sub
    End If

    'Display an input box to ask user for the amount of space between shapes.
    dSPACE = Application.InputBox("Enter the space between shapes in points.", "Vertical Shape Grid", Type:=1)

    'Exit macro if user presses cancel
    If TypeName(dSPACE) = "Boolean" Then
        Exit Sub
    End If

    'Set variables
    lCnt = 1
  
    'Loop through selected shapes (charts, slicers, timelines, etc.)
    For Each shp In Selection.ShapeRange
        With shp
            'If first shape then store left position
            If lCnt = 1 Then
                dStart = .Left
            Else
                If lCnt Mod lRowCnt = 1 Or lRowCnt = 1 Then
                    'New row, move shape down
                    .Top = dTop + dMaxHeight + dSPACE
                    .Left = dStart
                    dMaxHeight = .Height
                Else
                    'Same row, move shape right
                    .Top = dTop
                    .Left = dLeft + dWidth + dSPACE
                End If
        
            End If
      
            'Store properties of shape for use in moving next shape in the collection.
            dTop = .Top
            dLeft = .Left
            dHeight = .Height
            dWidth = .Width
            dMaxHeight = WorksheetFunction.Max(dMaxHeight, .Height)
        End With
    
        'Add to shape counter
        lCnt = lCnt + 1
    Next shp

End Sub

Sub GridHorizontal()
    'Automatically space and align shapes to create horizontal grid.

    Dim shp As Shape
    Dim lCnt As Long
    Dim dTop As Double
    Dim dLeft As Double
    Dim dHeight As Double
    Dim dWidth As Double
    Dim dSPACE As Variant
    Dim lColCnt As Variant
    Dim lCol As Long
    Dim dStart As Double
    Dim lRow As Double
    Dim dMaxWidth As Double

    'Check if shapes are selected
    If TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If

    'Display an input box to ask user for the number of rows in the horizontal grid.
    lColCnt = Application.InputBox("Enter the number of rows for the horizontal shape grid.", "Horizontal Shape Grid", Type:=1)
  
    'Exit macro if user presses cancel
    If lColCnt <= 0 Or lColCnt = False Then
        Exit Sub
    End If

    'Display an input box to ask user for the amount of space between shapes.
    dSPACE = Application.InputBox("Enter the space between shapes in points.", "Horizontal Shape Grid", Type:=1)

    'Exit macro if user presses cancel
    If TypeName(dSPACE) = "Boolean" Then
        Exit Sub
    End If

    'Set variables
    lCnt = 1
  
    'Loop through selected shapes (charts, slicers, timelines, etc.)
    For Each shp In Selection.ShapeRange
        With shp
            'If first shape then store top position
            If lCnt = 1 Then
                dStart = .Top
            Else
                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
                    'New column, move shape right
                    .Top = dStart
                    .Left = dLeft + dMaxWidth + dSPACE
                    dMaxWidth = .Width
                Else
                    'Same column, move shape down
                    .Top = dTop + dHeight + dSPACE
                    .Left = dLeft
                End If
            End If
      
            'Store properties of shape for use in moving next shape in the collection.
            dTop = .Top
            dLeft = .Left
            dHeight = .Height
            dWidth = .Width
            dMaxWidth = WorksheetFunction.Max(dMaxWidth, .Width)
        End With
    
        'Add to shape counter
        lCnt = lCnt + 1
    Next shp

End Sub

'-----END ALIGN-------
