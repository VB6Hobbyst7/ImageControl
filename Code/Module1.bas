Attribute VB_Name = "Module1"
Public Shrink As Double

Function Clipboard(Optional StoreText As String) As String
    'PURPOSE: Read/Write to Clipboard
    'Source: ExcelHero.com (Daniel Ferry)

    Dim X As Variant

    'Store as variant for 64-bit VBA support
    X = StoreText

    'Create HTMLFile Object
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                'Write to the clipboard
                .SetData "text", X
            Case Else
                'Read from the clipboard (no variable passed through)
                Clipboard = .GetData("text")
            End Select
        End With
    End With

End Function


Sub TextBoxResizeTB()
    'Auto resize all text boxes to fit the content in a worksheet
    Dim xShape As Shape
    Dim xSht As Worksheet
    On Error Resume Next
    For Each xSht In ActiveWorkbook.Worksheets
        For Each xShape In xSht.Shapes
            '            If xShape.Type = 17 Then
            xShape.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            xShape.TextFrame2.WordWrap = False
            '            End If
        Next
    Next
End Sub

Sub PicturesFitCenter()
If TypeName(Selection) = "Range" Then
        MsgBox "Please select shapes before running the macro."
Exit Sub
End If
Dim ans As Long
    ans = MsgBox("Lock Aspect Ratio?", vbYesNoCancel)
    If ans = vbCancel Then Exit Sub
    
    Dim P As Shape
    'Set p=ActiveSheet.Shapes(Application.Caller)
    For Each P In Selection.ShapeRange 'ActiveSheet.Shapes
        Dim cell As Range: Set cell = Cells(P.TopLeftCell.Row, P.TopLeftCell.Column)
        With P
            If ans = vbYes Then
                .LockAspectRatio = True
                If .Height > .Width Then
                    .Height = cell.Height - (cell.Height * Shrink)
                Else
                    .Width = cell.Width - (cell.Width * Shrink)
                End If
            Else
                .LockAspectRatio = False
                If .Height > .Width Then
                    .Width = cell.Width - (cell.Width * Shrink)
                    .Height = cell.Height - (cell.Height * Shrink)
                Else
                    .Height = cell.Height - (cell.Height * Shrink)
                    .Width = cell.Width - (cell.Width * Shrink)
                End If
            End If
            '//aspectratio locked,set only height or width
            .Top = cell.MergeArea.Top + (cell.MergeArea.Height - .Height) / 2
            .Left = cell.MergeArea.Left + (cell.MergeArea.Width - .Width) / 2
        End With
    Next

End Sub

Sub ShapesOutsideVisibleRange()
    If ActiveSheet.Shapes.Count = 0 Then
        MsgBox "No shapes in active sheet"
        Exit Sub
    End If
    
    Dim s As Shape
    Dim rngholder As String
    For Each s In ActiveSheet.Shapes
        If Range(s.BottomRightCell.Address).Row > ActiveWindow.VisibleRange.Rows.Count Then
            rngholder = _
                      rngholder & Chr(10) & s.BottomRightCell.Address
        End If
    Next s
    
    If rngholder = "" Then
        MsgBox "No shape out of range"
        Exit Sub
    End If
    
    Dim Arr
    Arr = Split(rngholder, Chr(10))
    Dim lastSposition As String
    lastSposition = Arr(UBound(Arr))
    If Range(lastSposition).Row > ActiveWindow.VisibleRange.Rows.Count Then
        MsgBox "There are shapes after the last visible row." _
             & Chr(10) & "Their BottomRight cells span the following ranges: " _
             & rngholder
    Else
        MsgBox "All shapes positioned inside visible range"
    End If
End Sub


Sub PasteAsPicture()
If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a range before running the macro."
Exit Sub
End If
    For i = 1 To Selection.Areas.Count
        Application.CutCopyMode = False
        Selection.Areas(i).Copy
        ActiveSheet.Pictures.Paste
    Next
        Application.CutCopyMode = False
End Sub

Sub PasteAsLinkedPicture()
 If TypeName(Selection) <> "Range" Then
         MsgBox "Please select a range before running the macro."
Exit Sub
End If
Dim coll As New Collection
    For i = 1 To Selection.Areas.Count
        coll.Add Selection.Areas(i).Address
    Next
    Dim element As Variant
    Range(coll(1)).Select
    For Each element In coll
           Application.CutCopyMode = False
        Range(element).Copy
        ActiveSheet.Pictures.Paste link:=True
    Next
        Application.CutCopyMode = False
End Sub

