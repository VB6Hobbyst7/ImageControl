Attribute VB_Name = "ExportImport"



Sub ExportShapeAsPicture()
'PURPOSE: Save a selected shape/icon as a PNG file to computer's desktop
'SOURCE: www.thespreadsheetguru.com

Dim cht As ChartObject
Dim ActiveShape As Shape

Dim ext As String
ext = uImageControl.ComboBox1.Text

Dim Path As String
Path = "C:\Users\" & Environ$("Username") & "\Pictures" & "\ExportedImages\"
On Error Resume Next
MkDir Path
On Error GoTo 0

'Ensure a Shape is selected
'Dim UserSelection As Variant
'  On Error GoTo NoShapeSelected
'    Set UserSelection = ActiveWindow.Selection
'    Set ActiveShape = ActiveSheet.Shapes(UserSelection.Name)
'  On Error GoTo 0
If TypeName(Selection) = "Range" Then GoTo NoShapeSelected


For Each ActiveShape In Selection.ShapeRange
'Create a temporary chart object (same size as shape)
  Set cht = ActiveSheet.ChartObjects.Add( _
    Left:=ActiveCell.Left, _
    Width:=ActiveShape.Width, _
    Top:=ActiveCell.Top, _
    Height:=ActiveShape.Height)

'Format temporary chart to have a transparent background
  cht.ShapeRange.Fill.Visible = msoFalse
  cht.ShapeRange.Line.Visible = msoFalse
    
'Copy/Paste Shape inside temporary chart
  ActiveShape.Copy
  cht.Activate
  ActiveChart.Paste
  
'Save chart to User's Desktop as PNG File
  cht.Chart.Export Path & ActiveShape.Name & "." & ext

'Delete temporary Chart
  cht.Delete

'Re-Select Shape (appears like nothing happened!)
  ActiveShape.Select

Next ActiveShape

Exit Sub

'ERROR HANDLERS
NoShapeSelected:
        MsgBox "Please select shapes before running the macro."
  Exit Sub

End Sub


Sub ExportAsImage()
    If Not TypeName(Selection) = "Range" Then
            MsgBox "Please select shapes before running the macro."
        Exit Sub
    End If
    Dim cell As Range
    Dim ext As String
    ext = uImageControl.ComboBox1.Text

    Dim action As Long
    action = MsgBox("(YES) = for each area in selection" & Chr(10) & _
                    "(NO) = for each cell in selection", vbYesNoCancel)
    If action = vbCancel Then Exit Sub
    
    Dim ExportFolder As String
    ExportFolder = "C:\Users\" & Environ$("Username") & "\Pictures" & "\ExportedImages\" 'Environ$ ("USERPROFILE") & "\Downloads\ExportedImages\"

        On Error Resume Next
        MkDir ExportFolder
        On Error GoTo 0
        
    On Error Resume Next                         'goto 0
    Application.DisplayAlerts = False

    Select Case action
    Case Is = vbNo
        For Each cell In Selection
            Call ExportRangeAsImage(ActiveSheet, cell, ExportFolder, cell.Value, ext)
            Application.Wait (Now + TimeValue("0:00:01"))
        Next cell

    Case Is = vbYes
    Dim result As String
        For i = 1 To Selection.Areas.Count
            result = ""
            result = InputBox("name for image of area: " & Selection.Areas(i).Address)
            If CStr(result) = "" Then result = Format(Now, "hhmmss")
            Call ExportRangeAsImage(ActiveSheet, Selection.Areas(i), ExportFolder, result, ext)
            Application.Wait (Now + TimeValue("0:00:01"))
        Next i
    End Select
    Application.DisplayAlerts = True
    
    Shell "explorer.exe" & " " & ExportFolder, vbNormalFocus
End Sub

' Procedure : ExportRangeAsImage
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Capture a picture of a worksheet range and save it to disk
'               Returns True if the operation is successful
' Note      : *** Overwrites files, if already exists, without any warning! ***
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' ws            : Worksheet to capture the image of the range from
' rng           : Range to capture an image of
' sPath         : Fully qualified path where to export the image to
' sFileName     : filename to save the image to WITHOUT the extension, just the name
' sImgExtension : The image file extension, commonly: JPG, GIF, PNG, BMP
'                   If omitted will be JPG format
'
' Usage:
' ~~~~~~
' ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01". "JPG")
' ? ExportRangeAsImage(Sheets("Products"), Range("D5:F23"), "C:\Temp\Charts", "test02")
' ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01", "PNG")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2020-04-06              Initial Release
'---------------------------------------------------------------------------------------
Function ExportRangeAsImage(ws As Worksheet, _
                            rng As Range, _
                            sPath As String, _
                            sFileName As String, _
                            Optional sImgExtension As String = "JPG") As Boolean
    Dim oChart                As ChartObject
 
    On Error GoTo Error_Handler
 
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
 
    Application.ScreenUpdating = False
    ws.Activate
    rng.CopyPicture xlScreen, xlPicture          'Copy Range Content
    Set oChart = ws.ChartObjects.Add(0, 0, rng.Width, rng.Height) 'Add chart
    oChart.Activate
    With oChart.Chart
        .Paste                                   'Paste our Range
        .Export sPath & sFileName & "." & LCase(sImgExtension), sImgExtension 'Export the chart as an image
    End With
    oChart.Delete                                'Delete the chart
    ExportRangeAsImage = True
 
Error_Handler_Exit:
    On Error Resume Next
    Application.ScreenUpdating = True
    If Not oChart Is Nothing Then Set oChart = Nothing
    Exit Function
 
Error_Handler:
    '76 - Path not found
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: ExportRangeAsImage" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function



Sub InsertPictures()
    'Update 20140513
    Dim PicList() As Variant
    Dim PicFormat As String
    Dim rng As Range
    Dim sShape As Shape

    On Error Resume Next
    PicList = Application.GetOpenFilename(PicFormat, multiSelect:=True)
    xColIndex = Application.ActiveCell.Column
    If IsArray(PicList) Then
        xRowIndex = Application.ActiveCell.Row
        For lLoop = LBound(PicList) To UBound(PicList)
            Set rng = Cells(xRowIndex, xColIndex)
            Set sShape = ActiveSheet.Shapes.AddPicture(PicList(lLoop), msoFalse, msoCTrue, _
                                                       rng.Left, rng.Top, -1, -1)
            With sShape
                .LockAspectRatio = msoTrue       'set only height or width
                If .Height > .Width Then
                    .Height = rng.Height - (rng.Height * Shrink)
                Else
                    .Width = rng.Width - (rng.Width * Shrink)
                End If
                .Top = rng.MergeArea.Top + (rng.MergeArea.Height - .Height) / 2
                .Left = rng.MergeArea.Left + (rng.MergeArea.Width - .Width) / 2
            End With
            '           /to change cell height and width instead of pic h and w
            '            Rows(cell.Row).RowHeight = .height + (2 * Shrink)
            '            Columns(cell.Column).ColumnWidth = .width + (2 * Shrink)

            xRowIndex = xRowIndex + 1
        Next
    End If
End Sub


Sub InsertImageInActivecellComment()
If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell before running the macro."
Exit Sub
End If
Dim cell As Range
    Dim cmt As Comment
    Dim PicPath As String
    Dim str As String
    
    Dim myObj As Object
    Dim myDirString As String
    Set myObj = Application.FileDialog(msoFileDialogFilePicker)
    With myObj
        .InitialFileName = "C:\Users\" & Environ$("Username") & "\Pictures"
        .Filters.Add "Images", "*.png, *jpeg, *.jpg, *.gif, *.ico, *.cur, *.wmf"
        .FilterIndex = 2
        If .Show = False Then MsgBox "No picture selected", vbExclamation: Exit Sub
        PicPath = .SelectedItems(1)
    End With

    On Error Resume Next
    Set cell = Selection.MergeArea
    
    With cell
        If .Comment Is Nothing Then
            Set cmt = .AddComment
            str = cmt.Text
        Else
        Set cmt = .Comment
        str = cmt.Text
        End If
    End With
        With cmt
            .Text ((Replace(str, Application.UserName & ":", "")))
            .Shape.Fill.UserPicture PicPath
            .Visible = False
        End With

End Sub
