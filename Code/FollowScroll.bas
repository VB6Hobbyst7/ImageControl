Attribute VB_Name = "FollowScroll"
Dim eTime As Variant

Sub ScreenRefresh()
    Dim s As Shape
    For Each s In Workbooks("").Sheets("Sheet1")
        '        s.left = ThisWorkbook.Windows(1).VisibleRange.left ' .(2, 2).left
        s.Top = ThisWorkbook.Windows(1).VisibleRange.Top ' .(2, 2).top
    Next s

    '    With ThisWorkbook.Worksheets("Sheet1").Shapes(1)
    '        .left = ThisWorkbook.Windows(1).VisibleRange(2, 2).left
    '        .top = ThisWorkbook.Windows(1).VisibleRange(2, 2).top
    '    End With
End Sub

Sub StartTimedRefresh()
    Call ScreenRefresh
    eTime = Now + TimeValue("00:00:01")
    Application.OnTime eTime, "StartTimedRefresh"
End Sub

Sub StopTimer()
    Application.OnTime eTime, "StartTimedRefresh", , False
End Sub

'Add the following code in Sheet1 (where the shapes are in)
'
'Private Sub Worksheet_Activate()
'    Call StartTimedRefresh
'End Sub
'
'Private Sub Worksheet_Deactivate()
'    Call StopTimer
'End Sub
