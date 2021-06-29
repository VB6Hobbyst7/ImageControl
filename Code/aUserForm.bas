Attribute VB_Name = "aUserForm"
Sub ShowImageControl()
    If Not IsLoaded("me") Then
        Call OpenUserForm("uImageControl")
    End If
End Sub

Sub OpenUserForm(formName As String)
    Dim oUserForm As Object
    On Error GoTo err
    Set oUserForm = UserForms.Add(formName)
    oUserForm.Show (vbModeless)
    Exit Sub
err:
    Select Case err.Number
    Case 424:
        MsgBox "The Userform with the name " & formName & " was not found.", vbExclamation, "Load userforn by name"
    Case Else:
        MsgBox err.Number & ": " & err.Description, vbCritical, "Load userforn by name"
    End Select
End Sub

Sub AddCommandbarImageControl()
    On Error Resume Next                         'Just in case
    'Delete any existing menu item that may have been left.
    Dim bar As CommandBarControl
    For Each bar In Application.CommandBars("Worksheet Menu Bar").Controls
        If bar.Caption = "ImageControl" Then bar.Delete
        'Debug.Print bar.Caption
    Next

    '    Application.CommandBars("Worksheet Menu Bar").Controls("uNotes").Delete
    
    'Add the new menu item and set a CommandBarButton variable to it
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
    With cControl
        .Caption = "ImageControl"
        .Style = msoButtonIconAndCaption
        .FaceId = 2619
        .OnAction = "ShowImageControl"               'Macro stored in a Standard Module
    End With
    On Error GoTo 0
End Sub

Function IsLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

