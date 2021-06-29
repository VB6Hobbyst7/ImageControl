VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uImageControl 
   Caption         =   "ImageControl"
   ClientHeight    =   5256
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2544
   OleObjectBlob   =   "uImageControl.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uImageControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Call SelectShapesByName
End Sub

Private Sub CommandButton10_Click()
    Call InsertPictures
End Sub

Private Sub CommandButton11_Click()
    Call ExportAsImage
End Sub

Private Sub CommandButton12_Click()

End Sub

Private Sub CommandButton13_Click()
    Call InsertImageInActivecellComment
End Sub

Private Sub CommandButton14_Click()
    Call CircleBoxADD
End Sub

Private Sub CommandButton15_Click()
    Call CircleBoxREMOVE
End Sub

Private Sub CommandButton16_Click()
    Call PasteAsPicture
End Sub

Private Sub CommandButton17_Click()
    Call PasteAsLinkedPicture
End Sub

Private Sub CommandButton18_Click()
    Call ExportShapeAsPicture
End Sub

Private Sub CommandButton19_Click()
Call ShapesOutsideVisibleRange
End Sub

Private Sub CommandButton2_Click()
    Call SelectShapesByText
End Sub

Private Sub CommandButton3_Click()
    Call SelectShapesWithinSelectedRange
End Sub

Private Sub CommandButton4_Click()
    Call PicturesFitCenter
End Sub

Private Sub CommandButton5_Click()
    Call TextBoxResizeTB
End Sub

Private Sub CommandButton6_Click()
    Call GridHorizontal
End Sub

Private Sub CommandButton7_Click()
    Call GridVertical
End Sub



Private Sub SpinButton1_Change()
lbShrink.Caption = SpinButton1.Value
Shrink = CInt(lbShrink.Caption) * 0.1
End Sub


Private Sub UserForm_Initialize()
    ComboBox1.List = ThisWorkbook.Sheets("SETTINGS").Range("$D$2:$D$7").Value
    ComboBox1.ListIndex = 0
    Shrink = CInt(lbShrink.Caption) * 0.1
End Sub

