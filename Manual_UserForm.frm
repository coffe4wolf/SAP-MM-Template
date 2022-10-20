VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Manual_UserForm 
   ClientHeight    =   10155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15885
   OleObjectBlob   =   "Manual_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Manual_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub BackwardInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    BackwardInactive_ImageButton.Visible = False
End Sub

Private Sub ForwardInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardInactive_ImageButton.Visible = False
End Sub

Private Sub ForwardActive_ImageButton_Click()

    If SlidesCounter < 9 Then
        SlidesCounter = SlidesCounter + 1
    End If
    
    Call ShowPicture(SlidesCounter)
    
End Sub

Private Sub BackwardActive_ImageButton_Click()

    If SlidesCounter > 0 Then
        SlidesCounter = SlidesCounter - 1
    End If
    
    Call ShowPicture(SlidesCounter)
    
End Sub

Public Function PictureFromShape(ByVal S As Shape) As IPicture

  S.CopyPicture xlScreen, xlBitmap
  Set PictureFromShape = PictureFromClipboard
  
End Function

Private Sub ShowPicture(PictureNum As Integer)

    With Manual_UserForm
        Select Case PictureNum
            Case 1
                .Slide1_Image.Visible = True
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = False
            Case 2
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = True
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = False
            Case 3
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = True
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = False
            Case 4
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = True
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = False
            Case 5
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = True
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = False
            Case 6
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = True
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = False
            Case 7
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = True
                .Slide8_Image.Visible = False
            Case 8
                .Slide1_Image.Visible = False
                .Slide2_Image.Visible = False
                .Slide3_Image.Visible = False
                .Slide4_Image.Visible = False
                .Slide5_Image.Visible = False
                .Slide6_Image.Visible = False
                .Slide7_Image.Visible = False
                .Slide8_Image.Visible = True
        End Select
    End With
    
End Sub

Private Sub Slide1_Image_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub UserForm_Initialize()
    Dim SlidesCounter As Integer: SlidesCounter = 1
    Call ShowPicture(SlidesCounter)
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ForwardInactive_ImageButton.Visible = True
    BackwardInactive_ImageButton.Visible = True
End Sub
