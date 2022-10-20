VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseLanguage_UserForm 
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3495
   OleObjectBlob   =   "ChooseLanguage_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChooseLanguage_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Language_Russian As String = "Russian"
Const Language_English As String = "English"


Private Sub UserForm_Initialize()

    With ChooseLanguage_ComboBox
    
        .AddItem Language_Russian
        .AddItem Language_English
    
    End With

End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = vbFormControlMenu Then
    
        Cancel = True
        MsgBox ("Choose language.")
    
    End If
    
End Sub

Private Sub CommandButton1_Click()

    MsgBox ("Choose language.")
    
End Sub


Private Sub ChooseLanguage_CancelActive_ImageButton_Click()

    MsgBox ("Choose language.")

End Sub

Private Sub ChooseLanguage_OKActive_ImageButton_Click()

    If ChooseLanguage_ComboBox.Value <> "" Then
    
        Current_Language = ChooseLanguage_ComboBox.Value
        Unload ChooseLanguage_UserForm
        
    Else
    
        MsgBox ("Choose language.")
        
    End If

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    ChooseLanguage_OKInactive_ImageButton.Visible = True
    ChooseLanguage_CancelInactive_ImageButton.Visible = True

End Sub

Private Sub ChooseLanguage_OKInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    ChooseLanguage_OKInactive_ImageButton.Visible = False
    ChooseLanguage_CancelInactive_ImageButton.Visible = True
    
End Sub

Private Sub ChooseLanguage_CancelInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    ChooseLanguage_CancelInactive_ImageButton.Visible = False
    ChooseLanguage_OKInactive_ImageButton.Visible = True

End Sub


