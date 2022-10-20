VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RequestForm_UserForm 
   ClientHeight    =   9510
   ClientLeft      =   180
   ClientTop       =   690
   ClientWidth     =   10665
   OleObjectBlob   =   "RequestForm_UserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RequestForm_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Language_Russian As String = "Russian"
Const Language_English As String = "English"
Const HeaderRowsToSkip As Integer = 10

Private Sub ChangeRow_SpinButton_SpinUp()
    ' For spinbutton control.
    ' You can use this if you have to change
    ' up-down buttons to spinbutton instead.
    
    Call ChangeRow

End Sub
Private Sub ChangeRow_SpinButton_SpinDown()
    ' For spinbutton control.
    ' You can use this if you have to change
    ' up-down buttons to spinbutton instead.
    
    Call ChangeRow

End Sub

Private Sub Article_Label_Click()

End Sub

Private Sub Article_TextBox_Change()
    
    ChangedNotSaved = True
    
End Sub



Private Sub ClearCurrentRowInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    ClearCurrentRowInactive_ImageButton.Visible = False
    SaveInactive_ImageButton.Visible = True
    
End Sub

Private Sub ClearFormActive_ImageButton_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub ClearFullDescriptionActive_ImageButton_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub ClearFullDescriptionInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.

    ClearFullDescriptionInactive_ImageButton.Visible = False
    ClearFormInactive_ImageButton.Visible = True
    
End Sub
Private Sub ClearFormInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    ClearFormInactive_ImageButton.Visible = False
    ClearFullDescriptionInactive_ImageButton.Visible = True
    
End Sub




Private Sub CloseInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    CloseInactive_ImageButton.Visible = False
    
End Sub

Private Sub CopyNextRowInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    CopyNextRowInactive_ImageButton.Visible = False
    CopyPreviousRowInactive_ImageButton.Visible = True
    
End Sub

Private Sub CopyPreviousRowInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    CopyPreviousRowInactive_ImageButton.Visible = False
    CopyNextRowInactive_ImageButton.Visible = True
    
End Sub


Private Sub RowDownInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    RowDownInactive_ImageButton.Visible = False
    RowUpInactive_ImageButton.Visible = True
    
End Sub

Private Sub RowUpInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    RowUpInactive_ImageButton.Visible = False
    RowDownInactive_ImageButton.Visible = True
    
End Sub

Private Sub SaveActive_ImageButton_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub SaveInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Button hovering effect.
    
    SaveInactive_ImageButton.Visible = False
    ClearCurrentRowInactive_ImageButton.Visible = True
    
End Sub

Private Sub CriticalMaterial_CheckBox_Change()

    ChangedNotSaved = True
    
End Sub

Private Sub FullDescription_TextBox_Change()
    
    ChangedNotSaved = True
    
End Sub

Private Sub GroupCode_TextBox_Change()
    
    ChangedNotSaved = True
    
End Sub
Private Sub ClearCurrentRowActive_ImageButton_Click()

     Call ClearRequestForm(True)
     ChangedNotSaved = False
     
End Sub


Private Sub ClearFormActive_ImageButton_Click()
    
    Call ClearRequestForm(False)
    ChangedNotSaved = True
    
End Sub

Private Sub ClearFullDescriptionActive_ImageButton_Click()
    
    FullDescription_TextBox.Value = ""
    ChangedNotSaved = True

End Sub

Private Sub CloseActive_ImageButton_Click()
    
    AnswerCloseEditMaterialForm = MsgBox(CloseFormMessage, vbYesNo + vbQuestion, WarningMessage)
    
    If AnswerCloseEditMaterialForm = vbYes Then
        Unload Me
    End If
    
    Dim RowCounter As Integer, _
        ColCounter As Integer
        
    With Ws_Request
    
        For RowCounter = Constant_RowIndentConstant + 1 To GetBorders("LR", .Name)
        
            .Range("A" & RowCounter).EntireRow.AutoFit
            
        Next RowCounter
        
        For ColCounter = 2 To 20
        
            .Cells(1, ColCounter).EntireColumn.AutoFit
            
        Next ColCounter
        
        .Columns("R").Hidden = True
        .Columns("S").Hidden = True
        .Columns("T").Hidden = True
        
    End With
    
End Sub


Private Sub CopyPreviousRowActive_ImageButton_Click()

    Dim PrevRow As Integer: PrevRow = RowNumberForEdit - 1
    
    If RowNumberForEdit_TextBox.Value >= 2 Then
    
        If RowIsNotEmpty(PrevRow) Then
        
            Call CopyRow(PrevRow)
            
        Else
        
            MsgBox (NothingToCopyMessage)
            
        End If
        
    Else
    
        MsgBox (FirstRowMessage)
    
    End If

End Sub

Private Sub CopyNextRowActive_ImageButton_Click()

    Dim NextRow As Integer: NextRow = RowNumberForEdit + 1
    
    If RowIsNotEmpty(NextRow) = True Then
    
        Call CopyRow(RowNumberForEdit + 1)
        
    Else
    
        MsgBox (NothingToCopyMessage)
        
    End If

End Sub


Private Sub MaterialType_ComboBox_Change()

    If MaterialType_ComboBox.Value = MaterialTypeStandartized Then
    
        MaterialTypeNote_Label.Caption = MaterialTypeNoteLabelCaption
        ShortDescriptionTemplate_Label = ShortDescriptionTemplate1LabelCaption
        ProductCode_Label = ProductCodeLabelCaption
        ProductCode_Label.ForeColor = RGB(70, 70, 70)
        Article_Label.Caption = ArticleLabelCaption
        Article_Label.ForeColor = RGB(70, 70, 70)
        
    ElseIf MaterialType_ComboBox.Value = MaterialTypeManufactured Then
    
        MaterialTypeNote_Label.Caption = ""
        ShortDescriptionTemplate_Label = ShortDescriptionTemplate2LabelCaption
        ProductCode_Label = ProductCode1LabelCaption
        ProductCode_Label.ForeColor = RGB(255, 0, 0)
        Article_Label.Caption = Article1LabelCaption
        Article_Label.ForeColor = RGB(255, 0, 0)

    ElseIf MaterialType_ComboBox.Value = MaterialTypeFQT Then
    
        MaterialTypeNote_Label.Caption = ""
        ShortDescriptionTemplate_Label = MaxDescriptionLengthLabelCatpion
        ProductCode_Label = ProductCodeLabelCaption
        ProductCode_Label.ForeColor = RGB(70, 70, 70)
        Article_Label = ArticleLabelCaption
        Article_Label.ForeColor = RGB(70, 70, 70)
        
    Else
    
        MaterialTypeNote_Label.Caption = ""
        ShortDescriptionTemplate_Label = ""
        ProductCode_Label = ProductCodeLabelCaption
        ProductCode_Label.ForeColor = RGB(70, 70, 70)
        Article_Label = ArticleLabelCaption
        Article_Label.ForeColor = RGB(70, 70, 70)
        
    End If
    
    ChangedNotSaved = True
    
End Sub


Private Sub MaxPrice_TextBox_Change()

    ChangedNotSaved = True
    
End Sub


Private Sub PurchasingGroup_TextBox_Change()
    
    ChangedNotSaved = True
    
End Sub

Private Sub SelectManufacturerCodeActive_ImageButton_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub SelectManufacturerCodeActive_ImageButton_Click()

    If Current_Language = Language_Russian Then
    
        'MultiSelectionForm.Caption = MultiSelectionFormCaption
        MultiSelectionForm.ListBox1.MultiSelect = fmMultiSelectSingle
        Call Populate2DimListBox(Ws_MasterManufacturerCode.Name)
        MultiSelectionForm.Show
    
    ElseIf Current_Language = Language_English Then
    
        MultiSelectionForm_Eng.ListBox1.MultiSelect = fmMultiSelectSingle
        Call Populate2DimListBox(Ws_MasterManufacturerCode.Name, , Language_English)
        MultiSelectionForm_Eng.Show
    
    End If
    
End Sub

Private Sub SelectManufacturerCodeInactive_ImageButton_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    SelectManufacturerCodeInactive_ImageButton.Visible = False
    
End Sub

Private Sub Serialization_Combobox_Click()
    
    ChangedNotSaved = True
    
End Sub



Private Sub ShortNameRus_TextBox_AfterUpdate()

    If MaterialType_ComboBox.Value = MaterialTypeFQT Then
        ShortNameRus_TextBox.Value = MaterialShortNameStringToFQT(ShortNameRus_TextBox.Text, "Rus")
        ChangedNotSaved = True
    End If
    
End Sub

Private Sub ShortNameEng_TextBox_AfterUpdate()

    If MaterialType_ComboBox.Value = MaterialTypeFQT Then
        ShortNameEng_TextBox.Value = MaterialShortNameStringToFQT(ShortNameEng_TextBox.Text, "Eng")
        ChangedNotSaved = True
    End If
    
End Sub

Private Sub Unit_ComboBox_Change()
    
    ChangedNotSaved = True
    
End Sub

Private Sub UrgentRequest_CheckBox_Change()

    ChangedNotSaved = True
    
End Sub

Private Sub UserForm_Initialize()


    If Current_Language = Language_Russian Then
        Article_Label.Caption = Ws_localizationObjects.Range("Article_Label").Value
        ProductCode_Label.Caption = Ws_localizationObjects.Range("ProductCode_Label").Value
    ElseIf Current_Language = Language_English Then
        Article_Label.Caption = Ws_localizationObjects.Range("Article_LabelEn").Value
        ProductCode_Label.Caption = Ws_localizationObjects.Range("ProductCode_LabelEn").Value
    End If
 
    RequestForm_UserForm.ProductCode_Label.Visible = True

    ChangedNotSaved = False
    
    Dim counter As Long
    With RequestForm_UserForm
        .GroupCategory_Combobox.Clear
        .UrgentRequest_CheckBox = False
        .Unit_ComboBox.Clear
        
        .ShortNameRus_TextBox.MaxLength = 40
        .ShortNameEng_TextBox.MaxLength = 40
        .ShortDescriptionTemplate_Label.Caption = MaxDescriptionLengthLabelCatpion
        
        .ProductCode_TextBox.MaxLength = 10
    End With
    
    ' Init combobox "Group Category"
    With Ws_MasterGroupCategory
        For counter = 2 To GetBorders("LR", .Name, Wb_Current)
            RequestForm_UserForm.GroupCategory_Combobox.AddItem .Cells(counter, 2).Value & " | " & .Cells(counter, 1).Value
        Next counter
    End With
    
    With Ws_helpSheet
        ' Init combobox "Unit"
        For counter = 2 To 34
            RequestForm_UserForm.Unit_ComboBox.AddItem .Cells(counter, 2).Value
        Next counter
        
        ' Init combobox "Material type"
        For counter = 2 To 4
            MaterialType_ComboBox.AddItem .Cells(counter, 3)
        Next counter

        ' Init combobox "Serialization"
        For counter = 2 To 3
            RequestForm_UserForm.Serialization_ComboBox.AddItem .Cells(counter, 4).Value
        Next counter
    End With
    
    RowNumberForEdit_TextBox.Font.Size = 14
    RowNumberForEdit_TextBox.Font.Bold = True

End Sub

Private Sub GroupCategory_Combobox_Change()
    
    If GroupCategory_Combobox.Value <> "" Then

        Dim MasterGroup_LR As Long: MasterGroup_LR = GetBorders("LR", Ws_MasterGroup.Name, Wb_Current)
        
        GroupCategoryCode = RxMatch(GroupCategory_Combobox.Value, "Z\d+(?=$)", True, False)
        
        RequestForm_UserForm.Group_ComboBox.Clear
        GroupCode_TextBox.Value = ""
        PurchasingGroup_TextBox.Value = ""
        RequestForm_UserForm.Repaint
        
        With Ws_MasterGroup
            ' Init combobox "Group"
            For counter = 2 To MasterGroup_LR
                If Trim(.Cells(counter, 1).Value) = GroupCategoryCode Then
                    If Trim(.Cells(counter, 4).Value) <> "" And Trim(.Cells(counter, 3).Value) <> "" Then
                        RequestForm_UserForm.Group_ComboBox.AddItem .Cells(counter, 4).Value & " | " & .Cells(counter, 3).Value
                    End If
                End If
            Next counter
            
        End With

        ChangedNotSaved = True

    End If
    
    ' Reset "Full description".
    FullDescription_TextBox.Value = ""
    
End Sub

Private Sub Group_ComboBox_Change()

    Dim MasterGroup_LR As Long: MasterGroup_LR = GetBorders("LR", Ws_MasterGroup.Name, Wb_Current)
    Dim MasterGroupMap_LR As Long: MasterGroupMap_LR = GetBorders("LR", WS_MasterGroupMap.Name, Wb_Current)

    If Group_ComboBox.Value <> "" Then
    
        GroupCode = RxMatch(Group_ComboBox.Value, "Z\d+(?=$)", True, False)
        
        FullDescription = ""
        RequestForm_UserForm.Repaint
        
        
        
        ' Init textbox "Purchasing group" after changing Group combobox
         With WS_MasterGroupMap
         
            For counter = 2 To MasterGroupMap_LR
                If Trim(.Cells(counter, 1).Value) = GroupCode Then
                    PurchasingGroup_TextBox.Value = CStr(.Cells(counter, 2))
                    Exit For
                End If
            Next counter
            
        End With
        
        GroupCode_TextBox.Value = GroupCode
        
        ' Init textbox "Full description"
        FullDescription_TextBox.Value = ""
        
        With Ws_MasterAttributes
            For counter = 2 To GetBorders("LR", .Name, Wb_Current)
                If .Cells(counter, 1).Value = GroupCode Then
                    FullDescription = Replace(FullDescription & .Cells(counter, 2).Value, "\n", "") & ": " & vbNewLine & vbNewLine
                End If
            Next counter
        End With
        
        ' Filling "Full description" textbox with line breaks
        ' between values.
        If Len(FullDescription) > 2 Then
            FullDescription_TextBox.Value = Left(FullDescription, Len(FullDescription) - 2)
        End If
        
        ChangedNotSaved = True
        
    End If
    
End Sub

Private Sub RowUpActive_ImageButton_Click()

   Dim ChangeRowAnswer As Integer
    If ChangeRowAnswer = vbYes Then
    Else
    End If
    
    If ChangedNotSaved = True Then
    
        ChangeRowAnswer = MsgBox(SaveChangesWarningMessageCaption, vbYesNo + vbQuestion)
        If ChangeRowAnswer = vbYes Then
            If RowNumberForEdit_TextBox.Value > 1 Then
                RowNumberForEdit_TextBox.Value = RowNumberForEdit_TextBox.Value - 1
                RowNumberForEdit = RowNumberForEdit_TextBox.Value + Constant_RowIndentConstant
            End If
            
        'Load data from the row on the sheet into UserForm.
        InitRow (RowNumberForEdit)
        
        End If
        
    ElseIf ChangedNotSaved = False Then
    
        If RowNumberForEdit_TextBox.Value > 1 Then
            RowNumberForEdit_TextBox.Value = RowNumberForEdit_TextBox.Value - 1
            RowNumberForEdit = RowNumberForEdit_TextBox.Value + Constant_RowIndentConstant
        End If
    
        'Load data from the row on the sheet into UserForm.
        InitRow (RowNumberForEdit)
        
    End If
    
End Sub
Private Sub RowDownActive_ImageButton_Click()
    
   Dim ChangeRowAnswer As Integer
    If ChangeRowAnswer = vbYes Then
    Else
    End If
    
    If ChangedNotSaved = True Then
    
        ChangeRowAnswer = MsgBox(SaveChangesWarningMessageCaption, vbYesNo + vbQuestion)
        If ChangeRowAnswer = vbYes Then
            RowNumberForEdit_TextBox.Value = RowNumberForEdit_TextBox.Value + 1
            RowNumberForEdit = RowNumberForEdit_TextBox.Value + Constant_RowIndentConstant
            
            'Load data from the row on the sheet into UserForm.
            InitRow (RowNumberForEdit)
            
        End If
        
    ElseIf ChangedNotSaved = False Then
    
        RowNumberForEdit_TextBox.Value = RowNumberForEdit_TextBox.Value + 1
        RowNumberForEdit = RowNumberForEdit_TextBox.Value + Constant_RowIndentConstant
        
        'Load data from the row on the sheet into UserForm.
        InitRow (RowNumberForEdit)
        
    End If
     
End Sub

Private Sub SaveActive_ImageButton_Click()
    'Write data from UserForm to worksheet.
    
    If MaterialType_ComboBox.Value <> MaterialTypeFQT Then
    ' Check entered price is correct (contains digits only).
    If Not StringIsPrice(MaxPrice_TextBox.Value) Or MaxPrice_TextBox.Value = "" Then
        Debug.Print MaxPrice_TextBox.Value
        MsgBox (WrongEnteredPriceMessageCaption)
        Exit Sub
        
    End If
    
    ' Check Manufacturer code is numerical.
        If Not IsNumeric(ProductCode_TextBox.Value) And Trim(ProductCode_TextBox.Value) <> "" Then
            MsgBox (ProductCodeOnlyDigitsWarningMessage)
            ProductCode_TextBox.Value = ""
            ChangedNotSaved = True
            Exit Sub
        End If
    End If
    
    If MaterialType_ComboBox.Value = MaterialTypeManufactured And (Article_TextBox.Value = "" Or ProductCode_TextBox.Value = "") Then
    
        MsgBox (EmptyProductCodeAndArticleMessageCaption)
    
    ElseIf MaterialType_ComboBox.Value = MaterialTypeStandartized And FullDescription_TextBox.Value = "" Then
        
        MsgBox (EmptyFullDescriptionMessageCaption)
        
    ElseIf MaterialType_ComboBox.Value = "" Then

        MsgBox (EmptyMaterialTypeMessageCaption)
        
    ElseIf IsNull(RxMatch(ShortNameEng_TextBox.Value, "^[a-zA-Z0-9\(\)\*\_\-\#\$\%\^\&\*\,\.\'\[\]\{\}\?\<\>\|\\\/\=\@\!\+\s]*$", True, False)) Then
    
        MsgBox (NoCyrillicCharsInEngDescriptionCaption)
    
    Else
        
        ' Add new properties to save on sheet here.
        With Ws_Request
            .Cells(RowNumberForEdit, 1) = RowNumberForEdit - HeaderRowsToSkip
            .Cells(RowNumberForEdit, 2) = PriorityToString(UrgentRequest_CheckBox.Value)
            .Cells(RowNumberForEdit, 3) = ShortNameRus_TextBox.Value
            .Cells(RowNumberForEdit, 4) = FullDescription_TextBox.Value
            .Cells(RowNumberForEdit, 5) = ShortNameEng_TextBox.Value
            .Cells(RowNumberForEdit, 6) = Unit_ComboBox.Value
            .Cells(RowNumberForEdit, 7) = ProductCode_TextBox.Value
            .Cells(RowNumberForEdit, 8) = Article_TextBox.Value
            .Cells(RowNumberForEdit, 9) = MaxPrice_TextBox.Value
            .Cells(RowNumberForEdit, 10) = GroupCode_TextBox.Value
            .Cells(RowNumberForEdit, 11) = PurchasingGroup_TextBox.Value
            .Cells(RowNumberForEdit, 12) = CheckboxToCell(CriticalMaterial_CheckBox.Value)
            .Cells(RowNumberForEdit, 13) = Serialization_ComboBox.Value
            .Cells(RowNumberForEdit, 18) = GroupCategory_Combobox.Value
            .Cells(RowNumberForEdit, 19) = Group_ComboBox.Value
            .Cells(RowNumberForEdit, 20) = MaterialType_ComboBox.Value
            .Cells(RowNumberForEdit, 21) = CheckboxToCell(BatchManagement_CheckBox.Value)
            .Columns("R").Hidden = True
            .Columns("S").Hidden = True
            .Columns("T").Hidden = True
        End With
        
        MsgBox (ChangesSavedMessageCaption)
        
        ChangedNotSaved = False
    
    End If
    
End Sub


Sub ClearRequestForm(ClearRowData As Boolean)
    ' Clear form and slected row on a worksheet.

    With RequestForm_UserForm
        .GroupCategory_Combobox.Value = ""
        .Group_ComboBox.Value = ""
        .MaterialType_ComboBox.Value = ""
        .UrgentRequest_CheckBox.Value = False
        .ShortNameRus_TextBox.Value = ""
        .ShortNameEng_TextBox.Value = ""
        .Unit_ComboBox.Value = ""
        .ProductCode_TextBox.Value = ""
        .Article_TextBox.Value = ""
        .GroupCode_TextBox.Value = ""
        .MaxPrice_TextBox.Value = ""
        .FullDescription_TextBox.Value = ""
        .PurchasingGroup_TextBox.Value = ""
        .CriticalMaterial_CheckBox.Value = False
        .Serialization_ComboBox.Value = ""
        .BatchManagement_CheckBox.Value = False

    End With
    
    If ClearRowData = True Then
    
        With Ws_Request
            Dim ColumnCounter As Integer
            For ColumnCounter = 2 To GetBorders("LC", .Name, Wb_Current)
                .Cells(RowNumberForEdit, ColumnCounter).Value = vbNothing
            Next ColumnCounter
        End With
        
    End If
    
    ChangedNotSaved = True
    
End Sub


Public Function MaterialShortNameStringToFQT(MaterialShortName As String, Language As String) As String
    ' Adding postfix to Shortname if MatertialType FQT is chosen.
    
    Dim MaterialShortNameLen As Integer
    Dim FQTPostfix As String
    Dim MatchResult As String
    
    MaterialShortNameLen = Len(MaterialShortName)
    
    If Language = "Rus" Then
        FQTPostfix = FQTRusPostfix
    ElseIf Language = "Eng" Then
        FQTPostfix = FQTEngPostfix
    End If
    
    If IsNull(RxMatch(MaterialShortName, FQTRegEx, True, True)) Then
    
        MatchResult = ""
        
    Else
    
        MatchResult = RxMatch(MaterialShortName, FQTRegEx, True, True)
        
    End If
    
    If MatchResult = "" Then
    
        If MaterialShortNameLen <= 35 Then
        
            MaterialShortNameStringToFQT = MaterialShortName + FQTPostfix
            
        Else
        
            MaterialShortNameStringToFQT = Left(MaterialShortNameStringToFQT, 35) + FQTPostfix
            
        End If
        
    Else
    
        MaterialShortNameStringToFQT = MaterialShortName
        
    End If
    
End Function

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Reset button hovering effects.

    ClearFullDescriptionInactive_ImageButton.Visible = True
    ClearFormInactive_ImageButton.Visible = True
    CopyPreviousRowInactive_ImageButton.Visible = True
    CopyNextRowInactive_ImageButton.Visible = True
    ClearCurrentRowInactive_ImageButton.Visible = True
    SaveInactive_ImageButton.Visible = True
    CloseInactive_ImageButton.Visible = True
    RowUpInactive_ImageButton.Visible = True
    RowDownInactive_ImageButton.Visible = True
    SelectManufacturerCodeInactive_ImageButton.Visible = True
    
End Sub

