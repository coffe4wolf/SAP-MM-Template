Attribute VB_Name = "m_Main"
Option Explicit

Const Language_Russian As String = "Russian"
Const Language_English As String = "English"

' Start of localization variables
Public CloseFormMessage                         As String
Public WarningMessage                           As String
Public NothingToCopyMessage                     As String
Public FirstRowMessage                          As String
Public MaterialTypeNoteLabelCaption             As String
Public ShortDescriptionTemplate1LabelCaption    As String
Public ProductCodeLabelCaption                  As String
Public ArticleLabelCaption                      As String
Public ShortDescriptionTemplate2LabelCaption    As String
Public ProductCode1LabelCaption                 As String
Public Article1LabelCaption                     As String
Public ProductCodeOnlyDigitsWarningMessage      As String
Public MaxDescriptionLengthLabelCatpion         As String
Public EmptyProductCodeAndArticleMessageCaption As String
Public EmptyFullDescriptionMessageCaption       As String
Public EmptyMaterialTypeMessageCaption          As String
Public WrongEnteredPriceMessageCaption          As String
Public ChangesSavedMessageCaption               As String
Public FQTRusPostfix                            As String
Public FQTEngPostfix                            As String
Public FQTRegEx                                 As String
Public SaveChangesWarningMessageCaption         As String
Public RequestFormUserFormCaption               As String
Public NoCyrillicCharsInEngDescriptionCaption   As String

Public MaterialTypeStandartized                 As String
Public MaterialTypeManufactured                 As String
Public MaterialTypeFQT                          As String
Public MultiSelectionFormCaption                As String
Public ManualUserFormCaption                    As String
' End of localization variables

Public RequestForm                  As MSForms.UserForm

Public Current_Language             As String


Public Constant_RowIndentConstant   As Integer      ' The number of header rows must be passed on Request sheet.
Public Wb_Current                   As Workbook
Public Ws_Request                   As Worksheet
Public Ws_MasterGroupCategory       As Worksheet
Public Ws_MasterGroup               As Worksheet
Public Ws_MasterUnit                As Worksheet
Public Ws_MasterAttributes          As Worksheet
Public Ws_helpSheet                 As Worksheet
Public Ws_MasterManufacturerCode    As Worksheet
Public WS_settingsSheet             As Worksheet
Public WS_MasterGroupMap            As Worksheet
Public WS_template                  As Worksheet

Public Ws_localizationUserForm      As Worksheet
Public Ws_localizationObjects       As Worksheet    ' Contains titles of all sheets.

Public RowNmberForEdit     As Integer
Public GroupCategory       As String
Public Group               As String
Public UrgentRequest       As Boolean
Public ShortNameRus        As String
Public ShortNameEng        As String
Public Unit                As String
Public ProductCode         As String
Public Article             As String
Public MaxPrice            As String
Public PurchasingGroup     As String
Public CriticalMaterial    As Boolean
Public Serialization       As String
Public FullDescription     As String
Public Priority            As String
Public MaterialType        As String
Public BatchManagement     As Boolean
Public SlidesCounter       As Integer
Public ArraySlidesNames()  As Variant

Public GroupCategoryCode   As String
Public GroupCode           As String
Public RowNumberForEdit    As Integer ' Counter of current row for changing.
Public ChangedNotSaved     As Boolean ' Flag for unsaved changes on the form.

Public AnswerCloseEditMaterialForm As Integer

Public HaltFlag            As Boolean

Public GroupCategory_Combobox As ListBox


Sub main()
    'Start of the program.
    
    Call InitWork
    Call InitConsts
    Call InitRow

    If Current_Language = Language_Russian Then
    
        RequestForm_UserForm.Show
        
    ElseIf Current_Language = Language_English Then
    
        RequestForm_UserForm_Eng.Show
        
    End If
    
End Sub

Sub InitWork()

    Constant_RowIndentConstant = 10
    
    '  Init worksheets.
    Set Wb_Current = ThisWorkbook
    Set Ws_localizationObjects = Wb_Current.Sheets("localizationObjects")
    Set Ws_localizationUserForm = Wb_Current.Sheets("localizationUserForm")
    Set WS_settingsSheet = Wb_Current.Sheets("settingsSheet")
    Set WS_template = Wb_Current.Sheets("template")
    
    Dim Ws_Request_Name                 As String
    Dim Ws_MasterGroupCategory_Name     As String
    Dim Ws_MasterGroup_Name             As String
    Dim Ws_MasterAttributes_Name        As String
    Dim Ws_helpSheet_Name               As String
    Dim Ws_MasterManufacturerCode_Name  As String
    Dim WS_MasterGroupMap_Name          As String

    
    Ws_Request_Name = WS_settingsSheet.Range("Request_SheetName").Value
    Ws_MasterGroupCategory_Name = WS_settingsSheet.Range("MasterCategoryGoup_SheetName").Value
    Ws_MasterGroup_Name = WS_settingsSheet.Range("MasterGroup_SheetName").Value
    Ws_MasterAttributes_Name = WS_settingsSheet.Range("ClassAttributes_SheetName").Value
    Ws_helpSheet_Name = WS_settingsSheet.Range("settingsSheet_SheetName").Value
    Ws_MasterManufacturerCode_Name = WS_settingsSheet.Range("MasterManufacturerCode_SheetName").Value
    WS_MasterGroupMap_Name = WS_settingsSheet.Range("MasterGroupMap_SheetName").Value
    
    With Wb_Current
        Set Ws_Request = .Sheets(Ws_Request_Name)
        Set Ws_MasterGroupCategory = .Sheets(Ws_MasterGroupCategory_Name)
        Set Ws_MasterGroup = .Sheets(Ws_MasterGroup_Name)
        Set Ws_MasterAttributes = .Sheets(Ws_MasterAttributes_Name)
        Set Ws_helpSheet = .Sheets(Ws_helpSheet_Name)
        Set Ws_MasterManufacturerCode = .Sheets(Ws_MasterManufacturerCode_Name)
        Set WS_MasterGroupMap = .Sheets(WS_MasterGroupMap_Name)
    End With

End Sub
Sub InitRow(Optional RowForEdit As Integer)
    ' Load data into the userform.
    
    If RowForEdit < 11 Then
        RowNumberForEdit = 11
    Else
        RowNumberForEdit = RowForEdit
    End If

    With Ws_Request
        Priority = .Cells(RowNumberForEdit, 2)
        ShortNameRus = .Cells(RowNumberForEdit, 3)
        FullDescription = .Cells(RowNumberForEdit, 4)
        ShortNameEng = .Cells(RowNumberForEdit, 5)
        Unit = .Cells(RowNumberForEdit, 6)
        ProductCode = .Cells(RowNumberForEdit, 7)
        Article = .Cells(RowNumberForEdit, 8)
        MaxPrice = .Cells(RowNumberForEdit, 9)
        GroupCode = .Cells(RowNumberForEdit, 10)
        PurchasingGroup = .Cells(RowNumberForEdit, 11)
        CriticalMaterial = CellToCheckbox(.Cells(RowNumberForEdit, 12))
        Serialization = .Cells(RowNumberForEdit, 13)
        GroupCategory = .Cells(RowNumberForEdit, 18)
        Group = .Cells(RowNumberForEdit, 19)
        MaterialType = .Cells(RowNumberForEdit, 20)
        BatchManagement = CellToCheckbox(.Cells(RowNumberForEdit, 21))
    End With
    
    
    With RequestForm_UserForm
        .UrgentRequest_CheckBox = PriorityToBool(Priority)
        .ShortNameRus_TextBox.Value = ShortNameRus
        .FullDescription_TextBox.Value = FullDescription
        .ShortNameEng_TextBox.Value = ShortNameEng
        .Unit_ComboBox.Value = Unit
        .ProductCode_TextBox.Value = ProductCode
        .Article_TextBox.Value = Article
        .MaxPrice_TextBox.Value = MaxPrice
        .GroupCode_TextBox.Value = GroupCode
        .PurchasingGroup_TextBox.Value = PurchasingGroup
        .CriticalMaterial_CheckBox.Value = CBool(CriticalMaterial)
        .Serialization_ComboBox.Value = Serialization
        .GroupCategory_Combobox.Value = GroupCategory
        .Group_ComboBox.Value = Group
        .MaterialType_ComboBox.Value = MaterialType
        .BatchManagement_CheckBox.Value = BatchManagement
    End With
    
    
    With RequestForm_UserForm_Eng
        .UrgentRequest_CheckBox = PriorityToBool(Priority)
        .ShortNameRus_TextBox.Value = ShortNameRus
        .FullDescription_TextBox.Value = FullDescription
        .ShortNameEng_TextBox.Value = ShortNameEng
        .Unit_ComboBox.Value = Unit
        .ProductCode_TextBox.Value = ProductCode
        .Article_TextBox.Value = Article
        .MaxPrice_TextBox.Value = MaxPrice
        .GroupCode_TextBox.Value = GroupCode
        .PurchasingGroup_TextBox.Value = PurchasingGroup
        .CriticalMaterial_CheckBox.Value = CBool(CriticalMaterial)
        .Serialization_ComboBox.Value = Serialization
        .GroupCategory_Combobox.Value = GroupCategory
        .Group_ComboBox.Value = Group
        .MaterialType_ComboBox.Value = MaterialType
        .BatchManagement_CheckBox.Value = BatchManagement
    End With
    
    RequestForm_UserForm.Repaint
    RequestForm_UserForm_Eng.Repaint
    
    ChangedNotSaved = False

End Sub

Public Sub InitConsts()

    Dim LocalizationPostfix As String: LocalizationPostfix = ""

    If Current_Language = Language_English Then
        LocalizationPostfix = "En"
    End If

    With Ws_localizationObjects
        CloseFormMessage = .Range("CloseForm_Message" & LocalizationPostfix).Value
        WarningMessage = .Range("Warning_Message" & LocalizationPostfix).Value
        NothingToCopyMessage = .Range("NothingToCopy_Message" & LocalizationPostfix).Value
        FirstRowMessage = .Range("FirstRow_Message" & LocalizationPostfix).Value
        MaterialTypeNoteLabelCaption = .Range("MaterialTypeNote_Label" & LocalizationPostfix).Value
        ShortDescriptionTemplate1LabelCaption = .Range("ShortDescriptionTemplate1_Label" & LocalizationPostfix).Value
        ProductCodeLabelCaption = .Range("ProductCode_Label" & LocalizationPostfix).Value
        ArticleLabelCaption = .Range("Article_Label" & LocalizationPostfix).Value
        ShortDescriptionTemplate2LabelCaption = .Range("ShortDescriptionTemplate2_Label" & LocalizationPostfix).Value
        ProductCode1LabelCaption = .Range("ProductCode1_Label" & LocalizationPostfix).Value
        Article1LabelCaption = .Range("Article1_Label" & LocalizationPostfix).Value
        ProductCodeOnlyDigitsWarningMessage = .Range("ProductCodeOnlyDigitsWarning_Message" & LocalizationPostfix).Value
        MaxDescriptionLengthLabelCatpion = .Range("MaxDescriptionLength_Label" & LocalizationPostfix).Value
        EmptyProductCodeAndArticleMessageCaption = .Range("EmptyProductCodeAndArticle_Message" & LocalizationPostfix).Value
        EmptyFullDescriptionMessageCaption = .Range("EmptyFullDescription_Message" & LocalizationPostfix).Value
        EmptyMaterialTypeMessageCaption = .Range("EmptyMaterialType_Message" & LocalizationPostfix).Value
        ChangesSavedMessageCaption = .Range("ChangesSaved_Message" & LocalizationPostfix).Value
        FQTRusPostfix = .Range("FQTRus_Postfix" & LocalizationPostfix).Value
        FQTEngPostfix = .Range("FQTEng_Postfix" & LocalizationPostfix).Value
        FQTRegEx = .Range("FQT_RegEx" & LocalizationPostfix).Value
        SaveChangesWarningMessageCaption = .Range("SaveChangesWarning_Message" & LocalizationPostfix).Value
        WrongEnteredPriceMessageCaption = .Range("WrongEnteredPrice_Message" & LocalizationPostfix).Value
        NoCyrillicCharsInEngDescriptionCaption = .Range("NoCyrillicCharsInEngDescription_Message" & LocalizationPostfix).Value
    End With
    
    With Ws_localizationUserForm
        MultiSelectionFormCaption = .Range("MultiSelectionForm_Caption").Value
        RequestFormUserFormCaption = .Range("RequestForm_UserForm_Caption").Value
    End With
    
    With Ws_helpSheet
        MaterialTypeStandartized = .Range("MaterialType_Standartized")
        MaterialTypeManufactured = .Range("MaterialType_Manufactured")
        MaterialTypeFQT = .Range("MaterialType_FQT")
    End With

End Sub

Sub InitManual()

    SlidesCounter = 1
    Manual_UserForm.Show
    
End Sub

Function PriorityToBool(PriorityValue As String) As Boolean
    ' Converting prioriry value from cell to checkbox.
    
    Dim Priority As String
    Priority = Trim(PriorityValue)
    
    If Priority = "01" Then
        PriorityToBool = False
    ElseIf Priority = "02" Then
        PriorityToBool = True
    End If
        
End Function

Function PriorityToString(PriorityValue As Boolean) As String
    ' Converting prioriry value from checkbox to cell.
    
    If PriorityValue = True Then
        PriorityToString = "02"
    ElseIf PriorityValue = False Then
        PriorityToString = "01"
    End If
        
End Function

Sub CopyRow(RowNumToCopy As Integer)
    ' Copy data from specified row to userform.
    
    With RequestForm_UserForm
    
        .UrgentRequest_CheckBox.Value = PriorityToBool(Ws_Request.Cells(RowNumToCopy, 2))
        .ShortNameRus_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 3)
        .ShortNameEng_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 5)
        .Unit_ComboBox.Text = Ws_Request.Cells(RowNumToCopy, 6)
        .ProductCode_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 7)
        .Article_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 8)
        .MaxPrice_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 9)
        .GroupCode_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 10)
        .PurchasingGroup_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 11)
        .CriticalMaterial_CheckBox.Value = CellToCheckbox(Ws_Request.Cells(RowNumToCopy, 12))
        .Serialization_ComboBox.Value = Ws_Request.Cells(RowNumToCopy, 13)
        .GroupCategory_Combobox.Text = Ws_Request.Cells(RowNumToCopy, 18)
        .Group_ComboBox.Text = Ws_Request.Cells(RowNumToCopy, 19)
        .MaterialType_ComboBox.Text = Ws_Request.Cells(RowNumToCopy, 20)
        .FullDescription_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 4)
        .BatchManagement_CheckBox.Value = CellToCheckbox(Ws_Request.Cells(RowNumToCopy, 21))
        
    End With
    
    With RequestForm_UserForm_Eng
    
        .UrgentRequest_CheckBox.Value = PriorityToBool(Ws_Request.Cells(RowNumToCopy, 2))
        .ShortNameRus_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 3)
        .ShortNameEng_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 5)
        .Unit_ComboBox.Text = Ws_Request.Cells(RowNumToCopy, 6)
        .ProductCode_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 7)
        .Article_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 8)
        .MaxPrice_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 9)
        .GroupCode_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 10)
        .PurchasingGroup_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 11)
        .CriticalMaterial_CheckBox.Value = CellToCheckbox(Ws_Request.Cells(RowNumToCopy, 12))
        .Serialization_ComboBox.Value = Ws_Request.Cells(RowNumToCopy, 13)
        .GroupCategory_Combobox.Text = Ws_Request.Cells(RowNumToCopy, 18)
        .Group_ComboBox.Text = Ws_Request.Cells(RowNumToCopy, 19)
        .MaterialType_ComboBox.Text = Ws_Request.Cells(RowNumToCopy, 20)
        .FullDescription_TextBox.Value = Ws_Request.Cells(RowNumToCopy, 4)
        .BatchManagement_CheckBox.Value = CellToCheckbox(Ws_Request.Cells(RowNumToCopy, 21))
        
    End With
    
    RequestForm_UserForm.Repaint
    RequestForm_UserForm_Eng.Repaint
    ChangedNotSaved = True

End Sub

Sub ClearRequestSheet()
    ' Clear main sheet from entered materials.

    With ThisWorkbook.Sheets(ThisWorkbook.Sheets("settingsSheet").Range("Request_SheetName").Value)
        Dim RowCounter As Long
        For RowCounter = 11 To 230
            .Range("A" & RowCounter & ":U" & RowCounter).ClearContents
        Next RowCounter
    End With
    
End Sub

Function RowIsNotEmpty(RowToCheck As Integer, Optional Ws As Worksheet) As Boolean

    If Ws Is Nothing Then
        Set Ws = Ws_Request
    End If

     If WorksheetFunction.CountA(Ws.Range("B" & RowToCheck & ":P" & RowToCheck)) > 0 Then
        RowIsNotEmpty = True
     Else
        RowIsNotEmpty = False
     End If

End Function



Public Function CheckboxToCell(CheckBoxValue As Variant) As String

    If CheckBoxValue = True Then
        CheckboxToCell = "X"
    Else
        CheckboxToCell = ""
    End If
    
End Function

Public Function CellToCheckbox(CellValue As String) As Boolean

    If CellValue = "X" Then
        CellToCheckbox = True
    ElseIf CellValue = "" Then
        CellToCheckbox = False
    End If
    
End Function

Function StringIsPrice(Price As String) As Boolean
    
    Dim PricePattern        As String: PricePattern = "^[\d]+[\.][\d]+$|^[\d]+$"
    Dim ZeroPricePattern    As String: ZeroPricePattern = "^[0.,]+$"
    
    
    If RxMatch(Price, ZeroPricePattern, False, False) Then
    
        StringIsPrice = False
        
    ElseIf RxMatch(Price, PricePattern, False, False) Then
    
        StringIsPrice = True
        
    Else
    
        StringIsPrice = False
        
    End If

End Function

Function RepairCyrillicView(InputString)

  Dim Arr, i%, sTxt$, sSymb$
  Arr = Split(Replace(Replace(InputString, "&#", ";&#"), ";;", ";"), ";")
  
  If UBound(Arr) > LBound(Arr) Then
     On Error Resume Next
     For i = LBound(Arr) To UBound(Arr)
        If Left(Arr(i), 2) = "&#" And Len(Arr(i)) = 5 And IsNumeric(Right(Arr(i), 3)) Then
           Arr(i) = Chr(CInt(Right(Arr(i), 3)))
        End If
     Next
     sTxt = Join(Arr, "")
  Else
     For i = 1 To Len(InputString)
        sSymb = Mid(InputString, i, 1)
        If AscW(sSymb) > 255 Then
           sTxt = sTxt & sSymb
        Else
           sTxt = sTxt & Chr(AscW(sSymb))
        End If
     Next i
  End If
  
  RepairCyrillicView = sTxt
  
End Function
