Attribute VB_Name = "m_Test"
Option Explicit

Sub main_test()

    Dim BadString As String
    
    BadString = "12312,22"
    
    'Debug.Print StringIsPrice(tempOk) '����������
    Debug.Print m_Main.RepairCyrillicView("��� �")

End Sub



Sub ChooseLanguage_UserForm_test()


    
    RequestForm_UserForm_Eng.Show
    


End Sub
