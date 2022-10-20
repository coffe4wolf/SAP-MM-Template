Attribute VB_Name = "m_Test"
Option Explicit

Sub main_test()

    Dim BadString As String
    
    BadString = "12312,22"
    
    'Debug.Print StringIsPrice(tempOk) 'выаывацыва
    Debug.Print m_Main.RepairCyrillicView("Эта д")

End Sub



Sub ChooseLanguage_UserForm_test()


    
    RequestForm_UserForm_Eng.Show
    


End Sub
