Attribute VB_Name = "refreshDb"

Sub refresh()
    Dim exe_str As String
    
    exe_str = ThisWorkbook.Path & "\refresh_stampholder.exe "
    Shell (exe_str)
    
End Sub
