Attribute VB_Name = "exportMasterList"
Sub export()
    Dim exe_str As String
    
    exe_str = ThisWorkbook.Path & "\master_list_export.exe "
    Debug.Print exe_str
    Shell (exe_str)
End Sub
